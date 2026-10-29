using System.Globalization;
using nietras.SeparatedValues;
using SpreadCheetah;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadsheetConverter;

internal sealed record CsvToXlsxOptions
{
    public string? SchemaFile { get; init; }

    /// <summary>Name of the first worksheet; later ones get " (2)", " (3)" and so on.</summary>
    public string SheetName { get; init; } = "Sheet1";

    /// <summary>Rows per worksheet, header included. Further rows continue on a new sheet with the header repeated.</summary>
    public int SheetRows { get; init; } = ExcelLimits.MaxRows;
}

internal sealed record CsvToXlsxResult(int Rows, int Sheets, int ClippedCells)
{
    public string ToJson() => $$"""{"rows":{{Rows}},"sheets":{{Sheets}},"clipped_cells":{{ClippedCells}}}""";
}

/// <summary>
/// Streams an RFC 4180 CSV (UTF-8, comma separated, header row) into an xlsx workbook with constant memory.
/// Values are typed by the schema; a value that does not parse as its column's type is written as text.
/// Text is always written as a string cell, never as a formula.
/// </summary>
internal static class CsvToXlsx
{
    private static readonly string[] DateFormats = ["yyyy-MM-dd"];

    private static readonly string[] DateTimeFormats =
    [
        "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd HH:mm:ss.FFFFFFF", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd'T'HH:mm",
        "yyyy-MM-dd",
    ];

    /// <param name="input">Stream to read from, or null to open <paramref name="inputPath"/>.</param>
    public static async Task<CsvToXlsxResult> ConvertAsync(Stream? input, string inputPath, string outputPath, CsvToXlsxOptions options)
    {
        ValidateSheetName(options.SheetName);
        var schema = Schema.Load(options.SchemaFile);

        // The header is read as an ordinary row, because Sep's own header handling rejects repeated column names,
        // which exports of joined queries and blank header cells commonly have. Column counts are checked in
        // FillCells, so that blank lines can be skipped and the error names the header's column count.
        var readerOptions = Sep.New(',').Reader(o => o with { Unescape = true, HasHeader = false, DisableColCountCheck = true });
        using var reader = input is null ? readerOptions.FromFile(inputPath) : readerOptions.From(input);
        if (!reader.MoveNext())
        {
            throw new InvalidDataException("input is empty, expected a header row");
        }

        CheckQuotesClosed(reader.Current);
        var columnNames = Enumerable.Range(0, reader.Current.ColCount).Select(i => reader.Current[i].ToString()).ToArray();
        var width = columnNames.Length;
        if (width > ExcelLimits.MaxColumns)
        {
            throw new InvalidDataException($"header has {width} columns, Excel allows at most {ExcelLimits.MaxColumns}");
        }
        var columns = ColumnsFor(columnNames, schema);

        await using var output = new AtomicFile(outputPath);
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(output.Stream);
        var headerStyle = spreadsheet.AddStyle(new Style { Font = { Bold = true } });
        var types = columns.Select(column => column.Type).ToArray();
        var styles = ColumnStyles(spreadsheet, columns);

        var clippedCells = columnNames.Count(name => name.Length > ExcelLimits.MaxCellChars);
        var header = columnNames.Select(name => new Cell(ExcelLimits.Clip(name), headerStyle)).ToArray();
        var worksheetOptions = new WorksheetOptions
        {
            FrozenRows = 1,
            AutoFilter = new AutoFilterOptions($"A1:{ExcelLimits.ColumnName(width)}1"),
        };

        var sheets = 0;
        var sheetRow = 0;
        var rows = 0;
        var cells = new Cell[width];

        async Task StartSheetAsync()
        {
            sheets++;
            var name = sheets == 1 ? options.SheetName : SuffixedSheetName(options.SheetName, sheets);
            await spreadsheet.StartWorksheetAsync(name, worksheetOptions);
            await spreadsheet.AddRowAsync(header);
            sheetRow = 1;
        }

        await StartSheetAsync();

        while (reader.MoveNext())
        {
            // A blank line has one empty field. With a single column that is an empty value; otherwise it is a blank
            // line, such as an extra newline at the end of the file, and is skipped.
            if (width > 1 && reader.Current.Span.IsEmpty)
            {
                continue;
            }
            if (sheetRow == options.SheetRows)
            {
                await StartSheetAsync();
            }

            clippedCells += FillCells(reader.Current, cells, types, styles);
            await spreadsheet.AddRowAsync(cells);
            sheetRow++;
            rows++;
        }

        await spreadsheet.FinishAsync();
        await output.CommitAsync();
        return new CsvToXlsxResult(rows, sheets, clippedCells);
    }

    /// <summary>Converts one CSV row into <paramref name="cells"/> and returns how many values were clipped.</summary>
    private static int FillCells(SepReader.Row row, Cell[] cells, ColumnType[] types, StyleId?[] styles)
    {
        CheckQuotesClosed(row);
        if (row.ColCount != cells.Length)
        {
            throw new InvalidDataException(
                $"line {row.LineNumberFrom}: found {row.ColCount} fields, expected {cells.Length} like the header");
        }

        var clippedCells = 0;
        for (var i = 0; i < cells.Length; i++)
        {
            cells[i] = ToCell(row[i].ToString(), types[i], styles[i], ref clippedCells);
        }
        return clippedCells;
    }

    /// <summary>
    /// Sep accepts a quoted field left open at the end of the input; a complete RFC 4180 row always has an even
    /// number of quotes, so this catches input that was cut off mid-row (for example a producer that crashed).
    /// </summary>
    private static void CheckQuotesClosed(SepReader.Row row)
    {
        if (row.Span.Count('"') % 2 != 0)
        {
            throw new InvalidDataException($"line {row.LineNumberFrom}: quoted field is never closed");
        }
    }

    /// <summary>Repeated column names are fine, except for a column the schema lists, which would be ambiguous.</summary>
    private static Column[] ColumnsFor(string[] columnNames, Dictionary<string, Column> schema)
    {
        var ambiguous = columnNames
            .Where(schema.ContainsKey)
            .GroupBy(name => name, StringComparer.Ordinal)
            .FirstOrDefault(group => group.Count() > 1);
        if (ambiguous is not null)
        {
            throw new InvalidDataException(
                $"schema: column '{ambiguous.Key}' appears {ambiguous.Count()} times in the header, so its type is ambiguous");
        }
        return columnNames.Select(name => schema.GetValueOrDefault(name) ?? TextColumn).ToArray();
    }

    private static readonly Column TextColumn = new(ColumnType.Text);

    /// <summary>
    /// One style per column: the schema's format, else yyyy-mm-dd for dates and yyyy-mm-dd hh:mm:ss for date-times.
    /// Columns sharing a format share a style. A value that falls back to text keeps its column's style, which only
    /// affects how numbers display, so the text is shown unchanged.
    /// </summary>
    private static StyleId?[] ColumnStyles(Spreadsheet spreadsheet, Column[] columns)
    {
        var styles = new Dictionary<string, StyleId>(StringComparer.Ordinal);
        return columns.Select(column => (column.Format ?? DefaultFormat(column.Type)) switch
        {
            null => (StyleId?)null,
            var format => styles.TryGetValue(format, out var style)
                ? style
                : styles[format] = spreadsheet.AddStyle(new Style { Format = NumberFormat.Custom(format) }),
        }).ToArray();
    }

    private static string? DefaultFormat(ColumnType type) => type switch
    {
        ColumnType.Date => "yyyy-mm-dd",
        ColumnType.DateTime => "yyyy-mm-dd hh:mm:ss",
        _ => null,
    };

    private static Cell ToCell(string value, ColumnType type, StyleId? style, ref int clippedCells)
    {
        if (value.Length == 0)
        {
            return default;
        }

        switch (type)
        {
            case ColumnType.Number when double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number):
                return new Cell(number, style);
            case ColumnType.Bool when TryParseBool(value, out var boolean):
                return new Cell(boolean, style);
            case ColumnType.Date when TryParseExcelDate(value, DateFormats, out var date):
                return new Cell(date, style);
            case ColumnType.DateTime when TryParseExcelDate(value, DateTimeFormats, out var dateTime):
                return new Cell(dateTime, style);
        }

        if (value.Length > ExcelLimits.MaxCellChars)
        {
            value = ExcelLimits.Clip(value);
            clippedCells++;
        }
        return new Cell(value, style);
    }

    /// <summary>
    /// Parses an ISO date that Excel can show as a date. Excel's date system starts in 1900 and wrongly treats 1900 as
    /// a leap year, so only dates from 1900-03-01 convert reliably; earlier dates are kept as text.
    /// </summary>
    private static bool TryParseExcelDate(string value, string[] formats, out DateTime result) =>
        DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out result)
        && result >= FirstReliableExcelDate;

    private static readonly DateTime FirstReliableExcelDate = new(1900, 3, 1);

    /// <summary>Accepts what PHP and Laravel commonly produce: 1/0, true/false and t/f in any case.</summary>
    private static bool TryParseBool(string value, out bool result)
    {
        switch (value.ToLowerInvariant())
        {
            case "1" or "true" or "t":
                result = true;
                return true;
            case "0" or "false" or "f":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }

    private static string SuffixedSheetName(string name, int sheet)
    {
        var suffix = $" ({sheet})";
        return ExcelLimits.Clip(name, ExcelLimits.MaxSheetNameChars - suffix.Length) + suffix;
    }

    private static void ValidateSheetName(string name)
    {
        if (name.Length is 0 or > ExcelLimits.MaxSheetNameChars
            || name.IndexOfAny(['\\', '/', '?', '*', '[', ']', ':']) >= 0
            || name.StartsWith('\'') || name.EndsWith('\''))
        {
            throw new ArgumentException(
                $"invalid sheet name '{name}': 1-{ExcelLimits.MaxSheetNameChars} characters, none of \\ / ? * [ ] : and not starting or ending with '");
        }
    }
}
