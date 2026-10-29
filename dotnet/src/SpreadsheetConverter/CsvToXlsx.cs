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

    /// <summary>Rejects invalid UTF-8 and schema columns missing from the header, instead of passing the data through.</summary>
    public bool Strict { get; init; }

    /// <summary>Field separator, or null to let Sep detect it from the header row (comma, semicolon, tab or pipe).</summary>
    public char? Delimiter { get; init; }
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
        // Without --delimiter, Sep detects the separator from the first line (comma, semicolon, tab or pipe, falling
        // back to semicolon), so the semicolon "CSV" that Excel and Numbers save in many European locales converts as
        // is. The detection counts characters without regard to quotes.
        Func<SepReaderOptions, SepReaderOptions> configure = o => o with { Unescape = true, HasHeader = false, DisableColCountCheck = true };
        var readerOptions = options.Delimiter is { } delimiter ? Sep.New(delimiter).Reader(configure) : Sep.Reader(configure);
        using var reader = input is null ? readerOptions.FromFile(inputPath) : readerOptions.From(input);
        return await ConvertAsync(reader, outputPath, options, schema);
    }

    private static async Task<CsvToXlsxResult> ConvertAsync(SepReader reader, string outputPath, CsvToXlsxOptions options, Dictionary<string, Column> schema)
    {
        if (!reader.MoveNext())
        {
            throw new InvalidDataException("input is empty, expected a header row");
        }
        if (reader.Current.Span.IsEmpty && options.Delimiter is null)
        {
            // Sep detects from the first line; from a blank one it falls back to a semicolon and would quietly read a
            // comma-separated file as one column. With --delimiter, a blank first line is an empty one-column header.
            throw new InvalidDataException(
                "line 1 is blank, so the delimiter cannot be detected; remove the blank line or pass --delimiter");
        }

        // The header is line 1; Sep reports a wrong line number for the first row.
        CheckRow(reader.Current, options.Strict, line: 1);
        var columnNames = Enumerable.Range(0, reader.Current.ColCount).Select(i => reader.Current[i].ToString()).ToArray();
        var width = columnNames.Length;
        if (width > ExcelLimits.MaxColumns)
        {
            throw new InvalidDataException($"header has {width} columns, Excel allows at most {ExcelLimits.MaxColumns}");
        }
        var columns = ColumnsFor(columnNames, schema);
        if (options.Strict)
        {
            var missing = schema.Keys.Where(name => !columnNames.Contains(name, StringComparer.Ordinal)).ToArray();
            if (missing.Length > 0)
            {
                throw new InvalidDataException($"schema: {string.Join(", ", missing.Select(name => $"'{name}'"))} not in the header");
            }
        }

        await using var output = new AtomicFile(outputPath);
        // Without xml:space="preserve", readers (Excel included) drop leading and trailing spaces from text.
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(output.Stream, new SpreadCheetahOptions { PreserveStringWhitespace = true });
        var headerStyle = spreadsheet.AddStyle(new Style { Font = { Bold = true } });
        var wrappedHeaderStyle = spreadsheet.AddStyle(new Style { Font = { Bold = true }, Alignment = { WrapText = true } });
        var types = columns.Select(column => column.Type).ToArray();
        var styles = ColumnStyles(spreadsheet, columns);

        var clippedCells = columnNames.Count(name => name.Length > ExcelLimits.MaxCellChars);
        var header = columnNames
            .Select(name => new Cell(ExcelLimits.Clip(name), HasLineBreak(name) ? wrappedHeaderStyle : headerStyle))
            .ToArray();
        var worksheetOptions = new WorksheetOptions
        {
            FrozenRows = 1,
            AutoFilter = new AutoFilterOptions($"A1:{ExcelLimits.ColumnName(width)}1"),
        };
        // Excel shows ##### for a date or formatted number that does not fit, and its default width (about 8.4
        // characters) is too narrow for a date-time or an Accounting amount.
        for (var i = 0; i < width; i++)
        {
            if (ColumnWidth(columns[i]) is { } columnWidth)
            {
                worksheetOptions.Column(i + 1).Width = columnWidth;
            }
        }

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

            clippedCells += FillCells(reader.Current, cells, types, styles, options.Strict);
            await spreadsheet.AddRowAsync(cells);
            sheetRow++;
            rows++;
        }

        await spreadsheet.FinishAsync();
        await output.CommitAsync();
        return new CsvToXlsxResult(rows, sheets, clippedCells);
    }

    /// <summary>Converts one CSV row into <paramref name="cells"/> and returns how many values were clipped.</summary>
    private static int FillCells(SepReader.Row row, Cell[] cells, ColumnType[] types, ColumnStyle[] styles, bool strict)
    {
        CheckRow(row, strict, row.LineNumberFrom);
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
    ///
    /// Invalid UTF-8 is decoded as U+FFFD, so in strict mode a row containing it is rejected. Checking the decoded
    /// row, rather than failing in the decoder, which reads far ahead of the rows, names the exact line.
    /// </summary>
    private static void CheckRow(SepReader.Row row, bool strict, int line)
    {
        if (row.Span.Count('"') % 2 != 0)
        {
            throw new InvalidDataException($"line {line}: quoted field is never closed");
        }
        if (strict && row.Span.Contains('\uFFFD'))
        {
            throw new InvalidDataException($"line {line}: not valid UTF-8 (or contains the replacement character \uFFFD)");
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
    private static ColumnStyle[] ColumnStyles(Spreadsheet spreadsheet, Column[] columns)
    {
        var styles = new Dictionary<string, ColumnStyle>(StringComparer.Ordinal);
        return columns.Select(column =>
        {
            var format = column.Format ?? DefaultFormat(column.Type);
            if (!styles.TryGetValue(format ?? "", out var style))
            {
                var numberFormat = format is null ? (NumberFormat?)null : NumberFormat.Custom(format);
                style = styles[format ?? ""] = new ColumnStyle(
                    numberFormat is null ? null : spreadsheet.AddStyle(new Style { Format = numberFormat }),
                    spreadsheet.AddStyle(new Style { Format = numberFormat, Alignment = { WrapText = true } }));
            }
            return style;
        }).ToArray();
    }

    /// <summary>
    /// The schema's width, else one that fits the column's format: the default date formats (a custom date format is
    /// assumed to show about as many characters as its code), and formatted amounts up to "9.999.999,99 kr." with
    /// Accounting padding.
    /// Rows are streamed, so widths cannot be measured from the data.
    /// </summary>
    private static double? ColumnWidth(Column column) => column.Width ?? column.Type switch
    {
        ColumnType.Date => Math.Max(column.Format is null ? 10 : DisplayLength(column.Format), 10) + 2,
        ColumnType.DateTime => Math.Max(column.Format is null ? 19 : DisplayLength(column.Format), 10) + 2,
        ColumnType.Number when column.Format is not null && XlsxToCsv.FormatShowsDate(column.Format) =>
            Math.Max(DisplayLength(column.Format) + 2, FormattedNumberWidth),
        ColumnType.Number when column.Format is not null => FormattedNumberWidth,
        _ => null,
    };

    private const double FormattedNumberWidth = 18;

    /// <summary>
    /// About how many characters a date format shows: its code length, except that month and weekday names
    /// (mmmm, dddd: up to 9 characters, as in "September" and "Wednesday") and one-letter d, m, h and s (up to 2
    /// digits) are longer than their codes.
    /// </summary>
    internal static int DisplayLength(string format)
    {
        var length = 0;
        for (var i = 0; i < format.Length;)
        {
            var c = char.ToLowerInvariant(format[i]);
            var run = 1;
            while (i + run < format.Length && char.ToLowerInvariant(format[i + run]) == c)
            {
                run++;
            }
            length += c switch
            {
                'm' or 'd' when run >= 4 => 9,
                'd' or 'm' or 'h' or 's' when run == 1 => 2,
                _ => run,
            };
            i += run;
        }
        return length;
    }

    private static string? DefaultFormat(ColumnType type) => type switch
    {
        ColumnType.Date => "yyyy-mm-dd",
        ColumnType.DateTime => "yyyy-mm-dd hh:mm:ss",
        _ => null,
    };

    /// <param name="Wrapped">For text with a line break, which Excel only shows with wrap text on, as after Alt+Enter.</param>
    private sealed record ColumnStyle(StyleId? Plain, StyleId Wrapped);

    /// <summary>A lone \r counts too: XML readers, Excel included, turn it into a line break.</summary>
    private static bool HasLineBreak(string value) => value.AsSpan().IndexOfAny('\n', '\r') >= 0;

    private static Cell ToCell(string value, ColumnType type, ColumnStyle columnStyle, ref int clippedCells)
    {
        if (value.Length == 0)
        {
            return default;
        }

        // Surrounding spaces are ignored when parsing a typed value; a value that falls back to text keeps them.
        var trimmed = type == ColumnType.Text ? value : value.Trim();
        switch (type)
        {
            case ColumnType.Number when double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number):
                return new Cell(number, columnStyle.Plain);
            case ColumnType.Bool when TryParseBool(trimmed, out var boolean):
                return new Cell(boolean, columnStyle.Plain);
            case ColumnType.Date when TryParseExcelDate(trimmed, DateFormats, out var date):
                return new Cell(date, columnStyle.Plain);
            case ColumnType.DateTime when TryParseExcelDate(trimmed, DateTimeFormats, out var dateTime):
                return new Cell(dateTime, columnStyle.Plain);
        }

        if (value.Length > ExcelLimits.MaxCellChars)
        {
            value = ExcelLimits.Clip(value);
            clippedCells++;
        }
        return new Cell(value, HasLineBreak(value) ? columnStyle.Wrapped : columnStyle.Plain);
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
