using System.Globalization;
using System.IO.Compression;
using System.Text;
using Sylvan.Data.Excel;

namespace SpreadsheetConverter;

internal sealed record XlsxToCsvOptions
{
    /// <summary>Worksheet to read; the first visible worksheet when null.</summary>
    public string? Sheet { get; init; }
}

internal sealed record XlsxToCsvResult(int Rows, int Columns, int ErrorCells, int IgnoredCells)
{
    public string ToJson() =>
        $$"""{"rows":{{Rows}},"columns":{{Columns}},"error_cells":{{ErrorCells}},"ignored_cells":{{IgnoredCells}}}""";
}

/// <summary>
/// Streams one worksheet of an xlsx, xlsb or xls workbook into an RFC 4180 CSV (UTF-8 without BOM, comma separated,
/// \n line endings), readable by PHP's <c>fgetcsv($handle, null, ',', '"', '')</c>.
///
/// The first non-blank row is the header and sets the column count: shorter rows are padded, and non-empty cells beyond the
/// header are dropped and counted as ignored. Values are written in a culture-independent form:
/// numbers with a dot decimal separator and no grouping, booleans as 1/0, dates as yyyy-MM-dd, date-times as
/// yyyy-MM-dd HH:mm:ss, times as HH:mm:ss and durations (elapsed-time formats such as [h]:mm:ss) in total hours. Formula errors such as #N/A become empty values and are counted.
/// </summary>
internal static class XlsxToCsv
{
    /// <param name="output">Stream to write to, or null to write <paramref name="outputPath"/>.</param>
    public static async Task<XlsxToCsvResult> ConvertAsync(string inputPath, Stream? output, string outputPath, XlsxToCsvOptions options)
    {
        await using var input = File.OpenRead(inputPath);
        using var reader = ExcelDataReader.Create(input, DetectWorkbookType(input), new ExcelDataReaderOptions
        {
            // The header is found below rather than taken from the first row, which may be blank.
            Schema = ExcelSchema.DynamicNoHeaders,
            FormulaErrorHandling = FormulaErrorHandling.Null,
        });

        if (options.Sheet is not null && !reader.TryOpenWorksheet(options.Sheet))
        {
            throw new InvalidDataException(
                $"worksheet '{options.Sheet}' not found; available: {string.Join(", ", reader.WorksheetNames)}");
        }

        await using var file = output is null ? new AtomicFile(outputPath) : null;
        // Disposed exactly once below: a second dispose would flush into the already committed (closed) file.
        var writer = new StreamWriter(output ?? file!.Stream, new UTF8Encoding(false), 1 << 16, leaveOpen: true)
        {
            NewLine = "\n",
        };

        var errorCells = 0;
        while (true)
        {
            if (!await reader.ReadAsync())
            {
                throw new InvalidDataException("worksheet is empty, expected a header row");
            }
            if (!IsBlankRow(reader))
            {
                break;
            }
        }

        var width = reader.RowFieldCount;
        var values = new string[width];
        for (var i = 0; i < width; i++)
        {
            values[i] = CellText(reader, i, ref errorCells);
        }
        Csv.WriteRow(writer, values);

        var rows = 0;
        var ignoredCells = 0;

        while (await reader.ReadAsync())
        {
            for (var i = 0; i < width; i++)
            {
                values[i] = i < reader.RowFieldCount ? CellText(reader, i, ref errorCells) : "";
            }
            for (var i = width; i < reader.RowFieldCount; i++)
            {
                if (reader.GetExcelDataType(i) != ExcelDataType.Null)
                {
                    ignoredCells++;
                }
            }

            Csv.WriteRow(writer, values);
            rows++;
        }

        await writer.DisposeAsync();
        if (file is not null)
        {
            await file.CommitAsync();
        }
        return new XlsxToCsvResult(rows, width, errorCells, ignoredCells);
    }

    private static bool IsBlankRow(ExcelDataReader reader)
    {
        for (var i = 0; i < reader.RowFieldCount; i++)
        {
            if (reader.GetExcelDataType(i) != ExcelDataType.Null)
            {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Detects the workbook type from the file contents rather than the extension, because uploaded files often have
    /// none (PHP stores uploads as <c>/tmp/phpAbC123</c>). xls is an OLE compound file; xlsx and xlsb are both zip
    /// archives, told apart by the binary workbook part only xlsb has.
    /// </summary>
    internal static ExcelWorkbookType DetectWorkbookType(Stream input)
    {
        Span<byte> signature = stackalloc byte[8];
        var length = input.ReadAtLeast(signature, signature.Length, throwOnEndOfStream: false);
        input.Position = 0;

        if (signature[..length].SequenceEqual(OleSignature))
        {
            return ExcelWorkbookType.Excel;
        }
        if (length >= 4 && signature[..4].SequenceEqual(ZipSignature))
        {
            ExcelWorkbookType type;
            using (var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
            {
                type = archive.GetEntry("xl/workbook.bin") is not null ? ExcelWorkbookType.ExcelBinary : ExcelWorkbookType.ExcelXml;
            }
            input.Position = 0;
            return type;
        }
        throw new InvalidDataException("input is not an xlsx, xlsb or xls workbook");
    }

    private static ReadOnlySpan<byte> OleSignature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];

    private static string CellText(ExcelDataReader reader, int ordinal, ref int errorCells)
    {
        try
        {
            return ReadCell(reader, ordinal, ref errorCells);
        }
        catch (Exception exception)
        {
            throw new InvalidDataException(
                $"row {reader.RowNumber}, column {ExcelLimits.ColumnName(ordinal + 1)}: {exception.Message}", exception);
        }
    }

    private static string ReadCell(ExcelDataReader reader, int ordinal, ref int errorCells)
    {
        switch (reader.GetExcelDataType(ordinal))
        {
            case ExcelDataType.Null:
                return "";
            case ExcelDataType.Error:
                errorCells++;
                return "";
            case ExcelDataType.Boolean:
                return reader.GetBoolean(ordinal) ? "1" : "0";
            case ExcelDataType.String:
                return reader.GetString(ordinal);
            case ExcelDataType.DateTime:
                return FormatDateTime(reader.GetDateTime(ordinal));
            case ExcelDataType.Numeric when IsElapsedTimeFormat(reader.GetFormat(ordinal)?.Format):
                return FormatDuration(reader.GetDouble(ordinal));
            case ExcelDataType.Numeric:
                return reader.GetFormat(ordinal)?.Kind switch
                {
                    FormatKind.Date => FormatDate(reader, ordinal),
                    FormatKind.Time => FormatTime(reader.GetDouble(ordinal)),
                    _ => FormatNumber(reader.GetDouble(ordinal)),
                };
            default:
                return reader.GetString(ordinal);
        }
    }

    /// <summary>A date-formatted number Excel cannot show as a date (such as a negative serial) is written as the number.</summary>
    private static string FormatDate(ExcelDataReader reader, int ordinal)
    {
        try
        {
            return FormatDateTime(reader.GetDateTime(ordinal));
        }
        catch (InvalidCastException)
        {
            return FormatNumber(reader.GetDouble(ordinal));
        }
    }

    private static string FormatDateTime(DateTime value) =>
        value.TimeOfDay == TimeSpan.Zero
            ? value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    /// <summary>Rounds to whole seconds, capped at 23:59:59 so a value just below midnight does not wrap to 00:00:00.</summary>
    internal static string FormatTime(double value) =>
        TimeSpan.FromSeconds(Math.Min(Math.Round((value - Math.Floor(value)) * 86_400), 86_399))
            .ToString(@"hh\:mm\:ss", CultureInfo.InvariantCulture);

    /// <summary>
    /// Elapsed-time formats such as <c>[h]:mm:ss</c> show a duration in total hours (or minutes or seconds) rather than
    /// a time of day, so they must keep the whole days that <see cref="FormatTime"/> drops.
    /// </summary>
    internal static bool IsElapsedTimeFormat(string? format)
    {
        if (format is null)
        {
            return false;
        }

        var inQuotes = false;
        for (var i = 0; i < format.Length; i++)
        {
            switch (format[i])
            {
                case '"':
                    inQuotes = !inQuotes;
                    break;
                case '\\' when !inQuotes:
                    i++;
                    break;
                case '[' when !inQuotes:
                    var end = format.IndexOf(']', i);
                    if (end > i + 1 && format.AsSpan(i + 1, end - i - 1).ContainsAnyExcept("hHmMsS") is false)
                    {
                        return true;
                    }
                    break;
            }
        }
        return false;
    }

    /// <summary>A duration in total hours, such as 30:00:00 for 1.25 days or -01:30:00, rounded to whole seconds.</summary>
    internal static string FormatDuration(double days)
    {
        var seconds = (long)Math.Round(Math.Abs(days) * 86_400);
        var sign = days < 0 && seconds > 0 ? "-" : "";
        return string.Create(CultureInfo.InvariantCulture, $"{sign}{seconds / 3600:00}:{seconds / 60 % 60:00}:{seconds % 60:00}");
    }

    /// <summary>Whole numbers without exponent or decimals (so 12345678 stays 12345678), others in shortest round-trip form.</summary>
    private static string FormatNumber(double value) =>
        value == Math.Floor(value) && Math.Abs(value) < 1e15
            ? ((long)value).ToString(CultureInfo.InvariantCulture)
            : value.ToString("R", CultureInfo.InvariantCulture);
}
