using System.Reflection;

namespace SpreadsheetConverter;

/// <summary>
/// Command line entry point. Exit codes: 0 success, 1 conversion failed, 2 invalid usage.
/// On success with a file output, one JSON line with statistics is written to stdout.
/// </summary>
internal static class Cli
{
    public const int Success = 0;
    public const int Failure = 1;
    public const int InvalidUsage = 2;

    public const string Usage = """
        usage:
          spreadsheet-converter to-xlsx [--schema-file <schema.json>] [--sheet-name <name>] [--sheet-rows <n>] <input.csv|-> <output.xlsx>
          spreadsheet-converter to-csv [--sheet <name>] <input.xlsx|.xlsb|.xls> <output.csv|->
          spreadsheet-converter --version
        """;

    public static async Task<int> RunAsync(string[] args, Stream stdin, Stream stdout, TextWriter stderr)
    {
        if (args is ["--version"])
        {
            await WriteLineAsync(stdout, Version());
            return Success;
        }

        if (args.Length == 0)
        {
            await stderr.WriteLineAsync(Usage);
            return InvalidUsage;
        }

        if (!TryParse(args[1..], out var options, out var paths, out var usageError))
        {
            await stderr.WriteLineAsync($"spreadsheet-converter: {usageError}\n{Usage}");
            return InvalidUsage;
        }

        try
        {
            switch (args[0])
            {
                case "to-xlsx":
                    if (!TryBuildToXlsxOptions(options, out var xlsxOptions, out usageError))
                    {
                        break;
                    }
                    var xlsxResult = await CsvToXlsx.ConvertAsync(paths[0] == "-" ? stdin : null, paths[0], paths[1], xlsxOptions);
                    await WriteLineAsync(stdout, xlsxResult.ToJson());
                    return Success;

                case "to-csv":
                    if (!TryBuildToCsvOptions(options, out var csvOptions, out usageError))
                    {
                        break;
                    }
                    var csvResult = await XlsxToCsv.ConvertAsync(paths[0], paths[1] == "-" ? stdout : null, paths[1], csvOptions);
                    if (paths[1] != "-")
                    {
                        await WriteLineAsync(stdout, csvResult.ToJson());
                    }
                    return Success;

                default:
                    usageError = $"unknown command '{args[0]}'";
                    break;
            }
        }
        catch (Exception exception)
        {
            await stderr.WriteLineAsync(Environment.GetEnvironmentVariable("SPREADSHEET_CONVERTER_DEBUG") == "1"
                ? $"spreadsheet-converter: {exception}"
                : $"spreadsheet-converter: {exception.Message}");
            return Failure;
        }

        await stderr.WriteLineAsync($"spreadsheet-converter: {usageError}\n{Usage}");
        return InvalidUsage;
    }

    /// <summary>Splits arguments into "--name value" options and exactly two positional paths.</summary>
    private static bool TryParse(string[] args, out Dictionary<string, string> options, out string[] paths, out string error)
    {
        options = [];
        var positional = new List<string>();
        error = "";

        for (var i = 0; i < args.Length; i++)
        {
            if (args[i].StartsWith("--", StringComparison.Ordinal))
            {
                if (i + 1 >= args.Length)
                {
                    error = $"option {args[i]} needs a value";
                    paths = [];
                    return false;
                }
                options[args[i]] = args[++i];
            }
            else
            {
                positional.Add(args[i]);
            }
        }

        paths = [.. positional];
        if (paths.Length != 2)
        {
            error = "expected an input and an output path";
            return false;
        }
        return true;
    }

    private static bool TryBuildToXlsxOptions(Dictionary<string, string> options, out CsvToXlsxOptions result, out string error)
    {
        result = new CsvToXlsxOptions();
        error = "";

        foreach (var (name, value) in options)
        {
            switch (name)
            {
                case "--schema-file":
                    result = result with { SchemaFile = value };
                    break;
                case "--sheet-name":
                    result = result with { SheetName = value };
                    break;
                case "--sheet-rows":
                    if (!int.TryParse(value, out var sheetRows) || sheetRows < 2 || sheetRows > ExcelLimits.MaxRows)
                    {
                        error = $"--sheet-rows must be between 2 and {ExcelLimits.MaxRows}";
                        return false;
                    }
                    result = result with { SheetRows = sheetRows };
                    break;
                default:
                    error = $"unknown option {name} for to-xlsx";
                    return false;
            }
        }
        return true;
    }

    private static bool TryBuildToCsvOptions(Dictionary<string, string> options, out XlsxToCsvOptions result, out string error)
    {
        result = new XlsxToCsvOptions();
        error = "";

        foreach (var (name, value) in options)
        {
            switch (name)
            {
                case "--sheet":
                    result = result with { Sheet = value };
                    break;
                default:
                    error = $"unknown option {name} for to-csv";
                    return false;
            }
        }
        return true;
    }

    private static string Version() =>
        typeof(Cli).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? "unknown";

    private static async Task WriteLineAsync(Stream stream, string line)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(line + "\n");
        await stream.WriteAsync(bytes);
        await stream.FlushAsync();
    }
}
