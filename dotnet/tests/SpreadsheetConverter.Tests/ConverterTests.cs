using System.IO.Compression;
using System.Text;
using SpreadCheetah;
using Sylvan.Data.Excel;

namespace SpreadsheetConverter.Tests;

public sealed class ConverterTests : IDisposable
{
    private readonly string _directory = Directory.CreateTempSubdirectory("spreadsheet-converter-tests-").FullName;

    public void Dispose() => Directory.Delete(_directory, recursive: true);

    [Fact]
    public async Task ToXlsx_TypesValuesBySchema_AndKeepsEverythingElseAsText()
    {
        var input = WriteFile("in.csv", """
            id,employees,active,founded,modified,note
            007,12,true,2024-01-31,2024-01-31 13:45:00,"=HYPERLINK(""x"")"
            008,n/a,maybe,1899-03-05,not a date,"a,""b""
            c"
            009,,,,,
            """);
        var schema = WriteFile("schema.json", """{"employees":"number","active":"bool","founded":"date","modified":"datetime"}""");
        var output = PathFor("out.xlsx");

        var (exitCode, stdout, _) = await RunAsync("to-xlsx", "--schema-file", schema, input, output);

        Assert.Equal(Cli.Success, exitCode);
        Assert.Equal("""{"rows":3,"sheets":1,"clipped_cells":0}""", stdout.Trim());

        using var reader = ExcelDataReader.Create(output, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });
        Assert.True(reader.Read());
        Assert.Equal((ExcelDataType.String, "007"), (reader.GetExcelDataType(0), reader.GetString(0)));
        Assert.Equal((ExcelDataType.Numeric, 12d), (reader.GetExcelDataType(1), reader.GetDouble(1)));
        Assert.Equal((ExcelDataType.Boolean, true), (reader.GetExcelDataType(2), reader.GetBoolean(2)));
        Assert.Equal(new DateTime(2024, 1, 31), reader.GetDateTime(3));
        Assert.Equal(new DateTime(2024, 1, 31, 13, 45, 0), reader.GetDateTime(4));
        Assert.Equal((ExcelDataType.String, "=HYPERLINK(\"x\")"), (reader.GetExcelDataType(5), reader.GetString(5)));

        Assert.True(reader.Read());
        Assert.Equal((ExcelDataType.String, "n/a"), (reader.GetExcelDataType(1), reader.GetString(1)));
        Assert.Equal((ExcelDataType.String, "maybe"), (reader.GetExcelDataType(2), reader.GetString(2)));
        Assert.Equal((ExcelDataType.String, "1899-03-05"), (reader.GetExcelDataType(3), reader.GetString(3)));
        Assert.Equal((ExcelDataType.String, "not a date"), (reader.GetExcelDataType(4), reader.GetString(4)));
        Assert.Equal("a,\"b\"\nc", reader.GetString(5));

        Assert.True(reader.Read());
        Assert.All(Enumerable.Range(1, 5), i => Assert.Equal(ExcelDataType.Null, reader.GetExcelDataType(i)));

        Assert.DoesNotContain("<f>", ReadWorksheetXml(output, 1));
    }

    [Fact]
    public async Task ToXlsx_ContinuesOnNewSheetsWithTheHeaderRepeated()
    {
        var input = WriteFile("in.csv", "id,name\n1,a\n2,b\n3,c\n");
        var output = PathFor("out.xlsx");

        var (exitCode, stdout, _) = await RunAsync("to-xlsx", "--sheet-rows", "3", "--sheet-name", "Companies", input, output);

        Assert.Equal(Cli.Success, exitCode);
        Assert.Equal("""{"rows":3,"sheets":2,"clipped_cells":0}""", stdout.Trim());

        using var reader = ExcelDataReader.Create(output, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });
        Assert.Equal(["Companies", "Companies (2)"], reader.WorksheetNames);
        Assert.True(reader.NextResult());
        Assert.Equal("id", reader.GetName(0));
        Assert.True(reader.Read());
        Assert.Equal("c", reader.GetString(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public async Task ToXlsx_ClipsTextToExcelsCellLimit_WithoutSplittingSurrogatePairs()
    {
        var tooLong = new string('ø', ExcelLimits.MaxCellChars - 1) + "😀😀";
        var input = WriteFile("in.csv", $"note\n{tooLong}\n");
        var output = PathFor("out.xlsx");

        var (exitCode, stdout, _) = await RunAsync("to-xlsx", input, output);

        Assert.Equal(Cli.Success, exitCode);
        Assert.Contains("\"clipped_cells\":1", stdout);
        using var reader = ExcelDataReader.Create(output, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });
        Assert.True(reader.Read());
        Assert.Equal(ExcelLimits.MaxCellChars - 1, reader.GetString(0).Length);
    }

    [Fact]
    public async Task ToXlsx_ReadsCsvFromStdin()
    {
        var output = PathFor("out.xlsx");

        var (exitCode, stdout, _) = await RunAsync(["to-xlsx", "-", output], stdin: "id\n1\n2\n");

        Assert.Equal(Cli.Success, exitCode);
        Assert.Contains("\"rows\":2", stdout);
        Assert.True(File.Exists(output));
    }

    [Theory]
    [InlineData("a,b\n1,2\n3,\"unterminated\n")]
    [InlineData("a,b\n1,2\n3\n")]
    [InlineData("")]
    public async Task ToXlsx_FailsOnInvalidCsv_AndLeavesNoOutputBehind(string csv)
    {
        var input = WriteFile("in.csv", csv);
        var output = PathFor("out.xlsx");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", input, output);

        Assert.Equal(Cli.Failure, exitCode);
        Assert.StartsWith("spreadsheet-converter:", stderr);
        Assert.Equal(["in.csv"], Directory.GetFiles(_directory).Select(Path.GetFileName));
    }

    [Fact]
    public async Task ToXlsx_AppliesTheSchemasFormatCodes()
    {
        var input = WriteFile("in.csv", "revenue,margin,founded,note\n1234.5,0.25,2024-01-31,x\n");
        var schema = WriteFile("schema.json", """
            {
              "revenue": {"type": "number", "format": "_-* #,##0.00 \"kr.\"_-;-* #,##0.00 \"kr.\"_-;_-* \"-\"?? \"kr.\"_-;_-@_-"},
              "margin": {"type": "number", "format": "0.0%"},
              "founded": {"type": "date", "format": "dd-mm-yyyy"},
              "note": {"format": "@"}
            }
            """);
        var output = PathFor("out.xlsx");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", "--schema-file", schema, input, output);

        Assert.True(exitCode == Cli.Success, stderr);
        using var reader = ExcelDataReader.Create(output, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });
        Assert.True(reader.Read());
        Assert.Equal("_-* #,##0.00 \"kr.\"_-;-* #,##0.00 \"kr.\"_-;_-* \"-\"?? \"kr.\"_-;_-@_-", reader.GetFormat(0)?.Format);
        Assert.Equal(1234.5, reader.GetDouble(0));
        Assert.Equal("0.0%", reader.GetFormat(1)?.Format);
        Assert.Equal("dd-mm-yyyy", reader.GetFormat(2)?.Format);
        Assert.Equal(new DateTime(2024, 1, 31), reader.GetDateTime(2));
        Assert.Equal((ExcelDataType.String, "x"), (reader.GetExcelDataType(3), reader.GetString(3)));
    }

    [Theory]
    [InlineData("""{"a": {"type": "number", "format": ""}}""", "format must be")]
    [InlineData("""{"a": {"type": "number", "format": "0 \"kr"}}""", "unclosed quote")]
    [InlineData("""{"a": {"type": "number", "format": "[h:mm"}}""", "unclosed '['")]
    [InlineData("""{"a": {"type": "number", "format": "h]:mm"}}""", "unmatched ']'")]
    [InlineData("""{"a": {"type": "number", "fromat": "0"}}""", "unknown key 'fromat'")]
    [InlineData("""{"a": 1}""", "must be a type name or an object")]
    public async Task ToXlsx_RejectsInvalidSchemas(string json, string error)
    {
        var input = WriteFile("in.csv", "a\n1\n");
        var schema = WriteFile("schema.json", json);

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", "--schema-file", schema, input, PathFor("out.xlsx"));

        Assert.Equal(Cli.Failure, exitCode);
        Assert.Contains(error, stderr);
    }

    [Fact]
    public async Task ToXlsx_RejectsMoreColumnsThanExcelAllows()
    {
        var header = string.Join(",", Enumerable.Range(1, ExcelLimits.MaxColumns + 1).Select(i => $"c{i}"));
        var input = WriteFile("in.csv", $"{header}\n{new string(',', ExcelLimits.MaxColumns)}\n");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", input, PathFor("out.xlsx"));

        Assert.Equal(Cli.Failure, exitCode);
        Assert.Contains("Excel allows at most 16384", stderr);
    }

    [Fact]
    public async Task ToXlsx_ClipsHeaderNamesToExcelsCellLimit()
    {
        var input = WriteFile("in.csv", $"{new string('h', 40_000)},b\n1,2\n");
        var output = PathFor("out.xlsx");

        var (exitCode, stdout, stderr) = await RunAsync("to-xlsx", input, output);

        Assert.True(exitCode == Cli.Success, stderr);
        Assert.Contains("\"clipped_cells\":1", stdout);
        using var reader = ExcelDataReader.Create(output, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });
        Assert.Equal(ExcelLimits.MaxCellChars, reader.GetName(0).Length);
    }

    [Theory]
    [InlineData("a,b\n1,2\n\n3,4\n\n", "a,b\n1,2\n3,4\n")]
    [InlineData("a,b\r\n1,2\r\n\r\n", "a,b\n1,2\n")]
    [InlineData("a\n1\n\n2\n", "a\n1\n\n2\n")]
    public async Task ToXlsx_SkipsBlankLines_UnlessTheyAreAnEmptyValueInASingleColumn(string csv, string expected)
    {
        var output = PathFor("out.xlsx");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", WriteFile("in.csv", csv), output);

        Assert.True(exitCode == Cli.Success, stderr);
        Assert.Equal(expected, (await RunAsync("to-csv", output, "-")).Stdout);
    }

    [Theory]
    [InlineData("a,b\n1,2\n3\n", "line 3: found 1 fields, expected 2 like the header")]
    [InlineData("a,b\n1,2,3\n", "line 2: found 3 fields, expected 2 like the header")]
    public async Task ToXlsx_NamesTheLineWithTheWrongNumberOfFields(string csv, string error)
    {
        var (exitCode, _, stderr) = await RunAsync("to-xlsx", WriteFile("in.csv", csv), PathFor("out.xlsx"));

        Assert.Equal(Cli.Failure, exitCode);
        Assert.Contains(error, stderr);
    }

    [Fact]
    public async Task ToXlsx_ShortensLaterSheetNamesWithoutSplittingSurrogatePairs()
    {
        var name = new string('x', 26) + "😀yyy";
        var output = PathFor("out.xlsx");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", "--sheet-rows", "2", "--sheet-name", name, WriteFile("in.csv", "id\n1\n2\n"), output);

        Assert.True(exitCode == Cli.Success, stderr);
        using var reader = ExcelDataReader.Create(output, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });
        Assert.Equal([name, new string('x', 26) + " (2)"], reader.WorksheetNames);
    }

    [Fact]
    public async Task ToXlsx_AcceptsRepeatedColumnNames()
    {
        var input = WriteFile("in.csv", "id,name,name,,\n1,a,b,c,d\n");
        var schema = WriteFile("schema.json", """{"id":"number"}""");
        var output = PathFor("out.xlsx");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", "--schema-file", schema, input, output);

        Assert.True(exitCode == Cli.Success, stderr);
        Assert.Equal("id,name,name,,\n1,a,b,c,d\n", (await RunAsync("to-csv", output, "-")).Stdout);
    }

    [Fact]
    public async Task ToXlsx_RejectsARepeatedColumnNameTheSchemaTypes()
    {
        var input = WriteFile("in.csv", "id,id\n1,2\n");
        var schema = WriteFile("schema.json", """{"id":"number"}""");

        var (exitCode, _, stderr) = await RunAsync("to-xlsx", "--schema-file", schema, input, PathFor("out.xlsx"));

        Assert.Equal(Cli.Failure, exitCode);
        Assert.Contains("column 'id' appears 2 times", stderr);
    }

    [Fact]
    public async Task ToCsv_WritesCultureIndependentValuesReadableByFgetcsv()
    {
        var input = WriteFile("in.csv", """
            cvr,employees,ratio,active,founded,modified,note
            12345678,12,0.25,1,2024-01-31,2024-01-31 13:45:00,"a,""b""
            c"
            87654321,,,0,,,
            """);
        var schema = WriteFile("schema.json", """{"cvr":"number","employees":"number","ratio":"number","active":"bool","founded":"date","modified":"datetime"}""");
        var workbook = PathFor("data.xlsx");
        await RunAsync("to-xlsx", "--schema-file", schema, input, workbook);
        var output = PathFor("out.csv");

        var (exitCode, stdout, _) = await RunAsync("to-csv", workbook, output);

        Assert.Equal(Cli.Success, exitCode);
        Assert.Equal("""{"rows":2,"columns":7,"error_cells":0,"ignored_cells":0}""", stdout.Trim());
        Assert.Equal("""
            cvr,employees,ratio,active,founded,modified,note
            12345678,12,0.25,1,2024-01-31,2024-01-31 13:45:00,"a,""b""
            c"
            87654321,,,0,,,

            """.ReplaceLineEndings("\n"), File.ReadAllText(output));
    }

    [Fact]
    public async Task ToCsv_WritesToStdout_WithoutStatistics()
    {
        var workbook = await CreateWorkbookAsync("id,name\n1,a\n");

        var (exitCode, stdout, _) = await RunAsync("to-csv", workbook, "-");

        Assert.Equal(Cli.Success, exitCode);
        Assert.Equal("id,name\n1,a\n", stdout);
    }

    [Fact]
    public async Task ToCsv_ReadsTheRequestedSheet_AndNamesTheAvailableOnesWhenMissing()
    {
        var input = WriteFile("in.csv", "id\n1\n2\n");
        var workbook = PathFor("data.xlsx");
        await RunAsync("to-xlsx", "--sheet-rows", "2", "--sheet-name", "Data", input, workbook);

        var (exitCode, stdout, _) = await RunAsync("to-csv", "--sheet", "Data (2)", workbook, "-");
        Assert.Equal(Cli.Success, exitCode);
        Assert.Equal("id\n2\n", stdout);

        (exitCode, _, var stderr) = await RunAsync("to-csv", "--sheet", "Missing", workbook, "-");
        Assert.Equal(Cli.Failure, exitCode);
        Assert.Contains("available: Data, Data (2)", stderr);
    }

    [Fact]
    public async Task ToCsv_DetectsTheWorkbookTypeFromTheContents_NotTheExtension()
    {
        var upload = PathFor("phpAbC123");
        File.Copy(await CreateWorkbookAsync("id,name\n1,a\n"), upload);

        var (exitCode, stdout, stderr) = await RunAsync("to-csv", upload, "-");

        Assert.True(exitCode == Cli.Success, stderr);
        Assert.Equal("id,name\n1,a\n", stdout);
    }

    [Fact]
    public async Task ToCsv_RejectsInputThatIsNotAWorkbook()
    {
        var (exitCode, _, stderr) = await RunAsync("to-csv", WriteFile("in.xlsx", "id\n1\n"), "-");

        Assert.Equal(Cli.Failure, exitCode);
        Assert.Contains("not an xlsx, xlsb or xls workbook", stderr);
    }

    [Fact]
    public async Task ToCsv_WritesElapsedTimeFormatsAsTotalHours()
    {
        var input = WriteFile("in.csv", "duration,time\n1.25,0.5\n-0.0625,0.25\n");
        var schema = WriteFile("schema.json", """{"duration": {"type": "number", "format": "[h]:mm:ss"}, "time": {"type": "number", "format": "hh:mm:ss"}}""");
        var workbook = PathFor("data.xlsx");
        await RunAsync("to-xlsx", "--schema-file", schema, input, workbook);

        var (exitCode, stdout, stderr) = await RunAsync("to-csv", workbook, "-");

        Assert.True(exitCode == Cli.Success, stderr);
        Assert.Equal("duration,time\n30:00:00,12:00:00\n-01:30:00,06:00:00\n", stdout);
    }

    [Theory]
    [InlineData("[h]:mm:ss", true)]
    [InlineData("[mm]:ss", true)]
    [InlineData("[Red][h]:mm", true)]
    [InlineData("hh:mm:ss", false)]
    [InlineData("[Red]0.00", false)]
    [InlineData("\"[h]\" 0", false)]
    [InlineData("[$-409]h:mm AM/PM", false)]
    public void ToCsv_RecognisesElapsedTimeFormats(string format, bool expected)
    {
        Assert.Equal(expected, XlsxToCsv.IsElapsedTimeFormat(format));
    }

    [Fact]
    public async Task ToCsv_SkipsBlankRowsAboveTheHeader()
    {
        var workbook = PathFor("blank-top.xlsx");
        await using (var stream = File.Create(workbook))
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Report");
            await spreadsheet.AddRowAsync(Array.Empty<Cell>());
            await spreadsheet.AddRowAsync([new Cell("id"), new Cell("name")]);
            await spreadsheet.AddRowAsync([new Cell(1), new Cell("a")]);
            await spreadsheet.FinishAsync();
        }
        Assert.Contains("<row r=\"1\"></row><row r=\"2\">", ReadWorksheetXml(workbook, 1));

        var (exitCode, stdout, stderr) = await RunAsync("to-csv", workbook, "-");

        Assert.True(exitCode == Cli.Success, stderr);
        Assert.Equal("id,name\n1,a\n", stdout);
    }

    [Theory]
    [InlineData(0.5, "12:00:00")]
    [InlineData(45_000.75, "18:00:00")]
    [InlineData(86_399.4 / 86_400, "23:59:59")]
    [InlineData(86_399.6 / 86_400, "23:59:59")]
    public void ToCsv_FormatsTimesWithoutWrappingPastMidnight(double value, string expected)
    {
        Assert.Equal(expected, XlsxToCsv.FormatTime(value));
    }

    [Theory]
    [InlineData]
    [InlineData("to-pdf", "in", "out")]
    [InlineData("to-xlsx", "only-one-path")]
    [InlineData("to-xlsx", "--sheet-rows", "1", "in.csv", "out.xlsx")]
    [InlineData("to-csv", "--unknown", "x", "in.xlsx", "out.csv")]
    public async Task RejectsInvalidUsage(params string[] args)
    {
        var (exitCode, _, stderr) = await RunAsync(args);

        Assert.Equal(Cli.InvalidUsage, exitCode);
        Assert.Contains("usage:", stderr);
    }

    private async Task<string> CreateWorkbookAsync(string csv)
    {
        var workbook = PathFor($"{Guid.NewGuid():N}.xlsx");
        var (exitCode, _, stderr) = await RunAsync("to-xlsx", WriteFile($"{Guid.NewGuid():N}.csv", csv), workbook);
        Assert.True(exitCode == Cli.Success, stderr);
        return workbook;
    }

    private static Task<(int ExitCode, string Stdout, string Stderr)> RunAsync(params string[] args) => RunAsync(args, stdin: "");

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunAsync(string[] args, string stdin)
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(stdin));
        using var output = new MemoryStream();
        using var error = new StringWriter();

        var exitCode = await Cli.RunAsync(args, input, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private string PathFor(string name) => Path.Combine(_directory, name);

    private string WriteFile(string name, string content)
    {
        var path = PathFor(name);
        File.WriteAllText(path, content.ReplaceLineEndings("\n"), new UTF8Encoding(false));
        return path;
    }

    private static string ReadWorksheetXml(string workbook, int sheet)
    {
        using var archive = ZipFile.OpenRead(workbook);
        using var reader = new StreamReader(archive.GetEntry($"xl/worksheets/sheet{sheet}.xml")!.Open());
        return reader.ReadToEnd();
    }
}
