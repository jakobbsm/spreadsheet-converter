# spreadsheet-converter

Fast, streaming CSV ⇄ XLSX conversion for PHP. Composer installs a native executable for macOS (Apple Silicon),
Windows (x64) and Linux (x64/arm64). No .NET runtime or other system dependency is needed.

Built on [SpreadCheetah](https://github.com/sveinungf/spreadcheetah) (writing), [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel)
(reading) and [Sep](https://github.com/nietras/Sep) (CSV), compiled with NativeAOT. 1M rows × 56 columns convert
from CSV to XLSX in about 7 seconds with about 20 MB of memory.

## Installation

```sh
composer require datapeeps/spreadsheet-converter
```

## Usage from PHP

```php
use Datapeeps\SpreadsheetConverter\SpreadsheetConverter;
use Illuminate\Support\Facades\Process;

// Export: stream CSV rows into the converter's stdin.
$result = Process::input($csvGenerator)
    ->run(SpreadsheetConverter::toXlsxCommand('-', $xlsxPath, schemaFile: $schemaPath, sheetName: 'Companies'))
    ->throw();
$stats = json_decode($result->output(), true); // ['rows' => …, 'sheets' => …, 'clipped_cells' => …]

// Import: read any xlsx/xlsb/xls file as CSV.
Process::run(SpreadsheetConverter::toCsvCommand($uploadedPath, $csvPath))->throw();
```

Write CSV for the converter with `fputcsv($handle, $row, ',', '"', '')` and read its output with
`fgetcsv($handle, null, ',', '"', '')`.

Set `SPREADSHEET_CONVERTER_PATH` to use another executable, such as a local build while working on the converter.

## Command line

```
spreadsheet-converter to-xlsx [--schema-file <schema.json>] [--sheet-name <name>] [--sheet-rows <n>] <input.csv|-> <output.xlsx>
spreadsheet-converter to-csv [--sheet <name>] <input.xlsx|.xlsb|.xls> <output.csv|->
spreadsheet-converter --version
```

The exit code is 0 on success, 1 when the conversion fails (the reason is on stderr) and 2 on invalid usage. On
success, one JSON line of statistics is printed, unless `to-csv` writes the CSV itself to stdout. Output files are
written to a temporary name and moved into place at the end, so a failed run never leaves a partial file. Set
`SPREADSHEET_CONVERTER_DEBUG=1` to print stack traces.

### to-xlsx

The input is RFC 4180 CSV: UTF-8, comma separated, with a header row. Every row must have as many fields as the
header; blank lines are skipped, except in a single-column file, where they are empty values. A quoted field that is
never closed is rejected, which catches input cut off mid-row. Column names may repeat, except for a name the
schema lists.

The schema is a JSON object that maps column names to types:

| type       | accepted values                                                   | cell        |
|------------|-------------------------------------------------------------------|-------------|
| `text`     | anything (the default for unlisted columns)                       | string      |
| `number`   | `12`, `-0.5`, `1e6` (dot decimal separator, no grouping)          | number      |
| `bool`     | `1`/`0`, `true`/`false`, `t`/`f`, in any case                    | boolean     |
| `date`     | `2024-01-31`                                                      | date        |
| `datetime` | `2024-01-31 13:45:00`, also with `T`, fractions or without seconds | date + time |

A column can also be given an Excel number format code, which is passed to Excel as is. The easiest way to get one
is from Excel itself: format a cell, then copy the code from **Format Cells → Custom**. `type` defaults to `text`.

```json
{
  "employees": "number",
  "revenue": {"type": "number", "format": "_-* #,##0.00 \"kr.\"_-;-* #,##0.00 \"kr.\"_-;_-* \"-\"?? \"kr.\"_-;_-@_-"},
  "margin": {"type": "number", "format": "0.0%"},
  "duration": {"type": "number", "format": "[h]:mm:ss"},
  "founded": {"type": "date", "format": "dd-mm-yyyy"}
}
```

Dates default to `yyyy-mm-dd` and date-times to `yyyy-mm-dd hh:mm:ss`; other types have no format. Times and
durations are numbers in days (`0.5` is 12:00, `1.25` is 30 hours). Excel reports a file with a malformed format
code as corrupt, so empty and overlong codes and unclosed quotes or brackets are rejected, but a code is not
otherwise checked.

A value that does not parse as its column's type is written as text, so nothing is lost. Dates before
1900-03-01 stay text, because Excel cannot represent them. Text is always written as a string cell and never as a
formula. The header row is bold, frozen and has an auto-filter. Text over Excel's 32,767-character cell limit,
header names included, is clipped and counted in `clipped_cells`. Rows beyond `--sheet-rows` (default 1,048,576, header included) continue on
a new sheet, `Name (2)`, with the header repeated. More than 16,384 columns is rejected.

### to-csv

The workbook type is detected from the file contents, so an input without an extension (such as a PHP upload's
temporary file) works. Reads the first visible worksheet, or the one named with `--sheet`. The first non-blank row is
the header and sets the column count: shorter rows are padded, and non-empty cells beyond the header are dropped
and counted in `ignored_cells`.
Values are culture independent: numbers use a dot decimal separator (whole numbers have no decimals), booleans are
`1`/`0`, dates are `Y-m-d`, date-times `Y-m-d H:i:s` and times `H:i:s`. Durations, meaning cells with an elapsed-time
format such as `[h]:mm:ss`, are written in total hours, such as `30:00:00`. Formula errors such as `#N/A` become empty
values and are counted in `error_cells`. The output is UTF-8 without BOM, with `\n` line endings.

### Known limitations

- Exported text cells are not marked `xml:space="preserve"`, so Excel and other readers may drop leading and
  trailing spaces.
- The Linux builds need glibc 2.35 or newer. Alpine and other musl-based hosts are not supported.

## Development

```sh
dotnet test dotnet/SpreadsheetConverter.slnx
dotnet publish dotnet/src/SpreadsheetConverter -c Release -r osx-arm64 -o artifacts/osx-arm64
composer install && vendor/bin/phpunit
```

To run the PHP round-trip test against a local build, copy it to `bin/spreadsheet-converter-<platform>` (create
`bin/` first; it is not part of the repository).

## Releasing

Run the **Release** workflow with a version such as `1.2.0`. It builds and tests on each platform, commits the four
binaries on a commit that only the `v1.2.0` tag points to (main stays free of binaries), and creates a GitHub
release with the binaries attached.
