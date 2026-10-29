# spreadsheet-converter

Fast, streaming CSV ⇄ XLSX conversion for PHP. Composer installs a native executable for macOS (Apple Silicon),
Windows (x64) and Linux (x64/arm64). No .NET runtime or other system dependency is needed.

Built on [SpreadCheetah](https://github.com/sveinungf/spreadcheetah) (writing), [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel)
(reading) and [Sep](https://github.com/nietras/Sep) (CSV), compiled with NativeAOT. 1M rows × 56 columns convert
from CSV to XLSX in about 7 seconds with about 20 MB of memory.

## Installation

```sh
composer require jakobbsm/spreadsheet-converter
```

## Usage from PHP

```php
use Jakobbsm\SpreadsheetConverter\SpreadsheetConverter;
use Illuminate\Support\Facades\Process;

// Export: stream rows into the converter's stdin. The first row is the header.
$rows = (function () {
    yield ['cvr', 'name', 'employees', 'founded'];
    foreach (Company::query()->lazy() as $company) {
        yield [$company->cvr, $company->name, $company->employees, $company->founded?->format('Y-m-d')];
    }
})();

$result = Process::input(SpreadsheetConverter::csvLines($rows))
    ->run(SpreadsheetConverter::toXlsxCommand('-', $xlsxPath, schemaFile: $schemaPath, sheetName: 'Companies', delimiter: ','))
    ->throw();
$stats = json_decode($result->output(), true); // ['rows' => …, 'sheets' => …, 'clipped_cells' => …]

// Import: read any xlsx/xlsb/xls file as CSV.
Process::run(SpreadsheetConverter::toCsvCommand($uploadedPath, $csvPath))->throw();
```

`csvLines()` encodes each row with `fputcsv($handle, $row, ',', '"', '')`, writes booleans as `1`/`0`, null as an
empty value and floats in their shortest exact form (where `fputcsv` alone rounds to 14 digits); `INF` and `NAN` are
written as that text. A one-column header is always quoted, so an empty one is not a blank line. Pass
`delimiter: ','` with `csvLines()` output, as above, rather than relying on detection. Format dates yourself as `Y-m-d` or `Y-m-d H:i:s`. Read the converter's CSV output with
`fgetcsv($handle, null, ',', '"', '')`.

To use another executable, such as a local build while working on the converter, pass `binary:` to the command
methods or set the `SPREADSHEET_CONVERTER_PATH` environment variable. The package reads the variable with `getenv()`,
so in Laravel a value in `.env` stops working after `php artisan config:cache`; pass `binary: config('…')` instead.

## Command line

```
spreadsheet-converter to-xlsx [--schema-file <schema.json>] [--sheet-name <name>] [--sheet-rows <n>] [--delimiter <,|;|tab>] [--strict] <input.csv|-> <output.xlsx>
spreadsheet-converter to-csv [--sheet <name>] <input.xlsx|.xlsb|.xls> <output.csv|->
spreadsheet-converter --version
```

The exit code is 0 on success, 1 when the conversion fails (the reason is on stderr) and 2 on invalid usage. On
success, one JSON line of statistics is printed, unless `to-csv` writes the CSV itself to stdout. Output files are
written to a temporary name and moved into place at the end, so a failed run never leaves a partial file. Set
`SPREADSHEET_CONVERTER_DEBUG=1` to print stack traces.

### to-xlsx

The input is RFC 4180 CSV: UTF-8, with a header row. Every row must have as many fields as the
header; blank lines below the header are skipped, except in a single-column file, where they are empty
values. A line of only spaces
is not blank: it is a row with one field. A quoted field that is never closed is rejected, which catches input cut
off mid-row. Column names may repeat, except for a name the schema lists.

The delimiter is detected from the first line (comma, semicolon, tab or pipe), so the semicolon-separated "CSV" that
Excel and Numbers save in many European locales, Danish included, converts as is. The detection counts characters
without regard to quotes and falls back to a semicolon when the first line has none, so a one-column file whose
values contain a semicolon needs `--delimiter ','`. A blank first line is rejected, since nothing can be detected
from it; with `--delimiter` it is an empty one-column header. Pass `--delimiter` (`delimiter:` in PHP) with `,`, `;`, `|` or `tab` whenever you know the separator, such as
for CSV your app writes itself. Decimal commas (`1234,5`) and local dates (`15/03/2001`) in such files are not parsed
and stay text.

By default the data is passed through as it is: invalid UTF-8 becomes `�` (U+FFFD), and schema columns that are not
in the header are ignored. With `--strict` (`strict: true` in PHP), both fail the conversion instead, which catches
an input in the wrong encoding and a typo in the schema. The UTF-8 error names the line. Because the check looks for
`�` after decoding, strict mode also rejects a `�` that was already in the input, which usually means the data was
damaged by an earlier conversion. Use `--strict` for CSV that users upload: Excel's own "CSV" format is
not UTF-8 (only "CSV UTF-8" is), and without `--strict` its æøå silently become `�`.

Without a schema every cell is text, including numbers, so Excel shows them as "number stored as text" and won't sum
them. Excel keeps 15 significant digits for numbers, so keep longer identifiers, such as 64-bit IDs, as `text`.

The schema is a JSON object that maps column names to types:

| type       | accepted values                                                   | cell        |
|------------|-------------------------------------------------------------------|-------------|
| `text`     | anything (the default for unlisted columns)                       | string      |
| `number`   | `12`, `-0.5`, `1e6` (dot decimal separator, no grouping)          | number      |
| `bool`     | `1`/`0`, `true`/`false`, `t`/`f`, in any case                    | boolean     |
| `date`     | `2024-01-31`                                                      | date        |
| `datetime` | `2024-01-31 13:45:00`, also with `T`, fractions or without seconds | date + time |

A column can also be given an Excel number format code, which is passed to Excel as is, and a `width` in characters
(Excel's default is about 8.4). The easiest way to get a format code is from Excel itself: format a cell, then copy
the code from **Format Cells → Custom**. `type` defaults to `text`.

```json
{
  "employees": "number",
  "revenue": {"type": "number", "format": "_-* #,##0.00 \"kr.\"_-;-* #,##0.00 \"kr.\"_-;_-* \"-\"?? \"kr.\"_-;_-@_-"},
  "margin": {"type": "number", "format": "0.0%"},
  "duration": {"type": "number", "format": "[h]:mm:ss"},
  "founded": {"type": "date", "format": "dd-mm-yyyy"},
  "description": {"width": 40}
}
```

Dates default to `yyyy-mm-dd` and date-times to `yyyy-mm-dd hh:mm:ss`; other types have no format. Times and
durations are numbers in days (`0.5` is 12:00, `1.25` is 30 hours). Excel reports a file with a malformed format
code as corrupt, so empty and overlong codes and unclosed quotes or brackets are rejected, but a code is not
otherwise checked.

Excel shows `#####` for a date or formatted number that does not fit, so date and date-time columns get a width that
fits their format, and number columns with a format get 18, enough for `9.999.999,99 kr.` in Accounting. Rows are
streamed, so widths cannot be measured from the data; set `width` for anything wider.

A value that does not parse as its column's type is written as text, so nothing is lost. Dates before
1900-03-01 stay text, because Excel cannot represent them. Text is always written as a string cell and never as a
formula. Text keeps its leading and trailing spaces, and text with a line break (`\n`, `\r\n` or `\r`), header
included, gets wrap text, so the break shows. Surrounding spaces are ignored when parsing a typed value (`" 12 "` is
the number 12); a value that falls back to text keeps them.
The header row is bold, frozen and has an auto-filter. Text over Excel's 32,767-character cell limit, header names
included, is clipped and counted in `clipped_cells`. Rows beyond `--sheet-rows` (default 1,048,576, header included)
continue on a new sheet, `Name (2)`, with the header repeated. More than 16,384 columns is rejected.

### to-csv

The workbook type is detected from the file contents, so an input without an extension (such as a PHP upload's
temporary file) works. Reads the first visible worksheet, or the one named with `--sheet` (ignoring case, like
Excel). The first non-blank row is
the header and sets the column count: shorter rows are padded, and non-empty cells beyond the header are dropped
and counted in `ignored_cells`.
Values are culture independent: numbers use a dot decimal separator (whole numbers have no decimals), booleans are
`1`/`0`, dates are `Y-m-d`, date-times `Y-m-d H:i:s` and times `H:i:s`. Whether a cell is a date or a date-time
follows its format, so a date-time column stays `Y-m-d H:i:s` at midnight; a date format that hides a time other
than midnight still gets the time. A value below 1 in a format that shows a time is written as a time; a negative time, which Excel shows as `#####`,
is written as the number. A
date-formatted number that is not a real date (negative, in Excel's fictional early 1900, or after 9999-12-31,
which Excel shows as `#####`) is written as the number. Durations, meaning cells with an elapsed-time
format such as `[h]:mm:ss`, are written in total hours, such as `30:00:00`. Formula errors such as `#N/A` become empty
values and are counted in `error_cells`. The output is UTF-8 without BOM, with `\n` line endings.

### Known limitations

- A carriage return in text (`\r` or `\r\n`) comes back from xlsx as `\n`, because XML readers normalise line
  endings, Excel included. Line breaks in Excel cells are `\n` anyway.
- Numbers are written in the shortest form that reads back exactly, so a value such as `0.1 + 0.2` comes out as
  `0.30000000000000004`.
- `to-csv` leaves out empty rows at the end of a sheet, as Sylvan does by default, so a `to-xlsx` export ending in
  empty rows comes back with fewer rows. Empty rows elsewhere are kept.
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
