<?php

declare(strict_types=1);

namespace Jakobbsm\SpreadsheetConverter;

use Generator;
use RuntimeException;

/**
 * Locates the bundled converter executable for the current platform and builds its command lines.
 *
 * The commands are argv arrays (no shell involved) for Laravel's Process facade, Symfony Process or proc_open.
 * On success the converter exits with 0 and, unless it writes CSV to stdout, prints one JSON line of statistics.
 * Exit code 1 means the conversion failed (message on stderr), 2 means invalid usage.
 */
final class SpreadsheetConverter
{
    /** Environment variable pointing at a converter executable to use instead of the bundled one. */
    public const PATH_ENV = 'SPREADSHEET_CONVERTER_PATH';

    /** @var list<string> */
    public const SUPPORTED_PLATFORMS = ['linux-arm64', 'linux-x64', 'osx-arm64', 'win-x64'];

    /**
     * Absolute path of the executable to run: $override, then SPREADSHEET_CONVERTER_PATH, then the bundled binary.
     *
     * @throws RuntimeException when the platform is unsupported or the executable is missing.
     */
    public static function binaryPath(?string $override = null): string
    {
        $configured = $override ?: (getenv(self::PATH_ENV) ?: null);
        $path = $configured ?? self::bundledBinaryPath(self::platform());

        if (! is_file($path)) {
            throw new RuntimeException($configured === null
                ? "The bundled spreadsheet converter is missing at {$path}. Install a tagged release of jakobbsm/spreadsheet-converter, or set ".self::PATH_ENV.'.'
                : "The spreadsheet converter configured at {$path} does not exist.");
        }

        if (PHP_OS_FAMILY !== 'Windows' && ! is_executable($path)) {
            // Composer may extract dist archives without the executable bit. A configured path is the user's own
            // file, so it is left alone.
            if ($configured === null) {
                @chmod($path, 0755);
                clearstatcache(true, $path);
            }

            if (! is_executable($path)) {
                throw new RuntimeException("The spreadsheet converter at {$path} is not executable.");
            }
        }

        return $path;
    }

    /**
     * The .NET runtime identifier of the binary for a machine, such as "osx-arm64" or "linux-x64".
     *
     * @throws RuntimeException when no binary is built for the platform.
     */
    public static function platform(string $osFamily = PHP_OS_FAMILY, ?string $machine = null): string
    {
        $machine = strtolower($machine ?? php_uname('m'));

        $os = match ($osFamily) {
            'Darwin' => 'osx',
            'Linux' => 'linux',
            'Windows' => 'win',
            default => null,
        };

        $architecture = match ($machine) {
            'arm64', 'aarch64' => 'arm64',
            'x86_64', 'amd64', 'x64' => 'x64',
            default => null,
        };

        // Windows on ARM runs the x64 build through emulation.
        if ($os === 'win') {
            $architecture = $architecture === null ? null : 'x64';
        }

        $platform = "{$os}-{$architecture}";

        if ($os === null || $architecture === null || ! in_array($platform, self::SUPPORTED_PLATFORMS, true)) {
            throw new RuntimeException("The spreadsheet converter has no build for {$osFamily} on {$machine}. Supported: ".implode(', ', self::SUPPORTED_PLATFORMS).'.');
        }

        return $platform;
    }

    public static function bundledBinaryPath(string $platform): string
    {
        $extension = str_starts_with($platform, 'win-') ? '.exe' : '';

        return dirname(__DIR__)."/bin/spreadsheet-converter-{$platform}{$extension}";
    }

    /**
     * Command converting CSV (RFC 4180, UTF-8, with a header row; the delimiter is detected unless given) to xlsx.
     *
     * Write the CSV with fputcsv($handle, $row, ',', '"', '') and pass "-" as $input to stream it through stdin.
     * The schema file is a JSON object mapping column names to text, number, bool, date (Y-m-d) or
     * datetime (Y-m-d H:i:s), or to an object that adds an Excel number format code, such as
     * {"type": "number", "format": "#,##0.00"}. Unlisted columns, and values that do not parse as their type, are
     * written as text. Column names may repeat, except for names the schema lists; blank lines are skipped.
     * With $strict, invalid UTF-8 and schema columns missing from the header fail the conversion instead of being
     * passed through (invalid bytes otherwise become U+FFFD, and a missing column is ignored).
     * $delimiter is ',', ';', '|' or "\t". Without it, the converter detects the delimiter from the first line, which
     * suits uploaded files; pass ',' for CSV you write yourself, such as csvLines() output.
     *
     * @return list<string>
     */
    public static function toXlsxCommand(
        string $input,
        string $output,
        ?string $schemaFile = null,
        ?string $sheetName = null,
        ?int $sheetRows = null,
        ?string $binary = null,
        bool $strict = false,
        ?string $delimiter = null,
    ): array {
        return [
            self::binaryPath($binary),
            'to-xlsx',
            ...($schemaFile === null ? [] : ['--schema-file', $schemaFile]),
            ...($sheetName === null ? [] : ['--sheet-name', $sheetName]),
            ...($sheetRows === null ? [] : ['--sheet-rows', (string) $sheetRows]),
            ...($strict ? ['--strict'] : []),
            ...($delimiter === null ? [] : ['--delimiter', $delimiter === "\t" ? 'tab' : $delimiter]),
            $input,
            $output,
        ];
    }

    /**
     * Encodes rows as CSV lines in the form to-xlsx reads, for streaming into its stdin:
     * Process::input(SpreadsheetConverter::csvLines($rows))->run(SpreadsheetConverter::toXlsxCommand('-', $path, delimiter: ',')).
     *
     * The first row is the header; a one-column header is always quoted, so an empty one is not a blank line.
     * Booleans become 1/0, null an empty value and floats their shortest exact form
     * (0.1 + 0.2 stays 0.30000000000000004, where fputcsv would round to PHP's 14-digit precision); INF and NAN are
     * written as the text INF and NAN. Format dates yourself as Y-m-d or Y-m-d H:i:s so a date or datetime schema
     * column recognises them.
     *
     * @param  iterable<iterable<scalar|\Stringable|null>>  $rows
     * @return Generator<int, string>
     */
    public static function csvLines(iterable $rows): Generator
    {
        $buffer = fopen('php://memory', 'w+b');
        $header = true;

        try {
            foreach ($rows as $row) {
                $fields = [];
                foreach ($row as $value) {
                    $fields[] = match (true) {
                        $value === true => '1',
                        $value === false => '0',
                        is_float($value) => var_export($value, true),
                        default => $value,
                    };
                }

                if ($header && count($fields) === 1) {
                    // fputcsv writes an empty one-field row as a blank line, which to-xlsx would skip, making the
                    // first data row the header. Quoting the field keeps it.
                    $header = false;
                    yield '"'.str_replace('"', '""', (string) $fields[0])."\"\n";

                    continue;
                }
                $header = false;

                fputcsv($buffer, $fields, ',', '"', '', "\n");
                rewind($buffer);
                yield (string) stream_get_contents($buffer);
                ftruncate($buffer, 0);
                rewind($buffer);
            }
        } finally {
            fclose($buffer);
        }
    }

    /**
     * Command converting one worksheet of an xlsx, xlsb or xls file to CSV readable with fgetcsv($handle, null, ',', '"', '').
     *
     * The workbook type is detected from the contents, so upload paths without an extension work. The first non-blank
     * row is the header. Pass "-" as $output to stream the CSV through stdout. Values are culture independent: numbers
     * use a dot, booleans are 1/0, dates Y-m-d, date-times Y-m-d H:i:s (chosen by the cell's format), times H:i:s and
     * durations (elapsed-time formats such as [h]:mm:ss) in total hours, such as 30:00:00.
     *
     * @return list<string>
     */
    public static function toCsvCommand(string $input, string $output, ?string $sheet = null, ?string $binary = null): array
    {
        return [
            self::binaryPath($binary),
            'to-csv',
            ...($sheet === null ? [] : ['--sheet', $sheet]),
            $input,
            $output,
        ];
    }
}
