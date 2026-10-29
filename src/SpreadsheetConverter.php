<?php

declare(strict_types=1);

namespace Jakobbsm\SpreadsheetConverter;

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
            // Composer may extract dist archives without the executable bit.
            @chmod($path, 0755);
            clearstatcache(true, $path);

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
     * Command converting CSV (RFC 4180, UTF-8, comma separated, with a header row) to xlsx.
     *
     * Write the CSV with fputcsv($handle, $row, ',', '"', '') and pass "-" as $input to stream it through stdin.
     * The schema file is a JSON object mapping column names to text, number, bool, date (Y-m-d) or
     * datetime (Y-m-d H:i:s); unlisted columns, and values that do not parse as their type, are written as text.
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
    ): array {
        return [
            self::binaryPath($binary),
            'to-xlsx',
            ...($schemaFile === null ? [] : ['--schema-file', $schemaFile]),
            ...($sheetName === null ? [] : ['--sheet-name', $sheetName]),
            ...($sheetRows === null ? [] : ['--sheet-rows', (string) $sheetRows]),
            $input,
            $output,
        ];
    }

    /**
     * Command converting one worksheet of an xlsx, xlsb or xls file to CSV readable with fgetcsv($handle, null, ',', '"', '').
     *
     * Pass "-" as $output to stream the CSV through stdout. Values are culture independent: numbers use a dot,
     * booleans are 1/0, dates Y-m-d, date-times Y-m-d H:i:s and times H:i:s.
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
