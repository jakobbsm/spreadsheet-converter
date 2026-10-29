<?php

declare(strict_types=1);

namespace Jakobbsm\SpreadsheetConverter\Tests;

use Jakobbsm\SpreadsheetConverter\SpreadsheetConverter;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use RuntimeException;

final class SpreadsheetConverterTest extends TestCase
{
    private string $directory;

    protected function setUp(): void
    {
        $this->directory = sys_get_temp_dir().'/spreadsheet-converter-'.bin2hex(random_bytes(6));
        mkdir($this->directory);
        putenv(SpreadsheetConverter::PATH_ENV);
    }

    protected function tearDown(): void
    {
        putenv(SpreadsheetConverter::PATH_ENV);
        array_map(unlink(...), glob($this->directory.'/*') ?: []);
        rmdir($this->directory);
    }

    /** @return iterable<string, array{string, string, string}> */
    public static function platforms(): iterable
    {
        yield 'apple silicon' => ['Darwin', 'arm64', 'osx-arm64'];
        yield 'linux arm' => ['Linux', 'aarch64', 'linux-arm64'];
        yield 'linux intel' => ['Linux', 'x86_64', 'linux-x64'];
        yield 'windows' => ['Windows', 'AMD64', 'win-x64'];
        yield 'windows on arm' => ['Windows', 'ARM64', 'win-x64'];
    }

    #[DataProvider('platforms')]
    public function test_it_maps_machines_to_bundled_builds(string $osFamily, string $machine, string $expected): void
    {
        $this->assertSame($expected, SpreadsheetConverter::platform($osFamily, $machine));
    }

    public function test_it_rejects_platforms_without_a_build(): void
    {
        $this->expectException(RuntimeException::class);
        $this->expectExceptionMessage('no build for Darwin on x86_64');

        SpreadsheetConverter::platform('Darwin', 'x86_64');
    }

    public function test_windows_binaries_have_the_exe_extension(): void
    {
        $this->assertStringEndsWith('/bin/spreadsheet-converter-win-x64.exe', SpreadsheetConverter::bundledBinaryPath('win-x64'));
        $this->assertStringEndsWith('/bin/spreadsheet-converter-linux-x64', SpreadsheetConverter::bundledBinaryPath('linux-x64'));
    }

    public function test_the_environment_variable_overrides_the_bundled_binary(): void
    {
        $binary = $this->executable('custom-converter');
        putenv(SpreadsheetConverter::PATH_ENV.'='.$binary);

        $this->assertSame($binary, SpreadsheetConverter::binaryPath());
    }

    public function test_an_explicit_override_wins_over_the_environment_variable(): void
    {
        putenv(SpreadsheetConverter::PATH_ENV.'='.$this->executable('from-env'));
        $explicit = $this->executable('explicit');

        $this->assertSame($explicit, SpreadsheetConverter::binaryPath($explicit));
    }

    public function test_it_fails_clearly_when_the_configured_binary_is_missing(): void
    {
        $this->expectException(RuntimeException::class);
        $this->expectExceptionMessage('does not exist');

        SpreadsheetConverter::binaryPath($this->directory.'/missing');
    }

    public function test_it_leaves_a_configured_binary_without_the_executable_bit_alone(): void
    {
        if (PHP_OS_FAMILY === 'Windows') {
            $this->markTestSkipped('Windows has no executable bit.');
        }

        $binary = $this->executable('converter');
        chmod($binary, 0644);

        try {
            SpreadsheetConverter::binaryPath($binary);
            $this->fail('Expected the non-executable binary to be rejected.');
        } catch (RuntimeException $exception) {
            $this->assertStringContainsString('is not executable', $exception->getMessage());
        }

        clearstatcache(true, $binary);
        $this->assertSame(0644, fileperms($binary) & 0777);
    }

    public function test_it_builds_command_lines_without_a_shell(): void
    {
        $binary = $this->executable('converter');

        $this->assertSame(
            [$binary, 'to-xlsx', '--schema-file', 'schema.json', '--sheet-name', 'Companies', '--sheet-rows', '500000', '-', 'out.xlsx'],
            SpreadsheetConverter::toXlsxCommand('-', 'out.xlsx', 'schema.json', 'Companies', 500000, $binary),
        );
        $this->assertSame([$binary, 'to-xlsx', 'in.csv', 'out.xlsx'], SpreadsheetConverter::toXlsxCommand('in.csv', 'out.xlsx', binary: $binary));
        $this->assertSame([$binary, 'to-xlsx', '--strict', '-', 'out.xlsx'], SpreadsheetConverter::toXlsxCommand('-', 'out.xlsx', binary: $binary, strict: true));
        $this->assertSame([$binary, 'to-xlsx', '--delimiter', ';', 'in.csv', 'out.xlsx'], SpreadsheetConverter::toXlsxCommand('in.csv', 'out.xlsx', binary: $binary, delimiter: ';'));
        $this->assertSame([$binary, 'to-xlsx', '--delimiter', 'tab', 'in.csv', 'out.xlsx'], SpreadsheetConverter::toXlsxCommand('in.csv', 'out.xlsx', binary: $binary, delimiter: "\t"));
        $this->assertSame([$binary, 'to-csv', '--sheet', 'Data', 'in.xlsx', '-'], SpreadsheetConverter::toCsvCommand('in.xlsx', '-', 'Data', $binary));
    }

    public function test_csv_lines_encode_rows_the_way_the_converter_reads_them(): void
    {
        $stringable = new class implements \Stringable
        {
            public function __toString(): string
            {
                return 'stringable';
            }
        };

        $lines = SpreadsheetConverter::csvLines((function () use ($stringable) {
            yield ['id', 'name', 'active', 'note'];
            yield [7, 'Æble, "Ø" & Co', true, null];
            yield [0.5, "two\nlines", false, $stringable];
            yield ['back\\slash"', '', 0, '007'];
            yield [1234567.891234567, 0.1 + 0.2, 1.0, 1.5e-7];
        })());

        $this->assertSame([
            "id,name,active,note\n",
            "7,\"Æble, \"\"Ø\"\" & Co\",1,\n",
            "0.5,\"two\nlines\",0,stringable\n",
            "\"back\\slash\"\"\",,0,007\n",
            "1234567.891234567,0.30000000000000004,1.0,1.5E-7\n",
        ], iterator_to_array($lines, false));
    }

    public function test_csv_lines_quote_a_one_column_header(): void
    {
        $this->assertSame(
            ["\"id;ref|x\"\"y\"\n", "a;b\n"],
            iterator_to_array(SpreadsheetConverter::csvLines([['id;ref|x"y'], ['a;b']]), false),
        );
        $this->assertSame(["\"\"\n", "1\n"], iterator_to_array(SpreadsheetConverter::csvLines([[null], [1]]), false));
    }

    public function test_the_bundled_binary_accepts_a_one_column_export_with_a_semicolon_in_its_header(): void
    {
        $bundled = SpreadsheetConverter::bundledBinaryPath(SpreadsheetConverter::platform());
        if (! is_file($bundled)) {
            $this->markTestSkipped('No bundled binary for this platform; release builds add them.');
        }

        $xlsx = $this->directory.'/out.xlsx';
        [$exitCode] = $this->runConverter(
            SpreadsheetConverter::toXlsxCommand('-', $xlsx, strict: true, delimiter: ','),
            SpreadsheetConverter::csvLines([['id;ref'], ['a;b'], ['c']]),
        );
        $this->assertSame(0, $exitCode);

        [, $stdout] = $this->runConverter(SpreadsheetConverter::toCsvCommand($xlsx, '-'));
        $this->assertSame("id;ref\na;b\nc\n", $stdout);
    }

    public function test_the_bundled_binary_round_trips_through_xlsx(): void
    {
        $bundled = SpreadsheetConverter::bundledBinaryPath(SpreadsheetConverter::platform());
        if (! is_file($bundled)) {
            $this->markTestSkipped('No bundled binary for this platform; release builds add them.');
        }

        $schema = $this->directory.'/schema.json';
        file_put_contents($schema, '{"employees":"number","active":"bool"}');
        $xlsx = $this->directory.'/out.xlsx';
        $rows = [['cvr', 'name', 'employees', 'active'], ['12345678', 'Æble, "Ø" & Co', 12, true]];

        [$exitCode, $stdout] = $this->runConverter(
            SpreadsheetConverter::toXlsxCommand('-', $xlsx, $schema, strict: true, delimiter: ','),
            SpreadsheetConverter::csvLines($rows),
        );
        $this->assertSame(0, $exitCode);
        $this->assertSame(['rows' => 1, 'sheets' => 1, 'clipped_cells' => 0], json_decode($stdout, true));

        [$exitCode, $stdout] = $this->runConverter(SpreadsheetConverter::toCsvCommand($xlsx, '-'));
        $this->assertSame(0, $exitCode);
        $this->assertSame("cvr,name,employees,active\n12345678,\"Æble, \"\"Ø\"\" & Co\",12,1\n", $stdout);
    }

    /**
     * @param  list<string>  $command
     * @param  iterable<string>  $stdin
     * @return array{int, string}
     */
    private function runConverter(array $command, iterable $stdin = []): array
    {
        $process = proc_open($command, [0 => ['pipe', 'r'], 1 => ['pipe', 'w'], 2 => ['pipe', 'w']], $pipes);
        foreach ($stdin as $chunk) {
            fwrite($pipes[0], $chunk);
        }
        fclose($pipes[0]);
        $stdout = stream_get_contents($pipes[1]);
        $stderr = stream_get_contents($pipes[2]);
        fclose($pipes[1]);
        fclose($pipes[2]);
        $exitCode = proc_close($process);
        $this->assertSame('', $stderr);

        return [$exitCode, $stdout];
    }

    private function executable(string $name): string
    {
        $path = $this->directory.'/'.$name;
        file_put_contents($path, '');
        chmod($path, 0755);

        return $path;
    }
}
