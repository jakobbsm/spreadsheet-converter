namespace SpreadsheetConverter;

/// <summary>
/// Writes to a temporary file next to the target and only moves it into place on <see cref="CommitAsync"/>,
/// so a failed or killed conversion never leaves a half-written output file.
/// </summary>
internal sealed class AtomicFile : IAsyncDisposable
{
    private readonly string _targetPath;
    private readonly string _temporaryPath;
    private bool _committed;

    public AtomicFile(string targetPath)
    {
        _targetPath = Path.GetFullPath(targetPath);
        _temporaryPath = Path.Combine(
            Path.GetDirectoryName(_targetPath)!,
            $".partial-{Guid.NewGuid():N}-{Path.GetFileName(_targetPath)}");
        Stream = new FileStream(_temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1 << 16, useAsync: false);
    }

    public FileStream Stream { get; }

    public async Task CommitAsync()
    {
        await Stream.FlushAsync();
        await Stream.DisposeAsync();
        File.Move(_temporaryPath, _targetPath, overwrite: true);
        _committed = true;
    }

    public async ValueTask DisposeAsync()
    {
        if (_committed)
        {
            return;
        }
        await Stream.DisposeAsync();
        File.Delete(_temporaryPath);
    }
}
