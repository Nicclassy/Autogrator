namespace Autogrator.Data;

public sealed record FileDownloadInfo {
    public required string FileName { get; init; }
    public required string? DestinationFileName { get; init; }
    public required string DestinationFolder { get; init; }
    public required string DriveName { get; init; }
    public required string SitePath { get; init; }
    public string? DownloadPath { get; init; }
    public bool AlwaysDownload { get; init; } = false;
}