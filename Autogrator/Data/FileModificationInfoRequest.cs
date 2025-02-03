namespace Autogrator.Data;

public sealed class FileModificationInfoRequest {
    public required string FileName { get; init; }
    public required string DriveName { get; init; }
    public required string SitePath { get; init; }
    public string? FileDirectory { get; init; }
}