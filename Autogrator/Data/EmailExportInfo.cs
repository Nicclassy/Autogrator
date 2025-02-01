namespace Autogrator.Data;

public sealed record EmailExportInfo {
    public required string FileName { get; init; }
    public required string FileDirectory { get; init; }
    public required string SenderEmailAddress { get; init; }
}