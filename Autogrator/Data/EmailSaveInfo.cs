namespace Autogrator.Data;

public sealed record EmailSaveInfo {
    public required string FileName { get; init; }
    public required string FileDirectory { get; init; }
    public required string FilePath { get; init; }

    public EmailSaveInfo() { }
}