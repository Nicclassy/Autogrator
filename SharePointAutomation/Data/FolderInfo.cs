namespace Autogrator.SharePointAutomation;

public sealed record FolderInfo(
    string Name,
    string? Directory,
    string DriveName,
    string SitePath
);