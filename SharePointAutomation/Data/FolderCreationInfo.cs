namespace Autogrator.SharePointAutomation;

public readonly record struct FolderCreationInfo(
    string Name,
    string SiteId,
    string DriveId,
    string? Path
);