namespace Autogrator.SharePointAutomation;

public readonly record struct FolderUploadInfo(
    string Name,
    string SiteId,
    string DriveId,
    string? Path
);