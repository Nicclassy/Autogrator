namespace Autogrator.SharePointAutomation;

public readonly record struct FileUploadInfo(
    string FileName,
    string LocalFileDir,
    string ParentId,
    string SiteId,
    string DriveId
);