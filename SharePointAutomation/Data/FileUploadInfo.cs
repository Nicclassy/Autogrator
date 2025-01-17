namespace Autogrator.SharePointAutomation;

public readonly record struct FileUploadInfo(
    string FileName,
    string LocalFileDirectory,
    string LocalFilePath,
    string ParentId,
    string SiteId,
    string DriveId
);