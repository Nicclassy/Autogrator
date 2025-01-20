namespace Autogrator.SharePointAutomation;

public sealed record FileUploadInfo(
    string FileName,
    string LocalFileDirectory,
    string UploadDirectory,
    string DriveName,
    string SitePath
);