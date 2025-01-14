namespace Autogrator.SharePointAutomation;

public readonly record struct FileUploadInfo(
    string SiteName, 
    string DriveName, 
    string LocalFilePath,
    string UploadFilePath
);