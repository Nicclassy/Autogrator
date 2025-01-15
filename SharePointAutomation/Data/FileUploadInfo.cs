namespace Autogrator.SharePointAutomation;

public readonly record struct FileUploadInfo(
    string FileName,
    string FilePath,
    string ParentName,
    string DriveName
);