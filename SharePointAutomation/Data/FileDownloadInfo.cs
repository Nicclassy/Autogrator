namespace Autogrator.SharePointAutomation;

public readonly record struct FileDownloadInfo(
    string FileName,
    string DestinationFolder,
    string DestinationPath,
    string DriveId,
    string ItemId
);