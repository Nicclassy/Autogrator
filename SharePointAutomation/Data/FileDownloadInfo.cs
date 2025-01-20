namespace Autogrator.SharePointAutomation;

public sealed record FileDownloadInfo(
    string FileName,
    string DestinationFolder,
    string? DownloadPath,
    string DriveName,
    string SitePath
);