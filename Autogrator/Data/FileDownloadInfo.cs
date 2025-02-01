namespace Autogrator.Data;

public sealed record FileDownloadInfo {
    public required string FileName { get; init; }
    public required string? DestinationFileName { get; init; }
    public required string DestinationFolder { get; init; }
    public required string DriveName { get; init; }
    public required string SitePath { get; init; }
    public string? DownloadPath { get; init; }

    public FileDownloadInfo() { }

    public FileDownloadInfo(
        string filename, 
        string destinationFolder, 
        string driveName,
        string sitePath,
        string downloadFileName,
        string? downloadPath = null
    ) {
        FileName = filename;
        DestinationFolder = destinationFolder;
        DriveName = driveName;
        SitePath = sitePath;
        DestinationFileName = downloadFileName;
        DownloadPath = downloadPath;
    }
}