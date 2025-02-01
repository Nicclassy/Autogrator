namespace Autogrator.Data;

public sealed record FileUploadInfo {
    public required string FileName { get; init; }
    public required string LocalFileDirectory { get; init; }
    public required string UploadDirectory { get; init; }
    public required string DriveName { get; init; }
    public required string SitePath { get; init; }

    public FileUploadInfo() { }

    public FileUploadInfo(
        string filename, 
        string localFileDirectory,
        string uploadDirectory,
        string driveName,
        string sitePath
    ) {
        FileName = filename;
        LocalFileDirectory = localFileDirectory;
        UploadDirectory = uploadDirectory;
        DriveName = driveName;
        SitePath = sitePath;
    }
}