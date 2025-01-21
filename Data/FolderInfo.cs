namespace Autogrator.Data;

public sealed record FolderInfo {
    public required string Name { get; init; }
    public string? Directory { get; init; }
    public required string DriveName { get; init; }
    public required string SitePath { get; init; }

    public FolderInfo() { }

    public FolderInfo(
        string name, 
        string driveName, 
        string sitePath, 
        string? directory = null
    ) {
        Name = name; 
        Directory = driveName; 
        SitePath = sitePath; 
        Directory = directory; 
    }
}