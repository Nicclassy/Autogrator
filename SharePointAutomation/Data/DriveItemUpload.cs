namespace Autogrator.SharePointAutomation;

public sealed class DriveItemUpload(string _name) {
    public string? Name { get; set; } = _name;
    public object? Folder { get; set; } = new { };
}