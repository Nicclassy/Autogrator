namespace Autogrator.SharePointAutomation;

public sealed class DriveItemInfo(string _name, string _id) {
    public string Name { get; set; } = _name;
    public string Id { get; set; } = _id;
}