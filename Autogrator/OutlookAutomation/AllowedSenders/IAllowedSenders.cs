namespace Autogrator.OutlookAutomation;

public interface IAllowedSenders : IEnumerable<string> {
    void Load(string filepath);
    void Reload();
    bool IsAllowed(string emailAddress);
    string GetSenderFolder(string emailAddress);
}