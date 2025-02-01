namespace Autogrator.OutlookAutomation;

public interface IAllowedSenders : IEnumerable<string> {
    void Load(string filepath);
    bool IsAllowed(string emailAddress);
    string GetSenderFolder(string emailAddress);
}