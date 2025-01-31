namespace Autogrator.OutlookAutomation;

public interface IAllowedSenderList: IEnumerable<string> {
    void Load(string filepath);
    bool IsAllowed(string emailAddress);
    string GetSenderFolder(string emailAddress);
}