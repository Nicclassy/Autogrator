namespace Autogrator.OutlookAutomation;

public interface IAllowedSenderList {
    void Load(string filepath);

    bool IsAllowed(string emailAddress);

    string GetSenderFolder(string emailAddress);
}