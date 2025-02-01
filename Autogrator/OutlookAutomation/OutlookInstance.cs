using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

using Autogrator.Utilities;
using Autogrator.Extensions;

namespace Autogrator.OutlookAutomation;

public static partial class OutlookInstance {
    private const bool AutomateProfileCreation = true;

    public static Outlook.Application Application { get; } = new();
    public static Outlook.NameSpace NameSpace { get; } = Application.GetNamespace("MAPI");
    public static bool IsAuthenticated { get; private set; } = false;

    public static Outlook.MAPIFolder Inbox =>
        NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

    static OutlookInstance() => Login();

    public static Outlook.MAPIFolder? FolderWithName(string name, string? subfolderName = null) {
        Outlook.MAPIFolder? folder = NameSpace.Folders.GetSubFolder(name);
        return subfolderName is not null ? folder?.Folders.GetSubFolder(subfolderName) : folder;
    }

    public static void Login() {
        if (IsAuthenticated)
            return;

        NameSpace.Logon(
            Credentials.Outlook.Profile, Credentials.Outlook.Password,
            ShowDialog: false, NewSession: true
        );
        Log.Information("Succesfully logged in to profile {Profile}", Credentials.Outlook.Profile);
    }
}