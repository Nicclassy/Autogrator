using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public static class OutlookInstance {
    public static Outlook.Application Application { get; } = new();
    public static Outlook.NameSpace NameSpace { get; } = Application.GetNamespace("MAPI");
    public static bool IsAuthenticated { get; private set; } = false;

    public static Outlook.MAPIFolder Inbox =>
        NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

    private static string Email => Credentials.Outlook.Email;
    private static string Password => Credentials.Outlook.Password;

    public static void Login() {
        if (IsAuthenticated)
            return;

        bool retry = false;
        Log.Information("Logging in with email {Email}", Email);
        try {
            LoginWithOptions(showDialog: false, newSession: true);
        } catch (System.Runtime.InteropServices.COMException ex) {
            Log.Error(ex, "Initial login attempt failed. Retrying with dialog...");
            retry = true;
        }

        if (retry) {
            try {
                // Try again, but this time show dialog.
                // The error previously thrown may be a consequence
                // of no profile existing. Hence, showing the dialog box
                // enables the user to create a profile and thus avoid the error
                LoginWithOptions(showDialog: true, newSession: true);
            } catch (System.Runtime.InteropServices.COMException ex) {
                Log.Error(ex, "Login failed");
                throw;
            }
        }

        IsAuthenticated = true;
        Log.Information("Successfully logged in!");
    }

    private static void LoginWithOptions(bool showDialog, bool newSession) =>
        NameSpace.Logon(Email, Password, ShowDialog: showDialog, NewSession: newSession);
}