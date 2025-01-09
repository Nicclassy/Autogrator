using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public sealed class OutlookAuthenticator {
    public Outlook.Application Application { get; }
    public Outlook.NameSpace NameSpace { get; }
    public bool UseStorageEmail { get; init; } = true;
    public bool AutheticationComplete { get; private set; } = false;

    private string Email => 
        UseStorageEmail ? Credentials.Outlook.Storage.Email : Credentials.Outlook.Email;

    private string Password =>
        UseStorageEmail ? Credentials.Outlook.Storage.Password : Credentials.Outlook.Password;

    public OutlookAuthenticator() {
        Application = new();
        NameSpace = Application.GetNamespace("MAPI");
    }

    public void Login() {
        bool retry = false;
        Log.Information($"Logging in with email {Email}");
        try {
            LoginWithOptions(showDialog: false, newSession: true);
        } catch (System.Runtime.InteropServices.COMException) {
            Log.Warning("Initial login attempt failed. Retrying with dialog...");
            retry = true;
        }

        if (retry) {
            try {
                // TODO: Automate profile creation
                // Try again, but this time show dialog.
                // The error previously thrown may be a consequence
                // of no profile existing. Hence, showing the dialog box
                // enables the user to create a profile and thus avoid the error
                LoginWithOptions(showDialog: true, newSession: true);
            } catch (System.Runtime.InteropServices.COMException ex) {
                Log.Fatal($"Login failed: {ex.Message}");
                Environment.Exit(ex.ErrorCode);
            }
        }

        AutheticationComplete = true;
        Log.Information("Successfully logged in!");
    }

    private void LoginWithOptions(bool showDialog, bool newSession) =>
        NameSpace.Logon(Email, Password, ShowDialog: showDialog, NewSession: newSession);
}