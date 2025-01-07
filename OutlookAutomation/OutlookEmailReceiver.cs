using System.Collections.Concurrent;
using Outlook = Microsoft.Office.Interop.Outlook;

using Serilog;

using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public sealed class OutlookEmailReceiver(Outlook.Application application, Outlook.NameSpace ns, Outlook.MAPIFolder inbox) : IDisposable {
    private readonly ConcurrentQueue<Outlook.MailItem> emailsTodo = new();
    public Outlook.MAPIFolder Inbox => inbox;

    public static OutlookEmailReceiver Create() {
        Outlook.Application application = new();
        Outlook.NameSpace ns = application.GetNamespace("MAPI");

        bool retry = false;
        Log.Information($"Logging in with email {Credentials.Outlook.Email}");
        try {
            LoginWithOptions(ns, showDialog: false, newSession: true);
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
                LoginWithOptions(ns, showDialog: true, newSession: true);
            } catch (System.Runtime.InteropServices.COMException ex) {
                Log.Fatal($"Login failed: {ex.Message}");
                Environment.Exit(ex.ErrorCode);
            }
        }
        Log.Information("Successfully logged in!");

        Outlook.MAPIFolder inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        return new(application, ns, inbox);
    }

    public void Listen() {
        application.NewMailEx += delegate(string entryID) {
            Outlook.MailItem email = ns.GetItemFromID(entryID);
            emailsTodo.Enqueue(email);
            Log.Information($"Received email with subject {email.Subject}");
        };
        Log.Information("Started listening for emails");
    }

    public bool TryReceiveEmail(out Outlook.MailItem email) => emailsTodo.TryDequeue(out email!);

    private static void LoginWithOptions(Outlook.NameSpace ns, bool showDialog, bool newSession) =>
        ns.Logon(Credentials.Outlook.Email, Credentials.Outlook.Password, ShowDialog: showDialog, NewSession: newSession);

    public void Dispose() => ns.Logoff();
}
