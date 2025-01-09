using System.Collections.Concurrent;

using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace Autogrator.OutlookAutomation;

public sealed class OutlookEmailReceiver(OutlookAuthenticator authenticator) : IDisposable {
    private readonly bool authenticationComplete =
        authenticator.AutheticationComplete ? true : throw new InvalidOperationException("Authentication must be complete first.");
    private readonly Outlook.Application application = authenticator.Application;
    private readonly Outlook.NameSpace ns = authenticator.NameSpace;
    private readonly ConcurrentQueue<Outlook.MailItem> emailsTodo = new();

    public Outlook.MAPIFolder Inbox => ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

    internal static OutlookEmailReceiver Create() {
        OutlookAuthenticator authenticator = new();
        authenticator.Login();
        return new(authenticator);
    }

    public void Listen() {
        if (!authenticationComplete) {
            Log.Fatal("Authentication must be complete prior to listening for emails.");
            Environment.Exit(1);
        }

        application.NewMailEx += delegate(string entryID) {
            Outlook.MailItem email = ns.GetItemFromID(entryID);
            emailsTodo.Enqueue(email);
            Log.Information($"Received email with subject {email.Subject}");
        };
        Log.Information("Started listening for emails");
    }

    public bool TryReceiveEmail(out Outlook.MailItem email) => emailsTodo.TryDequeue(out email!);

    public void Dispose() => ns.Logoff();
}
