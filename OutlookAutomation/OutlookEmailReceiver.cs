using System.Collections.Concurrent;

using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace Autogrator.OutlookAutomation;

public sealed class OutlookEmailReceiver(OutlookAuthenticator authenticator, IAllowedSenderList? allowedSenders) : IDisposable {
    private const bool LogRejectedSenders = true;
    
    private readonly Outlook.Application application = authenticator.Application;
    private readonly Outlook.NameSpace ns = authenticator.NameSpace;
    private readonly ConcurrentQueue<Outlook.MailItem> emailsTodo = new();

    public Outlook.MAPIFolder Inbox => ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

    internal static OutlookEmailReceiver Create(IAllowedSenderList? allowedSenders = null) {
        OutlookAuthenticator authenticator = new();
        authenticator.Login();
        return new(authenticator, allowedSenders);
    }

    public void Listen() {
        bool SenderIsAllowed(Outlook.MailItem email) {
            if (allowedSenders is null) return true;
            return allowedSenders.IsAllowed(email.SenderEmailAddress);
        } 

        if (!authenticator.AutheticationComplete)
            Log.Error(new InvalidOperationException(), "Authentication must be completed before listening for emails.");

        application.NewMailEx += delegate(string entryID) {
            Outlook.MailItem email = ns.GetItemFromID(entryID);
            if (SenderIsAllowed(email)) {
                emailsTodo.Enqueue(email);
                Log.Information("Received email with subject {EmailSubject}", email.Subject);
            } else if (LogRejectedSenders) {
                Log.Information(
                    "An email from {Sender} was ignored because the sender is not allowed", 
                    email.SenderEmailAddress
                );
            }
        };

        Log.Information("Started listening for emails");
    }

    public bool TryReceiveEmail(out Outlook.MailItem email) => emailsTodo.TryDequeue(out email!);

    public void Dispose() => ns.Logoff();
}
