using System.Collections.Concurrent;

using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace Autogrator.OutlookAutomation;

public sealed class EmailReceiver {
    private const bool LogRejectedSenders = true;
    
    private readonly ConcurrentQueue<Outlook.MailItem> emailsTodo = new();

    public void Listen(IAllowedSenders? allowedSenders = null) {
        bool SenderIsAllowed(Outlook.MailItem email) {
            if (allowedSenders is null) return true;
            return allowedSenders.IsAllowed(email.SenderEmailAddress);
        }

        OutlookInstance.Application.NewMailEx += delegate (string entryID) {
            Outlook.MailItem email = OutlookInstance.NameSpace.GetItemFromID(entryID);
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
}
