using System.Collections.Concurrent;

using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;
using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public delegate void EmailReceivedHandler();

public sealed class EmailReceiver {
    private const char AllowedRecipientsDelimiter = ';';

    public required bool LogRejectedEmails { get; init; }
    
    private readonly ConcurrentQueue<Outlook.MailItem> emailsTodo = new();
    private readonly HashSet<string> allowedRecipients = [.. 
        Credentials.Outlook.AllowedRecipients.Split(AllowedRecipientsDelimiter)
    ];

    public event EmailReceivedHandler? OnEmailReceived;

    public void Listen(IAllowedSenders? allowedSenders = null) {
        bool SenderIsAllowed(Outlook.MailItem email) =>
            allowedSenders is null || allowedSenders.IsAllowed(email.SenderEmailAddress);

        // If any recipient is allowed, the email is allowed
        bool RecipientsAreAllowed(Outlook.MailItem email) =>
            allowedRecipients.Count == 0 || email.Recipients
                .OfType<Outlook.Recipient>()
                .Any(recipient => allowedRecipients.Contains(recipient.Name));

        OutlookInstance.Application.NewMailEx += delegate (string entryID) {
            OnEmailReceived?.Invoke();
            Outlook.MailItem email = OutlookInstance.NameSpace.GetItemFromID(entryID);
            if (!RecipientsAreAllowed(email)) {
                if (LogRejectedEmails) {
                    IEnumerable<string> recipientNames = email.Recipients
                        .OfType<Outlook.Recipient>()
                        .Select(recipient => recipient.Name);
                    Log.Information(
                        "An email from {Sender} was rejected because none of the recipients ({Recipients}) are allowed",
                        email.Sender, string.Join(", ", recipientNames)
                    );
                }
                return;
            }

            if (SenderIsAllowed(email)) {
                emailsTodo.Enqueue(email);
                Log.Information(
                    "Received email from {Sender} with subject '{EmailSubject}'", 
                    email.SenderName, email.Subject
                );
            } else if (LogRejectedEmails) {
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
