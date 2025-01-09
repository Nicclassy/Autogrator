using Outlook = Microsoft.Office.Interop.Outlook;

using Serilog;

using Autogrator.OutlookAutomation;
using Autogrator.Extensions;
using Autogrator.Helpers;

namespace Autogrator;

public static class Program {
    static Program() {
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .CreateLogger();
        Log.Information("Begin logging");
    }

    public static void EmailListener() {
        var receiver = OutlookEmailReceiver.Create();
        Log.Information("Waiting for emails...");
        while (true) {
            if (receiver.TryReceiveEmail(out var email)) {
                Log.Information($"Received email: {email.Subject}");
            }
        }
    }

    public static void ExportFirstEmail() {
        using OutlookEmailReceiver receiver = OutlookEmailReceiver.Create();
        List<Outlook.MailItem> emails = receiver.Inbox.Emails().ToList();

        var email = emails[0];
        OutlookEmailExporter.SaveAndExportEmail(email);
    }

    public static void PrintSenderAndRecipients() {
        using OutlookEmailReceiver receiver = OutlookEmailReceiver.Create();
        Outlook.MailItem email = receiver.Inbox.EmailsByLatest().ElementAt(4);
        Console.WriteLine(email.SenderEmailAddress);
        Console.WriteLine(email.SenderName);
        Console.WriteLine(email.ReplyRecipientNames);
        Console.WriteLine();

        email.Recipients.OfType<Outlook.Recipient>().Print(formatter: recipient => recipient.Address);
        Console.Write("Sender: ");
        Console.WriteLine(email.Sender.Name);

        Console.WriteLine(EmailHelper.GetEmailAddressDomain("person@domain.com"));
    }

    public static void PrintEmailAddresses() {
        using OutlookEmailReceiver receiver = OutlookEmailReceiver.Create();
        foreach (Outlook.MailItem mailItem in receiver.Inbox.EmailsByLatest()) {
            if (EmailHelper.IsEmailAddress(mailItem.SenderEmailAddress))
                Console.WriteLine(mailItem.SenderEmailAddress);
        }
    }

    public static void PrintRootSiteID() {
        using AutogratorApplicationClient client = new();
        var result = client.GetRootSite().Result;
        Console.WriteLine(result!.Id);
    }

    public static void Main() {}
}