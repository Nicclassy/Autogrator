using Outlook = Microsoft.Office.Interop.Outlook;

using Serilog;

using Autogrator.OutlookAutomation;
using Autogrator.Extensions;

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
        List<Outlook.MailItem> emails = receiver.Inbox.MailItems().ToList();
        
        var email = emails[0];
        OutlookEmailExporter.SaveAndExportEmail(email);
    }

    public static void Main() {}
}