using System.Globalization;
using System.Text.RegularExpressions;

using Outlook = Microsoft.Office.Interop.Outlook;

using Autogrator.Utilities;
using Autogrator.OutlookAutomation;

namespace Autogrator.Notifications;

public static partial class EmailExceptionNotifier {
    public const string LogFileName = "log.txt";

    [GeneratedRegex(@"\d+")]
    private static partial Regex TimeStampPattern();

    public static string LatestLogFilePath(string directory = ".") {
        string latestFileName = Directory
            .EnumerateFiles(directory)
            .Select(path => Path.GetFileName(path))
            .Where(filename => TimeStampPattern().IsMatch(filename))
            .Select(filename => {
                string match = TimeStampPattern().Match(filename).Value;
                DateTime timestamp = DateTime.ParseExact(match, "yyyyMMdd", CultureInfo.InvariantCulture);
                return (filename, timestamp);
            })
            .OrderByDescending(pair => pair.timestamp)
            .First()
            .filename;
        return Path.Combine(directory, latestFileName);
    }

    public static void SendEmail(ExceptionInfo? exceptionInfo, StackTraceInfo? stackTraceInfo, string body) {
        Outlook.NameSpace ns = OutlookInstance.NameSpace;
        Outlook.Application application = OutlookInstance.Application;

        Outlook.Account account = application
            .Session
            .Accounts
            .OfType<Outlook.Account>()
            .First(account => account.SmtpAddress == NotificationEmail.SenderEmailAddress);
        Outlook.MailItem email = 
            (Outlook.MailItem) application.CreateItem(Outlook.OlItemType.olMailItem);
        Outlook.Inspector inspector = email.GetInspector;
        string currentBody = email.HTMLBody;

        email.Subject = $"Autogrator crashed at {exceptionInfo?.TimeStamp("hh:mm tt") ?? ""}";
        email.SendUsingAccount = account;
        email.To = NotificationEmail.RecipientEmailAddress;
        email.Importance = Outlook.OlImportance.olImportanceHigh;
        email.HTMLBody = "" + currentBody;
        email.Display();
        // Console.WriteLine(email.HTMLBody);
    }
}