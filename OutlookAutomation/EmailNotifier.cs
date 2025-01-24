using System.Diagnostics;
using System.Globalization;

using Outlook = Microsoft.Office.Interop.Outlook;

using Autogrator.Utilities;
using System.Text.RegularExpressions;

namespace Autogrator.OutlookAutomation;

public sealed record StackTraceInfo(string Method, string FileName, int LineNumber) {
    public static StackTraceInfo OfIndex(int index) {
        StackTrace stackTrace = new(fNeedFileInfo: true);
        StackFrame frame = stackTrace.GetFrame(index)!;

        string method = frame.GetMethod()!.Name;
        string filename = Path.GetFileName(frame.GetFileName())!;
        int lineNumber = frame.GetFileLineNumber();
        return new(method, filename, lineNumber);
    }
}

public sealed record ExceptionInfo(string Name, DateTime DateTime) {
    private const string DefaultTimeStampFormat = "d T";

    public static ExceptionInfo Create(Exception ex, DateTime dateTime) => new(ex.GetType().Name, dateTime);

    public string TimeStamp(string format = DefaultTimeStampFormat) => 
        DateTime.ToString(format, CultureInfo.CurrentCulture);
}

public static partial class EmailNotifier {
    public const string LogFileName = "log.txt";

    [GeneratedRegex(@"\d+")]
    private static partial Regex TimeStampPattern();

    public static string LatestLogFilePath(string directory = ".") {
        string latestFileName = Directory
            .EnumerateFiles(directory)
            .Select(path => Path.GetFileName(path))
            .Where(filename => TimeStampPattern().IsMatch(filename))
            .Select(filename => {
                string timestamp = TimeStampPattern().Match(filename).Value;
                return (filename, timestamp: DateTime.ParseExact(timestamp, "yyyyMMdd", CultureInfo.InvariantCulture));
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