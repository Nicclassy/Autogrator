using System.Globalization;
using System.Text.RegularExpressions;

using Outlook = Microsoft.Office.Interop.Outlook;
using SmartFormat;
using Serilog;

using Autogrator.Utilities;
using Autogrator.OutlookAutomation;

namespace Autogrator.Notifications;

public static partial class EmailExceptionNotifier {
    public const string LogFileName = "log.txt";

    private const int ExceptionThrowerFrameIndex = 4;
    private readonly static bool ReviewSentEmails = true;

    [GeneratedRegex(@"\d+")]
    private static partial Regex TimeStampPattern();

    public static UnhandledExceptionEventHandler EventHandler() =>
        (_, e) => {
            DateTime now = DateTime.Now;
            StackTraceInfo stackTraceInfo = StackTraceInfo.OfFrameIndex(ExceptionThrowerFrameIndex);
            ExceptionInfo exceptionInfo = ExceptionInfo.Create((Exception) e.ExceptionObject, now);
            string emailContent = File.ReadAllText(NotificationEmail.ContentPath);
            Log.Fatal(
                "Application crashed at {TimeStamp} in {FileName} in {Method} on line {LineNumber}. Sending an email...",
                now.ToString("t"), stackTraceInfo.FileName, stackTraceInfo.Method, stackTraceInfo.LineNumber
            );
            SendEmail(exceptionInfo, stackTraceInfo, emailContent);
        };

    internal static string LatestLogFilePath(string directory = ".") {
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

    internal static void SendEmail(
        ExceptionInfo exceptionInfo, 
        StackTraceInfo stackTraceInfo, 
        string emailContent
    ) {
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

        HTMLBodyEditor bodyEditor = new(email.HTMLBody);
        var formatArgs = new {
            stackTraceInfo.LineNumber,
            stackTraceInfo.FileName,
            stackTraceInfo.Method,
            exceptionInfo.ExceptionType,
            TimeStamp = exceptionInfo.TimeStamp(),
        };
        string formattedText = Smart.Format(emailContent, formatArgs);
        bodyEditor.PrependText(formattedText);

        email.Subject = $"Autogrator crashed at {exceptionInfo.TimeStamp("t")}";
        email.SendUsingAccount = account;
        email.To = NotificationEmail.RecipientEmailAddress;
        email.Importance = Outlook.OlImportance.olImportanceHigh;
        email.HTMLBody = bodyEditor.Content();
        email.Attachments.Add(
            LatestLogFilePath(), Outlook.OlAttachmentType.olByValue, 
            Type.Missing, Type.Missing
        );
        
        if (ReviewSentEmails) {
            email.Display();
            Console.WriteLine(email.HTMLBody);
        } else {
            email.Send();
        }
    }
}