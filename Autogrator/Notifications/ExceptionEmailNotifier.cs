using System.Globalization;
using System.Text.RegularExpressions;

using Outlook = Microsoft.Office.Interop.Outlook;
using SmartFormat;
using Serilog;

using Autogrator.Utilities;
using Autogrator.OutlookAutomation;

namespace Autogrator.Notifications;

public partial class EmailExceptionNotifier {

    private const int ExceptionThrowerFrameIndex = 7;
    public required bool ReviewSentEmails { get; init; }
    public required string LogFileName { get; init; }
    public required string LoggingDirectory { get; init; }

    [GeneratedRegex(@"\d+")]
    private static partial Regex TimeStampPattern();

    public UnhandledExceptionEventHandler EventHandler() =>
        (_, e) => {
            DateTime now = DateTime.Now;
            Exception ex = (Exception) e.ExceptionObject;

            StackTraceInfo stackTraceInfo = StackTraceInfo.OfFrameIndex(ExceptionThrowerFrameIndex);
            ExceptionInfo exceptionInfo = ExceptionInfo.Create(ex, now);
            string emailContent = File.ReadAllText(NotificationEmail.ContentPath);

            Log.Fatal(
                "Application crashed at {TimeStamp} in {FileName} in {Method} on line {LineNumber}",
                now.ToString("t"), stackTraceInfo.FileName, stackTraceInfo.Method, stackTraceInfo.LineNumber
            );
            Log.CloseAndFlush();

            SendEmail(exceptionInfo, stackTraceInfo, emailContent);
        };

    private string LatestLogFilePath() {
        const string SerilogFileFormat = "yyyyMMdd";
        
        string loggingDirectory = Path.Combine(Directory.GetCurrentDirectory(), LoggingDirectory);
        string filename = Directory
            .EnumerateFiles(loggingDirectory)
            .Select(path => Path.GetFileName(path))
            .Where(filename => TimeStampPattern().IsMatch(filename))
            .Select(filename => {
                string match = TimeStampPattern().Match(filename).Value;
                DateTime timestamp = DateTime.ParseExact(match, SerilogFileFormat, CultureInfo.InvariantCulture);
                return (filename, timestamp);
            })
            .OrderByDescending(pair => pair.timestamp)
            .First()
            .filename;

        return Path.Combine(loggingDirectory, filename);
    }

    private void SendEmail(
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
        Outlook.Inspector _ = email.GetInspector;

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
            Log.Information("Displaying email content");
            email.Display();
        } else {
            Log.Information("Sending email");
            email.Send();
        }
    }
}