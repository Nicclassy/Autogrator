using System.Text;
using DotNetEnv;
using Microsoft.Extensions.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailAutomation;

public static class EnumerableExtensions
{
    public static string ToListString<T>(this IEnumerable<T?> source)
    {
        string FormatValue(T? value) =>
            value is T t ? t.ToString()! : "(null)";

        var builder = new StringBuilder();
        builder.Append('[');
        builder.Append(FormatValue(source.FirstOrDefault()));
        foreach (T? t in source.Skip(1))
        {
            builder.Append($", {FormatValue(t)}");
        }
        builder.AppendLine("]");
        return builder.ToString();
    }

    public static void Print<T>(this IEnumerable<T> source) where T : notnull =>
        Console.WriteLine(source.ToListString());
}

sealed class EmailAutomationTest
{
    private static readonly string EMAIL_ADDRESS;
    private static readonly string PASSWORD;
    private static readonly ILogger<EmailAutomationTest> logger;

    static EmailAutomationTest()
    {
        using var factory = LoggerFactory.Create(builder => builder.AddConsole());
        logger = factory.CreateLogger<EmailAutomationTest>();

        logger.LogInformation("Before loading env");
        Env.Load(".env");
        logger.LogInformation("After loading env");
        EMAIL_ADDRESS = Environment.GetEnvironmentVariable("AG_EMAIL")!;
        logger.LogInformation("After obtained email address");
        PASSWORD = Environment.GetEnvironmentVariable("AG_PASSWORD")!;
        logger.LogInformation("Static constructor finished");
    }

    public void ReadOutlookMail()
    {
        logger.LogInformation("Read mail starting");

        Outlook.Application oApp = new Outlook.Application();
        Outlook.NameSpace outlookNamespace = oApp.GetNamespace("mapi");
        logger.LogInformation($"Logging in with email ${EMAIL_ADDRESS}...");
        outlookNamespace.Logon(EMAIL_ADDRESS, PASSWORD, ShowDialog: true, NewSession: true);
        logger.LogInformation("Successfully logged in!");

        Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        Outlook.Items mailItems = inbox.Items;
        Console.WriteLine(mailItems.ToString());
        foreach (Outlook.MailItem mailItem in mailItems.OfType<Outlook.MailItem>())
        {
            Console.WriteLine("Subject: " + mailItem.Subject);
        }

        outlookNamespace.Logoff();
    }
}

public class Program
{
    public static void Main()
    {
        var emailAutomationTest = new EmailAutomationTest();
        emailAutomationTest.ReadOutlookMail();
    }
}
