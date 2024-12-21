using DotNetEnv;
using Microsoft.Extensions.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailAutomation;

public sealed class EmailAutomationTest
{
    private static readonly string EMAIL_ADDRESS;
    private static readonly string PASSWORD;
    private static readonly ILogger<EmailAutomationTest> logger;

    static EmailAutomationTest()
    {
        using var factory = LoggerFactory.Create(builder => builder.AddConsole());
        logger = factory.CreateLogger<EmailAutomationTest>();

        Env.Load(".env");
        EMAIL_ADDRESS = Environment.GetEnvironmentVariable("AG_EMAIL")!;
        PASSWORD = Environment.GetEnvironmentVariable("AG_PASSWORD")!;
    }

    public void ReadOutlookMail()
    {
        logger.LogInformation("Reading emails...");

        Outlook.Application outlookApp = new Outlook.Application();
        Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
        logger.LogInformation($"Logging in with email ${EMAIL_ADDRESS}...");
        outlookNamespace.Logon(EMAIL_ADDRESS, PASSWORD, ShowDialog: false, NewSession: true);
        logger.LogInformation("Successfully logged in!");

        Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        Outlook.Items mailItems = inbox.Items;
        foreach (Outlook.MailItem mailItem in mailItems.OfType<Outlook.MailItem>())
            Console.WriteLine("Email subject: " + mailItem.Subject);

        outlookNamespace.Logoff();
        logger.LogInformation("Logged off.");
    }
}

public sealed class Program
{
    public static void Main()
    {
        var emailAutomationTest = new EmailAutomationTest();
        emailAutomationTest.ReadOutlookMail();
    }
}
