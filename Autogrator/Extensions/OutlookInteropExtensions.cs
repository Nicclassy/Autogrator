using Outlook = Microsoft.Office.Interop.Outlook;

namespace Autogrator.Extensions;

public static class OutlookInteropExtensions {
    public static IEnumerable<Outlook.MailItem> Emails(this Outlook.MAPIFolder folder) =>
        folder.Items.OfType<Outlook.MailItem>();

    public static IEnumerable<Outlook.MailItem> EmailsByLatest(this Outlook.MAPIFolder folder) =>
        folder.Items.OfType<Outlook.MailItem>().OrderBy(item => item.ReceivedTime).Reverse();

    public static Outlook.MailItem? LatestEmail(this Outlook.MAPIFolder folder, Func<Outlook.MailItem, bool>? predicate = null) =>
        folder.EmailsByLatest().FirstOrDefault(predicate ?? (_ => true));

    public static Outlook.MAPIFolder? GetSubFolder(this Outlook.Folders folders, string name) =>
        folders.OfType<Outlook.MAPIFolder>().FirstOrDefault(folder => folder.Name == name);
}