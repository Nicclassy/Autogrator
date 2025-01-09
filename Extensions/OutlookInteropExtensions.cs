using System;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Autogrator.Extensions;

public static class OutlookInteropExtensions {
    public static IEnumerable<Outlook.MailItem> Emails(this Outlook.MAPIFolder folder) =>
        folder.Items.OfType<Outlook.MailItem>();

    public static IEnumerable<Outlook.MailItem> EmailsByLatest(this Outlook.MAPIFolder folder) =>
        folder.Items.OfType<Outlook.MailItem>().OrderBy(item => item.ReceivedTime).Reverse();

    public static Outlook.MailItem? LatestEmail(this Outlook.MAPIFolder folder) =>
        folder.EmailsByLatest().FirstOrDefault();
}