using System;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Autogrator.Extensions;

public static class OutlookInteropExtensions {
    public static IEnumerable<Outlook.MailItem> MailItems(this Outlook.MAPIFolder folder) =>
        folder.Items.OfType<Outlook.MailItem>();

    public static IEnumerable<Outlook.MailItem> EmailsByLatest(this Outlook.MAPIFolder source) =>
        source.Items.OfType<Outlook.MailItem>().OrderBy(item => item.ReceivedTime).Reverse();
}