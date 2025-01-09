using System.Globalization;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Autogrator.OutlookAutomation;

public readonly struct MailItemProcessor(Outlook.MailItem mailItem) {

    public string GetFormattedSenderName() => FormatName(mailItem.SenderName);

    private static string FormatName(string name) => 
        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(name);
}
