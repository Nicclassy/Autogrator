using System.Collections;
using System.Text.RegularExpressions;

namespace Autogrator.OutlookAutomation;

public sealed class AllEmailSendersAllowed : IAllowedSenders {
    private static readonly Regex EmailRegex =
        new(string.Join("",
            @"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|",
            @"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|",
            @"\\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@",
            @"(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[",
            @"(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}",
            @"(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:",
            @"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)])"
        ));

    public IEnumerator<string> GetEnumerator() => Enumerable.Empty<string>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Load(string filename) { }

    public bool IsAllowed(string emailAddress) => IsValidEmailAddress(emailAddress);

    public string GetSenderFolder(string emailAddress) => emailAddress;

    private bool IsValidEmailAddress(string emailAddress) => EmailRegex.IsMatch(emailAddress);
}