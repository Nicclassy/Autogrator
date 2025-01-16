using Newtonsoft.Json.Linq;

namespace Autogrator.SharePointAutomation;

public struct DriveItemInfo {
    public string Name { get; set; }
    public string Id { get; set; }

    public static DriveItemInfo Parse(JToken token) =>
        token.ToObject<DriveItemInfo>();
}