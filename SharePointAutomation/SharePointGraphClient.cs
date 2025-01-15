using System.Text.Json;

using Newtonsoft.Json.Linq;
using Serilog;
using System.Diagnostics;

using Autogrator.Extensions;
using Autogrator.Exceptions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class SharePointGraphClient(GraphHttpClient httpClient) {
    private static readonly JsonSerializerOptions SerializerOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    internal GraphHttpClient HttpClient { get; } = httpClient;

    internal async Task<List<string>> GetDriveItemNames(string? driveName = null, string? sitePath = null) {
        string driveId = await GetDriveId(driveName, sitePath);
        string content = await HttpClient.GetAsync($"/drives/{driveId}/root/children", default);

        List<string> items = [];
        JToken root = JToken.Parse(content);
        root.Walk(property => {
            if (property.Name == "name")
                items.Add(property.Value.ToString());
        });
        return items;
    } 

    internal async Task<IEnumerable<DriveItemInfo>> GetDriveItems(string? driveName = null, string? sitePath = null) {
        string driveId = await GetDriveId(driveName, sitePath);
        IEnumerable<JToken> items = await HttpClient.GetValuesByKeysAsync($"/drives/{driveId}/root/children", default, "name", "id");
        return items.Select(DriveItemInfo.Parse);
    }

    internal async Task<string> GetItemId(string itemName, Func<string, string> stringFormatter, string? driveName = null, string? sitePath = null) {
        IEnumerable<DriveItemInfo> driveItems = await GetDriveItems(driveName, sitePath);
        DriveItemInfo? result = driveItems.FirstOrDefault(item => stringFormatter(item.Name) == itemName);
        if (result is null) {
            Log.Fatal($"Could not find ID for item with name {itemName}");
            Environment.Exit(1);
        }

        return result.Value.Id;
    }

    internal async Task<string> GetSiteId(string? sitePath = null) =>
        await HttpClient.GetKeyAsync($"/sites/{SharePointSite.Hostname}:{sitePath ?? GraphAPI.DefaultSitePath}", "id", default);
    
    internal async Task<string> GetDriveId(string? driveName = null, string? sitePath = null) {
        string siteId = await GetSiteId(sitePath);
        string endpoint = $"/sites/{siteId}/drives";
        string content = await HttpClient.GetAsync(endpoint, default);
        driveName ??= GraphAPI.DefaultDriveName;

        JObject root = JObject.Parse(content);
        JToken drives = root["value"]!;
        JToken drive = drives.First(drive =>
            drive is not null
            && drive["name"] is JToken name
            && name.ToString() == driveName
        )!;

        string? id = drive["id"]?.Value<string>();
        if (id is null) {
            Log.Fatal($"ID for drive {driveName} was not found");
            throw new AppDataNotFoundException();
        }

        return id;
    }

    internal async Task<string> CreateFolder(FolderUploadInfo folderUpload) {
        if (folderUpload.Path is string path)
            Debug.Assert(path[0] == '/', "Path must start with '/'");

        string fullPath = folderUpload.Path is null ? "root" : $"root:{folderUpload.Path}:";
        string endpoint = $"/sites/{folderUpload.SiteId}/drives/{folderUpload.DriveId}/{fullPath}/children";

        DriveItemUpload driveItem = new(folderUpload.Name);
        string data = JsonSerializer.Serialize(driveItem, SerializerOptions);
        Log.Information($"Creating folder with endpoint {endpoint} and request body {data.Colourise(AnsiColours.Magenta)}");
        
        string response = await HttpClient.PostAsync(endpoint, data, default);
        Log.Information($"Succesfully created folder {folderUpload.Path ?? string.Empty}/{folderUpload.Name}");
        return response;
    }
}