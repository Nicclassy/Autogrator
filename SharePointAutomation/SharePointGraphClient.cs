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
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    internal GraphHttpClient HttpClient { get; } = httpClient;

    internal async Task<List<string>> GetItemsInDrive(string? driveName = null, string? sitePath = null) {
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
            Debug.Assert(path[0] == '/', "Path should start with '/'");
        string fullPath = folderUpload.Path is null ? "root" : $"root:{folderUpload.Path}:";
        string endpoint = $"/sites/{folderUpload.SiteId}/drives/{folderUpload.DriveId}/{fullPath}/children";

        DriveItemUpload driveItem = new(folderUpload.Name);
        string data = JsonSerializer.Serialize(driveItem, SerializerOptions);
        Log.Information($"Creating folder with endpoint {endpoint} and request body {data.Colourise(AnsiColours.Magenta)}");
        return await HttpClient.PostAsync(endpoint, data, default);
    }
}