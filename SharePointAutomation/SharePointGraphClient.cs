using System.Text.Json;

using Newtonsoft.Json.Linq;
using Serilog;

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

    internal async Task<IEnumerable<DriveItemInfo>> GetChildren(string fullpath, string? driveName = null, string? sitePath = null) {
        string driveId = await GetDriveId(driveName, sitePath);
        List<DriveItemInfo> items = await HttpClient
            .GetPaginatedAsync<DriveItemInfo>($"/drives/{driveId}/{fullpath}/children", default, "name", "id");
        return items;
    }

    internal async Task<string> GetItemId(
        string itemName, string itemPath, string? driveName = null, string? sitePath = null
    ) {
        string fullpath = SharePointUtils.FormatFilePath(itemPath);
        IEnumerable<DriveItemInfo> driveItems = await GetChildren(fullpath, driveName, sitePath);
        
        DriveItemInfo notFound = default;
        DriveItemInfo result = driveItems.FirstOrDefault(item => item.Name == itemName, notFound);
        if (result.Equals(notFound)) {
            Log.Fatal($"Item with name '{itemName}' at path '{fullpath}' was not found");
            Environment.Exit(1);
        }

        return result.Id;
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

    internal async Task<string> CreateFolder(FolderCreationInfo folderCreation) {
        string endpoint = 
            $"/sites/{folderCreation.SiteId}/drives/{folderCreation.DriveId}/{folderCreation.Path}/children";

        DriveItemUpload driveItem = new(folderCreation.Name);
        string data = JsonSerializer.Serialize(driveItem, SerializerOptions);
        Log.Information(
            $"Creating folder with endpoint {endpoint} and request body {data.Colourise(AnsiColours.Magenta)}"
        );
        
        string response = await HttpClient.PostAsync(endpoint, data, default);
        Log.Information($"Succesfully created folder {folderCreation.Path ?? string.Empty}/{folderCreation.Name}");
        return response;
    }

    internal async Task<string> UploadFile(FileUploadInfo upload) {
        // Documentation: https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http
        string endpoint = 
            $"/sites/{upload.SiteId}/drives/{upload.DriveId}/items/{upload.ParentId}:/{upload.FileName}:/content";
        string filePath = Path.Combine(upload.LocalFileDir, upload.FileName);

        byte[] data = File.ReadAllBytes(filePath);
        ByteArrayContent content = new(data);
        return await HttpClient.PutAsync(endpoint, content, default);
    }   
}