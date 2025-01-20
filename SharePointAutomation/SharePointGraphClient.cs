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

    internal async Task<IEnumerable<DriveItemInfo>> GetChildren(string fullpath, string driveId) {
        List<DriveItemInfo> items = await HttpClient
            .GetPaginatedAsync<DriveItemInfo>($"/drives/{driveId}/{fullpath}/children", default, "name", "id");
        return items;
    }

    internal async Task<string> GetItemId(string driveId, string name, string? path = null) {
        string itemPath = SharePointUtils.FormatPath(path);
        IEnumerable<DriveItemInfo> driveItems = await GetChildren(itemPath, driveId);
        
        DriveItemInfo notFound = default;
        DriveItemInfo result = driveItems.FirstOrDefault(item => item.Name == name, notFound);
        if (result.Equals(notFound)) {
            Log.Fatal(
                "Item with name '{ItemName}' at path '{ItemPath}' was not found", 
                name, itemPath
            );
            throw new AppDataNotFoundException();
        }

        return result.Id;
    }

    internal async Task<string> GetSiteId(string sitePath) =>
        await HttpClient.GetKeyAsync($"/sites/{SharePointSite.Hostname}:{sitePath}", "id", default);
    
    internal async Task<string> GetDriveId(string driveName, string sitePath) {
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

        string? idValue = drive["id"]?.Value<string>();
        if (idValue is not string id) {
            Log.Fatal($"ID for drive {driveName} was not found");
            throw new AppDataNotFoundException();
        }

        return id;
    }

    internal async Task<string> CreateFolder(FolderInfo folder, string driveId) {
        string folderPath = SharePointUtils.FormatPath(folder.Directory);
        string endpoint = $"/drives/{driveId}/{folderPath}/children";

        DriveItemUpload driveItem = new(folder.Name);
        string data = JsonSerializer.Serialize(driveItem, SerializerOptions);
        Log.Information(
            "Creating folder with endpoint {Endpoint} and request body {RequestBody}",
            endpoint, data.Colourise(AnsiColours.Magenta)
        );
        
        string response = await HttpClient.PostAsync(endpoint, data, default);
        Log.Information(
            "Succesfully created folder {Folder} at {Path}",
            folder.Name, folder.Directory
        );
        return response;
    }

    internal async Task<bool> FolderExists(FolderInfo folder, string driveId) {
        string folderPath = SharePointUtils.FormatPath($"{folder.Directory}/{folder.Name}");
        string endpoint = $"/drives/{driveId}/{folderPath}";
        return await HttpClient.IsSuccessfulResponse(endpoint, default);
    }

    internal async Task<string> UploadFile(FileUploadInfo upload, string driveId, string parentId) {
        // Documentation: https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http
        string endpoint = $"/drives/{driveId}/items/{parentId}:/{upload.FileName}:/content";
        string localFilePath = Path.Combine(upload.LocalFileDirectory, upload.FileName);
        
        byte[] data = File.ReadAllBytes(localFilePath);
        ByteArrayContent content = new(data);
        return await HttpClient.PutAsync(endpoint, content, default);
    }

    internal async Task DownloadFile(FileDownloadInfo download, string destinationPath, string driveId, string itemId) {
        string endpoint = $"/drives/{driveId}/items/{itemId}/content";

        await using Stream downloadStream = await HttpClient.GetStreamAsync(endpoint, default);
        await using FileStream destinationStream = new(destinationPath, FileMode.Create);
        await downloadStream.CopyToAsync(destinationStream);
        Log.Information(
            "Successfully downloaded {FileName} to {DestinationFolder}",
            download.FileName, download.DestinationFolder
        );
    }
}