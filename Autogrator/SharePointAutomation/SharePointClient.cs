using System.Text;
using System.Text.Json;
using System.Diagnostics;

using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using Serilog;

using Autogrator.Data;
using Autogrator.Extensions;
using Autogrator.Exceptions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class SharePointClient(GraphHttpClient _httpClient) {
    private static readonly JsonSerializerOptions SerializerOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public static SharePointClient Create(
        bool enableRequestLogging,
        bool useSeparateRequestLogger,
        ILogger<RequestLoggingHandler>? logger = null
    ) {
        IMemoryCache memoryCache = new MemoryCache(new MemoryCacheOptions());
        RequestLoggingHandler loggingHandler = new(logger ?? DefaultRequestLogger()) {
            LoggingEnabled = enableRequestLogging,
            UseSeparateRequestsLogger = useSeparateRequestLogger,
        };
        AuthenticationHandler authenticationHandler = new(memoryCache) {
            InnerHandler = loggingHandler
        };

        HttpClient httpClient = new(authenticationHandler);
        GraphHttpClient graphHttpClient = new(httpClient);
        return new(graphHttpClient);
    }

    internal GraphHttpClient HttpClient { get; } = _httpClient;

    internal async Task<FileModificationInfo> GetFileModificationInfoAsync(FileModificationInfoRequest fileInfo) {
        string driveId = await GetDriveIdAsync(fileInfo.DriveName, fileInfo.SitePath);
        string itemId = await GetItemIdAsync(fileInfo.FileName, driveId, fileInfo.FileDirectory);
        return await HttpClient.GetAsync<FileModificationInfo>($"/drives/{driveId}/items/{itemId}", default);
    }

    internal async Task<IEnumerable<DriveItemInfo>> GetChildrenAsync(string fullpath, string driveId) {
        List<DriveItemInfo> items = await HttpClient
            .GetPaginatedAsync<DriveItemInfo>($"/drives/{driveId}/{fullpath}/children", default, "name", "id");
        return items;
    }

    internal async Task<string> GetItemIdAsync(string itemName, string driveId, string? path = null) {
        string itemPath = FormatPath(path);
        IEnumerable<DriveItemInfo> driveItems = await GetChildrenAsync(itemPath, driveId);
        
        DriveItemInfo? result = driveItems.FirstOrDefault(item => item.Name == itemName);
        if (result is not DriveItemInfo { Id: string id }) {
            Log.Fatal(
                "Item with name '{ItemName}' at path '{ItemPath}' was not found", 
                itemName, itemPath
            );
            throw new AppDataNotFoundException();
        }

        return id;
    }

    internal async Task<string> GetSiteIdAsync(string sitePath) =>
        await HttpClient.GetKeyAsync($"/sites/{SharePoint.Hostname}:{sitePath}", "id", default);
    
    internal async Task<string> GetDriveIdAsync(string driveName, string sitePath) {
        string siteId = await GetSiteIdAsync(sitePath);
        string endpoint = $"/sites/{siteId}/drives";
        string content = await HttpClient.GetAsync(endpoint, default);

        JObject root = JObject.Parse(content);
        JToken drives = root["value"]!;
        JToken drive = drives.First(drive =>
            drive is not null
            && drive["name"] is JToken name
            && name.ToString() == driveName
        )!;

        if (drive["id"]?.ToString() is not string id) {
            Log.Fatal("ID for drive {DriveName} was not found", driveName);
            throw new AppDataNotFoundException();
        }

        return id;
    }

    internal async Task<string> CreateFolderAsync(FolderInfo folder, string driveId) {
        string folderPath = FormatPath(folder.Directory);
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

    internal async Task CreateFolderRecursivelyAsync(FolderInfo folder, string driveId) {
        if (folder.Directory is not { } directory) {
            if (!await FolderExistsAsync(folder, driveId)) {
                Log.Information("Folder /{Name} does not exist. Creating...", folder.Name);
                await CreateFolderAsync(folder, driveId);
            } else {
                Log.Information("Folder /{Name} already exists", folder.Name);
            }
            return;
        }

        string[] dirnames = [..directory.TrimStart('/').Split('/'), folder.Name];
        StringBuilder builder = new();
        foreach (string dirname in dirnames) {
            FolderInfo parentFolder = folder with {
                Name = dirname,
                Directory = builder.ToString().NullIfWhiteSpace()
            };
            
            if (!await FolderExistsAsync(parentFolder, driveId)) {
                Log.Information(
                    "Folder {Directory}/{Name} does not exist. Creating...",
                    parentFolder.Directory ?? string.Empty, parentFolder.Name
                );
                await CreateFolderAsync(parentFolder, driveId);
            } else {
                Log.Information(
                    "Folder {Directory}/{Name} already exists",
                    parentFolder.Directory ?? string.Empty, parentFolder.Name
                );
            }

            builder.Append('/');
            builder.Append(dirname);
        }
    }

    internal async Task<bool> FolderExistsAsync(FolderInfo folder, string driveId) {
        string folderPath = FormatPath($"{folder.Directory}/{folder.Name}");
        string endpoint = $"/drives/{driveId}/{folderPath}";
        return await HttpClient.IsSuccessfulResponseÁsync(endpoint, default);
    }

    internal async Task<string> UploadFileAsync(FileUploadInfo upload, string driveId, string parentId) {
        string endpoint = $"/drives/{driveId}/items/{parentId}:/{upload.FileName}:/content";
        string localFilePath = Path.Combine(upload.LocalFileDirectory, upload.FileName);
        
        byte[] data = File.ReadAllBytes(localFilePath);
        ByteArrayContent content = new(data);
        return await HttpClient.PutAsync(endpoint, content, default);
    }

    internal async Task DownloadFileAsync(FileDownloadInfo download, string destinationPath, string driveId, string itemId) {
        string endpoint = $"/drives/{driveId}/items/{itemId}/content";

        await using Stream downloadStream = await HttpClient.GetStreamAsync(endpoint, default);
        await using FileStream destinationStream = new(destinationPath, FileMode.Create);
        await downloadStream.CopyToAsync(destinationStream);
        Log.Information(
            "Successfully downloaded {FileName} to {DestinationFolder}",
            Path.GetFileName(destinationPath), download.DestinationFolder
        );
    }

    private static string FormatPath(string? itemPath) {
        if (string.IsNullOrWhiteSpace(itemPath)) return "root";

        Debug.Assert(itemPath[0] == '/', "Path must start with '/'");
        return $"root:{itemPath}:";
    }

    private static ILogger<RequestLoggingHandler> DefaultRequestLogger() {
        using ILoggerFactory factory = LoggerFactory.Create(builder =>
            builder.AddConsole().SetMinimumLevel(LogLevel.Debug)
        );
        return factory.CreateLogger<RequestLoggingHandler>();
    }
}