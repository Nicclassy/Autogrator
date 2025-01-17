using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Serilog;

using Autogrator.Extensions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorAppClient(SharePointGraphClient _graphClient) {
    public SharePointGraphClient GraphClient { get; } = _graphClient;

    public async Task CreateFolder(string name, string sitePath, string driveName, string? path = null) {
        string siteId = await GraphClient.GetSiteId(sitePath);
        string driveId = await GraphClient.GetDriveId(driveName, sitePath);
        string folderPath = SharePointUtils.FormatItemPath(path);
        FolderCreationInfo uploadInfo = new(name, siteId, driveId, folderPath);
        
        string response = await GraphClient.CreateFolder(uploadInfo);
        Log.Information($"Folder creation responded with response {response}");
    }

    public async Task UploadFile(
        string fileName, string localFileDirectory, string uploadFolderPath, 
        string? sitePath = null, string? driveName = null
    ) {
        // TODO: Check if the folder that is being uploaded to exists. If not create it
        string siteId = await GraphClient.GetSiteId(sitePath);
        string driveId = await GraphClient.GetDriveId(driveName, sitePath);

        (string parentFolder, string parentName) = uploadFolderPath.RightSplitOnce('/');
        string parentId = await GraphClient.GetItemId(parentName, parentFolder);

        string localFilePath = Path.Combine(localFileDirectory, fileName);
        FileUploadInfo uploadInfo = new(
            fileName, localFileDirectory, localFilePath, 
            parentId, siteId, driveId
        );

        string response = await GraphClient.UploadFile(uploadInfo);
        Log.Information($"File creation responded with response {response}");
    }

    public async Task DownloadFile(string fileName, string destinationFolder, string driveName) {
        string driveId = await GraphClient.GetDriveId(driveName);
        string itemId = await GraphClient.GetItemId(fileName, driveName: driveName);

        string destinationPath = Path.Combine(destinationFolder, fileName);
        FileDownloadInfo download = new(
            fileName, destinationFolder, 
            destinationPath, driveId, itemId
        );
        await GraphClient.DownloadFile(download);
    }

    public static AutogratorAppClient Create() {
        using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
        ILogger<LoggingHandler> logger = factory.CreateLogger<LoggingHandler>();

        IMemoryCache memoryCache = new MemoryCache(new MemoryCacheOptions());
        LoggingHandler loggingHandler = new(logger);
        AuthenticationHandler authenticationHandler = new(memoryCache) {
            InnerHandler = loggingHandler
        };

        HttpClient httpClient = new(authenticationHandler);
        GraphHttpClient graphHttpClient = new(httpClient);
        SharePointGraphClient graphClient = new(graphHttpClient);
        return new(graphClient);
    }
}
