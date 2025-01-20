using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Serilog;

using Autogrator.Extensions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorAppClient(SharePointGraphClient _graphClient) {
    public SharePointGraphClient GraphClient { get; } = _graphClient;
    
    public async Task CreateFolder(FolderInfo folder) {
        string siteId = await GraphClient.GetSiteId(folder.SitePath);
        string driveId = await GraphClient.GetDriveId(folder.DriveName, folder.SitePath);
        
        string response = await GraphClient.CreateFolder(folder, driveId);
        Log.Information(
            "Folder creation responded with response {Response}", 
            response.PrettyJson().Colourise(AnsiColours.Magenta)
        );
    }

    public async Task<bool> FolderExists(FolderInfo folder) {
        string siteId = await GraphClient.GetDriveId(folder.DriveName, folder.SitePath);
        return await GraphClient.FolderExists(folder, siteId);
    }

    public async Task UploadFile(FileUploadInfo fileUpload) {
        string siteId = await GraphClient.GetSiteId(fileUpload.SitePath);
        string driveId = await GraphClient.GetDriveId(fileUpload.DriveName, fileUpload.SitePath);

        (string parentFolder, string parentName) = fileUpload.UploadDirectory.RightSplitOnce('/');
        string parentId = await GraphClient.GetItemId(driveId, parentName, parentFolder);

        string response = await GraphClient.UploadFile(fileUpload, driveId, parentId);
        Log.Information(
            "File creation responded with response {Response}",
            response.PrettyJson().Colourise(AnsiColours.Magenta)
        );
    }

    public async Task DownloadFile(FileDownloadInfo downloadInfo) {
        string destinationPath = Path.Combine(downloadInfo.DestinationFolder, downloadInfo.FileName);
        if (File.Exists(destinationPath)) {
            Log.Information(
                "File '{FileName}' already exists in {DestinationFolder}",
                downloadInfo.FileName, downloadInfo.DestinationFolder
            );
            return;
        }

        string driveId = await GraphClient.GetDriveId(downloadInfo.DriveName, downloadInfo.SitePath);
        string itemId = await GraphClient.GetItemId(driveId, downloadInfo.FileName, downloadInfo.DownloadPath);
        await GraphClient.DownloadFile(downloadInfo, destinationPath, driveId, itemId);
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
