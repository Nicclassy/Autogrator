using Autogrator.Extensions;
using Autogrator.Utilities;
using Microsoft.Extensions.Logging;
using Serilog;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorAppClient(SharePointGraphClient _graphClient) {
    public SharePointGraphClient GraphClient { get; } = _graphClient;

    public async Task CreateFolder(string name, string sitePath, string driveName, string? path = null) {
        string siteId = await GraphClient.GetSiteId(sitePath);
        string driveId = await GraphClient.GetDriveId(driveName, sitePath);
        string folderPath = SharePointUtils.FormatFilePath(path);
        FolderCreationInfo uploadInfo = new(name, siteId, driveId, folderPath);
        
        string response = await GraphClient.CreateFolder(uploadInfo);
        Log.Information($"Folder creation responded with response {response}");
    }

    public async Task UploadFile(string fileName, string localFileDir, string uploadFolderPath, string? sitePath = null, string? driveName = null) {
        string siteId = await GraphClient.GetSiteId(sitePath);
        string driveId = await GraphClient.GetDriveId(driveName, sitePath);

        (string parentFolder, string parentName) = uploadFolderPath.RightSplitOnce('/');
        string parentId = await GraphClient.GetItemId(parentName, parentFolder);
        Log.Information($"Item ID of '{parentName}' is {parentId}");

        FileUploadInfo uploadInfo = new(fileName, localFileDir, parentId, siteId, driveId);

        string response = await GraphClient.UploadFile(uploadInfo);
        Log.Information($"File creation responded with response {response}");
    }

    public static AutogratorAppClient Create() {
        using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
        ILogger<LoggingHandler> logger = factory.CreateLogger<LoggingHandler>();

        LoggingHandler loggingHandler = new(logger);
        AuthenticationHandler authenticationHandler = new() {
            InnerHandler = loggingHandler
        };

        HttpClient httpClient = new(authenticationHandler);
        GraphHttpClient graphHttpClient = new(httpClient);
        SharePointGraphClient graphClient = new(graphHttpClient);
        return new(graphClient);
    }
}
