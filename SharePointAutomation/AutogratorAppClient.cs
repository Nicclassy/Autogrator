using Microsoft.Extensions.Logging;
using Serilog;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorAppClient(SharePointGraphClient _graphClient) {
    public SharePointGraphClient GraphClient { get; } = _graphClient;

    public async Task CreateFolder(string name, string sitePath, string driveName, string? path = null) {
        string siteId = await GraphClient.GetSiteId(sitePath);
        string driveId = await GraphClient.GetDriveId(driveName, sitePath);
        FolderUploadInfo uploadInfo = new(name, siteId, driveId, path);
        
        string response = await GraphClient.CreateFolder(uploadInfo);
        Log.Information($"Folder creation responded with response {response}");
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
