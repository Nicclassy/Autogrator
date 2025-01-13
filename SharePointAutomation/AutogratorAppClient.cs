using Microsoft.Graph;
using Azure.Identity;
using Graph = Microsoft.Graph.Models;
using Serilog;

using Autogrator.Utilities;
using Autogrator.Exceptions;
using Microsoft.Extensions.Logging;
using Autogrator.Extensions;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorAppClient : IDisposable {
    private static readonly string[] Scopes = { AutogratorApplication.Scope };

    internal GraphServiceClient GraphClient { get; }
    internal GraphRequestHandler RequestHandler { get; }

    public AutogratorAppClient() {
        RequestHandler = CreateRequestHandler();
        GraphClient = CreateGraphClient(RequestHandler);
    }

    public async Task<Graph.Site?> GetRootSiteAsync() =>
        await GraphClient.Sites[SharePointSite.Hostname].GetAsync();

    public async Task<Graph.Drive?> GetRootDriveAsync() {
        Graph.Site? rootSite = await GetRootSiteAsync();
        return rootSite is Graph.Site { Id: string id } ?
            await GraphClient.Sites[id].Drive.GetAsync() : null;
    }

    public async Task<Graph.Site?> GetSiteAtPathAsync(string? sitePath = null) =>
        await GraphClient
            .Sites[$"{SharePointSite.Hostname}:{sitePath ?? SharePointSite.DefaultSitePath}"]
            .GetAsync();

    public async Task<Graph.Drive?> GetDriveFromSiteAsync(string driveName, string? siteId = null) {
        string validSiteId = siteId ?? await GetSiteIdAsync();
        var driveCollection = await GraphClient
            .Sites[validSiteId]
            .Drives
            .GetAsync();
        List<Graph.Drive> drives = driveCollection?.Value
            ?? throw new AppDataNotFoundException($"No drives found for site ID {siteId}");
        return drives.Find(drive => drive.Name == driveName);
    }

    public async Task<List<Graph.DriveItem>?> GetDriveItemsAsync(string? siteId = null, string? driveId = null) {
        string validSiteId = siteId ?? await GetSiteIdAsync();
        string validDriveId = driveId ?? await GetDriveIdAsync();
        Console.WriteLine($"Site ID: {validSiteId}\nDrive ID: {validDriveId}");

        RequestHandler.ModifyNextRequest(request =>
            request.RequestUri = request.RequestUri!.Append("/root/children")
        );
        object? result = await GraphClient
            .Sites[validSiteId]
            .Drives[validDriveId]
            .GetAsync();
        return (List<Graph.DriveItem>?) result;
    }

    public async Task<string> GetSiteIdAsync(string? sitePath = null) {
        string path = sitePath ?? SharePointSite.DefaultSitePath;
        Graph.Site? site = await GetSiteAtPathAsync(path);
        return site is Graph.Site { Id: string id } ?
            id : throw new AppDataNotFoundException($"Site at path {path} not found");
    }

    public async Task<string> GetDriveIdAsync(string? driveName = null) {
        string name = driveName ?? SharePointSite.DefaultDrive;
        Graph.Drive? drive = await GetDriveFromSiteAsync(name);
        return drive is Graph.Drive { Id: string id } ?
            id : throw new AppDataNotFoundException($"Drive with name {name} not found");
    }
    private static GraphRequestHandler CreateRequestHandler() {
        using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
        ILogger<GraphRequestHandler> logger = factory.CreateLogger<GraphRequestHandler>();
        var handler = new GraphRequestHandler(logger);
        return handler;
    }

    private static GraphServiceClient CreateGraphClient(GraphRequestHandler handler) {
        var httpClient = new HttpClient(handler);
        var options = new TokenCredentialOptions {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        ClientSecretCredential tokenCredential = new(
            AutogratorApplication.TenantID,
            AutogratorApplication.ClientID,
            AutogratorApplication.ClientSecret,
            options
        );

        Log.Information("Created application client");
        GraphServiceClient client = new(httpClient, tokenCredential, Scopes);
        return client;
    }

    public void Dispose() => GraphClient.Dispose();
}
