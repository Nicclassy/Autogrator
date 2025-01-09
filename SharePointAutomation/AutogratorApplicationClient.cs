using Microsoft.Graph;
using Azure.Identity;
using Graph = Microsoft.Graph.Models;

using Autogrator.Utilities;
using Serilog;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorApplicationClient(GraphServiceClient graphClient) : IDisposable {
    private static readonly string[] Scopes = { AutogratorApplication.Scope };

    public GraphServiceClient GraphClient { get; } = graphClient;

    public AutogratorApplicationClient() : this(CreateGraphClient()) {}

    public async Task<Graph.Site?> GetRootSiteAsync() =>
        await GraphClient.Sites[Sites.Hostname].GetAsync();

    public async Task<Graph.Drive?> GetRootDriveAsync() {
        Graph.Site? rootSite = await GetRootSiteAsync();
        return rootSite is Graph.Site { Id: string id } ? 
            await GraphClient.Sites[id].Drive.GetAsync() : null;
    }

    public async Task<Graph.DriveItem?> GetSiteAtPathAsync(string sitePath) {
        Graph.Drive? rootDrive = await GetRootDriveAsync();
        Log.Information($"Root drive name is {rootDrive?.Name} and is {rootDrive?.Id}");
        return rootDrive is Graph.Drive { Id: string id } ?
            await GraphClient.Drives[rootDrive.Id].Root.ItemWithPath(sitePath).GetAsync() : null;
    }

    public async Task<List<Graph.DriveItem>?> GetSiteSubfoldersAsync(string sitePath) {
        Graph.DriveItem? siteDrive = await GetSiteAtPathAsync(sitePath);
        return siteDrive?.Children;
    }

    private static GraphServiceClient CreateGraphClient() {
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
        return new(tokenCredential, Scopes);
    }

    public void Dispose() => GraphClient.Dispose();
}
