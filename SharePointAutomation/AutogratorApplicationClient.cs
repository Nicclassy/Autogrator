using Microsoft.Graph;
using Azure.Identity;
using Serilog;

using Autogrator.Utilities;

namespace Autogrator.Helpers;

public sealed class AutogratorApplicationClient(GraphServiceClient graphClient) : IDisposable {
    private static readonly string[] Scopes = { AutogratorApplication.Scope };

    public GraphServiceClient GraphClient { get; } = graphClient;

    public AutogratorApplicationClient() : this(CreateGraphClient()) {}

    public async Task<Microsoft.Graph.Models.Site?> GetRootSite() =>
        await GraphClient.Sites[Sites.Hostname].GetAsync();

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

        return new(tokenCredential, Scopes);
    }

    public void Dispose() => GraphClient.Dispose();
}
