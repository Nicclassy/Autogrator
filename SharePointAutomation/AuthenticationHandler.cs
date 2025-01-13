using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

using Serilog;

using Autogrator.Exceptions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

internal sealed class AuthenticationHandler : DelegatingHandler {
    private const string UrlFormat = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";
    private const string ContentFormat = "grant_type=client_credentials&client_id={0}&client_secret={1}&scope={2}";
    private const string MediaType = "application/x-www-form-urlencoded";

    internal AuthenticationHandler() : base(new HttpClientHandler()) { }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        string accessToken = await GetAccessToken();
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<string> GetAccessToken() {
        string authUrl = string.Format(UrlFormat, AutogratorApplication.TenantID);
        string dataContent = string.Format(
            ContentFormat,
            AutogratorApplication.ClientID,
            AutogratorApplication.ClientSecret,
            AutogratorApplication.Scope
        );
        StringContent data = new(dataContent, Encoding.UTF8, MediaType);

        using HttpClient httpClient = new();
        HttpResponseMessage authResponse = await httpClient.PostAsync(authUrl, data);
        if (!authResponse.IsSuccessStatusCode) {
            Log.Fatal($"Access token request returned an unsuccessful status code of {authResponse.StatusCode}");
            throw new AccessTokenRequestFailedException();
        }

        string content = await authResponse.Content.ReadAsStringAsync();
        using JsonDocument json = JsonDocument.Parse(content);
        string token =
            json.RootElement.GetProperty("access_token").GetString()
            ?? throw new InvalidDataException("Access token not found");
        Log.Information("Token succesfully obtained.");
        return token;
    }
}