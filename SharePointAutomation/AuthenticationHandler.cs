using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.IdentityModel.Tokens.Jwt;

using Microsoft.Extensions.Caching.Memory;
using Microsoft.IdentityModel.Tokens;
using Serilog;

using Autogrator.Exceptions;
using Autogrator.Utilities;
using Newtonsoft.Json.Linq;

namespace Autogrator.SharePointAutomation;

internal sealed class AuthenticationHandler : DelegatingHandler {
    private const string UrlFormat = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";
    private const string ContentFormat = "grant_type=client_credentials&client_id={0}&client_secret={1}&scope={2}";
    private const string MediaType = "application/x-www-form-urlencoded";
    private const string AccessTokenKey = "access_token";

    private readonly JwtSecurityTokenHandler TokenHandler = new();
    private readonly IMemoryCache memoryCache;

    internal AuthenticationHandler(IMemoryCache memoryCache)
        : base(new HttpClientHandler()) => this.memoryCache = memoryCache;

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        string accessToken = await GetAccessToken();
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
    }

    private async Task<string> GetAccessToken() {
        if (GetCachedAccessToken() is string cachedToken) {
            Log.Information("Access token was found in the cache.");
            return cachedToken;
        }

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
        JObject json = JObject.Parse(content);
        string accessToken = json[AccessTokenKey]?.ToString()
            ?? throw new InvalidDataException("Access token not found");
        
        CacheAccessToken(accessToken);
        Log.Information("Token succesfully cached and obtained.");
        return accessToken;
    }

    private string? GetCachedAccessToken() =>
        memoryCache.TryGetValue(AccessTokenKey, out object? value) && value is string accessToken
            ? accessToken 
            : null;

    private void CacheAccessToken(string accessToken) {
        SecurityToken token = TokenHandler.ReadToken(accessToken)!;
        TimeSpan duration = token.ValidTo - token.ValidFrom;
        Log.Information($"Token duration is {duration.Minutes} minutes");
        memoryCache.Set(AccessTokenKey, accessToken, duration);
    }
}