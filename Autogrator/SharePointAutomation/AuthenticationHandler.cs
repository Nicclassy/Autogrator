using System.Net.Http.Headers;
using System.Text;
using System.IdentityModel.Tokens.Jwt;

using Microsoft.Extensions.Caching.Memory;
using Microsoft.IdentityModel.Tokens;
using Serilog;

using Autogrator.Exceptions;
using Autogrator.Utilities;
using Newtonsoft.Json.Linq;
using Autogrator.Extensions;

namespace Autogrator.SharePointAutomation;

internal sealed class AuthenticationHandler : DelegatingHandler {
    // TODO: Check if JWT Token has appropriate permissions
    private const string UrlFormat = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";
    private const string ContentFormat = "grant_type=client_credentials&client_id={0}&client_secret={1}&scope={2}";
    private const string MediaType = "application/x-www-form-urlencoded";
    private const string AccessTokenKey = "access_token";

    private readonly JwtSecurityTokenHandler TokenHandler = new();
    private readonly IMemoryCache memoryCache;

    internal AuthenticationHandler(IMemoryCache memoryCache)
        : base(new HttpClientHandler()) => this.memoryCache = memoryCache;

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        string accessToken = await GetAccessTokenAsync();
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
    }

    internal async Task<string> GetAccessTokenAsync() {
        if (GetCachedAccessToken() is string cachedToken) {
            Log.Information("Access token was found in the cache.");
            return cachedToken;
        }

        string authUrl = string.Format(UrlFormat, ApplicationRegistration.TenantID);
        string dataContent = string.Format(
            ContentFormat,
            ApplicationRegistration.ClientID,
            ApplicationRegistration.ClientSecret,
            ApplicationRegistration.Scope
        );
        StringContent data = new(dataContent, Encoding.UTF8, MediaType);

        using HttpClient httpClient = new();
        HttpResponseMessage authResponse = await httpClient.PostAsync(authUrl, data);
        if (!authResponse.IsSuccessStatusCode) {
            Log.Fatal("Access token request returned an unsuccessful status code of {StatusCode}", authResponse.StatusCode);
            throw new AccessTokenRetrievalFailedException();
        }

        string content = await authResponse.Content.ReadAsStringAsync();
        JObject json = JObject.Parse(content);
        if (json[AccessTokenKey]?.ToString() is not string accessToken) {
            Log.Fatal("The key '{AccessTokenKey}' was not found in the JSON response", AccessTokenKey);
            throw new AccessTokenRetrievalFailedException();
        }
            
        
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
        Log.Information("Cachked token duration is {Minutes} minutes", duration.Minutes);
        memoryCache.Set(AccessTokenKey, accessToken, duration);
    }
}