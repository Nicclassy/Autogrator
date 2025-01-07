using System.Text.Json;
using System.Text;

using Serilog;

using Autogrator.Utilities;
using Autogrator.Extensions;

namespace Autogrator.SharePointAutomation;

public static class SharePointAuthenticator {
    private const string UrlFormat = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";
    private const string ContentFormat = "grant_type=client_credentials&client_id={0}&client_secret={1}&scope={2}/.default";
    private const string MediaType = "application/x-www-form-urlencoded";

    public static string GetAccessToken(bool useStored) {
        if (useStored)
            return "AG_SHAREPOINT_ACCESS_TOKEN".Env();
        else
            return Task.Run(GetAccessToken).Result;
    }
	public static async Task<string> GetAccessToken() {
        string authUrl = string.Format(UrlFormat, Credentials.SharePoint.TenantID);
        string dataContent = string.Format(
            ContentFormat, 
            Credentials.SharePoint.ClientID, 
            Credentials.SharePoint.ClientSecret, 
            Credentials.SharePoint.SiteURL
        );
        StringContent data = new(dataContent, Encoding.UTF8, MediaType);

        using HttpClient httpClient = new();
        HttpResponseMessage authResponse = await httpClient.PostAsync(authUrl, data);
        authResponse.EnsureSuccessStatusCode();

        string content = await authResponse.Content.ReadAsStringAsync();
        using JsonDocument json = JsonDocument.Parse(content);
        string token = 
            json.RootElement.GetProperty("access_token").GetString()
            ?? throw new InvalidDataException("Access token not found");
        Log.Information("Token succesfully obtained.");
        return token;
    }
}
