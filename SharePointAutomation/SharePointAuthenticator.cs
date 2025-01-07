using System.Text.Json;
using System.Text;

using Serilog;

using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public static class SharePointAuthenticator {
    public static string GetAccessToken(bool useStored) {
        if (useStored)
            return Credentials.EnvironmentVariable("AG_SHAREPOINT_ACCESS_TOKEN");
        else
            return Task.Run(GetAccessToken).Result;
    }
	public static async Task<string> GetAccessToken() {
        string authUrl = $"https://login.microsoftonline.com/{Credentials.SharePoint.TenantID}/oauth2/v2.0/token";
        StringContent data = new(
            $"grant_type=client_credentials&client_id={Credentials.SharePoint.ClientID}&client_secret={Credentials.SharePoint.ClientSecret}&scope={Credentials.SharePoint.SiteURL}/.default",
            Encoding.UTF8,
            "application/x-www-form-urlencoded"
        );

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
