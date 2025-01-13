using System.Text.Json;

using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using Serilog;

using Autogrator.Extensions;
using Autogrator.Exceptions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class SharePointGraphClient(HttpClient httpClient) {
    internal async Task<string> GetSiteId(string? sitePath = null) =>
        await GetContentValueAsync($"/sites/{SharePointSite.Hostname}:{sitePath ?? GraphAPI.DefaultSitePath}", "id", default);
    
    internal async Task<string> GetDriveId(string? driveName = null, string? sitePath = null) {
        string siteId = await GetSiteId(sitePath);
        string endpoint = $"/sites/{siteId}/drives";
        string content = await GetContentAsync(endpoint, default);
        driveName ??= GraphAPI.DefaultDriveName;

        JObject root = JObject.Parse(content);
        JToken drives = root["value"]!;
        JToken drive = drives.First(drive =>
            drive is not null
            && drive["name"] is JToken name
            && name.ToString() == driveName
        )!;

        string? id = drive["id"]?.Value<string>();
        if (id is null) {
            Log.Fatal($"ID for drive {driveName} was not found");
            throw new AppDataNotFoundException();
        }

        return id;
    }

    internal async Task<T> GetAsync<T>(string endpoint, CancellationToken cancellationToken) {
        string content = await GetContentAsync(endpoint, cancellationToken);
        return JsonSerializer.Deserialize<T>(content)!;
    }

    internal async Task<string> GetContentValueAsync(string endpoint, string key, CancellationToken cancellationToken) {
        string content = await GetContentAsync($"{endpoint}?$select={key}", cancellationToken);
        using JsonDocument document = JsonDocument.Parse(content);
        return document.RootElement.GetProperty(key).GetString()!;
    }

    internal async Task<string> GetContentAsync(string endpoint, CancellationToken cancellationToken) {
        string requestUri = RequestUri(endpoint);
        HttpResponseMessage message = await httpClient.GetAsync(requestUri, cancellationToken);
        if (!message.IsSuccessStatusCode) {
            string failedRequest = $"GET {endpoint}".Colourise(AnsiColours.Green);
            Log.Fatal(
                $"Request {failedRequest} failed with status code {(int) message.StatusCode} " +
                $"with reason {message.ReasonPhrase}"
            );
            throw new RequestUnsuccessfulException();
        }

        return await message.Content.ReadAsStringAsync();
    }

    private string RequestUri(string endpoint) => GraphAPI.URL + endpoint;
}