using System.Text;
using System.Text.Json;

using Newtonsoft.Json.Linq;
using Serilog;

using Autogrator.Exceptions;
using Autogrator.Extensions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class GraphHttpClient(HttpClient httpClient) {
    private static readonly Encoding PostEncoding = Encoding.UTF8;
    private const string PostMediaType = "application/json";

    internal async Task<string> GetKeyAsync(string endpoint, string key, CancellationToken cancellationToken) {
        string content = await GetAsync($"{endpoint}?{FormatRequestKeys([key])}", cancellationToken);
        using JsonDocument document = JsonDocument.Parse(content);
        return document.RootElement.GetProperty(key).GetString()!;
    }

    internal async Task<IEnumerable<JToken>> GetValuesByKeysAsync(
        string endpoint, CancellationToken cancellationToken, params string[] keys
    ) {
        string content = await GetAsync($"{endpoint}?{FormatRequestKeys(keys)}", cancellationToken);
        JObject root = JObject.Parse(content);
        return root["value"]!.Children();
    }

    internal async Task<string> GetAsync(string endpoint, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        HttpResponseMessage message = await httpClient.GetAsync(requestUri, cancellationToken);
        if (!message.IsSuccessStatusCode) {
            string errorMessage = $"Request GET {endpoint} failed with status code {(int)message.StatusCode}. "
                + $"Reason: {message.ReasonPhrase}";
            Log.Fatal(errorMessage.Colourise(AnsiColours.Red));
            throw new RequestUnsuccessfulException();
        }

        return await message.Content.ReadAsStringAsync(cancellationToken);
    }

    internal async Task<string> PostAsync(string endpoint, string data, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        StringContent content = new(data, PostEncoding, PostMediaType);
        HttpResponseMessage message = await httpClient.PostAsync(requestUri, content, cancellationToken);
        if (!message.IsSuccessStatusCode) {
            string errorMessage = $"Request POST {endpoint} failed with status code {(int)message.StatusCode}. "
                + $"Reason: {message.ReasonPhrase}";
            Log.Fatal(errorMessage.Colourise(AnsiColours.Red));
            throw new RequestUnsuccessfulException();
        }

        return await message.Content.ReadAsStringAsync(cancellationToken);
    }

    //internal async Task UploadAsync(FileUploadInfo uploadInfo, CancellationToken cancellationToken) {
    //    byte[] content = File.ReadAllBytes(uploadInfo.LocalFilePath);
    //    ByteArrayContent uploadData = new(content);
    //}

    private static string CreateRequestUri(string endpoint) => GraphAPI.URL + endpoint;

    private static string FormatRequestKeys(string[] keys, string query = "select") => $"${query}={string.Join(',', keys)}";
}