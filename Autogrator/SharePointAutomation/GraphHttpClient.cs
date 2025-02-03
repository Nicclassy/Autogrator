using System.Text;
using System.Text.Json;
using System.Net.Http.Headers;

using Newtonsoft.Json.Linq;
using Serilog;

using Autogrator.Exceptions;
using Autogrator.Utilities;
namespace Autogrator.SharePointAutomation;

public sealed class GraphHttpClient(HttpClient httpClient) {
    private const string PostMediaType = "application/json";
    private const string PaginationKey = "@odata.nextLink";

    private static readonly Encoding PostEncoding = Encoding.UTF8;
    private static readonly JsonSerializerOptions SerializerOptions = new() {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    internal async Task<bool> IsSuccessfulResponseÁsync(string endpoint, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        HttpResponseMessage response = await httpClient.GetAsync(requestUri, cancellationToken);
        return response.IsSuccessStatusCode;
    }

    internal async Task<string> GetKeyAsync(string endpoint, string key, CancellationToken cancellationToken) {
        string content = await GetAsync($"{endpoint}{FormatRequestKeys([key])}", cancellationToken);
        JObject json = JObject.Parse(content);
        if (json[key]?.ToString() is not string value) {
            Log.Fatal($"Key '{key}' was not found in the response.");
            throw new AppDataNotFoundException();
        }

        return value;
    }

    internal async Task<List<T>> GetPaginatedAsync<T>(
        string endpoint, CancellationToken cancellationToken, params string[] keys
    ) {
        List<T> items = [];
        string? currentEndpoint = $"{endpoint}{FormatRequestKeys(keys)}";
        do {
            string content = await GetAsync(currentEndpoint, cancellationToken);
            JObject json = JObject.Parse(content);
            IEnumerable<T> values = json["value"]!
                .Children()
                .Select(token => token.ToObject<T>()!);
            items.AddRange(values);
            currentEndpoint = json[PaginationKey]?.ToString();
        } while (currentEndpoint is not null);

        return items;
    }

    internal async Task<Stream> GetStreamAsync(string endpoint, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        HttpResponseMessage response = await httpClient.GetAsync(requestUri, cancellationToken);
        if (!response.IsSuccessStatusCode)
            LogFailureAndThrow("GET", endpoint, response);

        return await response.Content.ReadAsStreamAsync(cancellationToken);
    }

    internal async Task<string> GetAsync(string endpoint, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        HttpResponseMessage response = await httpClient.GetAsync(requestUri, cancellationToken);
        if (!response.IsSuccessStatusCode)
            LogFailureAndThrow("GET", endpoint, response);

        return await response.Content.ReadAsStringAsync(cancellationToken);
    }

    internal async Task<T> GetAsync<T>(string endpoint, CancellationToken cancellationToken) {
        string response = await GetAsync(endpoint, cancellationToken);
        return JsonSerializer.Deserialize<T>(response, SerializerOptions)!;
    }

    internal async Task<string> PostAsync(string endpoint, string data, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        StringContent content = new(data, PostEncoding, PostMediaType);
        HttpResponseMessage response = await httpClient.PostAsync(requestUri, content, cancellationToken);
        if (!response.IsSuccessStatusCode)
            LogFailureAndThrow("POST", endpoint, response);

        return await response.Content.ReadAsStringAsync(cancellationToken);
    }

    internal async Task<string> PutAsync(string endpoint, HttpContent content, CancellationToken cancellationToken) {
        string requestUri = CreateRequestUri(endpoint);
        HttpRequestMessage request = new(HttpMethod.Put, requestUri) {
            Content = content,
        };
        request.Content!.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

        HttpResponseMessage response = await httpClient.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
            LogFailureAndThrow("PUT", endpoint, response);

        return await response.Content.ReadAsStringAsync(cancellationToken);
    }

    private static void LogFailureAndThrow(string method, string endpoint, HttpResponseMessage response) {
        Log.Fatal(
            "{ErrorColour}Request {Method} {Endpoint} failed with status code {StatusCode}. Reason: {Reason}{Reset}",
            AnsiColours.Red, method, endpoint, (int) response.StatusCode, response.ReasonPhrase, AnsiColours.Reset
        );
        throw new RequestUnsuccessfulException();
    }

    private static string CreateRequestUri(string endpoint) =>
        endpoint.StartsWith("https") ? endpoint : GraphAPI.URL + endpoint;

    private static string FormatRequestKeys(string[] keys, string query = "select") =>
        keys.Length > 0 ? $"?${query}={string.Join(',', keys)}" : string.Empty;
}