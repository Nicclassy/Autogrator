using Microsoft.Extensions.Logging;

using Autogrator.Utilities;
using Serilog;

namespace Autogrator.SharePointAutomation;

public sealed class RequestLoggingHandler(ILogger<RequestLoggingHandler> logger) : DelegatingHandler(new HttpClientHandler()) {
    public required bool LoggingEnabled { get; init; }
    public required bool UseSeparateRequestsLogger { get; init; }

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        void log(string requestUri, IAnsiSequence ansi) {
            const string template = "{RequestColour}{Method} {RequestUri}{Reset}";
            if (UseSeparateRequestsLogger)
                logger.LogDebug(
                    template, ansi, request.Method.Method, requestUri, AnsiColours.Reset
                );
            else
                Log.Debug(
                    template, ansi, request.Method.Method, requestUri, AnsiColours.Reset
                );
        }
        
        if (LoggingEnabled) {
            AnsiColour requestColour = RequestColour(request.Method);
            string requestUri = request
                .RequestUri!
                .ToString()
                .Replace("%2C", ",");

            log(requestUri, requestColour);
        }
            
        return base.SendAsync(request, cancellationToken);
    }

    private static AnsiColour RequestColour(HttpMethod method) {
        if (method == HttpMethod.Get) return AnsiColours.Green;
        if (method == HttpMethod.Post) return AnsiColours.Blue;
        if (method == HttpMethod.Put) return AnsiColours.Yellow;
        return AnsiColours.White;
    }
}
