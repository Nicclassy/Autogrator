using Microsoft.Extensions.Logging;

using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

public sealed class RequestLoggingHandler : DelegatingHandler {
    private const bool LoggingEnabled = true;

    private readonly ILogger<RequestLoggingHandler> logger;

    public RequestLoggingHandler(ILogger<RequestLoggingHandler> logger) : base(new HttpClientHandler()) => this.logger = logger;
    
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        if (LoggingEnabled) {
            string requestColour = RequestColour(request.Method);
            string requestUri = request
                .RequestUri!
                .ToString()
                .Replace("%2C", ",");

            logger.LogInformation(
                "{RequestColour}{Method} {RequestUri}{Reset}",
                requestColour, request.Method.Method, requestUri, AnsiColours.Reset
            );
        }
            
        return base.SendAsync(request, cancellationToken);
    }

    private static string RequestColour(HttpMethod method) {
        if (method == HttpMethod.Get) return AnsiColours.Green;
        if (method == HttpMethod.Post) return AnsiColours.Blue;
        if (method == HttpMethod.Put) return AnsiColours.Yellow;
        return AnsiColours.White;
    }
}
