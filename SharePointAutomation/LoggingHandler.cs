using System;
using Microsoft.Extensions.Logging;

using Autogrator.Extensions;
using Autogrator.Utilities;

namespace Autogrator.SharePointAutomation;

internal sealed class LoggingHandler : DelegatingHandler {
    private const bool LoggingEnabled = true;

    private readonly ILogger<LoggingHandler> logger;

    internal LoggingHandler(ILogger<LoggingHandler> logger) : base(new HttpClientHandler()) => this.logger = logger;
    
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        if (LoggingEnabled)
            logger.LogInformation(FormatLoggedRequest(request));

        return base.SendAsync(request, cancellationToken);
    }

    private static string FormatLoggedRequest(HttpRequestMessage request) {
        string requestColour = RequestColour(request.Method);
        string requestUri = request
            .RequestUri!
            .ToString()
            .Replace("%2C", ",");
        return $"{request.Method.Method} {requestUri}".Colourise(requestColour);
    }
    private static string RequestColour(HttpMethod method) {
        if (method == HttpMethod.Get) return AnsiColours.Green;
        if (method == HttpMethod.Post) return AnsiColours.Blue;
        return AnsiColours.White;
    }
}
