using System;
using Autogrator.Extensions;
using Autogrator.Utilities;
using Azure.Core;
using Microsoft.Extensions.Logging;

internal delegate void HttpRequestMessageAction(HttpRequestMessage request);

internal sealed class GraphRequestHandler : DelegatingHandler {
    private const bool IsEnabled = true;

    private readonly ILogger _logger;
    private HttpRequestMessageAction? _messageAction = null;

    internal GraphRequestHandler(ILogger<GraphRequestHandler> logger) : base(new HttpClientHandler()) => _logger = logger;

    internal void ModifyNextRequest(HttpRequestMessageAction action) => _messageAction = action;

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken
    ) {
        if (_messageAction is HttpRequestMessageAction action) {
            action(request);
            _messageAction = null;
        }

        if (IsEnabled) {
            var requestUri = request.RequestUri!.ToString().Colourise(AnsiColours.Cyan).Replace("%2C", ",");
            _logger.LogInformation($"Sending {request.Method} request to {requestUri}");
        }

        var response = base.SendAsync(request, cancellationToken);
        return response;
    }
}