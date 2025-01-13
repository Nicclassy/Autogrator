using Microsoft.Extensions.Logging;

namespace Autogrator.SharePointAutomation;

public sealed class AutogratorAppClient(SharePointGraphClient _graphClient) {
    public SharePointGraphClient GraphClient { get; } = _graphClient;

    public static AutogratorAppClient Create() {
        using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
        ILogger<LoggingHandler> logger = factory.CreateLogger<LoggingHandler>();

        LoggingHandler loggingHandler = new(logger);
        AuthenticationHandler authenticationHandler = new() {
            InnerHandler = loggingHandler
        };

        HttpClient httpClient = new(authenticationHandler);
        SharePointGraphClient graphClient = new(httpClient);
        return new(graphClient);
    }
}
