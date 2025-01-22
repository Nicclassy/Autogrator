using System.Globalization;

using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

using Autogrator.OutlookAutomation;
using Autogrator.SharePointAutomation;

namespace Autogrator;

public partial class Autogrator {
    public partial class Builder {
        private ILogger<RequestLoggingHandler>? _requestLogger;
        private IAllowedSenderList? _allowedSenders;
        private EmailFileNameFormatter? _emailFileNameFormatter;

        public Builder WithRequestLogger(ILogger<RequestLoggingHandler> requestLogger) {
            _requestLogger = requestLogger;
            return this;
        }

        public Builder WithAllowedSenders(IAllowedSenderList allowedSenders) {
            _allowedSenders = allowedSenders;
            return this;
        }

        public Builder WithEmailFileNameFormatter(EmailFileNameFormatter formatter) {
            _emailFileNameFormatter = formatter;
            return this;
        }

        public Autogrator Build() {
            // TODO: Consider moving some of this logic to SharePointGraphClient
            IMemoryCache memoryCache = new MemoryCache(new MemoryCacheOptions());
            RequestLoggingHandler loggingHandler = new(_requestLogger ?? DefaultRequestLogger);
            AuthenticationHandler authenticationHandler = new(memoryCache) {
                InnerHandler = loggingHandler
            };

            HttpClient httpClient = new(authenticationHandler);
            GraphHttpClient graphHttpClient = new(httpClient);
            SharePointGraphClient graphClient = new(graphHttpClient);

            OutlookAuthenticator authenticator = new();
            EmailReceiver receiver = new(authenticator);

            return new(graphClient, receiver) {
                AllowedSenders = _allowedSenders ?? DefaultAllowedSenders,
                EmailFileNameFormatter = _emailFileNameFormatter ?? DefaultEmailFileNameFormatter
            };
        }

        private static ILogger<RequestLoggingHandler> DefaultRequestLogger {
            get {
                using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
                return factory.CreateLogger<RequestLoggingHandler>();
            }
        }

        private static EmailFileNameFormatter DefaultEmailFileNameFormatter =>
            delegate (Outlook.MailItem mailItem) {
                string creationTime =
                    mailItem.CreationTime.ToString("yyyyMMdd", CultureInfo.CurrentCulture);
                return $"{mailItem.SenderName} {creationTime}";
            };

        private static IAllowedSenderList DefaultAllowedSenders => new AllEmailSendersAllowed();
    }
}