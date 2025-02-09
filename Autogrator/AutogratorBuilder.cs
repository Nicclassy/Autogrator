﻿using System.Globalization;

using Outlook = Microsoft.Office.Interop.Outlook;

using Autogrator.OutlookAutomation;
using Autogrator.SharePointAutomation;
using Autogrator.Notifications;

namespace Autogrator;

public partial class Autogrator {
    public partial class Builder {
        private IAllowedSenders? _allowedSenders;
        private EmailFileNameFormatter? _emailFileNameFormatter;
        private TimeNotifier? _timeNotifier;
        private AutogratorOptions? _options;

        public Builder WithAllowedSenders(IAllowedSenders allowedSenders) {
            _allowedSenders = allowedSenders;
            return this;
        }

        public Builder WithEmailFileNameFormatter(EmailFileNameFormatter formatter) {
            _emailFileNameFormatter = formatter;
            return this;
        }

        public Builder WithOptions(AutogratorOptions options) {
            _options = options;
            return this;
        }

        public Builder WithTimeNotifier(TimeNotifier timeNotifier) {
            _timeNotifier = timeNotifier;
            return this;
        }

        public Autogrator Build() {
            AutogratorOptions options = _options ?? new();
            SharePointClient client = SharePointClient.Create(
                enableRequestLogging: options.EnableRequestLogging,
                useSeparateRequestLogger: options.UseSeparateRequestLogger
            );

            EmailReceiver receiver = new() { LogRejectedEmails = options.LogRejectedEmails };

            return new(client, receiver) {
                Options = options,
                AllowedSenders = _allowedSenders ?? DefaultAllowedSenders,
                EmailFileNameFormatter = _emailFileNameFormatter ?? DefaultEmailFileNameFormatter,
                TimeNotifier = _timeNotifier
            };
        }

        private static EmailFileNameFormatter DefaultEmailFileNameFormatter =>
            delegate (Outlook.MailItem mailItem) {
                string creationTime =
                    mailItem.CreationTime.ToString("yyyyMMddTHHmmss", CultureInfo.CurrentCulture);
                return $"{mailItem.SenderName} {creationTime}";
            };

        private static IAllowedSenders DefaultAllowedSenders => new AllEmailSendersAllowed();
    }
}