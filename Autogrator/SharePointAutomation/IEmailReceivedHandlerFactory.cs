using Autogrator.OutlookAutomation;

namespace Autogrator.SharePointAutomation;

public interface IEmailReceivedHandlerFactory {
    public Task<EmailReceivedHandler> CreateHandler();
}