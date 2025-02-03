using System.Globalization;

using Serilog;

using Autogrator.Data;
using Autogrator.OutlookAutomation;

namespace Autogrator.SharePointAutomation;

using ModificationInfoProvider = Func<FileModificationInfoRequest, Task<FileModificationInfo>>;
using FileDownloader = Func<FileDownloadInfo, Task<string>>;

public sealed class DownloadFileOnChange(FileDownloadInfo downloadInfo) : IEmailReceivedHandlerFactory {
    public required ModificationInfoProvider GetFileModificationInfo { get; init; }
    public required FileDownloader DownloadFile { get; init; }
    public Action? PostDownload { get; init; }

    private FileModificationInfo? currentFileInfo;
    private readonly FileModificationInfoRequest fileInfoRequest = new() {
        FileName = downloadInfo.FileName,
        FileDirectory = downloadInfo.DownloadPath,
        DriveName = downloadInfo.DriveName,
        SitePath = downloadInfo.SitePath
    };

    public async Task<EmailReceivedHandler> CreateHandler() {
        currentFileInfo = await GetFileModificationInfo(fileInfoRequest);
        return async delegate {
            FileModificationInfo fileInfo = await GetFileModificationInfo(fileInfoRequest);
            if (fileInfo is not FileModificationInfo { LastModifiedDateTime: DateTime modificationTime })
                throw new KeyNotFoundException(
                    "The attribute 'LastModifiedDateTime' was not found in the API response"
                );

            if (fileInfo != currentFileInfo) {
                DateTime previousModificationTime = currentFileInfo.LastModifiedDateTime
                    ?? throw new InvalidDataException("The previous LastModifiedDateTime was null");
                Log.Information(
                    "File {FileName} (previously modified at {PreviousTime}) was modified at {Time}. Redownloading", 
                    downloadInfo.FileName,
                    previousModificationTime.ToString("t", CultureInfo.CurrentCulture),
                    modificationTime.ToString("t", CultureInfo.CurrentCulture)
                );
                await DownloadFile(downloadInfo);
                PostDownload?.Invoke();
                currentFileInfo = fileInfo;
            }
        };
    }
}