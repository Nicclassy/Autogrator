using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

using Autogrator.Extensions;
using Autogrator.Utilities;
using Autogrator.Data;
using Autogrator.SharePointAutomation;
using Autogrator.OutlookAutomation;
using Autogrator.Notifications;

namespace Autogrator;

public sealed partial class Autogrator(SharePointClient _client, EmailReceiver _emailReceiver) {
    public required IAllowedSenders AllowedSenders { get; set; }
    public required EmailFileNameFormatter EmailFileNameFormatter { get; set; }
    public required AutogratorOptions Options {
        get => field;
        init {
            field = value;
            if (value.UseDefaultLoggingConfiguration)
                SetDefaultLoggingConfiguration();
        }
    }

    internal SharePointClient Client { get; } = _client;
    internal EmailReceiver EmailReceiver { get; } = _emailReceiver;

    public sealed partial class Builder;

    private partial void SetDefaultLoggingConfiguration();

    public async Task AddEmailReceivedHandler(IEmailReceivedHandlerFactory factory) =>
        EmailReceiver.OnEmailReceived += await factory.CreateHandler();

    public DownloadFileOnChange GetFileChangeDownloader(FileDownloadInfo download, Action? postDownload = null) =>
        new(download) {
            GetFileModificationInfo = Client.GetFileModificationInfoAsync,
            DownloadFile = DownloadFileAsync,
            PostDownload = postDownload
        };
      
    public async Task CreateFolderAsync(FolderInfo folder) {
        string driveId = await Client.GetDriveIdAsync(folder.DriveName, folder.SitePath);

        string response = await Client.CreateFolderAsync(folder, driveId);
        if (Options.LogGraphJSONResponses)
            Log.Information(
                "Folder creation responded with response {Response}",
                response.PrettyJson().Colourise(AnsiColours.Magenta)
            );
    }
    public async Task CreateFolderRecursivelyAsync(FolderInfo folder) {
        string driveId = await Client.GetDriveIdAsync(folder.DriveName, folder.SitePath);
        await Client.CreateFolderRecursivelyAsync(folder, driveId);
    }

    public async Task<bool> FolderExistsAsync(FolderInfo folder) {
        string siteId = await Client.GetDriveIdAsync(folder.DriveName, folder.SitePath);
        return await Client.FolderExistsAsync(folder, siteId);
    }

    public async Task UploadFileAsync(FileUploadInfo fileUpload) {
        string driveId = await Client.GetDriveIdAsync(fileUpload.DriveName, fileUpload.SitePath);

        (string parentFolder, string parentName) = fileUpload.UploadDirectory.RightSplitOnce('/');
        string parentId = await Client.GetItemIdAsync(parentName, driveId, parentFolder);

        string response = await Client.UploadFileAsync(fileUpload, driveId, parentId);
        if (Options.LogGraphJSONResponses)
            Log.Information(
                "File upload responded with {Response}",
                response.PrettyJson().Colourise(AnsiColours.Magenta)
            );
    }

    public async Task<string> DownloadFileAsync(FileDownloadInfo downloadInfo) {
        string destinationPath = Path.Combine(
            downloadInfo.DestinationFolder,
            downloadInfo.DestinationFileName ?? downloadInfo.FileName
        );
        if (!downloadInfo.AlwaysDownload && !Options.OverwriteDownloads && File.Exists(destinationPath)) {
            Log.Information(
                "File '{FileName}' already exists in {DestinationFolder}",
                downloadInfo.FileName, downloadInfo.DestinationFolder
            );
            return destinationPath;
        }

        string driveId = 
            await Client.GetDriveIdAsync(downloadInfo.DriveName, downloadInfo.SitePath);
        string itemId = 
            await Client.GetItemIdAsync(downloadInfo.FileName, driveId, downloadInfo.DownloadPath);
        await Client.DownloadFileAsync(downloadInfo, destinationPath, driveId, itemId);
        return destinationPath;
    }

    public async Task ProcessEmailAsync(Outlook.MailItem email) {
        EmailExporter.SaveAndExportEmail(
            email, 
            out EmailExportInfo emailInfo,
            fileNameFormatter: EmailFileNameFormatter
        );

        string senderFolderName = AllowedSenders.GetSenderFolder(emailInfo.SenderEmailAddress);
        string uploadDirectory = $"/{senderFolderName}/{Options.EmailsFolderName}";
        FolderInfo folder = new() {
            Name = Options.EmailsFolderName,
            Directory = senderFolderName,
            DriveName = SharePoint.UploadDriveName,
            SitePath = SharePoint.UploadSitePath
        };

        FileUploadInfo uploadInfo = new() {
            FileName = emailInfo.FileName,
            LocalFileDirectory = emailInfo.FileDirectory,
            UploadDirectory = uploadDirectory,
            DriveName = SharePoint.UploadDriveName,
            SitePath = SharePoint.UploadSitePath
        };

        await CreateFolderRecursivelyAsync(folder);
        await UploadFileAsync(uploadInfo);
    }

    public async Task Run() {
        string destinationFileName =
            AllowedSendersFile.Name.FileNameWithSuffix(Options.CopiedFileSuffix);
        FileDownloadInfo download = new() {
            FileName = AllowedSendersFile.Name,
            DestinationFileName = destinationFileName,
            DestinationFolder = AllowedSendersFile.DownloadDestination,
            DriveName = AllowedSendersFile.DriveName,
            SitePath = AllowedSendersFile.SitePath,
            AlwaysDownload = true
        };

        string downloadedFilePath = await DownloadFileAsync(download);
        AllowedSenders.Load(downloadedFilePath);

        if (Options.AutoDownloadAllowedSenders) {
            DownloadFileOnChange downloader = GetFileChangeDownloader(download, postDownload: delegate {
                Log.Information(
                    "Finished downloading file {DownloadedFilePath}. Reloading allowed senders",
                    downloadedFilePath
                );
                AllowedSenders.Reload();
            });
            await AddEmailReceivedHandler(downloader);
        }

        if (Options.SendExceptionNotificationEmails) {
            EmailExceptionNotifier emailNotifier = new() {
                LogFileName = Options.LogFileName,
                LoggingDirectory = Options.LoggingFolder,
                ReviewSentEmails = Options.ReviewSentEmails
            };
            AppDomain.CurrentDomain.UnhandledException += emailNotifier.EventHandler();
        }
        EmailReceiver.Listen(AllowedSenders);

        while (true) {
            if (EmailReceiver.TryReceiveEmail(out Outlook.MailItem email))
                await ProcessEmailAsync(email);
            else
                await Task.Delay(Options.ExecutionInterval);
        }
    }

    public static async Task Main() {
        Autogrator autogrator = new Builder().Build();
        await autogrator.Run();
    }
}
