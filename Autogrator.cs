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
    public required IAllowedSenderList AllowedSenders { get; set; }
    public required EmailFileNameFormatter EmailFileNameFormatter { get; set; }
    public required AutogratorOptions Options {
        get => field;
        init {
            if (value.UseDefaultLoggingConfiguration)
                SetDefaultLoggingConfiguration();
            field = value;
        }
    }

    internal SharePointClient Client { get; } = _client;
    internal EmailReceiver EmailReceiver { get; } = _emailReceiver;

    public sealed partial class Builder;

    private static partial void SetDefaultLoggingConfiguration();
      
    public async Task CreateFolderAsync(FolderInfo folder) {
        string driveId = await Client.GetDriveIdAsync(folder.DriveName, folder.SitePath);

        string response = await Client.CreateFolderAsync(folder, driveId);
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
        string parentId = await Client.GetItemIdAsync(driveId, parentName, parentFolder);

        string response = await Client.UploadFileAsync(fileUpload, driveId, parentId);
        Log.Information(
            "File creation responded with response {Response}",
            response.PrettyJson().Colourise(AnsiColours.Magenta)
        );
    }

    public async Task<string> DownloadFileAsync(FileDownloadInfo downloadInfo) {
        string destinationPath = Path.Combine(
            downloadInfo.DestinationFolder, 
            downloadInfo.DestinationFileName ?? downloadInfo.FileName
        );
        if (!Options.OverwriteDownloads && File.Exists(destinationPath)) {
            Log.Information(
                "File '{FileName}' already exists in {DestinationFolder}",
                downloadInfo.FileName, downloadInfo.DestinationFolder
            );
            return destinationPath;
        }

        string driveId = 
            await Client.GetDriveIdAsync(downloadInfo.DriveName, downloadInfo.SitePath);
        string itemId = 
            await Client.GetItemIdAsync(driveId, downloadInfo.FileName, downloadInfo.DownloadPath);
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
            SitePath = AllowedSendersFile.SitePath
        };

        string downloadedFilePath = await DownloadFileAsync(download);
        AllowedSenders.Load(downloadedFilePath);

        AppDomain.CurrentDomain.UnhandledException += EmailExceptionNotifier.EventHandler();
        EmailReceiver.Listen(AllowedSenders);

        while (true) {
            if (EmailReceiver.TryReceiveEmail(out Outlook.MailItem email))
                await ProcessEmailAsync(email);
            await Task.Delay(1000);
        }
    }
}
