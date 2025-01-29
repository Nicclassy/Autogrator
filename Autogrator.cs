using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;
using Serilog.Sinks.SystemConsole.Themes;

using Autogrator.Extensions;
using Autogrator.Utilities;
using Autogrator.Data;
using Autogrator.SharePointAutomation;
using Autogrator.OutlookAutomation;
using Autogrator.Notifications;

namespace Autogrator;

public sealed partial class Autogrator(
    SharePointGraphClient _graphClient,
    EmailReceiver _emailReceiver
) {
    public required IAllowedSenderList AllowedSenders { get; set; }
    public required EmailFileNameFormatter EmailFileNameFormatter { get; set; }
    public required AutogratorOptions Options {
        get => field;
        set {
            if (value.UseDefaultLoggingConfiguration)
                SetDefaultLoggingConfiguration();
            field = value;
        }
    }

    internal SharePointGraphClient GraphClient { get; } = _graphClient;
    internal EmailReceiver EmailReceiver { get; } = _emailReceiver;

    public sealed partial class Builder;

    private static void SetDefaultLoggingConfiguration() {
        Dictionary<ConsoleThemeStyle, string> styles = new() {
            { ConsoleThemeStyle.LevelWarning, "\u001b[38;5;214m" },
            { ConsoleThemeStyle.LevelError, "\u001b[38;5;196m" },
            { ConsoleThemeStyle.LevelFatal, "\u001b[38;5;161m" },
            { ConsoleThemeStyle.LevelDebug, "\u001b[38;5;249m" },
            { ConsoleThemeStyle.LevelVerbose, "\u001b[38;5;245m" }
        };
        AnsiConsoleTheme theme = new(styles);

        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console(theme: theme)
            .WriteTo.File(
                new StylelessTextFormatter(), 
                EmailExceptionNotifier.LogFileName, 
                rollingInterval: RollingInterval.Day
            )
            .CreateLogger();
        Log.Information("Begin logging");
    }

    public async Task CreateFolderAsync(FolderInfo folder) {
        string driveId = await GraphClient.GetDriveIdAsync(folder.DriveName, folder.SitePath);

        string response = await GraphClient.CreateFolderAsync(folder, driveId);
        Log.Information(
            "Folder creation responded with response {Response}",
            response.PrettyJson().Colourise(AnsiColours.Magenta)
        );
    }

    public async Task CreateFolderRecursivelyAsync(FolderInfo folder) {
        string driveId = await GraphClient.GetDriveIdAsync(folder.DriveName, folder.SitePath);
        await GraphClient.CreateFolderRecursivelyAsync(folder, driveId);
    }

    public async Task<bool> FolderExistsAsync(FolderInfo folder) {
        string siteId = await GraphClient.GetDriveIdAsync(folder.DriveName, folder.SitePath);
        return await GraphClient.FolderExistsAsync(folder, siteId);
    }

    public async Task UploadFileAsync(FileUploadInfo fileUpload) {
        string driveId = await GraphClient.GetDriveIdAsync(fileUpload.DriveName, fileUpload.SitePath);

        (string parentFolder, string parentName) = fileUpload.UploadDirectory.RightSplitOnce('/');
        string parentId = await GraphClient.GetItemIdAsync(driveId, parentName, parentFolder);

        string response = await GraphClient.UploadFileAsync(fileUpload, driveId, parentId);
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

        string driveId = await GraphClient.GetDriveIdAsync(downloadInfo.DriveName, downloadInfo.SitePath);
        string itemId = await GraphClient.GetItemIdAsync(driveId, downloadInfo.FileName, downloadInfo.DownloadPath);
        await GraphClient.DownloadFileAsync(downloadInfo, destinationPath, driveId, itemId);
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
        FileDownloadInfo download = new() {
            FileName = AllowedSendersFile.Name,
            DestinationFileName = AllowedSendersFile.Name.FileNameWithSuffix(Options.CopiedFileSuffix),
            DestinationFolder = AllowedSendersFile.DownloadDestination,
            DriveName = AllowedSendersFile.DriveName,
            SitePath = AllowedSendersFile.SitePath
        };

        string downloadedFilePath = await DownloadFileAsync(download);
        AllowedSenders.Load(downloadedFilePath);
        AllowedSenders.Print();

        EmailReceiver.Listen(AllowedSenders);

        while (true) {
            if (EmailReceiver.TryReceiveEmail(out Outlook.MailItem email))
                await ProcessEmailAsync(email);
            await Task.Delay(1000);
        }
    }
}
