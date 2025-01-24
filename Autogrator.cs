using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

using Autogrator.Extensions;
using Autogrator.Utilities;
using Autogrator.Data;
using Autogrator.SharePointAutomation;
using Autogrator.OutlookAutomation;

namespace Autogrator;

public sealed partial class Autogrator(
    SharePointGraphClient _graphClient,
    EmailReceiver _emailReceiver
) {
    public required IAllowedSenderList AllowedSenders { get; set; }
    public required EmailFileNameFormatter EmailFileNameFormatter { get; set; }

    internal SharePointGraphClient GraphClient { get; } = _graphClient;
    internal EmailReceiver EmailReceiver { get; } = _emailReceiver;

    // TODO: Make options
    private const bool OverwriteDownloads = false;
    private const string CopiedFileSuffix = " (Copy)";
    private const string EmailsFolderName = "Emails";

    public sealed partial class Builder;

    public async Task CreateFolder(FolderInfo folder) {
        string driveId = await GraphClient.GetDriveId(folder.DriveName, folder.SitePath);

        string response = await GraphClient.CreateFolder(folder, driveId);
        Log.Information(
            "Folder creation responded with response {Response}",
            response.PrettyJson().Colourise(AnsiColours.Magenta)
        );
    }

    public async Task CreateFolderRecursively(FolderInfo folder) {
        string driveId = await GraphClient.GetDriveId(folder.DriveName, folder.SitePath);
        await GraphClient.CreateFolderRecursively(folder, driveId);
    }

    public async Task<bool> FolderExists(FolderInfo folder) {
        string siteId = await GraphClient.GetDriveId(folder.DriveName, folder.SitePath);
        return await GraphClient.FolderExists(folder, siteId);
    }

    public async Task UploadFile(FileUploadInfo fileUpload) {
        string driveId = await GraphClient.GetDriveId(fileUpload.DriveName, fileUpload.SitePath);

        (string parentFolder, string parentName) = fileUpload.UploadDirectory.RightSplitOnce('/');
        string parentId = await GraphClient.GetItemId(driveId, parentName, parentFolder);

        string response = await GraphClient.UploadFile(fileUpload, driveId, parentId);
        Log.Information(
            "File creation responded with response {Response}",
            response.PrettyJson().Colourise(AnsiColours.Magenta)
        );
    }

    public async Task<string> DownloadFile(FileDownloadInfo downloadInfo, bool overwrite = OverwriteDownloads) {
        string destinationPath = Path.Combine(
            downloadInfo.DestinationFolder, 
            downloadInfo.DestinationFileName ?? downloadInfo.FileName
        );
        if (!overwrite && File.Exists(destinationPath)) {
            Log.Information(
                "File '{FileName}' already exists in {DestinationFolder}",
                downloadInfo.FileName, downloadInfo.DestinationFolder
            );
            return destinationPath;
        }

        string driveId = await GraphClient.GetDriveId(downloadInfo.DriveName, downloadInfo.SitePath);
        string itemId = await GraphClient.GetItemId(driveId, downloadInfo.FileName, downloadInfo.DownloadPath);
        await GraphClient.DownloadFile(downloadInfo, destinationPath, driveId, itemId);
        return destinationPath;
    }

    public async Task ProcessEmail(Outlook.MailItem email) {
        EmailExporter.SaveAndExportEmail(
            email, 
            out EmailExportInfo emailInfo,
            fileNameFormatter: EmailFileNameFormatter
        );

        string senderFolderName = AllowedSenders.GetSenderFolder(emailInfo.SenderEmailAddress);
        string uploadDirectory = $"/{senderFolderName}/{EmailsFolderName}";
        FolderInfo folder = new() {
            Name = EmailsFolderName,
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

        await CreateFolderRecursively(folder);
        await UploadFile(uploadInfo);
    }

    public async Task Run() {
        FileDownloadInfo download = new() {
            FileName = AllowedSendersFile.Name,
            DestinationFileName = AllowedSendersFile.Name.FileNameWithSuffix(CopiedFileSuffix),
            DestinationFolder = AllowedSendersFile.DownloadDestination,
            DriveName = AllowedSendersFile.DriveName,
            SitePath = AllowedSendersFile.SitePath
        };

        string downloadedFilePath = await DownloadFile(download);
        AllowedSenders.Load(downloadedFilePath);
        AllowedSenders.Print();

        //var emails = EmailReceiver.Inbox.Emails();
        //Console.ForegroundColor = ConsoleColor.Cyan;
        //emails.ToList().ForEach(mailItem => Console.WriteLine(mailItem.SenderEmailAddress));
        //Console.ForegroundColor = ConsoleColor.DarkCyan;

        //Console.ForegroundColor = ConsoleColor.White;
        //var firstEmail = EmailReceiver.Inbox
        //    .Emails()
        //    .First(email => AllowedSenders.IsAllowed(email.SenderEmailAddress));
        //await ProcessEmail(firstEmail);
        EmailReceiver.Listen(AllowedSenders);

        while (true) {
            if (EmailReceiver.TryReceiveEmail(out Outlook.MailItem email))
                await ProcessEmail(email);
            await Task.Delay(1000);
        }
    }
}
