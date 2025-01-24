using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Serilog;

using Autogrator.Data;
using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public enum ExportResult {
    Ok = 0,
    FileExists = 1,
}

public delegate string EmailFileNameFormatter(Outlook.MailItem email);

public static class EmailExporter {
    private const bool DeleteSavedEmail = false;
    private const Outlook.OlSaveAsType SaveAsType = Outlook.OlSaveAsType.olHTML;
    private const Word.WdSaveFormat ExportFormat = Word.WdSaveFormat.wdFormatPDF;

    public static void SaveAndExportEmail(
        Outlook.MailItem email,
        out EmailExportInfo exportInfo,
        string? directory = null,
        EmailFileNameFormatter? fileNameFormatter = null,
        Outlook.OlSaveAsType saveAsType = SaveAsType,
        Word.WdSaveFormat exportFormat = ExportFormat
    ) {
        SaveEmail(email, out EmailSaveInfo emailInfo, directory, fileNameFormatter, saveAsType, infoOnly: true);
        ExportResult exportResult = 
            ExportEmail(emailInfo, email, saveAsType, exportFormat, out exportInfo);

        if (exportResult == ExportResult.FileExists)
            // Nothing was created therefore nothing needs deletion. The function can end here.
            return;

        Cleanup(emailInfo);
        Log.Information($"Successfully exported email.");
    }

    public static void SaveEmail(
        Outlook.MailItem email, 
        out EmailSaveInfo savedEmail,
        string? directory = null,
        EmailFileNameFormatter? fileNameFormatter = null,
        Outlook.OlSaveAsType saveAsType = SaveAsType,
        bool infoOnly = false
    ) {
        string fileName = fileNameFormatter is EmailFileNameFormatter formatter
            ? formatter(email)
            : email.EntryID;
        string fileExtension = GetFileExtensionWithoutPrefix(saveAsType, "ol");

        string savedFileName = $"{fileName}.{fileExtension}";
        string savedFileDirectory = directory ?? Directories.DownloadsFolder;
        string savedFilePath = Path.Combine(savedFileDirectory, savedFileName);
        savedEmail = new EmailSaveInfo { 
            FileName = savedFileName,
            FileDirectory = savedFileDirectory,
            FilePath = savedFilePath
        };

        if (infoOnly)
            return;

        Log.Information("Saving email to {FilePath}", savedFilePath);
        email.SaveAs(savedFilePath, saveAsType);
        Cleanup(savedEmail);
    }

    public static ExportResult ExportEmail(
        EmailSaveInfo emailInfo, 
        Outlook.MailItem email,
        Outlook.OlSaveAsType saveAsType,
        Word.WdSaveFormat exportFormat,
        out EmailExportInfo exportInfo

    ) {
        string fileName = Path.GetFileNameWithoutExtension(emailInfo.FileName);
        string fileExtension = GetFileExtensionWithoutPrefix(ExportFormat, "wdFormat");

        string exportedFileName = $"{fileName}.{fileExtension}";
        string exportedFileDirectory = emailInfo.FileDirectory;
        string exportedFilePath = Path.Combine(exportedFileDirectory, exportedFileName);
        exportInfo = new EmailExportInfo {
            FileName = exportedFileName,
            FileDirectory = exportedFileDirectory,
            SenderEmailAddress = email.SenderEmailAddress
        };

        if (File.Exists(exportedFilePath)) {
            Log.Information("The exported file {ExportedFilePath} already exists.", exportedFilePath);
            return ExportResult.FileExists;
        }

        if (!File.Exists(emailInfo.FilePath))
            email.SaveAs(emailInfo.FilePath, saveAsType);

        Word.Application word = new();
        Word.Document document = word.Documents.Open(emailInfo.FilePath);

        Log.Information(
          "Exporting email to {FilePath} with address {Address} and type {SaveAsType}",
           emailInfo.FilePath, email.SenderEmailAddress, saveAsType.ToString()
        );

        document.SaveAs2(exportedFilePath, exportFormat);
        Log.Information("Exported file to path '{ExportedFilePath}'", exportedFilePath);

        document.Close();
        word.Quit();
        return ExportResult.Ok;
    }

    private static void Cleanup(EmailSaveInfo emailInfo, bool deleteSavedEmail = DeleteSavedEmail) {
        // Delete the temporary email file. Only the exported verison is needed
        // A files folder is also generated containing attachments and other files;
        // this must also be removed.
        string filesFolder = 
            Path.Combine(emailInfo.FileDirectory, $"{Path.GetFileNameWithoutExtension(emailInfo.FileName)}_files");
        if (Directory.Exists(filesFolder))
            Directory.Delete(filesFolder, recursive: true);
        else
            Log.Warning("Files folder {FilesFolder} does not exist even though it should have been generated", filesFolder);

        if (!deleteSavedEmail)
            return;
        if (!File.Exists(emailInfo.FilePath)) {
            Log.Fatal("File {SavedFilePath} was expected to be generated but does not exist", emailInfo.FilePath);
            throw new InvalidOperationException();
        }
        File.Delete(emailInfo.FilePath);
    }

    private static string GetFileExtensionWithoutPrefix(Enum extensionValue, string prefix) =>
        extensionValue.ToString()[prefix.Length..].ToLower();
}
