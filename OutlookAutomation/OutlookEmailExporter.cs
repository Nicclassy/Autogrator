using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Serilog;

using Autogrator.Utilities;
using Autogrator.Extensions;

namespace Autogrator.OutlookAutomation;

internal enum ExportResult {
    Ok = 0,
    FileExists = 1,
}

public static class OutlookEmailExporter {
    private const Outlook.OlSaveAsType SaveAsType = Outlook.OlSaveAsType.olHTML;
    private const Word.WdSaveFormat ExportFormat = Word.WdSaveFormat.wdFormatPDF;

    public static void SaveAndExportEmail(
        Outlook.MailItem email,
        out string exportedFilePath,
        string? directory = null,
        Outlook.OlSaveAsType saveAsType = SaveAsType,
        Word.WdSaveFormat exportFormat = ExportFormat
    ) {
        string fileName = email.ExportFileName();
        string fileExtension = GetFileExtensionWithoutPrefix(SaveAsType, "ol");

        string savedFileDir = directory ?? Directories.DownloadsFolder;
        string savedFileName = $"{fileName}.{fileExtension}";
        string savedFilePath = Path.Combine(savedFileDir, savedFileName);

        ExportResult exportResult = 
            ExportEmail(savedFilePath, email, saveAsType, exportFormat, out exportedFilePath);
        if (exportResult == ExportResult.FileExists)
            // Nothing was created therefore nothing needs deletion. The function can stop here.
            return;

        // Delete the temporary email file. Only the exported verison is needed
        // A files folder is also generated containing attachments and other files;
        // this must also be removed.
        string filesFolder = Path.Combine(savedFileDir, $"{fileName}_files");
        if (Directory.Exists(filesFolder))
            Directory.Delete(filesFolder, recursive: true);
        else
            Log.Warning($"Files folder '{filesFolder}' does not exist even though it should be generated");

        if (!File.Exists(savedFilePath)) {
            Log.Fatal($"File '{savedFilePath}' was expected to be generated but does not exist");
            Environment.Exit(1);
        }
        File.Delete(savedFilePath);

        Log.Information($"Successfully exported email.");
    }

    private static ExportResult ExportEmail(
        string path, 
        Outlook.MailItem email,
        Outlook.OlSaveAsType saveAsType,
        Word.WdSaveFormat exportFormat,
        out string exportedFilePath
    ) {
        string fileName = Path.GetFileNameWithoutExtension(path)!;
        string fileExtension = GetFileExtensionWithoutPrefix(ExportFormat, "wdFormat");

        string exportedFileName = $"{fileName}.{fileExtension}";
        exportedFilePath = 
            Path.Combine(Path.GetDirectoryName(path)!, exportedFileName);

        if (File.Exists(exportedFilePath)) {
            Log.Information($"The file that will be exported {exportedFilePath} already exists.");
            return ExportResult.FileExists;
        }

        email.SaveAs(path, saveAsType);

        Word.Application word = new();
        Word.Document document = word.Documents.Open(path);

        document.SaveAs2(exportedFilePath, exportFormat);
        Log.Information($"Exported file to path '{exportedFilePath}'");

        document.Close();
        word.Quit();
        return ExportResult.Ok;
    }

    private static string GetFileExtensionWithoutPrefix(Enum extensionValue, string prefix) =>
        extensionValue.ToString()[prefix.Length..].ToLower();
}
