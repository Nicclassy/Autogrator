using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

using Serilog;

using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public static class OutlookEmailExporter {
    private const Outlook.OlSaveAsType SaveAsType = Outlook.OlSaveAsType.olHTML;
    private const Word.WdSaveFormat ExportFormat = Word.WdSaveFormat.wdFormatPDF;

    public static void SaveAndExportEmail(
        Outlook.MailItem email,
        string? directory = null,
        Outlook.OlSaveAsType saveAsType = SaveAsType,
        Word.WdSaveFormat exportFormat = ExportFormat
    ) {
        string fileName = email.EntryID.ToString();
        string fileExtension = GetFileExtensionWithoutPrefix(SaveAsType, "ol");

        string savedFileDir = directory ?? Directories.DownloadsFolder;
        string savedFileName = $"{fileName}.{fileExtension}";
        string savedFilePath = Path.Combine(savedFileDir, savedFileName);

        email.SaveAs(savedFilePath, saveAsType);

        Log.Information($"Saved email to {savedFilePath}");
        ExportEmail(savedFilePath, exportFormat);

        // Delete the temporary email file. Only the exported verison is needed
        // A files folder is also generated containing attachments and other files;
        // this must also be removed.
        string filesFolder = Path.Combine(savedFileDir, $"{fileName}_files");
        if (Directory.Exists(filesFolder))
            Directory.Delete(filesFolder, recursive: true);
        else
            Log.Warning($"Files folder '{filesFolder}' does not exist even though it is usually generated");

        if (!File.Exists(savedFilePath)) {
            Log.Fatal($"File '{savedFilePath}' was expected to be generated but does not exist");
            Environment.Exit(1);
        }
        File.Delete(savedFilePath);
    }

    private static void ExportEmail(string directory, Word.WdSaveFormat exportFormat) {
        // Using Word Interop to save the file
        Word.Application word = new();
        Word.Document document = word.Documents.Open(directory);

        string fileName = Path.GetFileNameWithoutExtension(directory)!;
        string fileExtension = GetFileExtensionWithoutPrefix(ExportFormat, "wdFormat");

        string exportedFileName = $"{fileName}.{fileExtension}";
        string exportedFilePath = 
            Path.Combine(Path.GetDirectoryName(directory)!, exportedFileName);

        Log.Information($"Exporting file to path '{exportedFilePath}'");
        document.SaveAs2(exportedFilePath, exportFormat);
        document.Close();
        word.Quit();
    }

    private static string GetFileExtensionWithoutPrefix(Enum extensionValue, string prefix) =>
        extensionValue.ToString()[prefix.Length..].ToLower();
}
