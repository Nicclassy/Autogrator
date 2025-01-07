using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

using Autogrator.Utilities;

namespace Autogrator.OutlookAutomation;

public static class OutlookEmailExporter {

    private const Outlook.OlSaveAsType SaveAsType = Outlook.OlSaveAsType.olHTML;
    private const Word.WdSaveFormat ExportFormat = Word.WdSaveFormat.wdFormatPDF;
    public static void SaveAndExportEmail(
        Outlook.MailItem email,
        string? dir = null,
        Outlook.OlSaveAsType saveAsType = SaveAsType,
        Word.WdSaveFormat exportFormat = ExportFormat
    ) {
        string path =
            Path.Combine(dir ?? Credentials.Outlook.DownloadsFolder, email.EntryID.ToString());
        email.SaveAs(path, saveAsType);
        ExportEmail(path, exportFormat);
    }

    private static void ExportEmail(string path, Word.WdSaveFormat exportFormat) {
        // Using Word Interop to save the file
        Word.Application word = new();
        Word.Document document = word.Documents.Open(path);

        string pathDir = Path.GetDirectoryName(path)!;
        string fileName = Path.GetFileNameWithoutExtension(path)!;
        string fileExtension = ExportFormatFileExtension(ExportFormat);

        string exportedFileName = $"{fileName}.{fileExtension}";
        string exportedFilePath = Path.Combine(pathDir, exportedFileName);

        document.SaveAs2(exportedFilePath, exportFormat);
        document.Close();
        word.Quit();
    }

    private static string ExportFormatFileExtension(Word.WdSaveFormat saveFormat) {
        // Assumes the correctness of the enum values of Word.WdSaveFormat, therefore
        // it may not be correct if the enum value does not reflect the extension accurately.
        const string wdFormat = "wdFormat";
        return saveFormat.ToString()[(wdFormat.Length + 1)..].ToLower();
    }
}
