namespace Autogrator;

public sealed class AutogratorOptions {
    public string CopiedFileSuffix { get; set; } = " (Copy)";
    public string EmailsFolderName { get; set; } = "Emails";
    public bool OverwriteDownloads { get; set; } = false;
    public bool UseDefaultLoggingConfiguration { get; set; } = true;
}