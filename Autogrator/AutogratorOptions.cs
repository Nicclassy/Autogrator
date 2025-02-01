namespace Autogrator;

public sealed class AutogratorOptions {
    public string CopiedFileSuffix { get; set; } = " (Copy)";
    public string EmailsFolderName { get; set; } = "Emails";
    public bool SendExceptionNotificationEmails { get; set; } = true;
    public bool OverwriteDownloads { get; set; } = false;
    public bool EnableRequestLogging { get; set; } = true;
    public bool UseSeparateRequestLogger { get; set; } = true;
    public bool UseDefaultLoggingConfiguration { get; set; } = true;
}