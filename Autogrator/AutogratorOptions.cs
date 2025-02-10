namespace Autogrator;

public sealed class AutogratorOptions {
    public int ExecutionInterval { get; set; } = 1000;
    public string CopiedFileSuffix { get; set; } = " (Copy)";
    public string EmailsFolderName { get; set; } = "Emails";
    public string LogFileName { get; init; } = "log.txt";
    public string LoggingFolder { get; init; } = "logs";
    public bool ReviewSentEmails { get; set; } = false;
    public bool OverwriteDownloads { get; set; } = false;
    public bool LogRejectedEmails { get; set; } = true;
    public bool EnableRequestLogging { get; set; } = true;
    public bool LogGraphJSONResponses { get; set; } = false;
    public bool UseSeparateRequestLogger { get; set; } = true;
    public bool AutoDownloadAllowedSenders { get; set; } = true;
    public bool UseDefaultLoggingConfiguration { get; set; } = true;
    public bool SendExceptionNotificationEmails { get; set; } = true;
}