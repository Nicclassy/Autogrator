using Autogrator.Extensions;

namespace Autogrator.Utilities;

public static class Directories {
    public static readonly string DownloadsFolder = Path.Combine("USERPROFILE".EnvVariable(), "Downloads");
}

public static class GraphAPI {
    public static readonly string URL = "https://graph.microsoft.com/v1.0";
}

public static class SharePoint {
    public static readonly string UploadSitePath = "AG_SHAREPOINT_UPLOAD_SITE_PATH".EnvVariable();
    public static readonly string UploadDriveName = "AG_SHAREPOINT_UPLOAD_DRIVE_NAME".EnvVariable();
    public static readonly string Hostname = "AG_SHAREPOINT_HOSTNAME".EnvVariable();
}

public static class Credentials {
    public static class Outlook {
        public static readonly string Profile = "AG_OUTLOOK_PROFILE".EnvVariable(allowEmpty: true);
        public static readonly string Email = "AG_OUTLOOK_PROFILE".EnvVariable(allowEmpty: true);
        public static readonly string Password = "AG_OUTLOOK_PASSWORD".EnvVariable(allowEmpty: true);
        public static readonly string AllowedRecipients = "AG_OUTLOOK_ALLOWED_RECIPIENTS".EnvVariable(allowEmpty: true);
    }
}

public static class AllowedSendersFile {
    public static readonly string Name = "AG_ALLOWED_SENDERS_FILENAME".EnvVariable(allowEmpty: true);
    public static readonly string DownloadDirectory = "AG_ALLOWED_SENDERS_DOWNLOAD_DIRECTORY".EnvVariable(allowEmpty: true);
    public static readonly string DownloadDestination = Directories.DownloadsFolder;
    public static readonly string SitePath = "AG_ALLOWED_SENDERS_SITE_PATH".EnvVariable(allowEmpty: true);
    public static readonly string DriveName = "AG_ALLOWED_SENDERS_DRIVE_NAME".EnvVariable(allowEmpty: true);
}

public static class ApplicationRegistration {
    public static readonly string ClientID = "AG_APPLICATION_CLIENT_ID".EnvVariable();
    public static readonly string TenantID = "AG_APPLICATION_TENANT_ID".EnvVariable();
    public static readonly string ClientSecret = "AG_APPLICATION_CLIENT_SECRET".EnvVariable();
    public static readonly string Scope = "https://graph.microsoft.com/.default";
}

public static class NotificationEmail {
    public static readonly string SenderEmailAddress = "AG_NOTIFICATION_EMAIL_SENDER_ADDRESS".EnvVariable(allowEmpty: true);
    public static readonly string RecipientEmailAddress = "AG_NOTIFICATION_EMAIL_RECIPIENT_ADDRESS".EnvVariable(allowEmpty: true);
    public static readonly string ContentPath = "AG_NOTIFICATION_EMAIL_CONTENT_PATH".EnvVariable(allowEmpty: true);
}
