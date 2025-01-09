using Autogrator.Extensions;

namespace Autogrator.Utilities;

public static class Directories {
    public static readonly string DownloadsFolder = Path.Combine("USERPROFILE".EnvVariable(), "Downloads");
    public static readonly string OutlookExecutable = "AG_OUTLOOK_EXE".EnvVariable();
}

public static class Sites {
    public static readonly string Hostname = "AG_SHAREPOINT_HOSTNAME".EnvVariable();
    public static readonly string DefaultSitePath = "AG_SHAREPOINT_DEFAULT_SITE_PATH".EnvVariable();
    public static readonly string SharePointURL = "AG_SHAREPOINT_URL".EnvVariable();
}

public static class Credentials {
    public static class Outlook {
        public static readonly string Email = "AG_OUTLOOK_EMAIL".EnvVariable();
        public static readonly string Password = "AG_OUTLOOK_PASSWORD".EnvVariable();

        public static class Storage {
            public static readonly string Email = "AG_OUTLOOK_STORAGE_EMAIL".EnvVariable();
            public static readonly string Password = "AG_OUTLOOK_STORAGE_PASSWORD".EnvVariable(allowEmpty: true);
        }
    }
}

public static class AutogratorApplication {
    public static readonly string ClientID = "AG_APPLICATION_CLIENT_ID".EnvVariable();
    public static readonly string TenantID = "AG_APPLICATION_TENANT_ID".EnvVariable();
    public static readonly string ClientSecret = "AG_APPLICATION_CLIENT_SECRET".EnvVariable();
    public static readonly string Scope = "AG_APPLICATION_SCOPE".EnvVariable();
    public static readonly string AccessToken = "AG_APPLICATION_ACCESS_TOKEN".EnvVariable();
}
