using Autogrator.Extensions;

namespace Autogrator.Utilities;

public static class Directories {
    public static readonly string DownloadsFolder = Path.Combine("USERPROFILE".Env(), "Downloads");
    public static readonly string OutlookExecutable = "AG_OUTLOOK_EXE".Env();
}

public static class Sites {
    public static readonly string Hostname = "AG_SHAREPOINT_HOSTNAME".Env();
    public static readonly string DefaultSitePath = "AG_SHAREPOINT_DEFAULT_SITE_PATH".Env();
    public static readonly string SharePointURL = "AG_SHAREPOINT_URL".Env();
}

public static class Credentials {
    public static class Outlook {
        public static readonly string Email = "AG_OUTLOOK_EMAIL".Env();
        public static readonly string Password = "AG_OUTLOOK_PASSWORD".Env();
    }
}

public static class AutogratorApplication {
    public static readonly string ClientID = "AG_APPLICATION_CLIENT_ID".Env();
    public static readonly string TenantID = "AG_APPLICATION_TENANT_ID".Env();
    public static readonly string ClientSecret = "AG_APPLICATION_CLIENT_SECRET".Env();
    public static readonly string Scope = "AG_APPLICATION_SCOPE".Env();
    public static readonly string AccessToken = "AG_APPLICATION_ACCESS_TOKEN".Env();
}
