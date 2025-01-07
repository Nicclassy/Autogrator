using Autogrator.Extensions;

namespace Autogrator.Utilities;

public static class Directories {
    public static readonly string DownloadsFolder = Path.Combine("USERPROFILE".Env(), "Downloads");
    public static readonly string OutlookExecutable = "AG_OUTLOOK_EXE".Env();
}

public static class Credentials {

    public static class Outlook {
        public static readonly string Email = "AG_OUTLOOK_EMAIL".Env();
        public static readonly string Password = "AG_OUTLOOK_PASSWORD".Env();
    }

    public static class SharePoint {
        public static readonly string SiteURL = "AG_SHAREPOINT_URL".Env();
        public static readonly string ClientID = "AG_SHAREPOINT_CLIENT_ID".Env();
        public static readonly string TenantID = "AG_SHAREPOINT_TENANT_ID".Env();
        public static readonly string ClientSecret = "AG_SHAREPOINT_CLIENT_SECRET".Env();
    }
}
