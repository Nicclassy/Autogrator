using DotNetEnv;

namespace Autogrator.Utilities;

public static class Credentials {
    static Credentials() {
        Env.Load(FindPathUpwards(".env"));
    }

    public static class Outlook {
        public static readonly string Email = EnvironmentVariable("AG_OUTLOOK_EMAIL");
        public static readonly string Password = EnvironmentVariable("AG_OUTLOOK_PASSWORD");
    }

    public static class SharePoint {
        public static readonly string SiteURL = EnvironmentVariable("AG_SHAREPOINT_URL");
        public static readonly string ClientID = EnvironmentVariable("AG_SHAREPOINT_CLIENT_ID");
        public static readonly string TenantID = EnvironmentVariable("AG_SHAREPOINT_TENANT_ID");
        public static readonly string ClientSecret = EnvironmentVariable("AG_SHAREPOINT_CLIENT_SECRET");
    }

    internal static string EnvironmentVariable(string name) {
        string? value = Environment.GetEnvironmentVariable(name);
        if (string.IsNullOrWhiteSpace(value))
            throw new EnvVariableNotFoundException($"Environment variable ${name} not found.", name);
        return value;
    }

    private static string FindPathUpwards(string originalPath) {
        if (File.Exists(originalPath))
            return originalPath;

        string dir = Directory.GetCurrentDirectory();
        string? file = Path.GetFileName(originalPath);
        string path = Path.Combine(dir, file);

        while (!File.Exists(path)) {
            var parent =
                Directory.GetParent(dir) 
                ?? throw new ArgumentException($"{originalPath} not found");
            dir = parent.FullName;
            path = Path.Combine(dir, file);
        }

        return path;
    }
}
