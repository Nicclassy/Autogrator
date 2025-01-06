using DotNetEnv;

namespace Autogrator;

public static class Credentials {
    static Credentials() {
        Env.Load(EnvPath());
    }
    public static class Outlook {
        public static readonly string Email = EnvironmentVariable("AG_EMAIL");
        public static readonly string Password = EnvironmentVariable("AG_PASSWORD");
    }

    private static string EnvPath() {
        string path = ".env";
        string dir = Directory.GetCurrentDirectory();
        string? file = Path.GetFileName(path);
        path = Path.Combine(dir, file);

        while (!File.Exists(path)) {
            var parent =
                Directory.GetParent(dir) ?? throw new ApplicationException(".env not found");
            dir = parent.FullName;
            path = Path.Combine(dir, file);
        }

        return path;
    }

    private static string EnvironmentVariable(string name) {
        string? value = Environment.GetEnvironmentVariable(name);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"Environment variable ${name} not found.");
        return value;
    }
}
