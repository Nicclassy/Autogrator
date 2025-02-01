using DotNetEnv;
using Serilog;

namespace Autogrator.Extensions;

public static class EnvironmentVariableExtensions {
    static EnvironmentVariableExtensions() => Env.Load(FindPathUpwards(".env"));

    public static string EnvVariable(this string name, bool allowEmpty = false) {
        string? value = Environment.GetEnvironmentVariable(name);
        if (value is null || (!allowEmpty && string.IsNullOrWhiteSpace(value))) {
            Log.Fatal("Environment variable '{name}' not found.", name);
            Environment.Exit(1);
        }
        return value;
    }

    private static string FindPathUpwards(string originalPath) {
        if (File.Exists(originalPath))
            return originalPath;

        string dir = Directory.GetCurrentDirectory();
        string? file = Path.GetFileName(originalPath);
        string path = Path.Combine(dir, file);

        while (!File.Exists(path)) {
            DirectoryInfo parent =
                Directory.GetParent(dir)
                ?? throw new ArgumentException($"{originalPath} not found");
            dir = parent.FullName;
            path = Path.Combine(dir, file);
        }

        return path;
    }
}