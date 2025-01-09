using DotNetEnv;

namespace Autogrator.Extensions;

public static class EnvironmentVariableExtensions {
    static EnvironmentVariableExtensions() {
        DotNetEnv.Env.Load(FindPathUpwards(".env"));
    }

    public static string EnvVariable(this string name, bool allowEmpty = false) {
        string? value = Environment.GetEnvironmentVariable(name);
        if (value is null || (!allowEmpty && string.IsNullOrWhiteSpace(value)))
            throw new EnvVariableNotFoundException("Environment variable not found.", name);
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