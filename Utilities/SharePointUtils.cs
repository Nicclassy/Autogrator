using System.Diagnostics;

namespace Autogrator.Utilities;

public static class SharePointUtils {
    public static string FormatPath(string? itemPath) {
        if (string.IsNullOrWhiteSpace(itemPath)) return "root";

        Debug.Assert(itemPath[0] == '/', "Path must start with '/'");
        return $"root:{itemPath}:";
    }
}