using System.Diagnostics;

namespace Autogrator.Utilities;

public static class SharePointUtils {
    public static string FormatFilePath(string? filepath) {
        if (string.IsNullOrWhiteSpace(filepath)) return "root";

        Debug.Assert(filepath[0] == '/', "Path must start with '/'");
        return $"root:{filepath}:";
    }

    public static string FormatNameForComparison(string name) {
        return name;
    }
}