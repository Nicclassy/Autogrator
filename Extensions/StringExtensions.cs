using Autogrator.Utilities;

namespace Autogrator.Extensions;

public static class StringExtensions {
    public static string Colourise(this string text, string ansiColour) => $"{ansiColour}{text}{AnsiColours.Reset}";

    public static (string, string) RightSplitOnce(this string text, char separator) {
        int index = text.LastIndexOf(separator);
        if (index == -1)
            throw new ArgumentException($"Separator '{separator}' was not found in {text}");
        return (text[..index], text[(index + 1)..]);
    }
}