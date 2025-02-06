using Autogrator.Utilities;

namespace Autogrator.Extensions;

public static class StringExtensions {
    public static (string, string) RightSplitOnce(this string value, char separator) {
        int index = value.LastIndexOf(separator);
        if (index == -1)
            throw new ArgumentException($"Separator '{separator}' was not found in {value}");
        return (value[..index], value[(index + 1)..]);
    }

    public static string? NullIfWhiteSpace(this string value) => string.IsNullOrWhiteSpace(value) ? null : value;

    public static string FileNameWithSuffix(this string value, string suffix) {
        string filename = Path.GetFileNameWithoutExtension(value);
        string extension = Path.GetExtension(value);
        return $"{filename}{suffix}{extension}";
    } 
}

public static class StringAnsiExtensions {
    public static string Colourise(this string value, IAnsiSequence ansi) => $"{ansi}{value}{AnsiColours.Reset}";

    public static string Red(this string value) => value.Colourise(AnsiColours.Red);

    public static string Blue(this string value) => value.Colourise(AnsiColours.Blue);

    public static string Green(this string value) => value.Colourise(AnsiColours.Green);

    public static string Yellow(this string value) => value.Colourise(AnsiColours.Yellow);

    public static string Cyan(this string value) => value.Colourise(AnsiColours.Cyan);

    public static string Magenta(this string value) => value.Colourise(AnsiColours.Magenta);
}