﻿using Autogrator.Utilities;

namespace Autogrator.Extensions;

public static class StringExtensions {
    public static string Colourise(this string value, string ansiColour) => $"{ansiColour}{value}{AnsiColours.Reset}";

    public static (string, string) RightSplitOnce(this string value, char separator) {
        int index = value.LastIndexOf(separator);
        if (index == -1)
            throw new ArgumentException($"Separator '{separator}' was not found in {value}");
        return (value[..index], value[(index + 1)..]);
    }

    public static string? NullIfWhiteSpace(this string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}