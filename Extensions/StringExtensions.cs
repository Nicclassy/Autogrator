using System;
using Autogrator.Utilities;

namespace Autogrator.Extensions;

public static class StringExtensions {
    public static string Colourise(this string text, string ansiColour) => $"{ansiColour}{text}{AnsiColours.Reset}";
}