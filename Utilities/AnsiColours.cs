namespace Autogrator.Utilities;

public static class AnsiColours {
    public static readonly string Reset = AnsiColour(0);
    public static readonly string Black = AnsiColour(30);
    public static readonly string Red = AnsiColour(31);
    public static readonly string Green = AnsiColour(32);
    public static readonly string Yellow = AnsiColour(33);
    public static readonly string Blue = AnsiColour(34);
    public static readonly string Magenta = AnsiColour(35);
    public static readonly string Cyan = AnsiColour(36);
    public static readonly string White = AnsiColour(37);

    public static readonly string BgBlack = AnsiColour(40);
    public static readonly string BgRed = AnsiColour(41);
    public static readonly string BgGreen = AnsiColour(42);
    public static readonly string BgYellow = AnsiColour(43);
    public static readonly string BgBlue = AnsiColour(44);
    public static readonly string BgMagenta = AnsiColour(45);
    public static readonly string BgCyan = AnsiColour(46);
    public static readonly string BgWhite = AnsiColour(47);

    public static readonly string BrightBlack = AnsiColour(90);
    public static readonly string BrightRed = AnsiColour(91);
    public static readonly string BrightGreen = AnsiColour(92);
    public static readonly string BrightYellow = AnsiColour(93);
    public static readonly string BrightBlue = AnsiColour(94);
    public static readonly string BrightMagenta = AnsiColour(95);
    public static readonly string BrightCyan = AnsiColour(96);
    public static readonly string BrightWhite = AnsiColour(97);

    public static readonly string BgBrightBlack = AnsiColour(100);
    public static readonly string BgBrightRed = AnsiColour(101);
    public static readonly string BgBrightGreen = AnsiColour(102);
    public static readonly string BgBrightYellow = AnsiColour(103);
    public static readonly string BgBrightBlue = AnsiColour(104);
    public static readonly string BgBrightMagenta = AnsiColour(105);
    public static readonly string BgBrightCyan = AnsiColour(106);
    public static readonly string BgBrightWhite = AnsiColour(107);

    public static string AnsiColour(int code) => $"\u001b[{code}m";

    public static string AnsiColour(int r, int g, int b, bool background = false) =>
        $"\u001b[{(background ? 48 : 38)};2;{r};{g};{b}m";
}