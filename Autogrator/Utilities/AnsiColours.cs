using System.Globalization;

namespace Autogrator.Utilities;

public interface IAnsiSequence {
    string ToString();
}

public readonly record struct AnsiColour(int Code) : IAnsiSequence {
    public override string ToString() => $"\u001b[{Code}m";

    public static implicit operator string(AnsiColour colour) => colour.ToString();

    public AnsiCombinedColours Combine(params AnsiColour[] colours) =>
        new(new[] { Code }.Concat(colours.Select(colour => colour.Code)));
}

public readonly record struct AnsiRgbColour(int R, int G, int B, bool Background = false) : IAnsiSequence {
    public override string ToString() => $"\u001b[{(Background ? 48 : 38)};2;{R};{G};{B}m";

    public static implicit operator string(AnsiRgbColour colour) => colour.ToString();
}

public readonly record struct AnsiHexColour(string Hex, bool Background = false) : IAnsiSequence {
    public override string ToString() {
        const int mask = 0xFF;

        int value = int.Parse(Hex.TrimStart('#'), NumberStyles.HexNumber);
        int r = (value >> 16) & mask;
        int g = (value >> 8) & mask;
        int b = value & mask;
        return $"\u001b[{(Background ? 48 : 38)};2;{r};{g};{b}m";
    }

    public static implicit operator string(AnsiHexColour colour) => colour.ToString();
}

public readonly record struct AnsiCombinedColours(IEnumerable<int> Codes): IAnsiSequence {
    public override string ToString() => $"\u001b[{string.Join(';', Codes)}m";

    public static implicit operator string(AnsiCombinedColours colours) => colours.ToString();
}

public static class AnsiColours {
    public static readonly AnsiColour Reset = new(0);
    public static readonly AnsiColour Black = new(30);
    public static readonly AnsiColour Red = new(31);
    public static readonly AnsiColour Green = new(32);
    public static readonly AnsiColour Yellow = new(33);
    public static readonly AnsiColour Blue = new(34);
    public static readonly AnsiColour Magenta = new(35);
    public static readonly AnsiColour Cyan = new(36);
    public static readonly AnsiColour White = new(37);

    public static readonly AnsiColour BgBlack = new(40);
    public static readonly AnsiColour BgRed = new(41);
    public static readonly AnsiColour BgGreen = new(42);
    public static readonly AnsiColour BgYellow = new(43);
    public static readonly AnsiColour BgBlue = new(44);
    public static readonly AnsiColour BgMagenta = new(45);
    public static readonly AnsiColour BgCyan = new(46);
    public static readonly AnsiColour BgWhite = new(47);

    public static readonly AnsiColour BrightBlack = new(90);
    public static readonly AnsiColour BrightRed = new(91);
    public static readonly AnsiColour BrightGreen = new(92);
    public static readonly AnsiColour BrightYellow = new(93);
    public static readonly AnsiColour BrightBlue = new(94);
    public static readonly AnsiColour BrightMagenta = new(95);
    public static readonly AnsiColour BrightCyan = new(96);
    public static readonly AnsiColour BrightWhite = new(97);

    public static readonly AnsiColour BgBrightBlack = new(100);
    public static readonly AnsiColour BgBrightRed = new(101);
    public static readonly AnsiColour BgBrightGreen = new(102);
    public static readonly AnsiColour BgBrightYellow = new(103);
    public static readonly AnsiColour BgBrightBlue = new(104);
    public static readonly AnsiColour BgBrightMagenta = new(105);
    public static readonly AnsiColour BgBrightCyan = new(106);
    public static readonly AnsiColour BgBrightWhite = new(107);
}