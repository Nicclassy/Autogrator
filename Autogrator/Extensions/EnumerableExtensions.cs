using Autogrator.Utilities;

namespace Autogrator.Extensions;

public static class EnumerableExtensions {
    public static void Print<T>(
        this IEnumerable<T> source, 
        Func<T, string>? formatter = null,
        string delimiter = ", ",
        IAnsiSequence? ansi = null
    ) where T: notnull {
        string formatValue(T value) => $"\"{formatter?.Invoke(value) ?? value.ToString()}\"";

        if (!source.Any()) {
            Console.WriteLine();
            return;
        }

        if (ansi is not null)
            Console.Write(ansi);
        Console.Write('[');
        Console.Write(formatValue(source.First()));

        foreach (T value in source.Skip(1)) {
            Console.Write(delimiter);
            Console.Write(formatValue(value));
        }

        Console.Write(']');
        if (ansi is not null)
            Console.Write(AnsiColours.Reset);

        Console.WriteLine();
    }

    public static void ForEachWriteLine<T>(
        this IEnumerable<T> source, 
        Func<T, string>? formatter = null
    ) where T : notnull {
        foreach (T value in source)
            Console.WriteLine(formatter is not null ? formatter(value) : value);
    }
}