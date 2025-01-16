using System;

namespace Autogrator.Extensions;

public static class EnumerableExtensions {
    public static void Print<T>(
        this IEnumerable<T> source, 
        Func<T, string>? formatter = null,
        string delimiter = ", "
    ) where T: notnull {
        string FormatValue(T value) => $"\"{formatter?.Invoke(value) ?? value.ToString()}\"";

        if (!source.Any()) {
            Console.WriteLine();
            return;
        }

        Console.Write('[');
        Console.Write(FormatValue(source.First()));

        foreach (T value in source.Skip(1)) {
            Console.Write(delimiter);
            Console.Write(FormatValue(value));
        }

        Console.Write(']');
        Console.WriteLine();
    }
}