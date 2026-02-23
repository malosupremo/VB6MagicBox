using System;

namespace VB6MagicBox;

public static class ConsoleOutput
{
    public static void WriteLineColored(this string message, ConsoleColor color)
    {
        var previous = Console.ForegroundColor;
        Console.ForegroundColor = color;
        Console.WriteLine(message);
        Console.ForegroundColor = previous;
    }
}
