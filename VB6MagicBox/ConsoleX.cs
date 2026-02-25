namespace VB6MagicBox;

public static class ConsoleX
{
    public static void WriteLineColor(this string message, ConsoleColor color)
    {
        WriteColor(message + Environment.NewLine, color);
    }

    public static void WriteColor(this string message, ConsoleColor color)
    {
        var previous = Console.ForegroundColor;
        Console.ForegroundColor = color;
        Console.Write(message);
        Console.ForegroundColor = previous;
    }
}
