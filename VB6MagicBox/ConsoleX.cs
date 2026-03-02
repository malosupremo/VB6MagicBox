namespace VB6MagicBox;

public static class ConsoleExtensions
{
    extension(Console)
    {
        public static void WriteLineColor(string message, ConsoleColor color)
        {
            WriteColor(message + Environment.NewLine, color);
        }

        public static void WriteColor(string message, ConsoleColor color)
        {
            var previous = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.Write(message);
            Console.ForegroundColor = previous;
        }
    }
}
