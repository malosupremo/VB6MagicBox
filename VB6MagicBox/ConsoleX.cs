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

        public static void WriteLineError(string message)
        {
            WriteLineColor(message, ConsoleColor.Red);
        }
        public static void WriteLineWarning(string message)
        {
            WriteLineColor(message, ConsoleColor.Yellow);
        }
        public static void WriteLineSuccess(string message)
        {
            WriteLineColor(message, ConsoleColor.Green);
        }
        public static void WriteLineInfo(string message)
        {
            WriteLineColor(message, ConsoleColor.Cyan);
        }
    }
}
