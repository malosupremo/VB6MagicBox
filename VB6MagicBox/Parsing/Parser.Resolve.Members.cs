using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    // ---------------------------------------------------------
    // REGEX COMPILATE PER HOT-PATH
    // ---------------------------------------------------------

    private static readonly Regex ReWithDotReplacement =
        new(@"(?<![\w)])\.(\s*[A-Za-z_]\w*(?:\([^)]*\))?)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReTokens =
        new(@"\b[A-Za-z_]\w*\b", RegexOptions.Compiled);

    // ---------------------------------------------------------
    // HELPER: normalizzazione nomi tipo modulo
    // ---------------------------------------------------------

    private static string NormalizeModuleTypeName(string typeName)
    {
        if (string.IsNullOrWhiteSpace(typeName))
            return string.Empty;

        var normalized = typeName;
        var parenIndex = normalized.IndexOf('(');
        if (parenIndex >= 0)
            normalized = normalized.Substring(0, parenIndex);

        if (normalized.Contains('.'))
            normalized = normalized.Split('.').Last();

        return normalized.Trim();
    }

    // ---------------------------------------------------------
    // HELPER: property bounds e overlap pruning
    // ---------------------------------------------------------

    private static void EnsurePropertyBounds(VbProperty prop, string[] fileLines)
    {
        if (prop == null || fileLines == null || fileLines.Length == 0)
            return;

        bool IsPropertyDeclarationLine(string line)
        {
            var pattern = $@"^\s*(Public|Private|Friend)?\s*(Static\s+)?Property\s+{Regex.Escape(prop.Kind ?? string.Empty)}\s+{Regex.Escape(prop.Name ?? string.Empty)}\b";
            return Regex.IsMatch(line, pattern, RegexOptions.IgnoreCase);
        }

        if (prop.StartLine <= 0 || prop.StartLine > fileLines.Length ||
            !IsPropertyDeclarationLine(fileLines[prop.StartLine - 1]))
        {
            for (int i = 0; i < fileLines.Length; i++)
            {
                if (IsPropertyDeclarationLine(fileLines[i]))
                {
                    prop.StartLine = i + 1;
                    break;
                }
            }
        }

        if (prop.EndLine <= prop.StartLine || prop.EndLine > fileLines.Length)
        {
            var startIndex = Math.Max(0, prop.StartLine - 1);
            for (int i = startIndex; i < fileLines.Length; i++)
            {
                if (IsProcedureEndLine(fileLines[i], "Property"))
                {
                    prop.EndLine = i + 1;
                    break;
                }
            }
        }
    }

    private static (int Start, int End)? TryGetPropertyRange(VbProperty prop, string[] fileLines)
    {
        if (prop == null || fileLines == null || fileLines.Length == 0)
            return null;

        var pattern = $@"^\s*(Public|Private|Friend)?\s*(Static\s+)?Property\s+{Regex.Escape(prop.Kind ?? string.Empty)}\s+{Regex.Escape(prop.Name ?? string.Empty)}\b";
        for (int i = 0; i < fileLines.Length; i++)
        {
            if (!Regex.IsMatch(fileLines[i], pattern, RegexOptions.IgnoreCase))
                continue;

            for (int j = i; j < fileLines.Length; j++)
            {
                if (IsProcedureEndLine(fileLines[j], "Property"))
                    return (i + 1, j + 1);
            }

            return (i + 1, fileLines.Length);
        }

        return null;
    }

    private static void PrunePropertyReferenceOverlaps(VbModule mod, string[] fileLines)
    {
        foreach (var prop in mod.Properties)
            EnsurePropertyBounds(prop, fileLines);

        foreach (var prop in mod.Properties)
        {
            var overlapRanges = mod.Properties
                .Where(p => p != prop &&
                            string.Equals(p.Name, prop.Name, StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(p.Kind, prop.Kind, StringComparison.OrdinalIgnoreCase))
                .Select(p => TryGetPropertyRange(p, fileLines))
                .Where(r => r.HasValue)
                .Select(r => r.Value)
                .ToList();

            if (overlapRanges.Count == 0)
                continue;

            foreach (var reference in prop.References)
            {
                for (int i = reference.LineNumbers.Count - 1; i >= 0; i--)
                {
                    var lineNumber = reference.LineNumbers[i];
                    if (overlapRanges.Any(r => lineNumber >= r.Start && lineNumber <= r.End))
                    {
                        reference.LineNumbers.RemoveAt(i);
                        if (i < reference.OccurrenceIndexes.Count)
                            reference.OccurrenceIndexes.RemoveAt(i);
                        if (i < reference.StartChars.Count)
                            reference.StartChars.RemoveAt(i);
                    }
                }
            }
        }
    }

    // ---------------------------------------------------------
    // HELPER: enumerazione token e catene
    // ---------------------------------------------------------

    private static IEnumerable<(string Text, int Index)> EnumerateParenContents(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        var starts = new Stack<int>();
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '(')
            {
                starts.Push(i + 1);
                continue;
            }

            if (line[i] != ')' || starts.Count == 0)
                continue;

            var start = starts.Pop();
            if (start > i)
                continue;

            yield return (line.Substring(start, i - start), start);
        }
    }

    private static IEnumerable<(string Text, int Index)> EnumerateDotChains(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        int i = 0;
        while (i < line.Length)
        {
            if (!IsIdentifierStart(line[i]) || (i > 0 && IsIdentifierChar(line[i - 1])))
            {
                i++;
                continue;
            }

            int start = i;
            i++;
            while (i < line.Length && IsIdentifierChar(line[i]))
                i++;

            int end = i;
            bool hasDot = false;

            while (true)
            {
                int index = SkipWhitespace(line, end);
                int afterParen = SkipOptionalParentheses(line, index);
                if (afterParen != index)
                    end = afterParen;

                index = SkipWhitespace(line, end);
                if (index >= line.Length || line[index] != '.')
                    break;

                hasDot = true;
                index++;
                index = SkipWhitespace(line, index);
                if (index >= line.Length || !IsIdentifierStart(line[index]))
                    break;

                end = index + 1;
                while (end < line.Length && IsIdentifierChar(line[end]))
                    end++;
            }

            if (hasDot && end > start)
                yield return (line.Substring(start, end - start), start);

            i = Math.Max(i, end + 1);
        }
    }

    private static void AddChainMatch(List<(string Text, int Index)> chainMatches, HashSet<string> chainMatchSet, string text, int index)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        var key = $"{index}:{text}";
        if (chainMatchSet.Add(key))
            chainMatches.Add((text, index));
    }

    private static IEnumerable<(string Token, int Index)> EnumerateTokens(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        for (int i = 0; i < line.Length; i++)
        {
            if (!IsIdentifierStart(line[i]))
                continue;

            if (i > 0 && IsIdentifierChar(line[i - 1]))
                continue;

            int start = i;
            i++;
            while (i < line.Length && IsIdentifierChar(line[i]))
                i++;

            yield return (line.Substring(start, i - start), start);
            i--;
        }
    }

    private static IEnumerable<(string Left, string Right)> EnumerateQualifiedTokens(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        for (int i = 0; i < line.Length; i++)
        {
            if (!IsIdentifierStart(line[i]))
                continue;

            if (i > 0 && IsIdentifierChar(line[i - 1]))
                continue;

            int leftStart = i;
            i++;
            while (i < line.Length && IsIdentifierChar(line[i]))
                i++;

            var left = line.Substring(leftStart, i - leftStart);

            int dotIndex = i;
            while (dotIndex < line.Length && char.IsWhiteSpace(line[dotIndex]))
                dotIndex++;

            if (dotIndex >= line.Length || line[dotIndex] != '.')
            {
                i--;
                continue;
            }

            dotIndex++;
            while (dotIndex < line.Length && char.IsWhiteSpace(line[dotIndex]))
                dotIndex++;

            if (dotIndex >= line.Length || !IsIdentifierStart(line[dotIndex]))
            {
                i--;
                continue;
            }

            int rightStart = dotIndex;
            dotIndex++;
            while (dotIndex < line.Length && IsIdentifierChar(line[dotIndex]))
                dotIndex++;

            var right = line.Substring(rightStart, dotIndex - rightStart);
            yield return (left, right);
            i = dotIndex - 1;
        }
    }

    // ---------------------------------------------------------
    // HELPER: caratteri e boundary
    // ---------------------------------------------------------

    private static bool IsIdentifierStart(char value)
        => char.IsLetter(value) || value == '_';

    private static int SkipWhitespace(string line, int index)
    {
        while (index < line.Length && char.IsWhiteSpace(line[index]))
            index++;

        return index;
    }

    private static int SkipOptionalParentheses(string line, int index)
    {
        if (index >= line.Length || line[index] != '(')
            return index;

        int depth = 0;
        for (int i = index; i < line.Length; i++)
        {
            if (line[i] == '(')
                depth++;
            else if (line[i] == ')')
            {
                depth--;
                if (depth == 0)
                    return i + 1;
            }
        }

        return line.Length;
    }

    private static bool IsWordBoundary(string line, int index, int length)
    {
        bool startOk = index == 0 || !IsIdentifierChar(line[index - 1]);
        int endIndex = index + length;
        bool endOk = endIndex >= line.Length || !IsIdentifierChar(line[endIndex]);
        return startOk && endOk;
    }

    private static bool IsIdentifierChar(char value)
        => char.IsLetterOrDigit(value) || value == '_';

    private static bool IsMemberAccessToken(string line, int tokenIndex)
    {
        if (tokenIndex <= 0)
            return false;

        var index = tokenIndex - 1;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        return index >= 0 && line[index] == '.';
    }

    private static int GetStartCharFromRaw(string rawLine, string token, int fallbackIndex)
    {
        if (string.IsNullOrEmpty(rawLine) || string.IsNullOrEmpty(token))
            return fallbackIndex;

        var startIndex = Math.Max(0, fallbackIndex);
        if (startIndex > rawLine.Length)
            startIndex = rawLine.Length;

        var index = rawLine.IndexOf(token, startIndex, StringComparison.OrdinalIgnoreCase);
        if (index < 0 && startIndex > 0)
            index = rawLine.IndexOf(token, StringComparison.OrdinalIgnoreCase);

        return index >= 0 ? index : fallbackIndex;
    }

    private static int GetTokenStartChar(string rawLine, string token, int occurrenceIndex, int fallbackIndex)
    {
        if (string.IsNullOrEmpty(rawLine) || string.IsNullOrEmpty(token))
            return fallbackIndex;

        var matches = Regex.Matches(rawLine, $@"\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase);
        if (matches.Count > 0)
        {
            if (occurrenceIndex > 0 && occurrenceIndex <= matches.Count)
                return matches[occurrenceIndex - 1].Index;

            return matches[0].Index;
        }

        return GetStartCharFromRaw(rawLine, token, fallbackIndex);
    }

    // ---------------------------------------------------------
    // HELPER: occurrence index e utility
    // ---------------------------------------------------------

    private static int GetOccurrenceIndex(string line, string token, int tokenIndex, int currentLineNumber = 0)
    {
        bool isDebug = false;

        if (tokenIndex < 0)
            return -1;

        var matches = Regex.Matches(line, $@"\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase);

        if (isDebug)
        {
            Console.WriteLine($"[DEBUG GetOccurrenceIndex] Line {currentLineNumber}, Token={token}, TokenIndex={tokenIndex}");
            Console.WriteLine($"[DEBUG]   Line: {line}");
            Console.WriteLine($"[DEBUG]   Total matches: {matches.Count}");
            for (int j = 0; j < matches.Count; j++)
                Console.WriteLine($"[DEBUG]     Match {j + 1} at index {matches[j].Index}: '{matches[j].Value}'");
        }

        for (int i = 0; i < matches.Count; i++)
        {
            if (matches[i].Index == tokenIndex)
            {
                if (isDebug)
                    Console.WriteLine($"[DEBUG]   ? Returning occurrence {i + 1}");
                return i + 1; // 1-based occurrence index
            }
        }

        if (isDebug)
            Console.WriteLine($"[DEBUG]   ? Token not found at specified index, returning -1");

        return -1;
    }

    private static bool TryUnwrapFunctionChain(string chainText, int chainIndex, out string unwrappedChain, out int unwrappedIndex)
    {
        unwrappedChain = null;
        unwrappedIndex = chainIndex;

        if (string.IsNullOrEmpty(chainText))
            return false;

        var parenIndex = chainText.IndexOf('(');
        if (parenIndex <= 0)
            return false;

        int depth = 0;
        int closeIndex = -1;
        for (int i = parenIndex; i < chainText.Length; i++)
        {
            if (chainText[i] == '(')
                depth++;
            else if (chainText[i] == ')')
            {
                depth--;
                if (depth == 0)
                {
                    closeIndex = i;
                    break;
                }
            }
        }

        if (closeIndex >= 0 && closeIndex + 1 < chainText.Length && chainText[closeIndex + 1] == '.')
            return false;

        var prefix = chainText.Substring(0, parenIndex).Trim();
        if (prefix.Contains('.'))
            return false;

        unwrappedChain = chainText.Substring(parenIndex + 1).TrimStart();
        unwrappedIndex = chainIndex + parenIndex + 1;
        return !string.IsNullOrEmpty(unwrappedChain);
    }

    private static bool MatchesName(string name, string conventionalName, string token)
    {
        return string.Equals(name, token, StringComparison.OrdinalIgnoreCase) ||
               string.Equals(conventionalName, token, StringComparison.OrdinalIgnoreCase);
    }

    private static string[] SplitChainParts(string chainText)
    {
        if (string.IsNullOrEmpty(chainText))
            return Array.Empty<string>();

        var parts = new List<string>();
        int depth = 0;
        int start = 0;

        for (int i = 0; i < chainText.Length; i++)
        {
            if (chainText[i] == '(')
                depth++;
            else if (chainText[i] == ')')
                depth = Math.Max(0, depth - 1);
            else if (chainText[i] == '.' && depth == 0)
            {
                var part = chainText.Substring(start, i - start).Trim();
                if (part.Length > 0)
                    parts.Add(part);
                start = i + 1;
            }
        }

        if (start <= chainText.Length)
        {
            var tail = chainText.Substring(start).Trim();
            if (tail.Length > 0)
                parts.Add(tail);
        }

        return parts.ToArray();
    }
}
