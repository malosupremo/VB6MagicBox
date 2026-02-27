using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    // -------------------------
    // PARAMETRI
    // -------------------------

    private static List<VbParameter> ParseParameters(string paramList, int originalLineNumber = 0)
    {
        var result = new List<VbParameter>();
        if (string.IsNullOrWhiteSpace(paramList))
            return result;

        var parts = paramList.Split(',');
        var reParam = new Regex(
            @"^(Optional\s+)?(ByVal|ByRef)?\s*(\w+)([$%&!#@]?)(\([^)]*\))?\s*(As\s+([\w\.\(\)]+))?",
            RegexOptions.IgnoreCase);

        foreach (var p in parts)
        {
            var s = p.Trim();
            if (string.IsNullOrEmpty(s))
                continue;

            var m = reParam.Match(s);
            if (!m.Success)
                continue;

            var typeName = NormalizeTypeName(m.Groups[7].Value);

            result.Add(new VbParameter
            {
                Name = m.Groups[3].Value,
                Passing = string.IsNullOrEmpty(m.Groups[2].Value) ? "ByRef" : m.Groups[2].Value,
                Type = typeName,
                Used = false,
                LineNumber = originalLineNumber
            });
        }

        return result;
    }

    private static string NormalizeTypeName(string typeName)
    {
        if (string.IsNullOrWhiteSpace(typeName))
            return typeName;

        var normalized = typeName.Trim();
        while (normalized.EndsWith(")", StringComparison.Ordinal) &&
               normalized.Count(c => c == '(') < normalized.Count(c => c == ')'))
        {
            normalized = normalized.Substring(0, normalized.Length - 1).TrimEnd();
        }

        return normalized;
    }

    /// <summary>
    /// Aggiunge automaticamente References per i parametri delle Declare Function/Sub
    /// che si estendono su più righe con il carattere di continuazione "_"
    /// </summary>
    private static void AddParameterReferencesForMultilineDeclaration(
        VbProcedure procedure,
        string moduleName,
        string[] originalLines,
        int startLineNumber,
        int[] lineMapping,
        int collapsedIndex)
    {
        if (procedure.Parameters == null || procedure.Parameters.Count == 0)
            return;

        // Trova tutte le righe originali che costituivano questa dichiarazione collapsed
        var originalStartIndex = startLineNumber - 1; // Convert to 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        if (!originalLines[originalStartIndex].TrimEnd().EndsWith("_"))
            return;
        var originalEndIndex = originalStartIndex;

        // Trova l'ultima riga della dichiarazione (seguendo i "_")
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        // Per ogni parametro, cerca in quale riga originale si trova
        foreach (var param in procedure.Parameters)
        {
            for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
            {
                var originalLine = originalLines[lineIdx];
                var noComment = StripInlineComment(originalLine);
                noComment = MaskStringLiterals(noComment);

                // Cerca il nome del parametro in questa riga (word boundary per evitare match parziali)
                var paramPattern = $@"\b{Regex.Escape(param.Name)}\b";
                var match = Regex.Match(noComment, paramPattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var resolvedLineNumber = lineIdx + 1;
                    var startChar = match.Index;
                    var occurrenceIndex = GetOccurrenceIndex(noComment, param.Name, startChar, resolvedLineNumber);
                    param.LineNumber = resolvedLineNumber;

                    // Trovato! Aggiungi una Reference a questa riga specifica
                    // ma solo se non l'ho già segnato
                    param.References.AddLineNumber(moduleName, procedure.Name, resolvedLineNumber, occurrenceIndex, startChar);

                    // Un parametro può apparire solo una volta, quindi esci dal loop
                    break;
                }
            }
        }
    }

    /// <summary>
    /// Aggiorna i LineNumber dei parametri per firme su più righe (con '_'),
    /// senza aggiungere References extra.
    /// </summary>
    private static void FixParameterLineNumbersForMultilineSignature(
        VbProcedure procedure,
        string[] originalLines,
        int startLineNumber)
    {
        if (procedure.Parameters == null || procedure.Parameters.Count == 0)
            return;

        var originalStartIndex = startLineNumber - 1; // 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        // Solo se la firma è realmente multilinea
        if (!originalLines[originalStartIndex].TrimEnd().EndsWith("_"))
            return;

        var originalEndIndex = originalStartIndex;
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        foreach (var param in procedure.Parameters)
        {
            for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
            {
                var originalLine = originalLines[lineIdx];
                var paramPattern = $@"\b{Regex.Escape(param.Name)}\b";
                if (Regex.IsMatch(originalLine, paramPattern, RegexOptions.IgnoreCase))
                {
                    param.LineNumber = lineIdx + 1;
                    break;
                }
            }
        }
    }

    private static void FixParameterLineNumbersForMultilineSignature(
        VbProperty property,
        string[] originalLines,
        int startLineNumber)
    {
        if (property.Parameters == null || property.Parameters.Count == 0)
            return;

        var originalStartIndex = startLineNumber - 1; // 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        if (!originalLines[originalStartIndex].TrimEnd().EndsWith("_"))
            return;

        var originalEndIndex = originalStartIndex;
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        foreach (var param in property.Parameters)
        {
            for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
            {
                var originalLine = originalLines[lineIdx];
                var paramPattern = $@"\b{Regex.Escape(param.Name)}\b";
                if (Regex.IsMatch(originalLine, paramPattern, RegexOptions.IgnoreCase))
                {
                    param.LineNumber = lineIdx + 1;
                    break;
                }
            }
        }
    }

    private static void FixReturnTypeLineNumberForMultilineSignature(
        VbProcedure procedure,
        string[] originalLines,
        int startLineNumber)
    {
        if (procedure == null || string.IsNullOrWhiteSpace(procedure.ReturnType))
            return;

        var returnType = NormalizeTypeName(procedure.ReturnType);
        if (string.IsNullOrWhiteSpace(returnType))
            return;

        var originalStartIndex = startLineNumber - 1; // 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        var originalEndIndex = originalStartIndex;
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
        {
            var originalLine = originalLines[lineIdx];
            var commentIndex = originalLine.IndexOf("'");
            if (commentIndex >= 0)
                originalLine = originalLine.Substring(0, commentIndex);

            if (Regex.IsMatch(originalLine, $@"\bAs\s+{Regex.Escape(returnType)}(?=\s|$)", RegexOptions.IgnoreCase))
            {
                procedure.ReturnTypeLineNumber = lineIdx + 1;
                return;
            }
        }

        procedure.ReturnTypeLineNumber = startLineNumber;
    }

    private static void FixReturnTypeLineNumberForMultilineSignature(
        VbProperty property,
        string[] originalLines,
        int startLineNumber)
    {
        if (property == null || string.IsNullOrWhiteSpace(property.ReturnType))
            return;

        var returnType = NormalizeTypeName(property.ReturnType);
        if (string.IsNullOrWhiteSpace(returnType))
            return;

        var originalStartIndex = startLineNumber - 1; // 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        var originalEndIndex = originalStartIndex;
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
        {
            var originalLine = originalLines[lineIdx];
            var commentIndex = originalLine.IndexOf("'");
            if (commentIndex >= 0)
                originalLine = originalLine.Substring(0, commentIndex);

            if (Regex.IsMatch(originalLine, $@"\bAs\s+{Regex.Escape(returnType)}(?=\s|$)", RegexOptions.IgnoreCase))
            {
                property.ReturnTypeLineNumber = lineIdx + 1;
                return;
            }
        }

        property.ReturnTypeLineNumber = startLineNumber;
    }

    private static void FixParameterTypeLineNumbersForMultilineSignature(
        VbProcedure procedure,
        string[] originalLines,
        int startLineNumber)
    {
        if (procedure.Parameters == null || procedure.Parameters.Count == 0)
            return;

        var originalStartIndex = startLineNumber - 1; // 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        var originalEndIndex = originalStartIndex;
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        foreach (var param in procedure.Parameters)
        {
            if (string.IsNullOrWhiteSpace(param.Type))
            {
                param.TypeLineNumber = param.LineNumber;
                continue;
            }

            var typeName = NormalizeTypeName(param.Type);
            int foundLine = -1;

            for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
            {
                var originalLine = originalLines[lineIdx];
                var commentIndex = originalLine.IndexOf("'");
                if (commentIndex >= 0)
                    originalLine = originalLine.Substring(0, commentIndex);

                if (!Regex.IsMatch(originalLine, $@"\b{Regex.Escape(param.Name)}\b", RegexOptions.IgnoreCase))
                    continue;

                if (Regex.IsMatch(originalLine, $@"\bAs\s+{Regex.Escape(typeName)}(?=\s|\)|$)", RegexOptions.IgnoreCase))
                {
                    foundLine = lineIdx;
                    break;
                }

                if (originalLine.TrimEnd().EndsWith("_") &&
                    Regex.IsMatch(originalLine, @"\bAs\b", RegexOptions.IgnoreCase))
                {
                    for (int nextIdx = lineIdx + 1; nextIdx <= originalEndIndex; nextIdx++)
                    {
                        var nextLine = originalLines[nextIdx];
                        var nextCommentIndex = nextLine.IndexOf("'");
                        if (nextCommentIndex >= 0)
                            nextLine = nextLine.Substring(0, nextCommentIndex);

                        if (Regex.IsMatch(nextLine, $@"\b{Regex.Escape(typeName)}(?=\s|\)|$)", RegexOptions.IgnoreCase))
                        {
                            foundLine = nextIdx;
                            break;
                        }
                    }
                }

                if (foundLine >= 0)
                    break;
            }

            if (foundLine < 0)
            {
                for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
                {
                    var originalLine = originalLines[lineIdx];
                    var commentIndex = originalLine.IndexOf("'");
                    if (commentIndex >= 0)
                        originalLine = originalLine.Substring(0, commentIndex);

                    if (Regex.IsMatch(originalLine, $@"\b{Regex.Escape(typeName)}(?=\s|\)|$)", RegexOptions.IgnoreCase))
                    {
                        foundLine = lineIdx;
                        break;
                    }
                }
            }

            param.TypeLineNumber = foundLine >= 0 ? foundLine + 1 : param.LineNumber;
        }
    }

    private static void FixParameterTypeLineNumbersForMultilineSignature(
        VbProperty property,
        string[] originalLines,
        int startLineNumber)
    {
        if (property.Parameters == null || property.Parameters.Count == 0)
            return;

        var originalStartIndex = startLineNumber - 1; // 0-based
        if (originalStartIndex < 0 || originalStartIndex >= originalLines.Length)
            return;

        var originalEndIndex = originalStartIndex;
        while (originalEndIndex < originalLines.Length - 1)
        {
            var line = originalLines[originalEndIndex].TrimEnd();
            if (!line.EndsWith("_"))
                break;
            originalEndIndex++;
        }

        foreach (var param in property.Parameters)
        {
            if (string.IsNullOrWhiteSpace(param.Type))
            {
                param.TypeLineNumber = param.LineNumber;
                continue;
            }

            var typeName = NormalizeTypeName(param.Type);
            int foundLine = -1;

            for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
            {
                var originalLine = originalLines[lineIdx];
                var commentIndex = originalLine.IndexOf("'");
                if (commentIndex >= 0)
                    originalLine = originalLine.Substring(0, commentIndex);

                if (!Regex.IsMatch(originalLine, $@"\b{Regex.Escape(param.Name)}\b", RegexOptions.IgnoreCase))
                    continue;

                if (Regex.IsMatch(originalLine, $@"\bAs\s+{Regex.Escape(typeName)}(?=\s|\)|$)", RegexOptions.IgnoreCase))
                {
                    foundLine = lineIdx;
                    break;
                }

                if (originalLine.TrimEnd().EndsWith("_") &&
                    Regex.IsMatch(originalLine, @"\bAs\b", RegexOptions.IgnoreCase))
                {
                    for (int nextIdx = lineIdx + 1; nextIdx <= originalEndIndex; nextIdx++)
                    {
                        var nextLine = originalLines[nextIdx];
                        var nextCommentIndex = nextLine.IndexOf("'");
                        if (nextCommentIndex >= 0)
                            nextLine = nextLine.Substring(0, nextCommentIndex);

                        if (Regex.IsMatch(nextLine, $@"\b{Regex.Escape(typeName)}(?=\s|\)|$)", RegexOptions.IgnoreCase))
                        {
                            foundLine = nextIdx;
                            break;
                        }
                    }
                }

                if (foundLine >= 0)
                    break;
            }

            if (foundLine < 0)
            {
                for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
                {
                    var originalLine = originalLines[lineIdx];
                    var commentIndex = originalLine.IndexOf("'");
                    if (commentIndex >= 0)
                        originalLine = originalLine.Substring(0, commentIndex);

                    if (Regex.IsMatch(originalLine, $@"\b{Regex.Escape(typeName)}(?=\s|\)|$)", RegexOptions.IgnoreCase))
                    {
                        foundLine = lineIdx;
                        break;
                    }
                }
            }

            param.TypeLineNumber = foundLine >= 0 ? foundLine + 1 : param.LineNumber;
        }
    }
}
