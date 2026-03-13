using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace VB6MagicBox.Models;

public class VbVariable
{
    [JsonPropertyOrder(0)]
    public required string Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonIgnore]
    [JsonPropertyOrder(3)]
    public string? Level { get; set; }

    [JsonPropertyOrder(4)]
    public bool IsStatic { get; set; }

    [JsonPropertyOrder(5)]
    public bool IsArray { get; set; }

    [JsonPropertyOrder(6)]
    public bool IsWithEvents { get; set; }

    [JsonPropertyOrder(7)]
    public string? Scope { get; set; }

    [JsonPropertyOrder(8)]
    public string? Type { get; set; }

    [JsonPropertyOrder(9)]
    public bool Used { get; set; }

    [JsonPropertyOrder(10)]
    public string? Visibility { get; set; }

    [JsonPropertyOrder(11)]
    public List<VbReference> References { get; set; } = new();

    public int LineNumber { get; set; }
}

public class VbConstant
{
    [JsonPropertyOrder(0)]
    public string? Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonPropertyOrder(3)]
    public string? Level { get; set; }

    [JsonPropertyOrder(4)]
    public string? Scope { get; set; }

    [JsonPropertyOrder(5)]
    public string? Type { get; set; }

    [JsonPropertyOrder(6)]
    public bool Used { get; set; }

    [JsonPropertyOrder(7)]
    public string? Value { get; set; }

    [JsonPropertyOrder(8)]
    public string? Visibility { get; set; }

    [JsonPropertyOrder(9)]
    public List<VbReference> References { get; set; } = new();

    [JsonIgnore]
    public int LineNumber { get; set; }
}

public class VbTypeDef
{
    [JsonPropertyOrder(0)]
    public string? Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonPropertyOrder(3)]
    public bool Used { get; set; }

    [JsonPropertyOrder(4)]
    [JsonIgnore]
    public int LineNumber { get; set; }

    [JsonPropertyOrder(5)]
    public List<VbField> Fields { get; set; } = new();

    [JsonPropertyOrder(6)]
    public List<VbReference> References { get; set; } = new();
}

public class VbField
{
    [JsonPropertyOrder(0)]
    public string? Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonPropertyOrder(3)]
    public bool IsArray { get; set; }

    [JsonPropertyOrder(4)]
    public string? Type { get; set; }

    [JsonPropertyOrder(5)]
    public bool Used { get; set; }

    [JsonIgnore]
    public int LineNumber { get; set; }

    [JsonPropertyOrder(6)]
    public List<VbReference> References { get; set; } = new();
}

public class VbEnumDef
{
    [JsonPropertyOrder(0)]
    public string? Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonPropertyOrder(3)]
    public bool Used { get; set; }

    [JsonPropertyOrder(4)]
    [JsonIgnore]
    public int LineNumber { get; set; }

    [JsonPropertyOrder(5)]
    public List<VbEnumValue> Values { get; set; } = new();

    [JsonPropertyOrder(6)]
    public List<VbReference> References { get; set; } = new();
}

public class VbEnumValue
{
    [JsonPropertyOrder(0)]
    public string? Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonPropertyOrder(3)]
    public bool Used { get; set; }

    [JsonIgnore]
    public int LineNumber { get; set; }

    [JsonPropertyOrder(4)]
    public List<VbReference> References { get; set; } = new();
}

public class VbControl
{
    [JsonPropertyOrder(0)]
    public string? Name { get; set; }

    [JsonPropertyOrder(1)]
    public string? ConventionalName { get; set; }

    [JsonPropertyOrder(2)]
    public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

    [JsonPropertyOrder(3)]
    public string? ControlType { get; set; }

    [JsonPropertyOrder(4)]
    public bool IsArray { get; set; }

    [JsonPropertyOrder(5)]
    public bool Used { get; set; }

    [JsonPropertyOrder(6)]
    [JsonIgnore]
    public Dictionary<string, string> Properties { get; set; } = new();

    [JsonPropertyOrder(7)]
    [JsonIgnore]
    public int LineNumber { get; set; }

    [JsonPropertyOrder(8)]
    public List<int> LineNumbers { get; set; } = new();

    [JsonPropertyOrder(9)]
    public List<VbReference> References { get; set; } = new();
}

public class VbReference
{
    [JsonPropertyOrder(0)]
    public string? Module { get; set; }

    [JsonPropertyOrder(1)]
    public string? Procedure { get; set; }

    [JsonPropertyOrder(2)]
    public List<int> LineNumbers { get; set; } = new();

    [JsonPropertyOrder(3)]
    public List<int> StartChars { get; set; } = new();
}

/// <summary>
/// Rappresenta una singola sostituzione da applicare a una riga di codice.
/// Traccia la posizione esatta (carattere start/end) per sostituzioni precise.
/// </summary>
public class LineReplace
{
    [JsonPropertyOrder(0)]
    public int LineNumber { get; set; }

    [JsonPropertyOrder(1)]
    public int StartChar { get; set; }

    [JsonPropertyOrder(2)]
    public int EndChar { get; set; }

    [JsonPropertyOrder(3)]
    public string? OldText { get; set; }

    [JsonPropertyOrder(4)]
    public string? NewText { get; set; }

    [JsonPropertyOrder(5)]
    public string? Category { get; set; }
}

public class DependencyEdge
{
    [JsonPropertyOrder(0)]
    public string? CallerModule { get; set; }

    [JsonPropertyOrder(1)]
    public string? CallerProcedure { get; set; }

    [JsonPropertyOrder(2)]
    public string? CalleeModule { get; set; }

    [JsonPropertyOrder(3)]
    public string? CalleeProcedure { get; set; }

    [JsonPropertyOrder(4)]
    public string? CalleeRaw { get; set; }
}

/// <summary>
/// Extension methods for <see cref="List{VbReference}"/>.
/// </summary>
public static class VbReferenceListExtensions
{
    private sealed class ReferenceDebugEntry
    {
        public string? Module { get; set; }
        public string? Procedure { get; set; }
        public int LineNumber { get; set; }
        public int StartChar { get; set; }
        public List<VbReference>? ReferenceList { get; set; }
        public string? SymbolKind { get; set; }
        public string? SymbolName { get; set; }
        public string? SourceMember { get; set; }
        public string? SourceFile { get; set; }
        public int SourceLine { get; set; }
    }

    private static System.Collections.Concurrent.ConcurrentBag<ReferenceDebugEntry> ReferenceDebugEntries = new();

    public static void ResetReferenceDebugEntries()
    {
        ReferenceDebugEntries = new System.Collections.Concurrent.ConcurrentBag<ReferenceDebugEntry>();
    }

    /// <summary>
    /// Adds <paramref name="lineNumber"/> to an existing reference entry keyed by
    /// Module+Procedure, or creates a new entry when none exists.
    /// </summary>
    public static void AddLineNumber(
        this List<VbReference> references,
        string module,
        string procedure,
        int lineNumber,
        int startChar = -1,
        object? owner = null,
        [System.Runtime.CompilerServices.CallerMemberName] string sourceMember = "",
        [System.Runtime.CompilerServices.CallerFilePath] string sourceFile = "",
        [System.Runtime.CompilerServices.CallerLineNumber] int sourceLine = 0)
    {
        lock (references)
        {
            var normalizedProcedure = procedure ?? string.Empty;

            var existing = references.FirstOrDefault(r =>
                string.Equals(r.Module, module, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.Procedure ?? string.Empty, normalizedProcedure, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                if (lineNumber > 0)
                {
                    while (existing.StartChars.Count < existing.LineNumbers.Count)
                        existing.StartChars.Add(-1);

                    if (startChar < 0 && existing.LineNumbers.Any(ln => ln == lineNumber))
                    {
                        for (int i = 0; i < existing.LineNumbers.Count; i++)
                        {
                            if (existing.LineNumbers[i] == lineNumber &&
                                i < existing.StartChars.Count &&
                                existing.StartChars[i] >= 0)
                            {
                                return;
                            }
                        }
                    }

                    bool alreadyExists = false;
                    for (int i = 0; i < existing.LineNumbers.Count; i++)
                    {
                        var existingStartChar = i < existing.StartChars.Count ? existing.StartChars[i] : -1;
                        if (existing.LineNumbers[i] == lineNumber &&
                            existingStartChar == -1 &&
                            startChar >= 0)
                        {
                            existing.StartChars[i] = startChar;
                            alreadyExists = true;
                            break;
                        }
                        if (existing.LineNumbers[i] == lineNumber &&
                            existingStartChar == startChar)
                        {
                            alreadyExists = true;
                            break;
                        }
                    }

                    if (!alreadyExists)
                    {
                        existing.LineNumbers.Add(lineNumber);
                        existing.StartChars.Add(startChar);
                    }
                }
            }
            else
            {
                var newRef = new VbReference { Module = module, Procedure = normalizedProcedure };
                if (lineNumber > 0)
                {
                    newRef.LineNumbers.Add(lineNumber);
                    newRef.StartChars.Add(startChar);
                }
                references.Add(newRef);
            }

            if (lineNumber > 0 && startChar < 0)
            {
                var symbolKind = string.Empty;
                var symbolName = string.Empty;
                if (owner != null)
                {
                    (symbolKind, symbolName) = GetOwnerInfo(owner);
                }

                ReferenceDebugEntries.Add(new ReferenceDebugEntry
                {
                    Module = module,
                    Procedure = normalizedProcedure,
                    LineNumber = lineNumber,
                    StartChar = startChar,
                    ReferenceList = references,
                    SymbolKind = symbolKind,
                    SymbolName = symbolName,
                    SourceMember = sourceMember,
                    SourceFile = string.IsNullOrWhiteSpace(sourceFile) ? string.Empty : Path.GetFileName(sourceFile),
                    SourceLine = sourceLine
                });
            }
        }
    }

    private sealed class ReferenceListComparer : IEqualityComparer<List<VbReference>>
    {
        public bool Equals(List<VbReference>? x, List<VbReference>? y)
          => ReferenceEquals(x, y);

        public int GetHashCode(List<VbReference> obj)
          => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }

    private static (string Kind, string Name) GetOwnerInfo(object owner)
    {
        return owner switch
        {
            VbModule m => ("Module", m.Name),
            VbVariable v => ("Variable", v.Name),
            VbConstant c => ("Constant", c.Name ?? string.Empty),
            VbTypeDef t => ("Type", t.Name ?? string.Empty),
            VbField f => ("Field", f.Name ?? string.Empty),
            VbEnumDef e => ("Enum", e.Name ?? string.Empty),
            VbEnumValue ev => ("EnumValue", ev.Name ?? string.Empty),
            VbControl c => ("Control", c.Name ?? string.Empty),
            VbProperty p => ($"Property{p.Kind}", p.Name ?? string.Empty),
            VbParameter p => ("Parameter", p.Name ?? string.Empty),
            VbProcedure p => (p.Kind ?? "Procedure", p.Name ?? string.Empty),
            VbEvent e => ("Event", e.Name ?? string.Empty),
            _ => (string.Empty, string.Empty)
        };
    }
}

/// <summary>
/// Extension methods per gestire la lista di sostituzioni (LineReplace).
/// </summary>
public static class LineReplaceListExtensions
{
    /// <summary>
    /// Aggiunge una sostituzione precisa alla lista Replaces di un modulo.
    /// Traccia posizione esatta (carattere start/end) per sostituzioni univoche.
    /// </summary>
    public static void AddReplace(
        this List<LineReplace> replaces,
        int lineNumber,
        int startChar,
        int endChar,
        string oldText,
        string newText,
        string category)
    {
        lock (replaces)
        {
            // Verifica che la sostituzione non sia già presente (stesso lineNumber + startChar)
            var existing = replaces.FirstOrDefault(r =>
                r.LineNumber == lineNumber &&
                r.StartChar == startChar);

            if (existing != null)
                return; // Già tracciato

            replaces.Add(new LineReplace
            {
                LineNumber = lineNumber,
                StartChar = startChar,
                EndChar = endChar,
                OldText = oldText,
                NewText = newText,
                Category = category
            });
        }
    }

    /// <summary>
    /// Aggiunge una sostituzione trovando automaticamente la posizione nel codice della riga.
    /// Cerca il token specificato nella riga e traccia la sua posizione esatta.
    /// </summary>
    public static void AddReplaceFromLine(
        this List<LineReplace> replaces,
        string lineCode,
        int lineNumber,
        string oldName,
        string newName,
        string category,
        bool skipStringLiterals = false)
    {
        if (oldName == newName)
            return;

        var codeToSearch = lineCode;
        var stringRanges = new List<(int start, int end)>();

        if (skipStringLiterals)
        {
            bool inString = false;
            int stringStart = -1;

            for (int i = 0; i < lineCode.Length; i++)
            {
                if (lineCode[i] == '"')
                {
                    if (!inString)
                    {
                        inString = true;
                        stringStart = i;
                    }
                    else if (i + 1 < lineCode.Length && lineCode[i + 1] == '"')
                    {
                        i++;
                    }
                    else
                    {
                        inString = false;
                        if (stringStart >= 0)
                            stringRanges.Add((stringStart, i + 1));
                    }
                }
            }
        }

        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codeToSearch, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        bool IsInsideString(int pos)
        {
            return stringRanges.Any(range => pos >= range.start && pos < range.end);
        }

        var effectiveMatches = skipStringLiterals
            ? matches.Cast<Match>().Where(m => !IsInsideString(m.Index)).ToList()
            : matches.Cast<Match>().ToList();

        if (effectiveMatches.Count == 0)
            return;

        foreach (var match in effectiveMatches)
        {
            replaces.AddReplace(
                lineNumber,
                match.Index,
                match.Index + match.Length,
                match.Value,
                newName,
                category);
        }
    }
}
