using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace VB6MagicBox.Models;

public class VbVariable
{
  [JsonPropertyOrder(0)]
  public required string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonIgnore]
  [JsonPropertyOrder(3)]
  public string Level { get; set; }

  [JsonPropertyOrder(4)]
  public bool IsStatic { get; set; }

  [JsonPropertyOrder(5)]
  public bool IsArray { get; set; }

  [JsonPropertyOrder(6)]
  public bool IsWithEvents { get; set; }

  [JsonPropertyOrder(7)]
  public string Scope { get; set; }

  [JsonPropertyOrder(8)]
  public string Type { get; set; }

  [JsonPropertyOrder(9)]
  public bool Used { get; set; }

  [JsonPropertyOrder(10)]
  public string Visibility { get; set; }

  [JsonPropertyOrder(11)]
  public List<VbReference> References { get; set; } = new();

  public int LineNumber { get; set; }
}

public class VbConstant
{
  [JsonPropertyOrder(0)]
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public string Level { get; set; }

  [JsonPropertyOrder(4)]
  public string Scope { get; set; }

  [JsonPropertyOrder(5)]
  public string Type { get; set; }

  [JsonPropertyOrder(6)]
  public bool Used { get; set; }

  [JsonPropertyOrder(7)]
  public string Value { get; set; }

  [JsonPropertyOrder(8)]
  public string Visibility { get; set; }

  [JsonPropertyOrder(9)]
  public List<VbReference> References { get; set; } = new();

  [JsonIgnore]
  public int LineNumber { get; set; }
}

public class VbTypeDef
{
  [JsonPropertyOrder(0)]
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

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
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public bool IsArray { get; set; }

  [JsonPropertyOrder(4)]
  public string Type { get; set; }

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
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

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
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

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
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public string ControlType { get; set; }

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
  public string Module { get; set; }

  [JsonPropertyOrder(1)]
  public string Procedure { get; set; }

  [JsonPropertyOrder(2)]
  public List<int> LineNumbers { get; set; } = new();
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
  public string OldText { get; set; }

  [JsonPropertyOrder(4)]
  public string NewText { get; set; }

  [JsonPropertyOrder(5)]
  public string Category { get; set; }
}

public class DependencyEdge
{
  [JsonPropertyOrder(0)]
  public string CallerModule { get; set; }

  [JsonPropertyOrder(1)]
  public string CallerProcedure { get; set; }

  [JsonPropertyOrder(2)]
  public string CalleeModule { get; set; }

  [JsonPropertyOrder(3)]
  public string CalleeProcedure { get; set; }

  [JsonPropertyOrder(4)]
  public string CalleeRaw { get; set; }
}

/// <summary>
/// Extension methods for <see cref="List{VbReference}"/>.
/// </summary>
public static class VbReferenceListExtensions
{
  /// <summary>
  /// Adds <paramref name="lineNumber"/> to an existing reference entry keyed by
  /// Module+Procedure, or creates a new entry when none exists.
  /// Note: occurrenceIndex parameter kept for backward compatibility but no longer used.
  /// </summary>
  public static void AddLineNumber(
      this List<VbReference> references,
      string module,
      string procedure,
      int lineNumber,
      int occurrenceIndex = -1)
  {
    var normalizedProcedure = procedure ?? string.Empty;

    var existing = references.FirstOrDefault(r =>
        string.Equals(r.Module, module, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(r.Procedure ?? string.Empty, normalizedProcedure, StringComparison.OrdinalIgnoreCase));

    if (existing != null)
    {
      if (lineNumber > 0 && !existing.LineNumbers.Contains(lineNumber))
      {
        existing.LineNumbers.Add(lineNumber);
      }
    }
    else
    {
      var newRef = new VbReference { Module = module, Procedure = normalizedProcedure };
      if (lineNumber > 0)
      {
        newRef.LineNumbers.Add(lineNumber);
      }
      references.Add(newRef);
    }
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
      int occurrenceIndex = -1,
      bool skipStringLiterals = false)
  {
    if (oldName == newName)
      return;

    // Se dobbiamo saltare le stringhe literals (es. per le costanti),
    // elimina temporaneamente le stringhe dal codice prima di cercare i match
    var codeToSearch = lineCode;
    var stringRanges = new List<(int start, int end)>();

    if (skipStringLiterals)
    {
      // Trova tutte le stringhe literals e segna i loro range
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
            i++; // Escaped double quote
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

    // Trova tutte le occorrenze del token nella riga
    var pattern = $@"\b{Regex.Escape(oldName)}\b";
    var matches = Regex.Matches(codeToSearch, pattern, RegexOptions.IgnoreCase);

    if (matches.Count == 0)
      return;

    // Funzione helper per verificare se una posizione è dentro una stringa
    bool IsInsideString(int pos)
    {
      return stringRanges.Any(range => pos >= range.start && pos < range.end);
    }

    // Se occurrenceIndex è specificato (1-based), usa solo quella
    if (occurrenceIndex > 0 && occurrenceIndex <= matches.Count)
    {
      var match = matches[occurrenceIndex - 1];

      // Salta se è dentro una stringa
      if (skipStringLiterals && IsInsideString(match.Index))
        return;

      replaces.AddReplace(
          lineNumber,
          match.Index,
          match.Index + match.Length,
          match.Value,
          newName,
          category);
    }
    else
    {
      // Altrimenti aggiungi tutte le occorrenze (escluse quelle nelle stringhe)
      foreach (Match match in matches)
      {
        // Salta se è dentro una stringa
        if (skipStringLiterals && IsInsideString(match.Index))
          continue;

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
}
