using System.Text.Json.Serialization;

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

  [JsonPropertyOrder(3)]
  public List<int> OccurrenceIndexes { get; set; } = new();
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
      if (lineNumber > 0)
      {
        if (occurrenceIndex >= 0)
        {
          bool alreadyTracked = existing.LineNumbers
              .Select((ln, idx) => new { ln, idx })
              .Any(x => x.ln == lineNumber &&
                        x.idx < existing.OccurrenceIndexes.Count &&
                        existing.OccurrenceIndexes[x.idx] == occurrenceIndex);

          if (!alreadyTracked)
          {
            existing.LineNumbers.Add(lineNumber);
            existing.OccurrenceIndexes.Add(occurrenceIndex);
          }
        }
        else if (!existing.LineNumbers.Contains(lineNumber))
        {
          existing.LineNumbers.Add(lineNumber);
          existing.OccurrenceIndexes.Add(occurrenceIndex);
        }
      }
    }
    else
    {
      var newRef = new VbReference { Module = module, Procedure = normalizedProcedure };
      if (lineNumber > 0)
      {
        newRef.LineNumbers.Add(lineNumber);
        newRef.OccurrenceIndexes.Add(occurrenceIndex);
      }
      references.Add(newRef);
    }
  }
}
