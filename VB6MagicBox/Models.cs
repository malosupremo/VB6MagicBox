using System.Text.Json.Serialization;

namespace VB6MagicBox.Models;

public class VbProject
{
  [JsonPropertyOrder(0)]
  public required string ProjectFile { get; set; }
  
  [JsonPropertyOrder(1)]
  public List<VbModule> Modules { get; set; } = new();
  
  [JsonIgnore]
  public List<DependencyEdge> Dependencies { get; set; } = new();
}

public class VbModule
{
  [JsonPropertyOrder(0)]
  public required string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public string Kind { get; set; }

  [JsonPropertyOrder(4)]
  public string Path { get; set; }

  [JsonPropertyOrder(5)]
  public bool Used { get; set; }

  [JsonPropertyOrder(6)]
  [JsonIgnore]
  public string FullPath { get; set; }

  [JsonIgnore]
  public VbProject Owner { get; set; }

  [JsonPropertyOrder(7)]
  public List<VbConstant> Constants { get; set; } = new();

  [JsonPropertyOrder(8)]
  public List<VbControl> Controls { get; set; } = new();

  [JsonPropertyOrder(9)]
  public List<VbEnumDef> Enums { get; set; } = new();

  [JsonPropertyOrder(10)]
  public List<VbEvent> Events { get; set; } = new();

  [JsonPropertyOrder(11)]
  public List<VbVariable> GlobalVariables { get; set; } = new();

  [JsonPropertyOrder(12)]
  public List<VbProcedure> Procedures { get; set; } = new();

  [JsonPropertyOrder(13)]
  public List<VbProperty> Properties { get; set; } = new();

  [JsonPropertyOrder(14)]
  public List<VbTypeDef> Types { get; set; } = new();

  [JsonPropertyOrder(15)]
  public List<VbReference> References { get; set; } = new();

  /// <summary>
  /// Moduli che referenziano questo modulo attraverso qualsiasi suo membro
  /// (costanti, tipi, enum, procedure, property, controlli, variabili).
  /// Popolato da BuildDependenciesAndUsage; usato per il grafo Mermaid.
  /// </summary>
  [JsonPropertyOrder(16)]
  public List<string> ModuleReferences { get; set; } = new();

  [JsonIgnore]
  public bool IsClass => Kind.Equals("cls", StringComparison.OrdinalIgnoreCase);
  
  [JsonIgnore]
  public bool IsForm => Kind.Equals("frm", StringComparison.OrdinalIgnoreCase);
  
  /// <summary>
  /// Trova la procedura che contiene il numero di riga specificato
  /// </summary>
  public VbProcedure? GetProcedureAtLine(int lineNumber)
  {
    return Procedures.FirstOrDefault(p => p.ContainsLine(lineNumber));
  }
}

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

public class VbProcedure
{
  [JsonPropertyOrder(0)]
  public required string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public bool IsStatic { get; set; }

  [JsonPropertyOrder(4)]
  public string Kind { get; set; }

  [JsonPropertyOrder(5)]
  public string ReturnType { get; set; }

  [JsonPropertyOrder(6)]
  public string Scope { get; set; }

  [JsonPropertyOrder(7)]
  public bool Used { get; set; }

  [JsonPropertyOrder(8)]
  public string Visibility { get; set; }

  [JsonPropertyOrder(9)]
  [JsonIgnore]
  public int LineNumber { get; set; }

  [JsonIgnore]
  public int StartLine { get; set; }

  [JsonIgnore]
  public int EndLine { get; set; }

  /// <summary>
  /// Controlla se un numero di riga è dentro questa procedura
  /// </summary>
  public bool ContainsLine(int lineNumber)
  {
    return lineNumber >= StartLine && lineNumber <= EndLine;
  }

  [JsonIgnore]
  public List<VbCall> Calls { get; set; } = new();

  [JsonPropertyOrder(10)]
  public List<VbConstant> Constants { get; set; } = new();

  [JsonPropertyOrder(11)]
  public List<VbVariable> LocalVariables { get; set; } = new();

  [JsonPropertyOrder(12)]
  public List<VbParameter> Parameters { get; set; } = new();

  [JsonPropertyOrder(13)]
  public List<VbReference> References { get; set; } = new();
}

public class VbParameter
{
  [JsonPropertyOrder(0)]
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public string Passing { get; set; }

  [JsonPropertyOrder(4)]
  public string Type { get; set; }

  [JsonPropertyOrder(5)]
  public bool Used { get; set; }

  [JsonPropertyOrder(6)]
  public List<VbReference> References { get; set; } = new();

  [JsonIgnore]
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

public class VbEvent
{
  [JsonPropertyOrder(0)]
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public string Scope { get; set; }

  [JsonPropertyOrder(4)]
  public bool Used { get; set; }

  [JsonPropertyOrder(5)]
  public string Visibility { get; set; }

  [JsonIgnore]
  public int LineNumber { get; set; }

  [JsonPropertyOrder(6)]
  public List<VbParameter> Parameters { get; set; } = new();

  [JsonPropertyOrder(7)]
  public List<VbReference> References { get; set; } = new();
}

public class VbProperty
{
  [JsonPropertyOrder(0)]
  public string Name { get; set; }

  [JsonPropertyOrder(1)]
  public string ConventionalName { get; set; }

  [JsonPropertyOrder(2)]
  public bool IsConventional => string.Equals(Name, ConventionalName, StringComparison.Ordinal);

  [JsonPropertyOrder(3)]
  public string Kind { get; set; } // "Get", "Let", "Set"

  [JsonPropertyOrder(4)]
  public string Scope { get; set; }

  [JsonPropertyOrder(5)]
  public bool Used { get; set; }

  [JsonPropertyOrder(6)]
  public string Visibility { get; set; }

  [JsonPropertyOrder(7)]
  public string ReturnType { get; set; }

  [JsonIgnore]
  public int LineNumber { get; set; }

  [JsonIgnore]
  public int StartLine { get; set; }

  [JsonIgnore]
  public int EndLine { get; set; }

  [JsonPropertyOrder(8)]
  public List<VbParameter> Parameters { get; set; } = new();

  [JsonPropertyOrder(9)]
  public List<VbReference> References { get; set; } = new();

  /// <summary>
  /// Verifica se la riga specificata è all'interno di questa proprietà
  /// </summary>
  public bool ContainsLine(int lineNumber)
  {
    return lineNumber >= StartLine && lineNumber <= EndLine;
  }
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

public class VbCall
{
  [JsonPropertyOrder(0)]
  public string Raw { get; set; }

  [JsonPropertyOrder(1)]
  public string MethodName { get; set; }

  [JsonPropertyOrder(2)]
  public string ObjectName { get; set; }

  [JsonPropertyOrder(3)]
  public string ResolvedKind { get; set; }

  [JsonPropertyOrder(4)]
  public string ResolvedModule { get; set; }

  [JsonPropertyOrder(5)]
  public string ResolvedProcedure { get; set; }

  [JsonPropertyOrder(6)]
  public string ResolvedType { get; set; }

  [JsonPropertyOrder(7)]
  [JsonIgnore]  // Dati interni per refactoring puntuale
  public int LineNumber { get; set; }
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
      int lineNumber)
  {
    var existing = references.FirstOrDefault(r =>
        string.Equals(r.Module, module, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(r.Procedure, procedure, StringComparison.OrdinalIgnoreCase));

    if (existing != null)
    {
      if (lineNumber > 0 && !existing.LineNumbers.Contains(lineNumber))
        existing.LineNumbers.Add(lineNumber);
    }
    else
    {
      var newRef = new VbReference { Module = module, Procedure = procedure };
      if (lineNumber > 0)
        newRef.LineNumbers.Add(lineNumber);
      references.Add(newRef);
    }
  }
}
