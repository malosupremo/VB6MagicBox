using System.Text.Json.Serialization;

namespace VB6MagicBox.Models;

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
