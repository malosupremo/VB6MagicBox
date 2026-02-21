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

  [JsonIgnore]
  public List<string> ImplementsInterfaces { get; set; } = new();

  /// <summary>
  /// Lista di tutte le sostituzioni da applicare a questo modulo durante il refactoring.
  /// Ordinata per LineNumber (desc) e StartChar (desc) per applicazione sicura da fine a inizio.
  /// </summary>
  [JsonIgnore]
  public List<LineReplace> Replaces { get; set; } = new();
}
