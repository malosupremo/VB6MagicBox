using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  // ---------------------------------------------------------
  // RISOLUZIONE RIFERIMENTI AI TIPI (As TypeName)
  // ---------------------------------------------------------

  /// <summary>
  /// Aggiunge References ai VbTypeDef per tutte le posizioni in cui il tipo
  /// appare in una clausola "As TypeName": campi di altri Type, variabili
  /// globali/locali, parametri di procedure e proprietà.
  /// Senza questo, il refactoring non sa quali righe aggiornare quando
  /// rinomina un tipo (es. DISPAT_HEADER_T ? DispatHeader_T).
  /// </summary>
  private static void ResolveTypeReferences(
      VbProject project,
      Dictionary<string, VbTypeDef> typeIndex)
  {
    foreach (var mod in project.Modules)
    {
      // 1. Campi di altri Type: "FieldName As OTHER_TYPE"
      foreach (var typeDef in mod.Types)
      {
        foreach (var field in typeDef.Fields)
        {
          AddTypeReference(field.Type, field.LineNumber, mod.Name, string.Empty, typeIndex);
        }
      }

      // 2. Variabili globali: "Public/Dim varName As TYPE"
      foreach (var variable in mod.GlobalVariables)
      {
        AddTypeReference(variable.Type, variable.LineNumber, mod.Name, string.Empty, typeIndex);
      }

      // 3. Parametri e variabili locali delle procedure
      foreach (var proc in mod.Procedures)
      {
        AddTypeReference(proc.ReturnType, proc.LineNumber, mod.Name, proc.Name, typeIndex);

        foreach (var param in proc.Parameters)
          AddTypeReference(param.Type, param.LineNumber, mod.Name, proc.Name, typeIndex);

        foreach (var localVar in proc.LocalVariables)
          AddTypeReference(localVar.Type, localVar.LineNumber, mod.Name, proc.Name, typeIndex);
      }

      // 4. Parametri delle proprietà
      foreach (var prop in mod.Properties)
      {
        AddTypeReference(prop.ReturnType, prop.LineNumber, mod.Name, prop.Name, typeIndex);

        foreach (var param in prop.Parameters)
          AddTypeReference(param.Type, param.LineNumber, mod.Name, prop.Name, typeIndex);
      }
    }
  }

  /// <summary>
  /// Se typeName è un tipo noto nel typeIndex, aggiunge lineNumber alle sue References.
  /// </summary>
  private static void AddTypeReference(
      string typeName,
      int lineNumber,
      string moduleName,
      string procedureName,
      Dictionary<string, VbTypeDef> typeIndex)
  {
    if (string.IsNullOrEmpty(typeName) || lineNumber <= 0)
      return;

    // Rimuovi eventuali parentesi per tipi array (es. "MY_TYPE()" ? "MY_TYPE")
    var baseTypeName = typeName.Contains('(')
        ? typeName.Substring(0, typeName.IndexOf('('))
        : typeName;

    if (!typeIndex.TryGetValue(baseTypeName, out var referencedType))
      return;

    referencedType.Used = true;
    referencedType.References.AddLineNumber(moduleName, procedureName, lineNumber);
  }

  // ---------------------------------------------------------
  // RISOLUZIONE RIFERIMENTI ALLE CLASSI (As [New] ClassName)
  // ---------------------------------------------------------

  /// <summary>
  /// Aggiunge References alla VbModule classe per ogni dichiarazione
  /// "Dim/Private x As [New] ClassName" dove ClassName è un modulo classe.
  /// Garantisce che le classi usate solo come tipo (senza chiamate risolte)
  /// compaiano comunque nelle References e nel grafo Mermaid.
  /// </summary>
  private static void ResolveClassModuleReferences(VbProject project)
  {
    var classIndex = project.Modules
        .Where(m => m.IsClass)
        .ToDictionary(
            m => Path.GetFileNameWithoutExtension(m.Name),
            m => m,
            StringComparer.OrdinalIgnoreCase);

    foreach (var mod in project.Modules)
    {
      foreach (var v in mod.GlobalVariables)
        AddClassModuleReference(v.Type, v.LineNumber, mod.Name, string.Empty, classIndex);

      foreach (var proc in mod.Procedures)
      {
        AddClassModuleReference(proc.ReturnType, proc.LineNumber, mod.Name, proc.Name, classIndex);

        foreach (var param in proc.Parameters)
          AddClassModuleReference(param.Type, param.LineNumber, mod.Name, proc.Name, classIndex);
        foreach (var lv in proc.LocalVariables)
          AddClassModuleReference(lv.Type, lv.LineNumber, mod.Name, proc.Name, classIndex);
      }

      foreach (var prop in mod.Properties)
      {
        AddClassModuleReference(prop.ReturnType, prop.LineNumber, mod.Name, prop.Name, classIndex);

        foreach (var param in prop.Parameters)
          AddClassModuleReference(param.Type, param.LineNumber, mod.Name, prop.Name, classIndex);
      }
    }
  }

  private static void AddClassModuleReference(
      string typeName,
      int lineNumber,
      string declaringModule,
      string procedureName,
      Dictionary<string, VbModule> classIndex)
  {
    if (string.IsNullOrEmpty(typeName) || lineNumber <= 0)
      return;

    // Prendi solo il nome base ignorando eventuali namespace (es. "PDxI.clsPDxI" ? "clsPDxI")
    var baseName = typeName.Contains('.') ? typeName.Split('.').Last() : typeName;

    if (!classIndex.TryGetValue(baseName, out var classModule))
      return;

    // Non aggiungere auto-referenze
    if (string.Equals(classModule.Name, declaringModule, StringComparison.OrdinalIgnoreCase))
      return;

    classModule.Used = true;
    classModule.References.AddLineNumber(declaringModule, procedureName, lineNumber);
  }

  // ---------------------------------------------------------
  // MARCATURA TIPI USATI
  // ---------------------------------------------------------

  private static void MarkUsedTypes(VbProject project)
  {
    var allTypes = project.Modules
        .SelectMany(m => m.Types)
        .GroupBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

    var allEnums = project.Modules
        .SelectMany(m => m.Enums)
        .GroupBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

    // Indicizza le classi per nome (senza estensione .cls e senza prefisso cls)
    var allClasses = new Dictionary<string, VbModule>(StringComparer.OrdinalIgnoreCase);
    foreach (var mod in project.Modules.Where(m => m.IsClass))
    {
      var className = Path.GetFileNameWithoutExtension(mod.Name);
      allClasses[className] = mod;

      // Aggiungi anche senza prefisso "cls" se presente
      if (className.StartsWith("cls", StringComparison.OrdinalIgnoreCase))
      {
        var withoutPrefix = className.Substring(3);
        if (!allClasses.ContainsKey(withoutPrefix))
          allClasses[withoutPrefix] = mod;
      }
    }

    void Mark(string typeName, string moduleName, string procedureName = null, int lineNumber = 0)
    {
      if (string.IsNullOrWhiteSpace(typeName))
        return;

      var clean = typeName.Trim();

      // Rimuovi eventuali namespace (es. "PDxI.clsPDxI" -> "clsPDxI")
      if (clean.Contains("."))
        clean = clean.Split('.').Last();

      if (allTypes.TryGetValue(clean, out var t))
      {
        t.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
          t.References.AddLineNumber(moduleName, procedureName, lineNumber);
      }

      if (allEnums.TryGetValue(clean, out var e))
      {
        e.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
          e.References.AddLineNumber(moduleName, procedureName, lineNumber);
      }

      // Traccia anche le classi usate come tipo
      if (allClasses.TryGetValue(clean, out var cls))
      {
        cls.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
          cls.References.AddLineNumber(moduleName, procedureName, lineNumber);
      }
    }

    foreach (var mod in project.Modules)
    {
      // Variabili globali usano Type/Enum/Class - riferimento a livello di modulo
      foreach (var v in mod.GlobalVariables)
        Mark(v.Type, mod.Name, lineNumber: v.LineNumber);

      // Campi dei Type: "Field As ENUM/TYPE"
      foreach (var typeDef in mod.Types)
      {
        foreach (var field in typeDef.Fields)
          Mark(field.Type, mod.Name, lineNumber: field.LineNumber);
      }

      foreach (var proc in mod.Procedures)
      {
        // Return type, parametri e variabili locali - riferimento da procedura
        Mark(proc.ReturnType, mod.Name, proc.Name, proc.LineNumber);

        foreach (var p in proc.Parameters)
          Mark(p.Type, mod.Name, proc.Name, p.LineNumber);

        foreach (var lv in proc.LocalVariables)
          Mark(lv.Type, mod.Name, proc.Name, lv.LineNumber);
      }

      foreach (var prop in mod.Properties)
      {
        Mark(prop.ReturnType, mod.Name, prop.Name, prop.LineNumber);

        foreach (var p in prop.Parameters)
          Mark(p.Type, mod.Name, prop.Name, p.LineNumber);
      }
    }
  }
}
