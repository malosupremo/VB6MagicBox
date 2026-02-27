using System.Text.RegularExpressions;
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
      Dictionary<string, VbTypeDef> typeIndex,
      Dictionary<string, string[]> fileCache)
  {
    foreach (var mod in project.Modules)
    {
      var fileLines = GetFileLines(fileCache, mod);

      // 1. Campi di altri Type: "FieldName As OTHER_TYPE"
      foreach (var typeDef in mod.Types)
      {
        for (int i = 0; i < typeDef.Fields.Count; i++)
        {
          var field = typeDef.Fields[i];
          var occurrence = typeDef.Fields.Take(i)
              .Count(f => f.LineNumber == field.LineNumber &&
                          f.Type?.Equals(field.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

          var lineText = field.LineNumber > 0 && field.LineNumber <= fileLines.Length
              ? fileLines[field.LineNumber - 1]
              : string.Empty;
          AddTypeReference(field.Type, lineText, field.LineNumber, mod.Name, string.Empty, typeIndex, occurrence);
        }
      }

      // 2. Variabili globali: "Public/Dim varName As TYPE"
      for (int i = 0; i < mod.GlobalVariables.Count; i++)
      {
        var variable = mod.GlobalVariables[i];
        var occurrence = mod.GlobalVariables.Take(i)
            .Count(v => v.LineNumber == variable.LineNumber &&
                        v.Type?.Equals(variable.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

        var lineText = variable.LineNumber > 0 && variable.LineNumber <= fileLines.Length
            ? fileLines[variable.LineNumber - 1]
            : string.Empty;
        AddTypeReference(variable.Type, lineText, variable.LineNumber, mod.Name, string.Empty, typeIndex, occurrence);
      }

      // 3. Parametri e variabili locali delle procedure
      foreach (var proc in mod.Procedures)
      {
        var procReturnLine = proc.ReturnTypeLineNumber > 0 ? proc.ReturnTypeLineNumber : proc.LineNumber;
        var procReturnText = procReturnLine > 0 && procReturnLine <= fileLines.Length
            ? fileLines[procReturnLine - 1]
            : string.Empty;
        AddTypeReference(proc.ReturnType, procReturnText, procReturnLine, mod.Name, proc.Name, typeIndex, -1);

        // Per i parametri, calcola occurrenceIndex per gestire più parametri dello stesso tipo sulla stessa riga
        for (int i = 0; i < proc.Parameters.Count; i++)
        {
          var param = proc.Parameters[i];
          var paramLine = param.TypeLineNumber > 0 ? param.TypeLineNumber : param.LineNumber;
          var isMultilineType = param.TypeLineNumber > 0 && param.TypeLineNumber != param.LineNumber;
          int occurrence = isMultilineType
            ? -1
            : proc.Parameters.Take(i)
                .Count(p => (p.TypeLineNumber > 0 ? p.TypeLineNumber : p.LineNumber) == paramLine &&
                            p.Type?.Equals(param.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

          var paramLineText = paramLine > 0 && paramLine <= fileLines.Length
              ? fileLines[paramLine - 1]
              : string.Empty;
          AddTypeReference(param.Type, paramLineText, paramLine, mod.Name, proc.Name, typeIndex, occurrence);
        }

        for (int i = 0; i < proc.LocalVariables.Count; i++)
        {
          var localVar = proc.LocalVariables[i];
          var occurrence = proc.LocalVariables.Take(i)
              .Count(v => v.LineNumber == localVar.LineNumber &&
                          v.Type?.Equals(localVar.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

          var localLineText = localVar.LineNumber > 0 && localVar.LineNumber <= fileLines.Length
              ? fileLines[localVar.LineNumber - 1]
              : string.Empty;
          AddTypeReference(localVar.Type, localLineText, localVar.LineNumber, mod.Name, proc.Name, typeIndex, occurrence);
        }
      }

      // 4. Parametri delle proprietà
      foreach (var prop in mod.Properties)
      {
        var propReturnLine = prop.ReturnTypeLineNumber > 0 ? prop.ReturnTypeLineNumber : prop.LineNumber;
        var propReturnText = propReturnLine > 0 && propReturnLine <= fileLines.Length
            ? fileLines[propReturnLine - 1]
            : string.Empty;
        AddTypeReference(prop.ReturnType, propReturnText, propReturnLine, mod.Name, prop.Name, typeIndex, -1);

        // Per i parametri, calcola occurrenceIndex
        for (int i = 0; i < prop.Parameters.Count; i++)
        {
          var param = prop.Parameters[i];
          var paramLine = param.TypeLineNumber > 0 ? param.TypeLineNumber : param.LineNumber;
          var isMultilineType = param.TypeLineNumber > 0 && param.TypeLineNumber != param.LineNumber;
          int occurrence = isMultilineType
            ? -1
            : prop.Parameters.Take(i)
                .Count(p => (p.TypeLineNumber > 0 ? p.TypeLineNumber : p.LineNumber) == paramLine &&
                            p.Type?.Equals(param.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

          var paramLineText = paramLine > 0 && paramLine <= fileLines.Length
              ? fileLines[paramLine - 1]
              : string.Empty;
          AddTypeReference(param.Type, paramLineText, paramLine, mod.Name, prop.Name, typeIndex, occurrence);
        }
      }
    }
  }

  /// <summary>
  /// Se typeName è un tipo noto nel typeIndex, aggiunge lineNumber alle sue References.
  /// occurrenceIndex (1-based) specifica quale occorrenza del tipo sulla riga (per parametri multipli).
  /// </summary>
  private static void AddTypeReference(
      string typeName,
      string lineText,
      int lineNumber,
      string moduleName,
      string procedureName,
      Dictionary<string, VbTypeDef> typeIndex,
      int occurrenceIndex = -1)
  {
        // Debug per PLC_POLL_WHAT_CMD_T
        bool isDebug = false;

    if (isDebug)
    {
      Console.WriteLine($"\n[DEBUG AddTypeReference] Type={typeName}, Line={lineNumber}, Module={moduleName}, Proc={procedureName}, Occ={occurrenceIndex}");
    }

    if (string.IsNullOrEmpty(typeName) || lineNumber <= 0)
    {
      if (isDebug)
        Console.WriteLine($"[DEBUG] SKIPPED: typeName empty or lineNumber <= 0");
      return;
    }

    // Rimuovi eventuali parentesi per tipi array (es. "MY_TYPE()" -> "MY_TYPE")
    var baseTypeName = typeName.Contains('(')
        ? typeName.Substring(0, typeName.IndexOf('('))
        : typeName;

    if (!typeIndex.TryGetValue(baseTypeName, out var referencedType))
    {
      if (!baseTypeName.EndsWith("_T", StringComparison.OrdinalIgnoreCase))
      {
        var suffixedName = baseTypeName + "_T";
        if (typeIndex.TryGetValue(suffixedName, out referencedType))
          baseTypeName = suffixedName;
      }
      else
      {
        var unsuffixedName = baseTypeName.Substring(0, baseTypeName.Length - 2);
        if (typeIndex.TryGetValue(unsuffixedName, out referencedType))
          baseTypeName = unsuffixedName;
      }

      if (referencedType == null)
      {
        if (isDebug)
          Console.WriteLine($"[DEBUG] SKIPPED: Type not in typeIndex");
        return;
      }
    }

    if (isDebug)
    {
      Console.WriteLine($"[DEBUG] ✅ Adding Reference to {baseTypeName}");
    }

    var startChar = -1;
    var effectiveOccurrenceIndex = occurrenceIndex;
    if (!string.IsNullOrEmpty(lineText))
    {
      var noComment = StripInlineComment(lineText);
      startChar = GetTokenIndex(noComment, baseTypeName, occurrenceIndex);
      if (startChar < 0 && !string.Equals(baseTypeName, typeName, StringComparison.OrdinalIgnoreCase))
        startChar = GetTokenIndex(noComment, typeName, occurrenceIndex);

      if (effectiveOccurrenceIndex < 0 && startChar >= 0)
        effectiveOccurrenceIndex = GetOccurrenceIndex(noComment, baseTypeName, startChar, lineNumber);
    }

    referencedType.Used = true;
    referencedType.References.AddLineNumber(moduleName, procedureName, lineNumber, effectiveOccurrenceIndex, startChar);
  }

  private static int GetTokenIndex(string line, string token, int occurrenceIndex)
  {
    if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(token))
      return -1;

    var matches = Regex.Matches(line, $@"\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase);
    if (matches.Count == 0)
      return -1;

    if (occurrenceIndex > 0 && occurrenceIndex <= matches.Count)
      return matches[occurrenceIndex - 1].Index;

    return matches[0].Index;
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
  private static void ResolveClassModuleReferences(VbProject project, Dictionary<string, string[]> fileCache)
  {
    var classIndex = project.Modules
        .Where(m => m.IsClass)
        .ToDictionary(
            m => Path.GetFileNameWithoutExtension(m.Name),
            m => m,
            StringComparer.OrdinalIgnoreCase);

    foreach (var mod in project.Modules)
    {
      var fileLines = GetFileLines(fileCache, mod);
      foreach (var v in mod.GlobalVariables)
      {
        var lineText = v.LineNumber > 0 && v.LineNumber <= fileLines.Length
            ? fileLines[v.LineNumber - 1]
            : string.Empty;
        AddClassModuleReference(v.Type, lineText, v.LineNumber, mod.Name, string.Empty, classIndex, -1);
      }

      foreach (var proc in mod.Procedures)
      {
        var procLineText = proc.LineNumber > 0 && proc.LineNumber <= fileLines.Length
            ? fileLines[proc.LineNumber - 1]
            : string.Empty;
        AddClassModuleReference(proc.ReturnType, procLineText, proc.LineNumber, mod.Name, proc.Name, classIndex, -1);

        // Per i parametri, calcola occurrenceIndex
        for (int i = 0; i < proc.Parameters.Count; i++)
        {
          var param = proc.Parameters[i];
          var paramLineText = param.LineNumber > 0 && param.LineNumber <= fileLines.Length
              ? fileLines[param.LineNumber - 1]
              : string.Empty;
          int occurrence = proc.Parameters.Take(i)
              .Count(p => p.LineNumber == param.LineNumber && 
                          p.Type?.Equals(param.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

          AddClassModuleReference(param.Type, paramLineText, param.LineNumber, mod.Name, proc.Name, classIndex, occurrence);
        }

        foreach (var lv in proc.LocalVariables)
        {
          var localLineText = lv.LineNumber > 0 && lv.LineNumber <= fileLines.Length
              ? fileLines[lv.LineNumber - 1]
              : string.Empty;
          AddClassModuleReference(lv.Type, localLineText, lv.LineNumber, mod.Name, proc.Name, classIndex, -1);
        }
      }

      foreach (var prop in mod.Properties)
      {
        var propLineText = prop.LineNumber > 0 && prop.LineNumber <= fileLines.Length
            ? fileLines[prop.LineNumber - 1]
            : string.Empty;
        AddClassModuleReference(prop.ReturnType, propLineText, prop.LineNumber, mod.Name, prop.Name, classIndex, -1);

        // Per i parametri, calcola occurrenceIndex
        for (int i = 0; i < prop.Parameters.Count; i++)
        {
          var param = prop.Parameters[i];
          var paramLineText = param.LineNumber > 0 && param.LineNumber <= fileLines.Length
              ? fileLines[param.LineNumber - 1]
              : string.Empty;
          int occurrence = prop.Parameters.Take(i)
              .Count(p => p.LineNumber == param.LineNumber && 
                          p.Type?.Equals(param.Type, StringComparison.OrdinalIgnoreCase) == true) + 1;

          AddClassModuleReference(param.Type, paramLineText, param.LineNumber, mod.Name, prop.Name, classIndex, occurrence);
        }
      }
    }
  }

  private static void AddClassModuleReference(
      string typeName,
      string lineText,
      int lineNumber,
      string declaringModule,
      string procedureName,
      Dictionary<string, VbModule> classIndex,
      int occurrenceIndex = -1)
  {
    if (string.IsNullOrEmpty(typeName) || lineNumber <= 0)
      return;

    // Prendi solo il nome base ignorando eventuali namespace (es. "PDxI.clsPDxI" -> "clsPDxI")
    var baseName = typeName.Contains('.') ? typeName.Split('.').Last() : typeName;

    if (!classIndex.TryGetValue(baseName, out var classModule))
      return;

    // Non aggiungere auto-referenze
    if (string.Equals(classModule.Name, declaringModule, StringComparison.OrdinalIgnoreCase))
      return;

    var startChar = -1;
    var effectiveOccurrenceIndex = occurrenceIndex;
    if (!string.IsNullOrEmpty(lineText))
    {
      var noComment = StripInlineComment(lineText);
      startChar = GetTokenIndex(noComment, baseName, occurrenceIndex);
      if (effectiveOccurrenceIndex < 0 && startChar >= 0)
        effectiveOccurrenceIndex = GetOccurrenceIndex(noComment, baseName, startChar, lineNumber);
    }

    classModule.Used = true;
    classModule.References.AddLineNumber(declaringModule, procedureName, lineNumber, effectiveOccurrenceIndex, startChar);
  }

  // ---------------------------------------------------------
  // MARCATURA TIPI USATI
  // ---------------------------------------------------------

  private static void MarkUsedTypes(VbProject project, Dictionary<string, string[]> fileCache)
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

    bool TryGetLineInfo(string moduleName, int lineNumber, out string line)
    {
      line = null;
      if (lineNumber <= 0)
        return false;

      var mod = project.Modules.FirstOrDefault(m =>
          string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase));
      if (mod == null)
        return false;

      var lines = GetFileLines(fileCache, mod);
      if (lineNumber > lines.Length)
        return false;

      line = lines[lineNumber - 1];
      return true;
    }

    void Mark(string typeName, string moduleName, string procedureName = null, int lineNumber = 0)
    {
      if (string.IsNullOrWhiteSpace(typeName))
        return;

      var clean = typeName.Trim();

      // Rimuovi eventuali namespace (es. "PDxI.clsPDxI" -> "clsPDxI")
      if (clean.Contains("."))
        clean = clean.Split('.').Last();

    var occurrenceIndex = -1;
    var startChar = -1;
    if (TryGetLineInfo(moduleName, lineNumber, out var lineText))
    {
      var noComment = StripInlineComment(lineText);
      startChar = GetTokenIndex(noComment, clean, occurrenceIndex);
      if (startChar >= 0)
        occurrenceIndex = GetOccurrenceIndex(noComment, clean, startChar, lineNumber);
    }

    if (startChar < 0)
      return;

    if (allTypes.TryGetValue(clean, out var t))
      {
        t.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
        t.References.AddLineNumber(moduleName, procedureName, lineNumber, occurrenceIndex, startChar);
      }

      if (allEnums.TryGetValue(clean, out var e))
      {
        e.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
        e.References.AddLineNumber(moduleName, procedureName, lineNumber, occurrenceIndex, startChar);
      }

      // Traccia anche le classi usate come tipo
      if (allClasses.TryGetValue(clean, out var cls))
      {
        cls.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
        cls.References.AddLineNumber(moduleName, procedureName, lineNumber, occurrenceIndex, startChar);
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
