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
  /// Calcola lo startChar direttamente cercando il tipo dopo il nome del simbolo.
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
        foreach (var field in typeDef.Fields)
        {
          var lineText = GetLineText(fileLines, field.LineNumber);
          var sc = FindTypeStartChar(lineText, field.Name, field.Type);
          AddTypeReferenceAt(field.Type, field.LineNumber, sc, lineText, mod.Name, string.Empty, typeIndex);
        }
      }

      // 2. Variabili globali: "Public/Dim varName As TYPE"
      foreach (var variable in mod.GlobalVariables)
      {
        var lineText = GetLineText(fileLines, variable.LineNumber);
        var sc = FindTypeStartChar(lineText, variable.Name, variable.Type);
        AddTypeReferenceAt(variable.Type, variable.LineNumber, sc, lineText, mod.Name, string.Empty, typeIndex);
      }

      // 3. Parametri e variabili locali delle procedure
      foreach (var proc in mod.Procedures)
      {
        // Return type: cerca dopo ")" nella riga
        var retLine = proc.ReturnTypeLineNumber > 0 ? proc.ReturnTypeLineNumber : proc.LineNumber;
        var retText = GetLineText(fileLines, retLine);
        var retSc = FindTypeStartChar(retText, null, proc.ReturnType, isReturnType: true);
        AddTypeReferenceAt(proc.ReturnType, retLine, retSc, retText, mod.Name, proc.Name, typeIndex);

        // Parametri: cerca il tipo dopo il nome del parametro
        foreach (var param in proc.Parameters)
        {
          var paramLine = param.TypeLineNumber > 0 ? param.TypeLineNumber : param.LineNumber;
          var paramText = GetLineText(fileLines, paramLine);
          int sc;
          if (param.TypeLineNumber > 0 && param.TypeLineNumber != param.LineNumber)
            sc = FindTypeStartChar(paramText, null, param.Type); // multiline: name on different line
          else
            sc = FindTypeStartChar(paramText, param.Name, param.Type);
          AddTypeReferenceAt(param.Type, paramLine, sc, paramText, mod.Name, proc.Name, typeIndex);
        }

        // Variabili locali
        foreach (var lv in proc.LocalVariables)
        {
          var lineText = GetLineText(fileLines, lv.LineNumber);
          var sc = FindTypeStartChar(lineText, lv.Name, lv.Type);
          AddTypeReferenceAt(lv.Type, lv.LineNumber, sc, lineText, mod.Name, proc.Name, typeIndex);
        }
      }

      // 4. Parametri e return type delle proprietà
      foreach (var prop in mod.Properties)
      {
        var retLine = prop.ReturnTypeLineNumber > 0 ? prop.ReturnTypeLineNumber : prop.LineNumber;
        var retText = GetLineText(fileLines, retLine);
        var retSc = FindTypeStartChar(retText, null, prop.ReturnType, isReturnType: true);
        AddTypeReferenceAt(prop.ReturnType, retLine, retSc, retText, mod.Name, prop.Name, typeIndex);

        foreach (var param in prop.Parameters)
        {
          var paramLine = param.TypeLineNumber > 0 ? param.TypeLineNumber : param.LineNumber;
          var paramText = GetLineText(fileLines, paramLine);
          int sc;
          if (param.TypeLineNumber > 0 && param.TypeLineNumber != param.LineNumber)
            sc = FindTypeStartChar(paramText, null, param.Type);
          else
            sc = FindTypeStartChar(paramText, param.Name, param.Type);
          AddTypeReferenceAt(param.Type, paramLine, sc, paramText, mod.Name, prop.Name, typeIndex);
        }
      }
    }
  }

  /// <summary>
  /// Trova lo startChar del tipo nella riga cercando dopo il nome del simbolo.
  /// Per return type, cerca dopo l'ultima ")".
  /// </summary>
  private static int FindTypeStartChar(string lineText, string? symbolName, string? typeName, bool isReturnType = false)
  {
    if (string.IsNullOrEmpty(lineText) || string.IsNullOrEmpty(typeName))
      return -1;

    var noComment = StripInlineComment(lineText);
    var baseTypeName = typeName.Contains('(')
        ? typeName.Substring(0, typeName.IndexOf('('))
        : typeName;

    int searchFrom = 0;

    if (isReturnType)
    {
      var lastParen = noComment.LastIndexOf(')');
      if (lastParen >= 0)
        searchFrom = lastParen + 1;
    }
    else if (!string.IsNullOrEmpty(symbolName))
    {
      var nameMatch = Regex.Match(noComment, $@"\b{Regex.Escape(symbolName)}\b", RegexOptions.IgnoreCase);
      if (nameMatch.Success)
        searchFrom = nameMatch.Index + nameMatch.Length;
    }

    if (searchFrom >= noComment.Length)
      return -1;

    var sub = noComment.Substring(searchFrom);
    var typeMatch = Regex.Match(sub, $@"\b{Regex.Escape(baseTypeName)}\b", RegexOptions.IgnoreCase);
    if (typeMatch.Success)
      return searchFrom + typeMatch.Index;

    // Prova con/senza suffisso _T
    if (!baseTypeName.EndsWith("_T", StringComparison.OrdinalIgnoreCase))
    {
      typeMatch = Regex.Match(sub, $@"\b{Regex.Escape(baseTypeName + "_T")}\b", RegexOptions.IgnoreCase);
      if (typeMatch.Success)
        return searchFrom + typeMatch.Index;
    }
    else
    {
      var unsuffixed = baseTypeName.Substring(0, baseTypeName.Length - 2);
      typeMatch = Regex.Match(sub, $@"\b{Regex.Escape(unsuffixed)}\b", RegexOptions.IgnoreCase);
      if (typeMatch.Success)
        return searchFrom + typeMatch.Index;
    }

    return -1;
  }

  /// <summary>
  /// Aggiunge una Reference al tipo usando lo startChar già calcolato.
  /// </summary>
  private static void AddTypeReferenceAt(
      string? typeName,
      int lineNumber,
      int startChar,
      string lineText,
      string moduleName,
      string procedureName,
      Dictionary<string, VbTypeDef> typeIndex)
  {
    if (string.IsNullOrEmpty(typeName) || lineNumber <= 0 || startChar < 0)
      return;

    var baseTypeName = typeName.Contains('(')
        ? typeName.Substring(0, typeName.IndexOf('('))
        : typeName;

    if (!typeIndex.TryGetValue(baseTypeName, out var referencedType))
    {
      if (!baseTypeName.EndsWith("_T", StringComparison.OrdinalIgnoreCase))
      {
        if (typeIndex.TryGetValue(baseTypeName + "_T", out referencedType))
          baseTypeName = baseTypeName + "_T";
      }
      else
      {
        var unsuffixed = baseTypeName.Substring(0, baseTypeName.Length - 2);
        if (typeIndex.TryGetValue(unsuffixed, out referencedType))
          baseTypeName = unsuffixed;
      }

      if (referencedType == null)
        return;
    }

    var noComment = StripInlineComment(lineText);
    var occIdx = GetOccurrenceIndex(noComment, baseTypeName, startChar, lineNumber);
    referencedType.Used = true;
    referencedType.References.AddLineNumber(moduleName, procedureName, lineNumber, startChar, owner: referencedType);
  }

  private static string GetLineText(string[] fileLines, int lineNumber)
  {
    return lineNumber > 0 && lineNumber <= fileLines.Length
        ? fileLines[lineNumber - 1]
        : string.Empty;
  }

  // ---------------------------------------------------------
  // RISOLUZIONE RIFERIMENTI ALLE CLASSI (As [New] ClassName)
  // ---------------------------------------------------------

  /// <summary>
  /// Aggiunge References alla VbModule classe per ogni dichiarazione
  /// "Dim/Private x As [New] ClassName" dove ClassName è un modulo classe.
  /// Calcola lo startChar direttamente cercando la classe dopo il nome del simbolo.
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
        var lineText = GetLineText(fileLines, v.LineNumber);
        var sc = FindTypeStartChar(lineText, v.Name, v.Type);
        AddClassReferenceAt(v.Type, v.LineNumber, sc, lineText, mod.Name, string.Empty, classIndex);
      }

      foreach (var proc in mod.Procedures)
      {
        var retLine = proc.ReturnTypeLineNumber > 0 ? proc.ReturnTypeLineNumber : proc.LineNumber;
        var retText = GetLineText(fileLines, retLine);
        var retSc = FindTypeStartChar(retText, null, proc.ReturnType, isReturnType: true);
        AddClassReferenceAt(proc.ReturnType, retLine, retSc, retText, mod.Name, proc.Name, classIndex);

        foreach (var param in proc.Parameters)
        {
          var paramLine = param.TypeLineNumber > 0 ? param.TypeLineNumber : param.LineNumber;
          var paramText = GetLineText(fileLines, paramLine);
          int sc;
          if (param.TypeLineNumber > 0 && param.TypeLineNumber != param.LineNumber)
            sc = FindTypeStartChar(paramText, null, param.Type);
          else
            sc = FindTypeStartChar(paramText, param.Name, param.Type);
          AddClassReferenceAt(param.Type, paramLine, sc, paramText, mod.Name, proc.Name, classIndex);
        }

        foreach (var lv in proc.LocalVariables)
        {
          var lineText = GetLineText(fileLines, lv.LineNumber);
          var sc = FindTypeStartChar(lineText, lv.Name, lv.Type);
          AddClassReferenceAt(lv.Type, lv.LineNumber, sc, lineText, mod.Name, proc.Name, classIndex);
        }
      }

      foreach (var prop in mod.Properties)
      {
        var retLine = prop.ReturnTypeLineNumber > 0 ? prop.ReturnTypeLineNumber : prop.LineNumber;
        var retText = GetLineText(fileLines, retLine);
        var retSc = FindTypeStartChar(retText, null, prop.ReturnType, isReturnType: true);
        AddClassReferenceAt(prop.ReturnType, retLine, retSc, retText, mod.Name, prop.Name, classIndex);

        foreach (var param in prop.Parameters)
        {
          var paramLine = param.TypeLineNumber > 0 ? param.TypeLineNumber : param.LineNumber;
          var paramText = GetLineText(fileLines, paramLine);
          int sc;
          if (param.TypeLineNumber > 0 && param.TypeLineNumber != param.LineNumber)
            sc = FindTypeStartChar(paramText, null, param.Type);
          else
            sc = FindTypeStartChar(paramText, param.Name, param.Type);
          AddClassReferenceAt(param.Type, paramLine, sc, paramText, mod.Name, prop.Name, classIndex);
        }
      }
    }
  }

  private static void AddClassReferenceAt(
      string? typeName,
      int lineNumber,
      int startChar,
      string lineText,
      string declaringModule,
      string procedureName,
      Dictionary<string, VbModule> classIndex)
  {
    if (string.IsNullOrEmpty(typeName) || lineNumber <= 0 || startChar < 0)
      return;

    var baseName = typeName.Contains('.') ? typeName.Split('.').Last() : typeName;

    if (!classIndex.TryGetValue(baseName, out var classModule))
      return;

    if (string.Equals(classModule.Name, declaringModule, StringComparison.OrdinalIgnoreCase))
      return;

    var noComment = StripInlineComment(lineText);
    var occIdx = GetOccurrenceIndex(noComment, baseName, startChar, lineNumber);
    classModule.Used = true;
    classModule.References.AddLineNumber(declaringModule, procedureName, lineNumber, startChar, owner: classModule);
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

    var allClasses = new Dictionary<string, VbModule>(StringComparer.OrdinalIgnoreCase);
    foreach (var mod in project.Modules.Where(m => m.IsClass))
    {
      var className = Path.GetFileNameWithoutExtension(mod.Name);
      allClasses[className] = mod;

      if (className.StartsWith("cls", StringComparison.OrdinalIgnoreCase))
      {
        var withoutPrefix = className.Substring(3);
        if (!allClasses.ContainsKey(withoutPrefix))
          allClasses[withoutPrefix] = mod;
      }
    }

    // Raccoglie tutti i simboli con tipo, calcola startChar una volta sola per ciascuno
    void MarkSymbol(string? typeName, string? symbolName, string moduleName, string procedureName,
                    int lineNumber, string[] fileLines, bool isReturnType = false)
    {
      if (string.IsNullOrWhiteSpace(typeName) || lineNumber <= 0)
        return;

      var clean = typeName.Trim();
      if (clean.Contains('.'))
        clean = clean.Split('.').Last();

      var lineText = GetLineText(fileLines, lineNumber);
      var startChar = FindTypeStartChar(lineText, symbolName, clean, isReturnType);
      if (startChar < 0)
        return;

      var noComment = StripInlineComment(lineText);
      var occIdx = GetOccurrenceIndex(noComment, clean, startChar, lineNumber);

      if (allTypes.TryGetValue(clean, out var t))
      {
        t.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
          t.References.AddLineNumber(moduleName, procedureName, lineNumber, startChar, owner: t);
      }

      if (allEnums.TryGetValue(clean, out var e))
      {
        e.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
          e.References.AddLineNumber(moduleName, procedureName, lineNumber, startChar, owner: e);
      }

      if (allClasses.TryGetValue(clean, out var cls))
      {
        cls.Used = true;
        if (!string.IsNullOrEmpty(moduleName))
          cls.References.AddLineNumber(moduleName, procedureName, lineNumber, startChar, owner: cls);
      }
    }

    foreach (var mod in project.Modules)
    {
      var fileLines = GetFileLines(fileCache, mod);

      foreach (var v in mod.GlobalVariables)
        MarkSymbol(v.Type, v.Name, mod.Name, string.Empty, v.LineNumber, fileLines);

      foreach (var typeDef in mod.Types)
      {
        foreach (var field in typeDef.Fields)
          MarkSymbol(field.Type, field.Name, mod.Name, string.Empty, field.LineNumber, fileLines);
      }

      foreach (var proc in mod.Procedures)
      {
        var retLine = proc.ReturnTypeLineNumber > 0 ? proc.ReturnTypeLineNumber : proc.LineNumber;
        MarkSymbol(proc.ReturnType, null, mod.Name, proc.Name, retLine, fileLines, isReturnType: true);

        foreach (var p in proc.Parameters)
        {
          var paramLine = p.TypeLineNumber > 0 ? p.TypeLineNumber : p.LineNumber;
          string? anchor = (p.TypeLineNumber > 0 && p.TypeLineNumber != p.LineNumber) ? null : p.Name;
          MarkSymbol(p.Type, anchor, mod.Name, proc.Name, paramLine, fileLines);
        }

        foreach (var lv in proc.LocalVariables)
          MarkSymbol(lv.Type, lv.Name, mod.Name, proc.Name, lv.LineNumber, fileLines);
      }

      foreach (var prop in mod.Properties)
      {
        var retLine = prop.ReturnTypeLineNumber > 0 ? prop.ReturnTypeLineNumber : prop.LineNumber;
        MarkSymbol(prop.ReturnType, null, mod.Name, prop.Name, retLine, fileLines, isReturnType: true);

        foreach (var p in prop.Parameters)
        {
          var paramLine = p.TypeLineNumber > 0 ? p.TypeLineNumber : p.LineNumber;
          string? anchor = (p.TypeLineNumber > 0 && p.TypeLineNumber != p.LineNumber) ? null : p.Name;
          MarkSymbol(p.Type, anchor, mod.Name, prop.Name, paramLine, fileLines);
        }
      }
    }
  }
}
