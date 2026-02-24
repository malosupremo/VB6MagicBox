using System.Text;
using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox;

/// <summary>
/// Modulo 3 — Aggiunta dei tipi mancanti al codice VB6.
///
/// Usa il modello analizzato dalla Fase 1 per identificare con precisione
/// i simboli privi di annotazione di tipo, poi applica le correzioni
/// alle righe esatte del sorgente:
///   - Variabili (Dim / Private / …) senza "As Tipo" → "As Object"
///     oppure tipo ricavato dal suffisso VB6 ($, %, &amp;, !, #, @)
///   - Costanti (Const) senza "As Tipo" → tipo inferito dal valore letterale
///   - Parametri di Sub/Function/Property senza "As Tipo" → "As Object"
///     oppure tipo ricavato dal suffisso VB6
///
/// Il parser (Fase 1) cattura già costanti e parametri senza tipo (Type = "").
/// Per le variabili, il parser è stato esteso con regex fallback che aggiungono
/// al modello anche le variabili prive di "As tipo" (ReGlobalVarNoType, ReLocalVarNoType).
/// </summary>
public static class TypeAnnotator
{
  // -------------------------
  // MAPPATURA SUFFISSI TIPO VB6
  // -------------------------

  /// <summary>Suffissi tipo VB6 → tipo esplicito corrispondente.</summary>
  private static readonly Dictionary<char, string> TypeSuffixMap = new()
  {
    ['$'] = "String",
    ['%'] = "Integer",
    ['&'] = "Long",
    ['!'] = "Single",
    ['#'] = "Double",
    ['@'] = "Currency",
  };

  // -------------------------
  // REGEX (solo per l'applicazione delle fix, non per la discovery)
  // -------------------------

  private static readonly Regex ReVarKeyword = new(
    @"^((?:Public|Private|Friend|Global|Dim|Static)\s+(?:WithEvents\s+)?)",
    RegexOptions.IgnoreCase);

  private static readonly Regex ReVarSegment = new(
    @"^(WithEvents\s+)?(\w+)([$%&!#@]?)(\([^)]*\))?\s*$",
    RegexOptions.IgnoreCase);

  private static readonly Regex ReConstNoAs = new(
    @"^((?:Public|Private|Friend|Global)?\s*Const\s+)(\w+[$%&!#@]?)\s*=\s*(.+)$",
    RegexOptions.IgnoreCase);

  private static readonly Regex ReConstHasAs = new(
    @"^(?:Public|Private|Friend|Global)?\s*Const\s+\w+[$%&!#@]?\s+As\s+",
    RegexOptions.IgnoreCase);

  // -------------------------
  // TIPI INTERNI
  // -------------------------

  private enum FixKind { VariableOrConstant, Parameter }
  private sealed record SymbolFix(int LineNumber, FixKind Kind, string Name, string Module, string Procedure);
  private sealed record MissingTypeInfo(string Module, string Procedure, string Name, string ConventionalName, string Kind);

  // -------------------------
  // API PUBBLICA
  // -------------------------

  /// <summary>
  /// Usa il modello già analizzato dalla Fase 1 per identificare i simboli
  /// privi di tipo e aggiunge le annotazioni mancanti ai file sorgente.
  /// Copre variabili globali/locali, costanti e parametri di Sub/Function/Property.
  /// </summary>
  public static void AddMissingTypes(VbProject project)
  {
    ArgumentNullException.ThrowIfNull(project);

    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    var enc = Encoding.GetEncoding(1252);

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine();
    Console.WriteLine("===========================================");
    Console.WriteLine("  2: Aggiunta Tipi Mancanti");
    Console.WriteLine("===========================================");
    Console.WriteLine();
    Console.ForegroundColor= ConsoleColor.Gray;

    var vbpDir        = Path.GetDirectoryName(Path.GetFullPath(project.ProjectFile))!;
    var vbpName       = Path.GetFileNameWithoutExtension(project.ProjectFile);
    var backupBaseDir = new DirectoryInfo(vbpDir).Parent?.FullName ?? vbpDir;
    var folderName    = new DirectoryInfo(backupBaseDir).Name;
    var backupDir     = Path.Combine(
      Path.GetDirectoryName(backupBaseDir)!,
      $"{folderName}.typefix{DateTime.Now:yyyyMMdd_HHmmss}");
    Directory.CreateDirectory(backupDir);

    Console.WriteLine($">> Backup in: {backupDir}");
    Console.WriteLine();

    int totalFiles   = 0;
    int totalChanges = 0;
    var missingTypes = new List<MissingTypeInfo>();

    foreach (var mod in project.Modules)
    {
      if (mod.IsSharedExternal)
        continue;

      var fixes = CollectFixes(mod, missingTypes);

      var filePath = mod.FullPath;
      if (!File.Exists(filePath)) continue;

      var originalContent = File.ReadAllText(filePath, enc);
      var (changes, newContent) = ProcessFileWithFixes(originalContent, fixes, missingTypes, mod);
      if (changes == 0) continue;

      // Backup del file originale
      var rel        = Path.GetRelativePath(backupBaseDir, filePath);
      var backupPath = Path.Combine(backupDir, rel);
      Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);
      File.Copy(filePath, backupPath, overwrite: true);

      File.WriteAllText(filePath, newContent, enc);
      Console.WriteLine($"   [MODIFICATO] {mod.Name}: {changes} tipo/i aggiunto/i");
      totalFiles++;
      totalChanges += changes;
    }

    var missingTypesPath = Path.Combine(vbpDir, $"{vbpName}.missingTypes.csv");
    ExportMissingTypesCsv(missingTypesPath, missingTypes);

    Console.WriteLine();
    if (totalChanges == 0)
      "[OK] Nessun tipo mancante trovato.".WriteLineColored(ConsoleColor.Green);
    else
      $"[OK] {totalChanges} tipo/i aggiunto/i in {totalFiles} file/i.".WriteLineColored(ConsoleColor.Green);

    if (missingTypes.Count > 0)
      $"[WARN] Tipi non deducibili: {missingTypesPath}".WriteLineColored(ConsoleColor.Yellow);
  }

  // -------------------------
  // RACCOLTA FIX DAL MODELLO
  // -------------------------

  /// <summary>
  /// Raccoglie tutti i simboli privi di tipo dal modello del modulo.
  /// Restituisce una lista di fix con numero di riga, tipo di fix e nome del simbolo.
  /// </summary>
  private static List<SymbolFix> CollectFixes(VbModule mod, List<MissingTypeInfo> missingTypes)
  {
    var fixes = new List<SymbolFix>();

    // Variabili globali/membro senza tipo (catturate dal parser fallback ReGlobalVarNoType)
    foreach (var v in mod.GlobalVariables.Where(v => string.IsNullOrEmpty(v.Type)))
    {
      fixes.Add(new SymbolFix(v.LineNumber, FixKind.VariableOrConstant, v.Name, mod.Name, string.Empty));
      if (!HasTypeSuffix(v.Name))
        missingTypes.Add(new MissingTypeInfo(mod.Name, string.Empty, v.Name, v.ConventionalName, "GlobalVariable"));
    }

    // Costanti di modulo senza tipo (il parser cattura già Type="" per Const senza As)
    foreach (var c in mod.Constants.Where(c => string.IsNullOrEmpty(c.Type)))
      fixes.Add(new SymbolFix(c.LineNumber, FixKind.VariableOrConstant, c.Name, mod.Name, string.Empty));

    foreach (var proc in mod.Procedures)
    {
      if (proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase) &&
          string.IsNullOrEmpty(proc.ReturnType))
      {
        missingTypes.Add(new MissingTypeInfo(mod.Name, proc.Name, proc.Name, proc.ConventionalName, "FunctionReturn"));
      }

      // Parametri senza tipo (il parser cattura già Type="" per params senza As)
      foreach (var p in proc.Parameters.Where(p => string.IsNullOrEmpty(p.Type)))
      {
        fixes.Add(new SymbolFix(p.LineNumber, FixKind.Parameter, p.Name, mod.Name, proc.Name));
        if (!HasTypeSuffix(p.Name))
          missingTypes.Add(new MissingTypeInfo(mod.Name, proc.Name, p.Name, p.ConventionalName, "Parameter"));
      }

      // Variabili locali senza tipo (catturate dal parser fallback ReLocalVarNoType)
      foreach (var v in proc.LocalVariables.Where(v => string.IsNullOrEmpty(v.Type)))
      {
        fixes.Add(new SymbolFix(v.LineNumber, FixKind.VariableOrConstant, v.Name, mod.Name, proc.Name));
        if (!HasTypeSuffix(v.Name))
          missingTypes.Add(new MissingTypeInfo(mod.Name, proc.Name, v.Name, v.ConventionalName, "LocalVariable"));
      }

      // Costanti locali senza tipo
      foreach (var c in proc.Constants.Where(c => string.IsNullOrEmpty(c.Type)))
        fixes.Add(new SymbolFix(c.LineNumber, FixKind.VariableOrConstant, c.Name, mod.Name, proc.Name));
    }

    // Parametri di Property senza tipo
    foreach (var prop in mod.Properties)
      foreach (var p in prop.Parameters.Where(p => string.IsNullOrEmpty(p.Type)))
      {
        fixes.Add(new SymbolFix(p.LineNumber, FixKind.Parameter, p.Name, mod.Name, prop.Name));
        if (!HasTypeSuffix(p.Name))
          missingTypes.Add(new MissingTypeInfo(mod.Name, prop.Name, p.Name, p.ConventionalName, "PropertyParameter"));
      }

    foreach (var prop in mod.Properties.Where(p =>
               p.Kind.Equals("Get", StringComparison.OrdinalIgnoreCase) &&
               string.IsNullOrEmpty(p.ReturnType)))
    {
      missingTypes.Add(new MissingTypeInfo(mod.Name, prop.Name, prop.Name, prop.ConventionalName, "PropertyReturn"));
    }

    return fixes;
  }

  // -------------------------
  // ELABORAZIONE FILE
  // -------------------------

  /// <summary>
  /// Applica le fix al testo del file, modificando solo le righe indicate dal modello.
  /// </summary>
  private static (int changes, string newContent) ProcessFileWithFixes(
    string content, List<SymbolFix> fixes, List<MissingTypeInfo> missingTypes, VbModule mod)
  {
    // Raggruppa le fix per numero di riga
    var varConstLines = fixes
      .Where(f => f.Kind == FixKind.VariableOrConstant)
      .Select(f => f.LineNumber)
      .ToHashSet();

    var fixesByLine = fixes
      .GroupBy(f => f.LineNumber)
      .ToDictionary(g => g.Key, g => g.ToList());

    var paramsByLine = fixes
      .Where(f => f.Kind == FixKind.Parameter)
      .GroupBy(f => f.LineNumber)
      .ToDictionary(g => g.Key, g => g.Select(f => f.Name).ToList());

    var procedureLines = mod.Procedures.Select(p => p.LineNumber).ToHashSet();
    var propertyLines = mod.Properties.Select(p => p.LineNumber).ToHashSet();
    var constantLines = mod.Constants.Select(c => c.LineNumber).ToHashSet();
    var globalVariableLines = mod.GlobalVariables.Select(v => v.LineNumber).ToHashSet();
    var memberRanges = mod.Procedures.Select(p => (p.StartLine, p.EndLine))
      .Concat(mod.Properties.Select(p => (p.StartLine, p.EndLine)))
      .Where(r => r.StartLine > 0 && r.EndLine > 0)
      .ToList();

    bool applyVisibilityFixes = mod.IsClass || mod.IsForm;

    var lines   = content.Split('\n');
    int changes = 0;

    for (int i = 0; i < lines.Length; i++)
    {
      var lineNumber = i + 1;
      var line       = lines[i];
      var cr         = line.EndsWith('\r') ? "\r" : "";
      var clean      = cr.Length > 0 ? line[..^1] : line;

      var processed = clean;

      // Fix variabili/costanti: il trasformatore testo gestisce suffissi e multi-var
      if (varConstLines.Contains(lineNumber))
      {
        var (lineResult, missingConstantName) = ProcessLine(processed);
        processed = lineResult;

        if (!string.IsNullOrEmpty(missingConstantName) &&
            fixesByLine.TryGetValue(lineNumber, out var lineFixes))
        {
          var lineFix = lineFixes.FirstOrDefault(f => f.Kind == FixKind.VariableOrConstant);
          if (lineFix != null)
          {
            var kind = string.IsNullOrEmpty(lineFix.Procedure) ? "Constant" : "LocalConstant";
            var conventionalName = ResolveConstantConventionalName(mod, lineFix.Procedure, missingConstantName);
            missingTypes.Add(new MissingTypeInfo(lineFix.Module, lineFix.Procedure, missingConstantName, conventionalName, kind));
          }
        }
      }

      // Fix parametri: aggiunta mirata per nome, nuova funzionalità grazie al modello
      if (paramsByLine.TryGetValue(lineNumber, out var paramNames))
        processed = ApplyParameterFixes(processed, paramNames);

      processed = ApplyVisibilityFixes(processed,
        lineNumber,
        applyVisibilityFixes,
        procedureLines,
        propertyLines,
        constantLines,
        globalVariableLines,
        memberRanges);

      processed = ApplyCallRemoval(processed);
      processed = ApplyForStepCleanup(processed);

      if (!string.Equals(clean, processed, StringComparison.Ordinal))
      {
        lines[i] = processed + cr;
        changes++;
      }
    }

    return (changes, string.Join("\n", lines));
  }

  // -------------------------
  // TRASFORMATORE TESTO — VARIABILI E COSTANTI
  // -------------------------

  /// <summary>Applica le fix di tipo a una singola riga VB6. Internal per i test.</summary>
  internal static (string processed, string? missingConstantName) ProcessLine(string line)
  {
    var (code, comment) = SplitCodeAndComment(line);
    var trimmed         = code.TrimStart();

    if (string.IsNullOrWhiteSpace(trimmed) || code.TrimEnd().EndsWith("_"))
      return (line, null);

    var (constResult, missingName) = TryFixConstant(line, trimmed, comment);
    if (constResult != null)
      return (constResult, missingName);

    var varResult = TryFixVariable(line, trimmed, comment);
    return varResult != null ? (varResult, null) : (line, missingName);
  }

  private static string ApplyVisibilityFixes(
    string line,
    int lineNumber,
    bool applyVisibilityFixes,
    HashSet<int> procedureLines,
    HashSet<int> propertyLines,
    HashSet<int> constantLines,
    HashSet<int> globalVariableLines,
    List<(int StartLine, int EndLine)> memberRanges)
  {
    if (!applyVisibilityFixes || string.IsNullOrWhiteSpace(line))
      return line;

    var (code, comment) = SplitCodeAndComment(line);
    var trimmed = code.TrimStart();
    if (string.IsNullOrWhiteSpace(trimmed))
      return line;

    var indent = code[..(code.Length - trimmed.Length)];

    if ((procedureLines.Contains(lineNumber) || propertyLines.Contains(lineNumber)) &&
        StartsWithProcedureKeyword(trimmed) &&
        !StartsWithVisibility(trimmed))
    {
      var updated = indent + "Public " + trimmed;
      return string.IsNullOrEmpty(comment) ? updated : updated + " " + comment;
    }

    if (globalVariableLines.Contains(lineNumber) &&
        trimmed.StartsWith("Dim ", StringComparison.OrdinalIgnoreCase) &&
        !IsInsideMember(lineNumber, memberRanges))
    {
      var updated = indent + "Private " + trimmed.Substring(4);
      return string.IsNullOrEmpty(comment) ? updated : updated + " " + comment;
    }

    if (constantLines.Contains(lineNumber) &&
        trimmed.StartsWith("Const ", StringComparison.OrdinalIgnoreCase) &&
        !StartsWithVisibility(trimmed) &&
        !IsInsideMember(lineNumber, memberRanges))
    {
      var updated = indent + "Private " + trimmed;
      return string.IsNullOrEmpty(comment) ? updated : updated + " " + comment;
    }

    return line;
  }

  private static string ApplyCallRemoval(string line)
  {
    if (string.IsNullOrWhiteSpace(line))
      return line;

    var (code, comment) = SplitCodeAndComment(line);
    var trimmed = code.TrimStart();

    if (!trimmed.StartsWith("Call ", StringComparison.OrdinalIgnoreCase))
      return line;

    var indent = code[..(code.Length - trimmed.Length)];
    var callPart = trimmed.Substring(5).TrimStart();
    if (string.IsNullOrEmpty(callPart))
      return line;

    var parenIndex = callPart.IndexOf('(');
    if (parenIndex < 0)
    {
      var updated = indent + callPart;
      return string.IsNullOrEmpty(comment) ? updated : updated + " " + comment;
    }

    var lastNonSpace = callPart.Length - 1;
    while (lastNonSpace >= 0 && char.IsWhiteSpace(callPart[lastNonSpace]))
      lastNonSpace--;

    if (lastNonSpace < 0 || callPart[lastNonSpace] != ')')
      return line;

    var endParenIndex = callPart.LastIndexOf(')', lastNonSpace);
    if (endParenIndex <= parenIndex)
      return line;

    var target = callPart.Substring(0, parenIndex).TrimEnd();
    var args = callPart.Substring(parenIndex + 1, endParenIndex - parenIndex - 1).Trim();
    var updatedCall = string.IsNullOrEmpty(args) ? target : $"{target} {args}";
    var updatedLine = indent + updatedCall;
    return string.IsNullOrEmpty(comment) ? updatedLine : updatedLine + " " + comment;
  }

  private static string ApplyForStepCleanup(string line)
  {
    if (string.IsNullOrWhiteSpace(line))
      return line;

    var (code, comment) = SplitCodeAndComment(line);
    var trimmed = code.TrimStart();
    if (!trimmed.StartsWith("For ", StringComparison.OrdinalIgnoreCase))
      return line;

    var updatedCode = Regex.Replace(code, @"\s+Step\s+1\s*$", "", RegexOptions.IgnoreCase);
    if (string.Equals(code, updatedCode, StringComparison.Ordinal))
      return line;

    return string.IsNullOrEmpty(comment) ? updatedCode : updatedCode + " " + comment;
  }

  private static bool StartsWithVisibility(string trimmed)
    => trimmed.StartsWith("Public ", StringComparison.OrdinalIgnoreCase) ||
       trimmed.StartsWith("Private ", StringComparison.OrdinalIgnoreCase) ||
       trimmed.StartsWith("Friend ", StringComparison.OrdinalIgnoreCase);

  private static bool StartsWithProcedureKeyword(string trimmed)
  {
    if (trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
        trimmed.StartsWith("Function ", StringComparison.OrdinalIgnoreCase) ||
        trimmed.StartsWith("Property ", StringComparison.OrdinalIgnoreCase))
      return true;

    if (trimmed.StartsWith("Static ", StringComparison.OrdinalIgnoreCase))
    {
      var afterStatic = trimmed.Substring("Static ".Length).TrimStart();
      return afterStatic.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
             afterStatic.StartsWith("Function ", StringComparison.OrdinalIgnoreCase) ||
             afterStatic.StartsWith("Property ", StringComparison.OrdinalIgnoreCase);
    }

    return false;
  }

  private static bool IsInsideMember(int lineNumber, List<(int StartLine, int EndLine)> memberRanges)
  {
    foreach (var (start, end) in memberRanges)
    {
      if (lineNumber >= start && lineNumber <= end)
        return true;
    }

    return false;
  }

  private static (string? result, string? missingConstantName) TryFixConstant(
    string originalLine, string trimmed, string comment)
  {
    if (ReConstHasAs.IsMatch(trimmed)) return (null, null);

    var m = ReConstNoAs.Match(trimmed);
    if (!m.Success) return (null, null);

    var keyword  = m.Groups[1].Value;
    var rawName  = m.Groups[2].Value;
    var rawValue = m.Groups[3].Value.TrimEnd();

    // Salta Const con lista di costanti sulla stessa riga (Const A=1, B=2)
    if (!rawValue.TrimStart().StartsWith('"') &&
        Regex.IsMatch(rawValue, @",\s*\w+\s*=", RegexOptions.IgnoreCase))
      return (null, null);

    string typeName;
    string cleanName;
    if (rawName.Length > 0 && TypeSuffixMap.TryGetValue(rawName[^1], out var suffixType))
    {
      cleanName = rawName[..^1];
      typeName  = suffixType;
    }
    else
    {
      cleanName = rawName;
      typeName  = InferConstantType(rawValue)??string.Empty;
      if (string.IsNullOrEmpty(typeName))
        return (null, cleanName);
    }

    var indent  = originalLine[..(originalLine.Length - originalLine.TrimStart().Length)];
    var newLine = $"{indent}{keyword}{cleanName} As {typeName} = {rawValue}";
    if (!string.IsNullOrEmpty(comment))
      newLine += " " + comment;

    return (newLine, null);
  }

  private static string? TryFixVariable(string originalLine, string trimmed, string comment)
  {
    var kwMatch = ReVarKeyword.Match(trimmed);
    if (!kwMatch.Success) return null;

    var keyword = kwMatch.Groups[1].Value;
    var rest    = trimmed[keyword.Length..].TrimEnd();

    if (string.IsNullOrWhiteSpace(rest)) return null;

    var segments    = SplitTopLevel(rest);
    bool anyChanged = false;
    var fixedSegs   = new List<string>(segments.Count);

    foreach (var seg in segments)
    {
      var fixedSeg = TryFixVarSegment(seg, out bool changed);
      fixedSegs.Add(fixedSeg);
      if (changed) anyChanged = true;
    }

    if (!anyChanged) return null;

    var indent  = originalLine[..(originalLine.Length - originalLine.TrimStart().Length)];
    var newLine = indent + keyword + string.Join(", ", fixedSegs);
    if (!string.IsNullOrEmpty(comment))
      newLine += " " + comment;

    return newLine;
  }

  private static string TryFixVarSegment(string segment, out bool changed)
  {
    var s = segment.Trim();

    if (Regex.IsMatch(s, @"\bAs\b", RegexOptions.IgnoreCase))
    {
      changed = false;
      return segment;
    }

    var m = ReVarSegment.Match(s);
    if (!m.Success)
    {
      changed = false;
      return segment;
    }

    var withEvents = m.Groups[1].Value;
    var name       = m.Groups[2].Value;
    var suffix     = m.Groups[3].Value;
    var arrayDims  = m.Groups[4].Value;

    if (suffix.Length == 0)
    {
      changed = false;
      return segment;
    }

    var typeName = TypeSuffixMap.TryGetValue(suffix[0], out var st) ? st : null;
    if (string.IsNullOrEmpty(typeName))
    {
      changed = false;
      return segment;
    }

    changed = true;
    return $"{withEvents}{name}{arrayDims} As {typeName}";
  }

  // -------------------------
  // TRASFORMATORE TESTO — PARAMETRI (nuova funzionalità)
  // -------------------------

  /// <summary>Aggiunge "As tipo" ai parametri elencati nella riga. Internal per i test.</summary>
  internal static string ApplyParameterFixes(string line, List<string> paramNames)
  {
    var (code, comment) = SplitCodeAndComment(line);

    // Salta righe di continuazione
    if (code.TrimEnd().EndsWith("_")) return line;

    var result = code;
    foreach (var paramName in paramNames)
      result = ApplySingleParameterFix(result, paramName);

    if (string.Equals(code, result, StringComparison.Ordinal))
      return line;

    return string.IsNullOrEmpty(comment) ? result : result + " " + comment;
  }

  /// <summary>
  /// Aggiunge "As tipo" a un singolo parametro non tipizzato nella firma di una procedura.
  /// Rileva il suffisso tipo VB6 (es. x$) e lo sostituisce con il tipo esplicito.
  /// Usa word-boundary per non toccare nomi che iniziano con lo stesso prefisso.
  /// </summary>
  private static string ApplySingleParameterFix(string code, string paramName)
  {
    var pattern = $@"\b{Regex.Escape(paramName)}\b([$%&!#@]?)(\([^)]*\))?(?=\s*[,)]|\s*$)";
    return Regex.Replace(code, pattern, m =>
    {
      var suffix   = m.Groups[1].Value;
      var dims     = m.Groups[2].Value;
      if (suffix.Length == 0)
        return m.Value;

      var typeName = TypeSuffixMap.TryGetValue(suffix[0], out var st) ? st : null;
      return string.IsNullOrEmpty(typeName) ? m.Value : $"{paramName}{dims} As {typeName}";
    }, RegexOptions.IgnoreCase);
  }

  private static bool HasTypeSuffix(string name)
  {
    if (string.IsNullOrEmpty(name)) return false;
    return TypeSuffixMap.ContainsKey(name[^1]);
  }

  private static void ExportMissingTypesCsv(string outputPath, List<MissingTypeInfo> missingTypes)
  {
    var lines = new List<string> { "Module,Procedure,Name,ConventionalName,Kind" };

    foreach (var item in missingTypes)
    {
      lines.Add($"\"{EscapeCsv(item.Module)}\",\"{EscapeCsv(item.Procedure)}\",\"{EscapeCsv(item.Name)}\",\"{EscapeCsv(item.ConventionalName)}\",\"{EscapeCsv(item.Kind)}\"");
    }

    File.WriteAllLines(outputPath, lines, Encoding.UTF8);
  }

  private static string EscapeCsv(string value)
  {
    if (string.IsNullOrEmpty(value)) return string.Empty;
    return value.Replace("\"", "\"\"");
  }

  private static string ResolveConstantConventionalName(VbModule mod, string procedureName, string constantName)
  {
    if (string.IsNullOrEmpty(constantName))
      return constantName;

    if (string.IsNullOrEmpty(procedureName))
    {
      var moduleConst = mod.Constants.FirstOrDefault(c =>
          c.Name.Equals(constantName, StringComparison.OrdinalIgnoreCase));
      return moduleConst?.ConventionalName ?? constantName;
    }

    var proc = mod.Procedures.FirstOrDefault(p =>
        p.Name.Equals(procedureName, StringComparison.OrdinalIgnoreCase));
    var procConst = proc?.Constants.FirstOrDefault(c =>
        c.Name.Equals(constantName, StringComparison.OrdinalIgnoreCase));

    return procConst?.ConventionalName ?? constantName;
  }

  // -------------------------
  // INFERENZA TIPO COSTANTE
  // -------------------------

  /// <summary>
  /// Inferisce il tipo VB6 di una costante dal valore letterale grezzo nel sorgente.
  /// Restituisce null se il tipo non è determinabile.
  /// </summary>
  private static string? InferConstantType(string rawValue)
  {
    var v = rawValue.Trim();
    if (string.IsNullOrEmpty(v)) return null;

    // Stringa letterale (con virgolette nel sorgente)
    if (v.StartsWith('"')) return "String";

    // Booleano
    if (v.Equals("True",  StringComparison.OrdinalIgnoreCase) ||
        v.Equals("False", StringComparison.OrdinalIgnoreCase))
      return "Boolean";

    // Suffisso tipo nel valore stesso (es. 1.5!, 100&, 3.14#)
    if (v.Length > 1 && TypeSuffixMap.TryGetValue(v[^1], out var suffixTypeFromValue))
      return suffixTypeFromValue;

    // Letterale esadecimale: &Hffff
    // VB6 classifica per ampiezza in bit (NON per valore signed):
    //   &H0000..&HFFFF     (16 bit) → Integer  (&HFFFF = -1 in VB6)
    //   &H10000..&HFFFFFFFF (32 bit) → Long
    // N.B. i suffissi tipo (&HFFFF& → Long, &HFFFF% → Integer) sono già
    //      gestiti sopra dal controllo TypeSuffixMap sul valore.
    if (v.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
    {
      var hexLiteral = v[2..].Trim();
      if (TryStripNumericSuffix(hexLiteral, out var hexStr) && IsPureHexLiteral(hexStr) &&
          long.TryParse(hexStr,
            System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture,
            out long hexVal) && hexVal >= 0)
      {
        if (hexVal <= 0xFFFFL)     return "Integer";  // ≤ &HFFFF  → 16 bit
        if (hexVal <= 0xFFFFFFFFL) return "Long";     // ≤ &HFFFFFFFF → 32 bit
      }
      return "Long";
    }

    // Letterale ottale: &O777
    // Stessa logica: 16-bit (≤ &O177777 = 65535) → Integer, oltre → Long
    if (v.StartsWith("&O", StringComparison.OrdinalIgnoreCase))
    {
      var octLiteral = v[2..].Trim();
      if (TryStripNumericSuffix(octLiteral, out var octStr) && IsPureOctLiteral(octStr))
      {
        try
        {
          var octVal = Convert.ToInt64(octStr, 8);
          if (octVal >= 0 && octVal <= 0xFFFFL)     return "Integer";
          if (octVal >= 0 && octVal <= 0xFFFFFFFFL) return "Long";
        }
        catch (Exception) { }
      }
      return null;
    }

    // Intero decimale: dimensione determina Integer vs Long
    if (long.TryParse(v,
          System.Globalization.NumberStyles.Integer,
          System.Globalization.CultureInfo.InvariantCulture,
          out long intVal))
    {
      if (intVal >= short.MinValue && intVal <= short.MaxValue) return "Integer";
      if (intVal >= int.MinValue   && intVal <= int.MaxValue)   return "Long";
      return "Double";
    }

    // Virgola mobile
    if (double.TryParse(v,
          System.Globalization.NumberStyles.Float,
          System.Globalization.CultureInfo.InvariantCulture,
          out _))
      return "Double";

    return null;
  }

  // -------------------------
  // HELPER
  // -------------------------

  /// <summary>
  /// Separa la parte di codice dalla parte di commento di una riga VB6.
  /// Gestisce correttamente le stringhe letterali (le "" di escape non terminano la stringa).
  /// </summary>
  private static (string code, string comment) SplitCodeAndComment(string line)
  {
    bool inString = false;
    for (int i = 0; i < line.Length; i++)
    {
      var ch = line[i];
      if (ch == '"')
      {
        if (!inString)
          inString = true;
        else if (i + 1 < line.Length && line[i + 1] == '"')
          i++; // Virgoletta doppia: escape, rimane in stringa
        else
          inString = false;
      }
      else if (!inString && ch == '\'')
        return (line[..i].TrimEnd(), line[i..]);
    }
    return (line, string.Empty);
  }

  private static bool TryStripNumericSuffix(string value, out string stripped)
  {
    stripped = value.TrimEnd('&', '%', '!', '#', '@', 'L', 'l');
    return !string.IsNullOrWhiteSpace(stripped) && stripped.Length != value.Length
      ? true
      : !string.IsNullOrWhiteSpace(stripped);
  }

  private static bool IsPureHexLiteral(string value)
  {
    foreach (var ch in value)
    {
      if (!Uri.IsHexDigit(ch))
        return false;
    }

    return value.Length > 0;
  }

  private static bool IsPureOctLiteral(string value)
  {
    foreach (var ch in value)
    {
      if (ch < '0' || ch > '7')
        return false;
    }

    return value.Length > 0;
  }

  /// <summary>
  /// Divide una stringa per le virgole al livello zero di parentesi.
  /// Es. "x(), y As Integer, z(1 To 3)" → ["x()", " y As Integer", " z(1 To 3)"]
  /// </summary>
  private static List<string> SplitTopLevel(string input)
  {
    var segments = new List<string>();
    int depth    = 0;
    int start    = 0;

    for (int i = 0; i < input.Length; i++)
    {
      switch (input[i])
      {
        case '(':  depth++;  break;
        case ')':  depth--;  break;
        case ',' when depth == 0:
          segments.Add(input[start..i]);
          start = i + 1;
          break;
      }
    }
    segments.Add(input[start..]);
    return segments;
  }
}
