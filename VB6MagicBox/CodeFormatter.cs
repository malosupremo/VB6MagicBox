using System.Text;
using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox;

/// <summary>
/// Modulo 5 — Riordino variabili locali di procedura.
///
/// Per ogni Sub/Function/Property:
///   1. Raccoglie tutte le dichiarazioni Dim/Static dal corpo della procedura
///   2. Le rimuove dalla posizione originale
///   3. Le reinserisce all'inizio del corpo (dopo firma + eventuale commento di intestazione)
///   4. Raggruppa: prima Static (ordine alfabetico), poi Dim volatili (ordine alfa)
///
/// Modulo 4 (FixSpacing) verrà applicato dopo per normalizzare le righe bianche.
/// </summary>
public static class CodeFormatter
{
  private static readonly Regex ReDimOrStatic = new(
    @"^(Dim|Static)\s+\w", RegexOptions.IgnoreCase);

  private static readonly Regex ReVarName = new(
    @"^(?:Dim|Static)\s+(\w+)", RegexOptions.IgnoreCase);

  private static readonly Regex ReConstLocal = new(
    @"^(?:Private\s+|Public\s+)?Const\s+\w", RegexOptions.IgnoreCase);

  // -------------------------
  // API PUBBLICA
  // -------------------------

  /// <summary>
  /// Riordina le variabili locali in tutte le procedure e proprietà del progetto.
  /// Per ogni procedura: Static (alfa) poi Dim (alfa), subito dopo la firma/commento.
  /// </summary>
  public static void ReorderLocalVariables(VbProject project)
  {
    ArgumentNullException.ThrowIfNull(project);

    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    var enc = Encoding.GetEncoding(1252);

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine();
    Console.WriteLine("===========================================");
    Console.WriteLine("  4: Riordino Variabili Locali");
    Console.WriteLine("===========================================");
    Console.WriteLine();
    Console.ForegroundColor = ConsoleColor.Gray;

    var vbpDir        = Path.GetDirectoryName(Path.GetFullPath(project.ProjectFile))!;
    var backupBaseDir = new DirectoryInfo(vbpDir).Parent?.FullName ?? vbpDir;
    var folderName    = new DirectoryInfo(backupBaseDir).Name;
    var backupDir     = Path.Combine(
      Path.GetDirectoryName(backupBaseDir)!,
      $"{folderName}.varorder{DateTime.Now:yyyyMMdd_HHmmss}");
    Directory.CreateDirectory(backupDir);

    Console.WriteLine($">> Backup in: {backupDir}");
    Console.WriteLine();

    int totalFiles = 0;
    int totalProcs = 0;

    foreach (var mod in project.Modules)
    {
      var filePath = mod.FullPath;
      if (!File.Exists(filePath)) continue;

      // Costruisci la lista unificata dei range procedura/proprietà
      var ranges = CollectProcedureRanges(mod);
      if (ranges.Count == 0) continue;

      var originalContent = File.ReadAllText(filePath, enc);
      var lineSeparator = originalContent.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
      var lines = originalContent.Contains("\r\n", StringComparison.Ordinal)
        ? originalContent.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList()
        : originalContent.Split('\n').Select(l => l.EndsWith('\r') ? l[..^1] : l).ToList();
      bool anyChanged = false;
      int procsChanged = 0;

      // Processa dall'ultima procedura alla prima per non invalidare gli indici
      foreach (var (startLine, endLine) in ranges)
      {
        int startIdx = startLine - 1;
        int endIdx   = endLine - 1;

        if (startIdx < 0 || endIdx >= lines.Count || startIdx >= endIdx)
          continue;

        var procLines    = lines.GetRange(startIdx, endIdx - startIdx + 1);
        var newProcLines = ReorderDimsInProcedure(procLines);

        if (newProcLines == null) continue;

        lines.RemoveRange(startIdx, endIdx - startIdx + 1);
        lines.InsertRange(startIdx, newProcLines);
        anyChanged = true;
        procsChanged++;
      }

      if (!anyChanged) continue;

      // Backup
      var rel        = Path.GetRelativePath(backupBaseDir, filePath);
      var backupPath = Path.Combine(backupDir, rel);
      Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);
      File.Copy(filePath, backupPath, overwrite: true);

      File.WriteAllText(filePath, string.Join(lineSeparator, lines), enc);
      Console.WriteLine($"   [MODIFICATO] {mod.Name}: {procsChanged} procedura/e riordinata/e");
      totalFiles++;
      totalProcs += procsChanged;
    }

    Console.WriteLine();
    if (totalProcs == 0)
      Console.WriteLine("[OK] Nessuna variabile da riordinare.");
    else
      Console.WriteLine($"[OK] {totalProcs} procedura/e riordinata/e in {totalFiles} file/i.");
  }

  /// <summary>
  /// Armonizza le righe bianche secondo le regole di spacing.
  /// </summary>
  public static void HarmonizeSpacing(VbProject project)
  {
    ArgumentNullException.ThrowIfNull(project);

    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    var enc = Encoding.GetEncoding(1252);

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine();
    Console.WriteLine("===========================================");
    Console.WriteLine("  5: Armonizzazione Spaziature");
    Console.WriteLine("===========================================");
    Console.WriteLine();
    Console.ForegroundColor = ConsoleColor.Gray;

    var vbpDir = Path.GetDirectoryName(Path.GetFullPath(project.ProjectFile))!;
    var backupBaseDir = new DirectoryInfo(vbpDir).Parent?.FullName ?? vbpDir;
    var folderName = new DirectoryInfo(backupBaseDir).Name;
    var backupDir = Path.Combine(
      Path.GetDirectoryName(backupBaseDir)!,
      $"{folderName}.spacing{DateTime.Now:yyyyMMdd_HHmmss}");
    Directory.CreateDirectory(backupDir);

    Console.WriteLine($">> Backup in: {backupDir}");
    Console.WriteLine();

    int totalFiles = 0;

    foreach (var mod in project.Modules)
    {
      var filePath = mod.FullPath;
      if (!File.Exists(filePath)) continue;

      var originalContent = File.ReadAllText(filePath, enc);
      var lineSeparator = originalContent.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
      var lines = originalContent.Contains("\r\n", StringComparison.Ordinal)
        ? originalContent.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList()
        : originalContent.Split('\n').Select(l => l.EndsWith('\r') ? l[..^1] : l).ToList();
      var newLines = ApplySpacingRules(lines, mod);

      if (newLines == null) continue;

      var rel = Path.GetRelativePath(backupBaseDir, filePath);
      var backupPath = Path.Combine(backupDir, rel);
      Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);
      File.Copy(filePath, backupPath, overwrite: true);

      File.WriteAllText(filePath, string.Join(lineSeparator, newLines), enc);
      Console.WriteLine($"   [MODIFICATO] {mod.Name}");
      totalFiles++;
    }

    Console.WriteLine();
    if (totalFiles == 0)
      Console.WriteLine("[OK] Nessuna modifica alle spaziature.");
    else
      Console.WriteLine($"[OK] Spaziature armonizzate in {totalFiles} file/i.");
  }

  // -------------------------
  // RACCOLTA RANGE PROCEDURA
  // -------------------------

  private static List<(int start, int end)> CollectProcedureRanges(VbModule mod)
  {
    var ranges = new List<(int start, int end)>();

    foreach (var proc in mod.Procedures)
      if (proc.StartLine > 0 && proc.EndLine > proc.StartLine)
        ranges.Add((proc.StartLine, proc.EndLine));

    foreach (var prop in mod.Properties)
      if (prop.StartLine > 0 && prop.EndLine > prop.StartLine)
        ranges.Add((prop.StartLine, prop.EndLine));

    // Ordine decrescente: ultima procedura prima → gli indici restano validi
    return ranges.OrderByDescending(r => r.start).ToList();
  }

  // -------------------------
  // RIORDINO DIM IN UNA PROCEDURA
  // -------------------------

  /// <summary>
  /// Riordina le dichiarazioni Dim/Static all'interno di una singola procedura.
  /// Restituisce null se nessuna modifica è necessaria (idempotente).
  /// </summary>
  internal static List<string>? ReorderDimsInProcedure(List<string> procLines)
  {
    // Servono almeno firma + 1 riga body + End
    if (procLines.Count < 3) return null;

    // --- Fase 1: Fine della firma (gestione continuazione _) ---
    int sigEnd = 0;
    while (sigEnd < procLines.Count - 1 && Clean(procLines[sigEnd]).TrimEnd().EndsWith("_"))
      sigEnd++;

    int endLineIdx = procLines.Count - 1;
    int bodyFrom   = sigEnd + 1;
    int bodyTo     = endLineIdx - 1;

    if (bodyFrom > bodyTo) return null;

    // --- Fase 2: Attribute e commenti di intestazione ---
    int attributeEnd = bodyFrom;
    while (attributeEnd <= bodyTo)
    {
      var trimmed = Clean(procLines[attributeEnd]).TrimStart();
      if (trimmed.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
        attributeEnd++;
      else
        break;
    }

    int headerEnd = attributeEnd;
    while (headerEnd <= bodyTo)
    {
      var trimmed = Clean(procLines[headerEnd]).TrimStart();
      if (trimmed.StartsWith("'"))
        headerEnd++;
      else
        break;
    }

    // --- Fase 3: Raccolta dichiarazioni Const/Static/Dim ---
    var declEntries = new List<DeclarationEntry>();
    var isDeclLineIdx = new HashSet<int>();

    for (int i = headerEnd; i <= bodyTo; i++)
    {
      var trimmed = Clean(procLines[i]).TrimStart();
      bool isConst = ReConstLocal.IsMatch(trimmed);
      bool isDimOrStatic = ReDimOrStatic.IsMatch(trimmed);

      if (!isConst && !isDimOrStatic)
        continue;

      int start = i;
      var declLines = new List<string> { procLines[i] };

      while (i < bodyTo && Clean(procLines[i]).TrimEnd().EndsWith("_"))
      {
        i++;
        declLines.Add(procLines[i]);
      }

      DeclarationKind kind;
      if (isConst)
        kind = DeclarationKind.Const;
      else if (trimmed.StartsWith("Static", StringComparison.OrdinalIgnoreCase))
        kind = DeclarationKind.Static;
      else
        kind = DeclarationKind.Dim;

      declEntries.Add(new DeclarationEntry(start, i, kind, declLines));
      for (int j = start; j <= i; j++)
        isDeclLineIdx.Add(j);
    }

    if (declEntries.Count == 0) return null;

    // --- Fase 5: Costruzione risultato ---
    var result = new List<string>();

    // a) Firma
    for (int i = 0; i <= sigEnd; i++)
      result.Add(procLines[i]);

    // b) Attribute e commento di intestazione (escludendo eventuali dichiarazioni mescolate)
    for (int i = bodyFrom; i < headerEnd; i++)
      if (!isDeclLineIdx.Contains(i))
        result.Add(procLines[i]);

    // c) Costanti locali
    foreach (var entry in declEntries.Where(d => d.Kind == DeclarationKind.Const))
      AddIndentedLines(result, entry.Lines);

    // d) Variabili Static
    foreach (var entry in declEntries.Where(d => d.Kind == DeclarationKind.Static))
      AddIndentedLines(result, entry.Lines);

    // e) Variabili Dim
    foreach (var entry in declEntries.Where(d => d.Kind == DeclarationKind.Dim))
      AddIndentedLines(result, entry.Lines);

    // f) Righe rimanenti del corpo (escludendo dichiarazioni già spostate)
    for (int i = headerEnd; i <= bodyTo; i++)
      if (!isDeclLineIdx.Contains(i))
        result.Add(procLines[i]);

    // g) End Sub / End Function / End Property
    result.Add(procLines[endLineIdx]);

    // --- Fase 6: Idempotenza ---
    if (result.Count == procLines.Count)
    {
      bool same = true;
      for (int i = 0; i < result.Count; i++)
      {
        if (!string.Equals(result[i], procLines[i], StringComparison.Ordinal))
        {
          same = false;
          break;
        }
      }
      if (same) return null;
    }

    return result;
  }

  private static List<string>? ApplySpacingRules(List<string> lines, VbModule mod)
  {
    if (lines.Count == 0) return null;

    int startIdx = 0;
    if (mod.IsForm)
    {
      for (int i = 0; i < lines.Count; i++)
      {
        var trimmed = Clean(lines[i]).TrimStart();
        if (trimmed.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Option ", StringComparison.OrdinalIgnoreCase))
        {
          startIdx = i;
          break;
        }
      }
    }

    var result = new List<string>();
    for (int i = 0; i < startIdx; i++)
      result.Add(lines[i]);

    bool insideEnum = false;
    bool insideSelect = false;
    bool selectHasCase = false;
    bool inProcedure = false;
    bool inProcHeaderGroup = false;
    HeaderGroupKind? currentGroup = null;
    string currentPropertyName = null;
    bool lastWasLabel = false;

    for (int i = startIdx; i < lines.Count; i++)
    {
      var line = Clean(lines[i]);
      var trimmed = line.TrimStart();
      var normalized = string.IsNullOrWhiteSpace(line) ? string.Empty : line;

      if (string.IsNullOrWhiteSpace(normalized))
      {
        if (insideEnum || lastWasLabel)
          continue;

        if (result.Count > 0 && string.IsNullOrWhiteSpace(result.Last()))
          continue;

        var prevNonBlank = LastNonBlank(result);
        var nextNonBlank = NextNonBlank(lines, i + 1);

        if (!string.IsNullOrEmpty(prevNonBlank) && IsBlockStart(prevNonBlank))
          continue;

        if (!string.IsNullOrEmpty(prevNonBlank) && IsBlockStart(prevNonBlank) &&
            !string.IsNullOrEmpty(nextNonBlank) && IsCommentLine(nextNonBlank))
          continue;

        if (!string.IsNullOrEmpty(nextNonBlank) && IsBlockEnd(nextNonBlank))
          continue;

        if (!string.IsNullOrEmpty(prevNonBlank) &&
            (IsDimOrStatic(prevNonBlank) || IsConst(prevNonBlank)) &&
            !string.IsNullOrEmpty(nextNonBlank) &&
            (IsDimOrStatic(nextNonBlank) || IsConst(nextNonBlank)))
          continue;

        if (!string.IsNullOrEmpty(prevNonBlank) && prevNonBlank.TrimEnd().EndsWith("Then", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrEmpty(nextNonBlank) && nextNonBlank.TrimStart().StartsWith("If ", StringComparison.OrdinalIgnoreCase))
          continue;

        if (insideSelect && !selectHasCase && !string.IsNullOrEmpty(nextNonBlank) && IsCaseLine(nextNonBlank))
          continue;

        result.Add(string.Empty);
        lastWasLabel = false;
        continue;
      }

      lastWasLabel = false;

      if (IsOptionLine(trimmed) && result.Count > 0 && !string.IsNullOrWhiteSpace(result.Last()) &&
          !(mod.IsForm && IsAttributeLine(result.LastOrDefault() ?? string.Empty)))
        result.Add(string.Empty);

      if (IsProcedureStart(trimmed, out var procPropertyName))
      {
        inProcedure = true;
        inProcHeaderGroup = true;
        currentGroup = null;
        currentPropertyName = procPropertyName;
      }

      if (insideEnum && IsEnumEnd(trimmed))
        insideEnum = false;

      if (IsEnumStart(trimmed))
        insideEnum = true;

      if (IsSelectCaseStart(trimmed))
      {
        insideSelect = true;
        selectHasCase = false;
      }

      if (insideSelect && IsSelectCaseEnd(trimmed))
      {
        insideSelect = false;
        selectHasCase = false;
      }

      if (insideSelect && IsCaseLine(trimmed))
      {
        if (selectHasCase && result.Count > 0 && !string.IsNullOrWhiteSpace(result.Last()))
          result.Add(string.Empty);
        selectHasCase = true;
      }

      if (IsLabelLine(trimmed))
      {
        if (result.Count > 0 && !string.IsNullOrWhiteSpace(result.Last()))
          result.Add(string.Empty);
        lastWasLabel = true;
      }

      if (IsCommentLine(trimmed) &&
          !string.IsNullOrEmpty(NextNonBlank(lines, i + 1)) &&
          IsBlockStart(NextNonBlank(lines, i + 1)) &&
          (string.IsNullOrWhiteSpace(PrevNonBlank(result)) || !IsCommentLine(PrevNonBlank(result))))
      {
        if (result.Count > 0 && !string.IsNullOrWhiteSpace(result.Last()))
          result.Add(string.Empty);
      }

      if (inProcedure && inProcHeaderGroup && !IsAttributeLine(trimmed))
      {
        var group = GetHeaderGroupKind(trimmed);
        if (group == null)
        {
          inProcHeaderGroup = false;
          currentGroup = null;
        }
        else
        {
          if (currentGroup != null && currentGroup != group && result.Count > 0 &&
              !string.IsNullOrWhiteSpace(result.Last()) && !IsProcedureStart(PrevNonBlank(result), out _))
            result.Add(string.Empty);
          currentGroup = group;
        }
      }

      if (IsSingleLineIf(trimmed))
      {
        var prevNonBlank = LastNonBlank(result);
        var nextNonBlank = NextNonBlank(lines, i + 1);
        if (!string.IsNullOrEmpty(prevNonBlank) &&
            !IsIfBoundary(prevNonBlank) &&
            !string.IsNullOrWhiteSpace(result.LastOrDefault()) &&
            !IsCommentLine(prevNonBlank))
        {
          result.Add(string.Empty);
        }

        result.Add(normalized);

        if (!string.IsNullOrEmpty(nextNonBlank))
          result.Add(string.Empty);
        continue;
      }

      result.Add(normalized);

      if (IsCommentLine(trimmed) && result.Count > 0)
      {
        var nextNonBlank = NextNonBlank(lines, i + 1);
        var prevNonBlank = PrevNonBlank(result);
        if (!string.IsNullOrEmpty(prevNonBlank) && IsBlockStart(prevNonBlank) &&
            !string.IsNullOrEmpty(nextNonBlank) && !IsCommentLine(nextNonBlank))
        {
          if (result.Count > 0 && !string.IsNullOrWhiteSpace(result.Last()))
            result.Add(string.Empty);
        }
      }

      if (IsBlockEnd(trimmed))
      {
        if (IsEndProperty(trimmed, currentPropertyName, lines, i + 1))
        {
          currentPropertyName = null;
        }
        else if (RequiresBlankAfterEnd(trimmed))
        {
          var nextNonBlank = NextNonBlank(lines, i + 1);
          if (!string.IsNullOrEmpty(nextNonBlank))
            result.Add(string.Empty);
        }
      }
    }

    if (result.Count == lines.Count)
    {
      bool same = true;
      for (int i = 0; i < result.Count; i++)
      {
        if (!string.Equals(result[i], Clean(lines[i]), StringComparison.Ordinal))
        {
          same = false;
          break;
        }
      }
      if (same) return null;
    }

    return result;
  }


  // -------------------------
  // HELPER
  // -------------------------

  /// <summary>Rimuove \r finale (per gestione \r\n dopo split su \n).</summary>
  private static string Clean(string s) => s.EndsWith('\r') ? s[..^1] : s;

  private static void AddIndentedLines(List<string> target, List<string> lines)
  {
    const string bodyIndent = "  ";
    foreach (var line in lines)
    {
      var cr = line.EndsWith('\r') ? "\r" : "";
      var trimmed = Clean(line).TrimStart();
      target.Add(bodyIndent + trimmed + cr);
    }
  }

  private static string NextNonBlank(List<string> lines, int startIdx)
  {
    for (int i = startIdx; i < lines.Count; i++)
    {
      var line = Clean(lines[i]);
      if (!string.IsNullOrWhiteSpace(line))
        return line;
    }
    return string.Empty;
  }

  private static string LastNonBlank(List<string> lines)
  {
    for (int i = lines.Count - 1; i >= 0; i--)
    {
      var line = Clean(lines[i]);
      if (!string.IsNullOrWhiteSpace(line))
        return line;
    }
    return string.Empty;
  }

  private static string PrevNonBlank(List<string> lines)
  {
    return LastNonBlank(lines);
  }

  private static bool IsCommentLine(string line) => line.TrimStart().StartsWith("'");

  private static bool IsAttributeLine(string line) => line.TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

  private static bool IsOptionLine(string line) => line.TrimStart().StartsWith("Option ", StringComparison.OrdinalIgnoreCase);

  private static bool IsDimOrStatic(string line) => ReDimOrStatic.IsMatch(line.TrimStart());

  private static bool IsConst(string line) => ReConstLocal.IsMatch(line.TrimStart());

  private static bool IsLabelLine(string line) => Regex.IsMatch(line.TrimStart(), @"^[A-Za-z_]\w*:\s*$");

  private static bool IsProcedureStart(string line, out string propertyName)
  {
    propertyName = null;
    var match = Regex.Match(line, @"^(Public|Private|Friend)?\s*(Static\s+)?(Sub|Function|Property)\s+(\w+)", RegexOptions.IgnoreCase);
    if (!match.Success)
      return false;

    if (match.Groups[3].Value.Equals("Property", StringComparison.OrdinalIgnoreCase))
      propertyName = match.Groups[4].Value;

    return true;
  }

  private static bool IsBlockStart(string line)
  {
    var trimmed = StripInlineComment(line).TrimStart();
    return Regex.IsMatch(trimmed, @"^(Public|Private|Friend)?\s*(Sub|Function|Property|Type|Enum)\b", RegexOptions.IgnoreCase) ||
           trimmed.StartsWith("If ", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("For ", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("Do ", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("Select Case", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("With ", StringComparison.OrdinalIgnoreCase);
  }

  private static bool IsBlockEnd(string line)
  {
    var trimmed = StripInlineComment(line).TrimStart();
    return trimmed.StartsWith("End Sub", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Function", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Property", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Type", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Enum", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End If", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("Next", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("Loop", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Select", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End With", StringComparison.OrdinalIgnoreCase);
  }

  private static bool RequiresBlankAfterEnd(string line)
  {
    var trimmed = line.TrimStart();
    return trimmed.StartsWith("End Sub", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Function", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Property", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Type", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End Enum", StringComparison.OrdinalIgnoreCase);
  }

  private static bool IsEnumStart(string line) => Regex.IsMatch(line.TrimStart(), @"^(Public|Private|Friend)?\s*Enum\b", RegexOptions.IgnoreCase);

  private static bool IsEnumEnd(string line) => line.TrimStart().StartsWith("End Enum", StringComparison.OrdinalIgnoreCase);

  private static bool IsSelectCaseStart(string line) => line.TrimStart().StartsWith("Select Case", StringComparison.OrdinalIgnoreCase);

  private static bool IsSelectCaseEnd(string line) => line.TrimStart().StartsWith("End Select", StringComparison.OrdinalIgnoreCase);

  private static bool IsCaseLine(string line) => line.TrimStart().StartsWith("Case ", StringComparison.OrdinalIgnoreCase);

  private static bool IsSingleLineIf(string line)
  {
    var trimmed = StripInlineComment(line).TrimStart();
    if (!trimmed.StartsWith("If ", StringComparison.OrdinalIgnoreCase))
      return false;

    if (!Regex.IsMatch(trimmed, @"\bThen\b", RegexOptions.IgnoreCase))
      return false;

    return !trimmed.TrimEnd().EndsWith("Then", StringComparison.OrdinalIgnoreCase);
  }

  private static bool IsIfBoundary(string line)
  {
    var trimmed = StripInlineComment(line).TrimStart();
    return trimmed.StartsWith("If ", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("Else", StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith("End If", StringComparison.OrdinalIgnoreCase);
  }

  private static string StripInlineComment(string line)
  {
    if (string.IsNullOrEmpty(line))
      return line;

    bool inString = false;
    for (int i = 0; i < line.Length; i++)
    {
      if (line[i] == '"')
      {
        if (!inString)
          inString = true;
        else if (i + 1 < line.Length && line[i + 1] == '"')
          i++;
        else
          inString = false;
      }
      else if (!inString && line[i] == '\'')
      {
        return line.Substring(0, i);
      }
    }

    return line;
  }

  private static HeaderGroupKind? GetHeaderGroupKind(string line)
  {
    var trimmed = line.TrimStart();
    if (IsCommentLine(trimmed))
      return HeaderGroupKind.Comments;

    if (IsConst(trimmed))
      return HeaderGroupKind.Const;
    if (trimmed.StartsWith("Static", StringComparison.OrdinalIgnoreCase))
      return HeaderGroupKind.Static;
    if (trimmed.StartsWith("Dim", StringComparison.OrdinalIgnoreCase))
      return HeaderGroupKind.Dim;

    return null;
  }

  private static bool IsEndProperty(string line, string currentPropertyName, List<string> lines, int nextIndex)
  {
    if (!line.TrimStart().StartsWith("End Property", StringComparison.OrdinalIgnoreCase))
      return false;

    if (string.IsNullOrEmpty(currentPropertyName))
      return false;

    var nextNonBlank = NextNonBlank(lines, nextIndex);
    if (TryGetPropertyName(nextNonBlank, out var nextPropertyName) &&
        string.Equals(nextPropertyName, currentPropertyName, StringComparison.OrdinalIgnoreCase))
      return true;

    return false;
  }

  private static bool TryGetPropertyName(string line, out string propertyName)
  {
    propertyName = null;
    var match = Regex.Match(line, @"^(Public|Private|Friend)?\s*(Static\s+)?Property\s+(Get|Let|Set)\s+(\w+)", RegexOptions.IgnoreCase);
    if (!match.Success)
      return false;

    propertyName = match.Groups[4].Value;
    return true;
  }

  private enum DeclarationKind
  {
    Const,
    Static,
    Dim
  }

  private enum HeaderGroupKind
  {
    Comments,
    Const,
    Static,
    Dim
  }

  private sealed record DeclarationEntry(
    int StartIdx, int EndIdx, DeclarationKind Kind, List<string> Lines);
}
