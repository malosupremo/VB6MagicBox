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
      var lines = originalContent.Split('\n').ToList();
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

      File.WriteAllText(filePath, string.Join("\n", lines), enc);
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

    // --- Fase 2: Blocco commento di intestazione ---
    // Righe consecutive di commento all'inizio del corpo, attaccate alla firma.
    int headerEnd = bodyFrom;
    while (headerEnd <= bodyTo)
    {
      var trimmed = Clean(procLines[headerEnd]).TrimStart();
      if (trimmed.StartsWith("'"))
        headerEnd++;
      else
        break;
    }

    // --- Fase 3: Raccolta dichiarazioni Dim/Static ---
    var dimEntries   = new List<DimEntry>();
    var isDimLineIdx = new HashSet<int>();

    for (int i = bodyFrom; i <= bodyTo; i++)
    {
      var trimmed = Clean(procLines[i]).TrimStart();
      if (!ReDimOrStatic.IsMatch(trimmed)) continue;

      int start    = i;
      var dimLines = new List<string> { procLines[i] };

      // Gestione continuazione _
      while (i < bodyTo && Clean(procLines[i]).TrimEnd().EndsWith("_"))
      {
        i++;
        dimLines.Add(procLines[i]);
      }

      bool isStatic = trimmed.StartsWith("Static", StringComparison.OrdinalIgnoreCase);
      var nameMatch = ReVarName.Match(trimmed);
      var varName   = nameMatch.Success ? nameMatch.Groups[1].Value : "";

      dimEntries.Add(new DimEntry(start, i, isStatic, varName, dimLines));
      for (int j = start; j <= i; j++)
        isDimLineIdx.Add(j);
    }

    if (dimEntries.Count == 0) return null;

    // --- Fase 4: Indentazione corpo (2 spazi fissi) ---
    const string bodyIndent = "  ";

    // --- Fase 5: Costruzione risultato ---
    var result = new List<string>();

    // a) Firma
    for (int i = 0; i <= sigEnd; i++)
      result.Add(procLines[i]);

    // b) Commento di intestazione (escludendo eventuali Dim mescolati)
    for (int i = bodyFrom; i < headerEnd; i++)
      if (!isDimLineIdx.Contains(i))
        result.Add(procLines[i]);

    // c) Dim/Static ordinati: prima Static (alfa), poi Dim (alfa)
    var sorted = dimEntries
      .Where(d => d.IsStatic)
      .OrderBy(d => d.VarName, StringComparer.OrdinalIgnoreCase)
      .Concat(dimEntries
        .Where(d => !d.IsStatic)
        .OrderBy(d => d.VarName, StringComparer.OrdinalIgnoreCase));

    foreach (var dim in sorted)
    {
      foreach (var line in dim.Lines)
      {
        var cr      = line.EndsWith('\r') ? "\r" : "";
        var trimmed = Clean(line).TrimStart();
        result.Add(bodyIndent + trimmed + cr);
      }
    }

    // d) Righe rimanenti del corpo (escludendo Dim/Static già spostati)
    for (int i = headerEnd; i <= bodyTo; i++)
      if (!isDimLineIdx.Contains(i))
        result.Add(procLines[i]);

    // e) End Sub / End Function / End Property
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


  // -------------------------
  // HELPER
  // -------------------------

  /// <summary>Rimuove \r finale (per gestione \r\n dopo split su \n).</summary>
  private static string Clean(string s) => s.EndsWith('\r') ? s[..^1] : s;

  private sealed record DimEntry(
    int StartIdx, int EndIdx, bool IsStatic, string VarName, List<string> Lines);
}
