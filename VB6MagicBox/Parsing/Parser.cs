using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;


/// <summary>
/// File collante che espone l'API pubblica del parser VB6.
/// Unisce parsing, risoluzione, ordinamento ed export.
/// </summary>
public static partial class VbParser
{
  /// <summary>
  /// Esegue l’intera pipeline:
  /// 1) Parsing del progetto
  /// 2) Risoluzione chiamate, tipi, campi
  /// 3) Costruzione dipendenze + marcatura Used
  /// 4) Ordinamento alfabetico completo
  /// </summary>
  public static VbProject ParseAndResolve(string vbpPath)
  {
    if (string.IsNullOrWhiteSpace(vbpPath))
      throw new ArgumentException("Percorso VBP non valido.", nameof(vbpPath));

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine();
    Console.WriteLine("===========================================");
    Console.WriteLine("  1: Analisi progetto .vbp");
    Console.WriteLine("===========================================");
    Console.WriteLine();
    Console.ForegroundColor= ConsoleColor.Gray;

    // 1) Parsing
    Console.WriteLine("Step 1/4: Parsing del progetto VB6...");
    var project = ParseProjectFromVbp(vbpPath);
    Console.WriteLine($"  -> {project.Modules.Count} moduli trovati");

    // 2) Risoluzione semantica
    Console.WriteLine("Step 2/4: Risoluzione di tipi e chiamate...");
    ResolveTypesAndCalls(project);

    // 3) Dipendenze + marcatura Used
    Console.WriteLine("Step 3/4: Costruzione dipendenze e marcatura simboli utilizzati...");
    BuildDependenciesAndUsage(project);

    // 4) Ordinamento
    Console.WriteLine("Step 4/4: Applicazione convenzioni di naming e ordinamento...");
    SortProject(project);

    return project;
  }

  /// <summary>
  /// Esegue l’intera pipeline e salva JSON + Mermaid.
  /// </summary>
  public static void ParseResolveAndExport(
      string vbpPath,
      string jsonOutputPath,
      string mermaidOutputPath)
  {
    var project = ParseAndResolve(vbpPath);

    Console.WriteLine("Step 5/5: Esportazione file di output...");
    ExportJson(project, jsonOutputPath);
    
    // Genera anche il file .rename.json
    var renameOutputPath = jsonOutputPath.Replace(".json", ".rename.json");
    ExportRenameJson(project, renameOutputPath);
    
    // Genera anche il file .rename.csv
    var renameCsvPath = jsonOutputPath.Replace(".json", ".rename.csv");
    ExportRenameCsv(project, renameCsvPath);
    
    ExportMermaid(project, mermaidOutputPath);
    Console.WriteLine("  -> Completato!");
  }
}
