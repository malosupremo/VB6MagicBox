using System.Diagnostics;
using System.IO;
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

        var stopwatch = Stopwatch.StartNew();

        Console.WriteLine();
        ConsoleX.WriteLineColor("===========================================", ConsoleColor.Yellow);
        ConsoleX.WriteLineColor("  1: Analisi progetto .vbp", ConsoleColor.Yellow);
        ConsoleX.WriteLineColor("===========================================", ConsoleColor.Yellow);
        Console.WriteLine();

        // 1) Parsing
        Console.WriteLine("Step 1/5: Parsing del progetto VB6...");
        var project = ParseProjectFromVbp(vbpPath);
        Console.WriteLine($"  -> {project.Modules.Count} moduli trovati");

        var fileCache = BuildFileCache(project);

        // 2) Risoluzione semantica
        Console.WriteLine("Step 2/5: Risoluzione di tipi e chiamate...");
        ResolveTypesAndCalls(project, fileCache);

        // 3) Dipendenze + marcatura Used
        Console.WriteLine("Step 3/5: Costruzione dipendenze e marcatura simboli utilizzati...");
        BuildDependenciesAndUsage(project, fileCache);

        // 4) Ordinamento e naming
        Console.WriteLine("Step 4/5: Applicazione convenzioni di naming e ordinamento...");
        SortProject(project);

        // 5) Costruzione sostituzioni
        Console.WriteLine("Step 5/5: Costruzione sostituzioni precise (Replaces)...");
        BuildReplaces(project, fileCache);

        stopwatch.Stop();
        Console.WriteLine($"Tempo totale: {stopwatch.Elapsed.TotalMilliseconds:0.000} ms");

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

        ExportMermaid(project, mermaidOutputPath);
        var enumPrefixTodoPath = Path.Combine(Path.GetDirectoryName(jsonOutputPath) ?? string.Empty, "_TODO_enumprefix.csv");
        ExportEnumPrefixTodoCsv(project, enumPrefixTodoPath);
        Console.WriteLine("  -> Completato!");
    }
}
