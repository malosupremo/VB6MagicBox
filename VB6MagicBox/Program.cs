using System.Diagnostics;
using VB6MagicBox.Models;
using VB6MagicBox.Parsing;

namespace VB6MagicBox;

public class Program
{
    private const string LastVbpFileName = "last.vbp.path";

    public static void Main(string[] args)
    {
        ConsoleX.WriteLineColor("===========================================", ConsoleColor.Yellow);
        ConsoleX.WriteLineColor("              VB6 Magic Box ", ConsoleColor.Yellow);
        ConsoleX.WriteLineColor("===========================================", ConsoleColor.Yellow);
        Console.ForegroundColor = ConsoleColor.Gray;

        // Se ci sono argomenti da riga di comando, usa la modalità legacy (analisi diretta)
        if (args.Length > 0)
        {
            RunAnalysis(args[0]);
            return;
        }

        // Altrimenti mostra il menu interattivo
        ShowMenu();
    }

    private static void ShowMenu()
    {
        while (true)
        {
            Console.WriteLine();
            Console.WriteLine("Opzioni:");
            Console.WriteLine("1. Analizza progetto VB6");
            Console.WriteLine("2. Applica refactoring automatico");
            Console.WriteLine("3. Aggiunta tipi mancanti e scope e rimozione Call");
            Console.WriteLine("4. Riordina le variabili di procedura");
            Console.WriteLine("5. Armonizza le spaziature");
            Console.Write("6. ");
            ConsoleX.WriteColor("BACCHETTA MAGICA", ConsoleColor.Yellow);
            Console.WriteLine(": tutto insieme!");
            Console.WriteLine("0. Esci");
            Console.WriteLine();
            Console.Write("Seleziona opzione: ");

            var choice = Console.ReadLine()?.Trim();

            switch (choice)
            {
                case "1":
                    RunAnalysisInteractive();
                    break;

                case "2":
                    RunRefactoringInteractive();
                    break;

                case "3":
                    RunTypeAnnotatorInteractive();
                    break;

                case "4":
                    RunVariableReorderInteractive();
                    break;

                case "5":
                    RunSpacingInteractive();
                    break;

                case "6":
                    RunMagicWandInteractive();
                    break;

                case "0":
                    Console.WriteLine();
                    Console.WriteLine("Arrivederci!");
                    return;

                default:
                    Console.WriteLine();
                    ConsoleX.WriteLineColor("[X] Opzione non valida. Riprova.", ConsoleColor.Red);
                    break;
            }
        }
    }

    private static void RunSpacingInteractive()
    {
        Console.WriteLine();
        var vbpPath = ReadVbpPath("Percorso del file .vbp");

        if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
        {
            ConsoleX.WriteLineColor("[X] File .vbp non trovato!", ConsoleColor.Red);
            return;
        }

        try
        {
            var project = VbParser.ParseAndResolve(vbpPath);
            CodeFormatter.HarmonizeSpacing(project);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            ConsoleX.WriteLineColor("[X] Errore durante l'armonizzazione spaziature:", ConsoleColor.Red);
            ConsoleX.WriteLineColor(ex.ToString(), ConsoleColor.Red);
        }
    }

    private static void RunAnalysisInteractive()
    {
        Console.WriteLine();
        var vbpPath = ReadVbpPath("Percorso del file .vbp");

        if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
        {
            ConsoleX.WriteLineColor("[X] File non trovato!", ConsoleColor.Red);
            return;
        }

        RunAnalysis(vbpPath);
    }

    private static void RunAnalysis(string vbpPath)
    {
        try
        {
            var project = VbParser.ParseAndResolve(vbpPath);
            ExportProjectFiles(project, vbpPath);
            Console.WriteLine();
            ConsoleX.WriteLineColor("[OK] Analisi completata.", ConsoleColor.Green);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            ConsoleX.WriteLineColor("[X] Errore durante l'analisi:", ConsoleColor.Red);
            ConsoleX.WriteLineColor(ex.ToString(), ConsoleColor.Red);
        }
    }

    private static void RunTypeAnnotatorInteractive()
    {
        Console.WriteLine();
        var vbpPath = ReadVbpPath("Percorso del file .vbp");

        if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
        {
            ConsoleX.WriteLineColor("[X] File .vbp non trovato!", ConsoleColor.Red);
            return;
        }

        try
        {
            // Fase 1: parsing + risoluzione (stessa pipeline dell'opzione 1)
            var project = VbParser.ParseAndResolve(vbpPath);

            // Fase 3: aggiunta tipi mancanti usando il modello analizzato
            TypeAnnotator.AddMissingTypes(project);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            ConsoleX.WriteLineColor("[X] Errore durante l'aggiunta dei tipi:", ConsoleColor.Red);
            ConsoleX.WriteLineColor(ex.ToString(), ConsoleColor.Red);
        }
    }

    private static void RunRefactoringInteractive()
    {
        Console.WriteLine();
        var vbpPath = ReadVbpPath("Percorso del file .vbp");

        if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
        {
            ConsoleX.WriteLineColor("[X] File .vbp non trovato!", ConsoleColor.Red);
            return;
        }

        try
        {
            // 1) Analisi completa (parsing + risoluzione + naming + ordinamento)
            var project = VbParser.ParseAndResolve(vbpPath);

            // 2) Scrittura dei file di output prima del refactoring
            //    (l'analisi riflette i nomi originali VB6; il refactoring agisce solo sul disco)
            ExportProjectFiles(project, vbpPath);

            // 3) Refactoring: rinomina i simboli nei file sorgente
            Refactoring.ApplyRenames(project);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            ConsoleX.WriteLineColor("[X] Errore durante il refactoring:", ConsoleColor.Red);
            ConsoleX.WriteLineColor(ex.ToString(), ConsoleColor.Red);
        }
    }

    private static void RunVariableReorderInteractive()
    {
        Console.WriteLine();
        var vbpPath = ReadVbpPath("Percorso del file .vbp");

        if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
        {
            ConsoleX.WriteLineColor("[X] File .vbp non trovato!", ConsoleColor.Red);
            return;
        }

        try
        {
            var project = VbParser.ParseAndResolve(vbpPath);
            CodeFormatter.ReorderLocalVariables(project);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            ConsoleX.WriteLineColor("[X] Errore durante il riordino:", ConsoleColor.Red);
            ConsoleX.WriteLineColor(ex.ToString(), ConsoleColor.Red);
        }
    }

    private static string ReadVbpPath(string label)
    {
        var lastPath = ReadLastVbpPath();
        if (!string.IsNullOrWhiteSpace(lastPath))
            Console.Write($"{label} ({lastPath}): ");
        else
            Console.Write($"{label}: ");

        var input = Console.ReadLine()?.Trim().Trim('"');
        var vbpPath = string.IsNullOrWhiteSpace(input) ? lastPath : input;

        if (!string.IsNullOrWhiteSpace(vbpPath))
            SaveLastVbpPath(vbpPath);

        return vbpPath;
    }

    private static string ReadLastVbpPath()
    {
        var lastPathFile = GetLastVbpPathFile();
        if (!File.Exists(lastPathFile))
            return string.Empty;

        return File.ReadAllText(lastPathFile).Trim();
    }

    private static void SaveLastVbpPath(string vbpPath)
    {
        var lastPathFile = GetLastVbpPathFile();
        File.WriteAllText(lastPathFile, vbpPath);
    }

    private static string GetLastVbpPathFile()
    {
        var basePath = AppContext.BaseDirectory;
        return Path.Combine(basePath, LastVbpFileName);
    }

    private static void RunMagicWandInteractive()
    {
        Console.WriteLine();
        var vbpPath = ReadVbpPath("Percorso del file .vbp");

        if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
        {
            ConsoleX.WriteLineColor("[X] File .vbp non trovato!", ConsoleColor.Red);
            return;
        }

        var stopwatch = Stopwatch.StartNew();

        try
        {
            // 1) Analisi completa — una sola esecuzione del parser per tutte le fasi
            var project = VbParser.ParseAndResolve(vbpPath);

            // 2) Export file di analisi (symbols.json, rename.json, rename.csv, linereplace.json, dependencies.md)
            ExportProjectFiles(project, vbpPath);

            // 3) Refactoring: rinomina i simboli secondo le convenzioni
            //    DEVE precedere TypeAnnotator perché dopo il rename i nomi nel sorgente
            //    corrispondono ai ConventionalName del modello
            Refactoring.ApplyRenames(project);

            // 4) Aggiunta tipi mancanti: usa i nomi convenzionali dopo il rename
            TypeAnnotator.AddMissingTypes(project);

            // 5) Riordino variabili locali: sposta Dim/Static in cima a ogni procedura
            //    Deve seguire tutto il resto perché opera sui file già rinominati e tipizzati
            CodeFormatter.ReorderLocalVariables(project);

            // 6) Armonizzazione spaziature
            CodeFormatter.HarmonizeSpacing(project);

            Console.WriteLine();
            ConsoleX.WriteLineColor("[OK] Bacchetta magica applicata!", ConsoleColor.Green);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            ConsoleX.WriteLineColor("[X] Errore durante la bacchetta magica:", ConsoleColor.Red);
            ConsoleX.WriteLineColor(ex.ToString(), ConsoleColor.Red);
        }
        finally
        {
            stopwatch.Stop();
            Console.WriteLine($"Tempo bacchetta magica: {stopwatch.Elapsed.TotalMilliseconds / 1000:0.000} s");
        }
    }

    /// <summary>
    /// Scrive tutti i file di output del progetto analizzato:
    /// symbols.json, rename.json, rename.csv, linereplace.json, dependencies.md (Mermaid).
    /// Presuppone che ParseAndResolve (che include SortProject e BuildReplaces) sia già stato chiamato.
    /// </summary>
    private static void ExportProjectFiles(VbProject project, string vbpPath)
    {
        var vbpDir = Path.GetDirectoryName(vbpPath)!;
        var vbpName = Path.GetFileNameWithoutExtension(vbpPath);

        var jsonOut = Path.Combine(vbpDir, $"{vbpName}.symbols.json");
        var renameJson = Path.Combine(vbpDir, $"{vbpName}.rename.json");
        var renameCsv = Path.Combine(vbpDir, $"{vbpName}.rename.csv");
        var shadowsCsv = Path.Combine(vbpDir, $"{vbpName}._TODO_shadows.csv");
        var lineReplaceJson = Path.Combine(vbpDir, $"{vbpName}.linereplace.json");
        var startCharCheckCsv = Path.Combine(vbpDir, $"{vbpName}._CHECK_startchars.csv");
        var mermaidOut = Path.Combine(vbpDir, $"{vbpName}.dependencies.md");

        Console.WriteLine();
        Console.WriteLine(">> Esportazione file di output...");

        VbParser.ExportJson(project, jsonOut);
        VbParser.ExportRenameJson(project, renameJson);
        VbParser.ExportRenameCsv(project, renameCsv);
        VbParser.ExportShadowsCsv(project, shadowsCsv);
        VbParser.ExportLineReplaceJson(project, lineReplaceJson);
        VbParser.ExportStartCharChecks(startCharCheckCsv);
        VbParser.ExportMermaid(project, mermaidOut);

        Console.WriteLine($"   JSON completo:     {jsonOut}");
        Console.WriteLine($"   JSON rename:       {renameJson}");
        Console.WriteLine($"   CSV rename:        {renameCsv}");
        Console.WriteLine($"   CSV shadows:       {shadowsCsv}");
        Console.WriteLine($"   JSON linereplace:  {lineReplaceJson}");
        Console.WriteLine($"   CSV startchars:    {startCharCheckCsv}");
        Console.WriteLine($"   Mermaid:           {mermaidOut}");
    }
}
