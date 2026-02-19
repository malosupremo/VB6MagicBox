using System;
using System.IO;
using VB6MagicBox.Models;
using VB6MagicBox.Parsing;

namespace VB6MagicBox;

public class Program
{
  public static void Main(string[] args)
  {
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("===========================================");
    Console.WriteLine("              VB6 Magic Box ");
    Console.WriteLine("===========================================");
    Console.WriteLine();
    Console.ForegroundColor= ConsoleColor.Gray;

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
      Console.WriteLine("2. Aggiunta tipi mancanti");
      Console.WriteLine("3. Applica refactoring automatico");
      Console.WriteLine("4. Armonizza le spaziature");
      Console.WriteLine("5. Riordina le variabili di procedura");
      Console.WriteLine("6. BACCHETTA MAGICA: tutto insieme!");
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
          RunTypeAnnotatorInteractive();
          break;

        case "3":
          RunRefactoringInteractive();
          break;

        case "4":
          Console.WriteLine();
          Console.WriteLine("[!] Armonizzazione spaziature non in armonia.");
          Console.WriteLine("    Coming soon!");
          break;

        case "5":
          Console.WriteLine();
          Console.WriteLine("[!] Ordinamento variabili di procedura in disordine.");
          Console.WriteLine("    Coming soon!");
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
          Console.WriteLine("[X] Opzione non valida. Riprova.");
          break;
      }
    }
  }

  private static void RunAnalysisInteractive()
  {
    Console.WriteLine();
    Console.Write("Percorso del file .vbp: ");
    var vbpPath = Console.ReadLine()?.Trim().Trim('"');

    if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
    {
      Console.WriteLine("[X] File non trovato!");
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
      Console.WriteLine("[OK] Analisi completata.");
    }
    catch (Exception ex)
    {
      Console.WriteLine();
      Console.WriteLine("[X] Errore durante l'analisi:");
      Console.WriteLine(ex.ToString());
    }
  }

  private static void RunTypeAnnotatorInteractive()
  {
    Console.WriteLine();
    Console.Write("Percorso del file .vbp: ");
    var vbpPath = Console.ReadLine()?.Trim().Trim('"');

    if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
    {
      Console.WriteLine("[X] File .vbp non trovato!");
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
      Console.WriteLine("[X] Errore durante l'aggiunta dei tipi:");
      Console.WriteLine(ex.ToString());
    }
  }

  private static void RunRefactoringInteractive()
  {
    Console.WriteLine();
    Console.Write("Percorso del file .vbp: ");
    var vbpPath = Console.ReadLine()?.Trim().Trim('"');

    if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
    {
      Console.WriteLine("[X] File .vbp non trovato!");
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
      Console.WriteLine("[X] Errore durante il refactoring:");
      Console.WriteLine(ex.ToString());
    }
  }

  private static void RunMagicWandInteractive()
  {
    Console.WriteLine();
    Console.Write("Percorso del file .vbp: ");
    var vbpPath = Console.ReadLine()?.Trim().Trim('"');

    if (string.IsNullOrEmpty(vbpPath) || !File.Exists(vbpPath))
    {
      Console.WriteLine("[X] File .vbp non trovato!");
      return;
    }

    try
    {
      // 1) Analisi completa — una sola esecuzione del parser per tutte le fasi
      var project = VbParser.ParseAndResolve(vbpPath);

      // Export file di analisi (symbols.json, rename.json, rename.csv, dependencies.md)
      ExportProjectFiles(project, vbpPath);

      // 2) Aggiunta tipi mancanti: usa i nomi originali dal modello.
      //    Deve precedere il rename: dopo il rename i nomi nel sorgente non
      //    corrisponderebbero più a quelli del modello (es. parametri rinominati).
      TypeAnnotator.AddMissingTypes(project);

      // 3) Refactoring: rinomina i simboli secondo le convenzioni
      Refactoring.ApplyRenames(project);

      Console.WriteLine();
      Console.WriteLine("[OK] Bacchetta magica applicata!");
    }
    catch (Exception ex)
    {
      Console.WriteLine();
      Console.WriteLine("[X] Errore durante la bacchetta magica:");
      Console.WriteLine(ex.ToString());
    }
  }

  /// <summary>
  /// Scrive tutti i file di output del progetto analizzato:
  /// symbols.json, rename.json, rename.csv, dependencies.md (Mermaid).
  /// Presuppone che ParseAndResolve (che include SortProject) sia già stato chiamato.
  /// </summary>
  private static void ExportProjectFiles(VbProject project, string vbpPath)
  {
    var vbpDir  = Path.GetDirectoryName(vbpPath)!;
    var vbpName = Path.GetFileNameWithoutExtension(vbpPath);

    var jsonOut     = Path.Combine(vbpDir, $"{vbpName}.symbols.json");
    var renameJson  = Path.Combine(vbpDir, $"{vbpName}.rename.json");
    var renameCsv   = Path.Combine(vbpDir, $"{vbpName}.rename.csv");
    var mermaidOut  = Path.Combine(vbpDir, $"{vbpName}.dependencies.md");

    Console.WriteLine();
    Console.WriteLine(">> Esportazione file di output...");

    VbParser.ExportJson(project, jsonOut);
    VbParser.ExportRenameJson(project, renameJson);
    VbParser.ExportRenameCsv(project, renameCsv);
    VbParser.ExportMermaid(project, mermaidOut);

    Console.WriteLine($"   JSON completo: {jsonOut}");
    Console.WriteLine($"   JSON rename:   {renameJson}");
    Console.WriteLine($"   CSV rename:    {renameCsv}");
    Console.WriteLine($"   Mermaid:       {mermaidOut}");
  }
}
