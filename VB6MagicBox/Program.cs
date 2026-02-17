using System;
using System.IO;
using VB6MagicBox.Parsing;

namespace VB6MagicBox;

public class Program
{
  public static void Main(string[] args)
  {
    Console.WriteLine("===========================================");
    Console.WriteLine("              VB6 Magic Box ");
    Console.WriteLine("===========================================");
    Console.WriteLine();

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
      Console.WriteLine("1. Analizza progetto VB6 (genera .json)");
      Console.WriteLine("2. Applica refactoring automatico");
      Console.WriteLine("3. Aggiunta tipi mancanti");
      Console.WriteLine("4. Armonizza le spaziature");
      Console.WriteLine("5. Riordina le variabili di procedura");
      Console.WriteLine("6. Genera codice C#");
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
          Console.WriteLine();
          Console.WriteLine("[!] Aggiunta tipi mancanti mancante.");
          Console.WriteLine("    Coming soon!");
          break;

        case "4":
          Console.WriteLine();
          Console.WriteLine("[!] Armonizzazione spaziature fuori armonia.");
          Console.WriteLine("    Coming soon!");
          break;

        case "5":
          Console.WriteLine();
          Console.WriteLine("[!] Ordinamento variabili di procedura in disordine.");
          Console.WriteLine("    Coming soon!");
          break;

        case "6":
          Console.WriteLine();
          Console.WriteLine("[!] Generazione C#... non esageriamo!");
          Console.WriteLine("    Coming soon?!");
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
      // Genera i nomi dei file di output basati sul .vbp
      var vbpDir = Path.GetDirectoryName(vbpPath)!;
      var vbpName = Path.GetFileNameWithoutExtension(vbpPath);
      
      var jsonOut = Path.Combine(vbpDir, $"{vbpName}.symbols.json");
      var mermaidOut = Path.Combine(vbpDir, $"{vbpName}.dependencies.mmd");

      VbParser.ParseResolveAndExport(vbpPath, jsonOut, mermaidOut);

      Console.WriteLine();
      Console.WriteLine("[OK] Analisi completata.");
      Console.WriteLine($"     JSON completo: {jsonOut}");
      Console.WriteLine($"     JSON rename: {Path.Combine(vbpDir, $"{vbpName}.rename.json")}");
      Console.WriteLine($"     CSV rename: {Path.Combine(vbpDir, $"{vbpName}.rename.csv")}");
      Console.WriteLine($"     Mermaid: {mermaidOut}");
    }
    catch (Exception ex)
    {
      Console.WriteLine();
      Console.WriteLine("[X] Errore durante l'analisi:");
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
      Console.WriteLine();
      Console.WriteLine(">> Avvio analisi completa del progetto...");
      Console.WriteLine("   (necessario per avere il contesto semantico completo)");
      Console.WriteLine();
      
      // Esegue l'analisi completa (parsing + risoluzione + dipendenze)
      // Questo ci dà accesso a References, Calls, Dependencies e tutto il contesto
      var project = VbParser.ParseAndResolve(vbpPath);
      
      Console.WriteLine();
      Console.WriteLine("[OK] Analisi completata, avvio refactoring...");
      
      // Applica i rename usando il progetto completo in memoria
      Refactoring.ApplyRenames(project);
      
      // Salva anche il file .rename.json come report di cosa è stato fatto
      var vbpDir = Path.GetDirectoryName(vbpPath)!;
      var vbpName = Path.GetFileNameWithoutExtension(vbpPath);
      var renameJsonPath = Path.Combine(vbpDir, $"{vbpName}.rename.json");
      var renameCsvPath = Path.Combine(vbpDir, $"{vbpName}.rename.csv");
      var jsonOut = Path.Combine(vbpDir, $"{vbpName}.symbols.json");
      
      Console.WriteLine();
      Console.WriteLine(">> Generazione file di output...");
      
      // Esporta il JSON aggiornato con il contesto completo
      VbParser.SortProject(project);
      VbParser.ExportJson(project, jsonOut);
      Console.WriteLine($"   JSON aggiornato: {jsonOut}");
      
      // Esporta il report dei rename
      VbParser.ExportRenameJson(project, renameJsonPath);
      Console.WriteLine($"   Report rename: {renameJsonPath}");
      
      // Esporta il report CSV dei rename
      VbParser.ExportRenameCsv(project, renameCsvPath);
      Console.WriteLine($"   Report CSV rename: {renameCsvPath}");
      
      Console.WriteLine();
      Console.WriteLine("[OK] Refactoring completato con successo!");
    }
    catch (Exception ex)
    {
      Console.WriteLine();
      Console.WriteLine("[X] Errore durante il refactoring:");
      Console.WriteLine(ex.ToString());
    }
  }
}
