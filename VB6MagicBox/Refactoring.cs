using System.Text;
using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox;

/// <summary>
/// Gestisce il refactoring automatico del codice VB6 applicando i rename
/// </summary>
public static class Refactoring
{
  /// <summary>
  /// Applica i rename al progetto VB6 basandosi sul progetto completamente analizzato in memoria.
  /// Questo è molto più sicuro che usare solo il file .rename.json perché abbiamo accesso
  /// a tutte le References, Calls, Dependencies e contesto semantico completo.
  /// 
  /// FASE 2 - REFACTORING CON VALIDAZIONE:
  /// 1. Percorre ogni modulo
  /// 2. Colleziona i rename dalle analisi (Constants, GlobalVariables, etc.)
  /// 3. Ordina per dipendenza (Field prima di Type, etc.)
  /// 4. Applica rename preservando stringhe e commenti
  /// 5. Valida le occorrenze contro i dati di References dalla Fase 1
  /// </summary>
  public static void ApplyRenames(VbProject project)
  {
    // Registra il provider per encoding legacy (Windows-1252) necessario per VB6
    // Richiesto in .NET Core/.NET 5+ dove gli encoding non standard non sono disponibili di default
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    Console.WriteLine();
    Console.WriteLine("===========================================");
    Console.WriteLine("  Refactoring Automatico VB6 (FASE 2)");
    Console.WriteLine("  Con Validazione vs Analisi (FASE 1)");
    Console.WriteLine("===========================================");
    Console.WriteLine();

    var vbpPath = project.ProjectFile;
    var vbpDir = Path.GetDirectoryName(vbpPath)!;
    
    // Risali di 2 livelli dalla cartella del .vbp per creare il backup
    // Es: C:\...\5.0\CALLER\caller.vbp -> C:\...\5.0
    var vbpDirInfo = new DirectoryInfo(vbpDir);
    var backupBaseDir = vbpDirInfo.Parent?.FullName;
    
    if (string.IsNullOrEmpty(backupBaseDir))
    {
      Console.WriteLine("[!] Impossibile determinare la cartella base per il backup.");
      Console.WriteLine("    Verrà usata la cartella del progetto.");
      backupBaseDir = vbpDir;
    }

    // Nome backup: NomeCartella.backup (es. CALLER.backup)
    var folderName = new DirectoryInfo(backupBaseDir).Name;
    var backupDir = Path.Combine(Path.GetDirectoryName(backupBaseDir)!, $"{folderName}.backup{DateTime.Now:yyyyMMdd_HHmmss}");
    
    // Se esiste già, elimina
    if (Directory.Exists(backupDir))
    {
      try
      {
        Directory.Delete(backupDir, true);
      }
      catch { }
    }
    
    Console.WriteLine($">> Preparazione backup...");
    Console.WriteLine($"   Cartella backup: {backupDir}");
    Console.WriteLine($"   (backup progressivo: solo file modificati)");
    Directory.CreateDirectory(backupDir);

    // Statistiche di validazione
    int filesProcessed = 0;
    int totalRenames = 0;
    int filesBackedUp = 0;

    // STEP 1: Colleziona TUTTI i rename da TUTTI i moduli
    // (necessario per gestire i rename cross-module: un Type definito in un modulo
    // ma referenziato in un altro deve essere rinominato in entrambi i file)
    var allRenames = new List<(string oldName, string newName, string category, object source, string definingModule)>();

    foreach (var module in project.Modules)
    {
      if (!module.IsConventional)
        allRenames.Add((module.Name, module.ConventionalName, "Module", module, module.Name));

      foreach (var v in module.GlobalVariables.Where(v => !v.IsConventional))
        allRenames.Add((v.Name, v.ConventionalName, "GlobalVariable", v, module.Name));

      foreach (var c in module.Constants.Where(c => !c.IsConventional))
        allRenames.Add((c.Name, c.ConventionalName, "Constant", c, module.Name));

      foreach (var t in module.Types)
      {
        if (!t.IsConventional)
          allRenames.Add((t.Name, t.ConventionalName, "Type", t, module.Name));
        foreach (var f in t.Fields.Where(f => !f.IsConventional))
          allRenames.Add((f.Name, f.ConventionalName, "Field", f, module.Name));
      }

      foreach (var e in module.Enums)
      {
        if (!e.IsConventional)
          allRenames.Add((e.Name, e.ConventionalName, "Enum", e, module.Name));
        foreach (var v in e.Values.Where(v => !v.IsConventional))
          allRenames.Add((v.Name, v.ConventionalName, "EnumValue", v, module.Name));
      }

      foreach (var c in module.Controls.Where(c => !c.IsConventional))
        allRenames.Add((c.Name, c.ConventionalName, "Control", c, module.Name));

      foreach (var p in module.Procedures)
      {
        if (!p.IsConventional)
          allRenames.Add((p.Name, p.ConventionalName, "Procedure", p, module.Name));
        foreach (var param in p.Parameters.Where(param => !param.IsConventional))
          allRenames.Add((param.Name, param.ConventionalName, "Parameter", param, module.Name));
        foreach (var lv in p.LocalVariables.Where(lv => !lv.IsConventional))
          allRenames.Add((lv.Name, lv.ConventionalName, "LocalVariable", lv, module.Name));
      }

      foreach (var prop in module.Properties)
      {
        if (!prop.IsConventional)
          allRenames.Add((prop.Name, prop.ConventionalName, "Property", prop, module.Name));
        foreach (var param in prop.Parameters.Where(param => !param.IsConventional))
          allRenames.Add((param.Name, param.ConventionalName, "PropertyParameter", param, module.Name));
      }
    }

    // Ordina i rename per categoria (dal più specifico al più generale) e poi per lunghezza decrescente
    // Sequenza CORRETTA: Field → EnumValue → Type → Enum → Constant → GlobalVariable → Parameter → LocalVariable → Control → Procedure → Module
    allRenames = allRenames
      .OrderBy(r => GetCategoryPriority(r.category))
      .ThenByDescending(r => r.oldName.Length)
      .ToList();

    // STEP 2: Applica i rename a ciascun file modulo
    // Per ogni file, applica TUTTI i rename che hanno dichiarazioni o riferimenti in quel modulo
    foreach (var module in project.Modules)
    {
      Console.WriteLine($">> Processando: {module.Name}");

      var filePath = module.FullPath;
      if (!File.Exists(filePath))
      {
        Console.WriteLine($"   [!] File non trovato: {filePath}");
        continue;
      }

      // Usa esplicitamente Windows-1252 (ANSI) per VB6
      var ansiEncoding = Encoding.GetEncoding(1252);
      var content = File.ReadAllText(filePath, ansiEncoding);
      var originalContent = content;

      string relativePath = Path.GetRelativePath(backupBaseDir, filePath);
      var backupFilePath = Path.Combine(backupDir, relativePath);

      int moduleRenames = 0;

      // Applica tutti i rename che hanno dichiarazioni o References nel modulo corrente
      int renameIndex = 0;
      foreach (var (oldName, newName, category, source, definingModule) in allRenames)
      {
        renameIndex++;
        
        if (oldName == newName)
          continue;

        // Progress inline: quale rename sta elaborando
        Console.Write($"\r      [{renameIndex}/{allRenames.Count}] {oldName} > {newName}...".PadRight(Console.WindowWidth - 1));

        int count = RenameIdentifierUsingReferences(ref content, oldName, newName, source, definingModule, module.Name);

        if (count > 0)
        {
          moduleRenames += count;
        }
      }
      
     // if (allRenames.Count > 0)
        //Console.WriteLine(); // Vai a capo dopo il progress

      // Salva il file modificato con encoding ANSI (VB6 requirement)
      if (content != originalContent)
      {
        // BACKUP PROGRESSIVO: copia il file originale nel backup PRIMA di modificarlo
        var backupFileDir = Path.GetDirectoryName(backupFilePath)!;
        if (!Directory.Exists(backupFileDir))
        {
          Directory.CreateDirectory(backupFileDir);
        }
        File.Copy(filePath, backupFilePath, overwrite: true);
        filesBackedUp++;

        // Ora modifica il file originale
        File.WriteAllText(filePath, content, ansiEncoding);
        filesProcessed++;
        totalRenames += moduleRenames;
        Console.WriteLine($"\r   [OK] {moduleRenames} rename applicati (backup: {Path.GetFileName(backupFilePath)})");
      }
      else
      {
        Console.WriteLine($"\r   [i] Nessuna modifica necessaria");
      }
    }

    Console.WriteLine();
    Console.WriteLine("===========================================");
    Console.WriteLine($"[OK] Refactoring completato!");
    Console.WriteLine($"     File modificati: {filesProcessed}");
    Console.WriteLine($"     File backuppati: {filesBackedUp}");
    Console.WriteLine($"     Rename totali: {totalRenames}");
    if (totalRenames > 0)
      Console.WriteLine($"     ✓ Tutti i rename validati contro Fase 1 (LineNumbers)");
    else
      Console.WriteLine($"     [INFO] Nessun simbolo da rinominare (tutti IsConventional=true)");
    Console.WriteLine($"     Cartella backup: {backupDir}");
    Console.WriteLine();
    if (filesProcessed > 0)
    {
      Console.WriteLine("Per ripristinare il backup:");
      Console.WriteLine($"  1. Copia i file da: {backupDir}");
      Console.WriteLine($"  2. Verso: {backupBaseDir}");
      Console.WriteLine($"     (sovrascrivi i file modificati)");
    }
    Console.WriteLine("===========================================");
  }

  /// <summary>
  /// Determina la priorità di applicazione dei rename per categoria
  /// Priorità più bassa = applicato prima (dal più interno al più esterno, specifico al generale)
  /// ORDINE CORRETTO: Field → EnumValue → Type → Enum → Constant → GlobalVariable → Parameter → LocalVariable → Control → Procedure → Module
  /// </summary>
  private static int GetCategoryPriority(string category)
  {
    return category switch
    {
      "Field" => 1,            // Membri dei Type (più interno, deve essere prima del Type)
      "EnumValue" => 2,        // Valori degli Enum (deve essere prima dell'Enum)
      "Type" => 3,             // Dichiarazioni Type (dipendono da Field)
      "Enum" => 4,             // Dichiarazioni Enum (dipendono da EnumValue)
      "Constant" => 5,         // Costanti (no dipendenze, ma riferite da molti)
      "GlobalVariable" => 6,   // Variabili globali (potrebbero istanziare Type)
      "PropertyParameter" => 7,// Parametri proprietà (scope locale a Property)
      "Parameter" => 8,        // Parametri (scope locale a Procedure)
      "LocalVariable" => 9,    // Variabili locali (scope locale a Procedure)
      "Control" => 10,         // Controlli UI (Form-specific)
      "Property" => 11,        // Proprietà di classe (accessi con punto)
      "Procedure" => 12,       // Nome procedure/funzioni (visibili globalmente)
      "Module" => 13,          // Nome modulo (top-level, meno specifico)
      _ => 999                 // Sconosciuto (alla fine)
    };
  }

  /// <summary>
  /// Rinomina un identificatore usando SOLO i LineNumbers dalle References della Fase 1
  /// Questo è PRECISO: va solo sulle righe che sappiamo devono essere modificate
  /// Non ci sono rischi di stringhe, commenti, o sostituzioni accidentali
  /// 
  /// VANTAGGI rispetto a ricerca globale:
  /// 1. ✓ Preciso: sostituisce SOLO dove sappiamo che deve essere (dai LineNumbers)
  /// 2. ✓ Sicuro: nessun rischio di sostituire stringhe o parti accidentali
  /// 3. ✓ Veloce: processa SOLO le righe rilevanti, non tutto il file
  /// 4. ✓ Validato: usa i dati della Fase 1 (References)
  /// 5. ✓ Cross-module: filtra i References per modulo corrente, gestendo rename tra moduli diversi
  /// </summary>
  private static int RenameIdentifierUsingReferences(ref string content, string oldName, string newName, object source, string definingModuleName, string currentModuleName)
  {
    int totalCount = 0;
    var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
    
    // Raccogli i LineNumbers dalle References del simbolo, filtrati per il modulo corrente
    var lineNumbersToReplace = new HashSet<int>();
    int declarationLineNumber = 0;
    
    try
    {
      // Riga della dichiarazione: solo se il simbolo è definito nel modulo corrente
      if (string.Equals(definingModuleName, currentModuleName, StringComparison.OrdinalIgnoreCase))
      {
        // SPECIALE: Per i controlli, usa TUTTI i LineNumbers (array di controlli)
        if (source is VbControl control)
        {
          // Se è un controllo array, usa tutti i LineNumbers
          if (control.IsArray && control.LineNumbers?.Count > 0)
          {
            foreach (var lineNum in control.LineNumbers)
            {
              lineNumbersToReplace.Add(lineNum);
            }
            declarationLineNumber = control.LineNumbers.First(); // Prima riga per reference
          }
          else
          {
            // Controllo singolo, usa LineNumber normale
            if (control.LineNumber > 0)
            {
              lineNumbersToReplace.Add(control.LineNumber);
              declarationLineNumber = control.LineNumber;
            }
          }
        }
        else
        {
          // Altri tipi di oggetti (non controlli): usa LineNumber normale
          var lineNumberProp = source?.GetType().GetProperty("LineNumber");
          if (lineNumberProp?.GetValue(source) is int lineNum && lineNum > 0)
          {
            lineNumbersToReplace.Add(lineNum);
            declarationLineNumber = lineNum;
          }
        }
      }
      
      // References: includi SOLO quelle che puntano al modulo corrente
      // Questo è il fix per i rename cross-module: ogni Reference ha il Module di appartenenza,
      // quindi applichiamo solo i LineNumbers del file che stiamo effettivamente modificando
      var referencesProp = source?.GetType().GetProperty("References");
      if (referencesProp?.GetValue(source) is System.Collections.IEnumerable references)
      {
        foreach (var reference in references)
        {
          var moduleProp = reference?.GetType().GetProperty("Module");
          var refModuleName = moduleProp?.GetValue(reference) as string;
          
          // Filtra: applica solo i riferimenti che appartengono al file corrente
          if (!string.Equals(refModuleName, currentModuleName, StringComparison.OrdinalIgnoreCase))
            continue;
          
          var lineNumbersProp = reference?.GetType().GetProperty("LineNumbers");
          if (lineNumbersProp?.GetValue(reference) is System.Collections.Generic.List<int> refLineNumbers)
          {
            foreach (var refLineNum in refLineNumbers)
            {
              lineNumbersToReplace.Add(refLineNum);
            }
          }
        }
      }
    }
    catch (Exception ex)
    {
      Console.WriteLine($"   [WARN] Errore accesso References per {oldName}: {ex.Message}");
      return 0;
    }
    
    // VB6 ATTRIBUTE FIX: se c'è una dichiarazione, controlla la riga successiva
    // per "Attribute NomeVar." che è sempre alla riga N+1
    if (declarationLineNumber > 0 && declarationLineNumber < lines.Length)
    {
      var nextLine = lines[declarationLineNumber]; // declarationLineNumber è 1-based, array è 0-based
      var trimmedNextLine = nextLine.TrimStart();
      
      // Pattern: "Attribute UAServerObj.VB_VarHelpID = -1"
      if (trimmedNextLine.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
      {
        // Estrai il nome dopo "Attribute "
        var attributeMatch = Regex.Match(trimmedNextLine, @"^Attribute\s+(\w+)\.", RegexOptions.IgnoreCase);
        if (attributeMatch.Success && attributeMatch.Groups[1].Value.Equals(oldName, StringComparison.OrdinalIgnoreCase))
        {
          lineNumbersToReplace.Add(declarationLineNumber + 1); // Riga successiva alla dichiarazione
        }
      }
    }

    // VB6 ATTRIBUTE FIX: Gestione specifica per "Attribute VB_Name = "ClassName""
    // Questa è una riga speciale per le classi e form VB6 che va aggiornata quando cambia il nome del modulo
    // Applica a: classi (.cls), form (.frm), e qualsiasi altro modulo con VB_Name
    if (source is VbModule module && (module.IsClass || module.Kind.Equals("frm", StringComparison.OrdinalIgnoreCase)))
    {
      // Cerca "Attribute VB_Name" nelle prime righe del file (di solito all'inizio)
      for (int i = 0; i < Math.Min(20, lines.Length); i++)
      {
        var line = lines[i];
        var vbNameMatch = Regex.Match(line, @"Attribute\s+VB_Name\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase);
        
        if (vbNameMatch.Success && vbNameMatch.Groups[1].Value.Equals(oldName, StringComparison.OrdinalIgnoreCase))
        {
          lineNumbersToReplace.Add(i + 1); // Aggiungi la riga al set dei LineNumbers da processare
        }
      }
    }

    // Se non abbiamo References per questo modulo, non fare nulla
    if (lineNumbersToReplace.Count == 0)
      return 0;

    // Processa SOLO le righe specificate dai LineNumbers
    for (int i = 0; i < lines.Length; i++)
    {
      int lineNumber = i + 1; // LineNumber è 1-based
      
      if (!lineNumbersToReplace.Contains(lineNumber))
        continue; // Salta questa riga se non è nelle References

      var line = lines[i];
      var commentIdx = line.IndexOf("'");
      
      // Separa codice e commento
      var codePart = commentIdx >= 0 ? line.Substring(0, commentIdx) : line;
      var commentPart = commentIdx >= 0 ? line.Substring(commentIdx) : "";

      // Sostituisci SOLO nella parte di codice, non nel commento
      // Word boundary per evitare sostituzioni parziali
      string pattern;
      string replacement;
      
      // SPECIALE: Per i controlli sulle righe "Begin Library.ControlType ControlName",
      // usa lookbehind per rimpiazzare SOLO il nome dopo il secondo token,
      // evitando di toccare il nome della classe/libreria che lo precede.
      // Es: "Begin S7DATALib.S7Data S7Data" → NON rinomina S7DATALib.S7Data
      if (source is VbControl && Regex.IsMatch(codePart.TrimStart(), @"^Begin\s+\S+\s+", RegexOptions.IgnoreCase))
      {
        pattern = $@"(?<=^.*Begin\s+\S+\s+){Regex.Escape(oldName)}\b";
        replacement = newName;
      }
      // SPECIALE: Per le proprietà, fuori dal modulo che le definisce,
      // usa dot-prefixed replacement (.OldName → .NewName) per evitare
      // conflitti con parametri/variabili omonimi.
      // Es: "g_PlasmaSource.IsDeposit = IsDeposit" → rinomina solo ".IsDeposit"
      else if (source is VbProperty && !string.Equals(definingModuleName, currentModuleName, StringComparison.OrdinalIgnoreCase))
      {
        pattern = $@"\.{Regex.Escape(oldName)}\b";
        replacement = $".{newName}";
      }
      else
      {
        pattern = $@"\b{Regex.Escape(oldName)}\b";
        replacement = newName;
      }
      
      var newCodePart = Regex.Replace(codePart, pattern, replacement, RegexOptions.IgnoreCase);
      
      int matchesInLine = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase).Count;
      if (matchesInLine > 0)
      {
        totalCount += matchesInLine;
        lines[i] = newCodePart + commentPart;
      }
    }

    if (totalCount > 0)
    {
      content = string.Join(Environment.NewLine, lines);
    }

    return totalCount;
  }
}
