using System.Text;
using VB6MagicBox.Models;

namespace VB6MagicBox;

/// <summary>
/// Gestisce il refactoring automatico del codice VB6 applicando le sostituzioni pre-calcolate.
/// NUOVA ARCHITETTURA: usa la lista Replaces costruita nella Fase 1 (BuildReplaces).
/// Nessun re-parsing, nessuna logica complessa di matching: solo applicazione meccanica delle sostituzioni.
/// </summary>
public static class Refactoring
{
    /// <summary>
    /// Applica i rename al progetto VB6 usando le sostituzioni pre-calcolate nella lista Replaces.
    /// 
    /// VANTAGGI:
    /// - ⚡ Velocissimo: nessun re-parsing, solo applicazione sostituzioni
    /// - ✓ Preciso: sostituzioni già calcolate con posizione esatta (carattere start/end)
    /// - 🛡️ Sicuro: nessun rischio di match accidentali (stringhe, commenti, etc.)
    /// - 📝 Verificabile: export .linereplace.json permette controllo manuale
    /// </summary>
    public static void ApplyRenames(VbProject project)
    {
        // Registra il provider per encoding legacy (Windows-1252) necessario per VB6
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        Console.WriteLineColor("===========================================", ConsoleColor.Yellow);
        Console.WriteLineColor("  2: Applica refactoring (da Replaces)", ConsoleColor.Yellow);
        Console.WriteLineColor("===========================================", ConsoleColor.Yellow);
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Gray;

        var vbpPath = project.ProjectFile;
        var vbpDir = Path.GetDirectoryName(vbpPath)!;

        // Risali alla cartella base per il backup
        var vbpDirInfo = new DirectoryInfo(vbpDir);
        var backupBaseDir = vbpDirInfo.Parent?.FullName ?? vbpDir;

        var folderName = new DirectoryInfo(backupBaseDir).Name;
        var backupDir = Path.Combine(Path.GetDirectoryName(backupBaseDir)!,
            $"{folderName}.backup{DateTime.Now:yyyyMMdd_HHmmss}");

        if (Directory.Exists(backupDir))
        {
            try { Directory.Delete(backupDir, true); } catch { }
        }

        Console.WriteLine($">> Preparazione backup...");
        Console.WriteLine($"   Cartella backup: {backupDir}");
        Directory.CreateDirectory(backupDir);

        int filesProcessed = 0;
        int totalReplaces = 0;
        int filesBackedUp = 0;

        // Usa esplicitamente Windows-1252 (ANSI) per VB6
        var ansiEncoding = Encoding.GetEncoding(1252);

        foreach (var module in project.Modules)
        {
            if (module.IsSharedExternal)
            {
                Console.WriteLine($">> {module.Name}: modulo condiviso, salto refactoring");
                continue;
            }

            if (module.Replaces.Count == 0)
            {
                Console.WriteLine($">> {module.Name}: nessuna sostituzione");
                continue;
            }

            Console.WriteLine($">> {module.Name}: {module.Replaces.Count} sostituzioni...");

            var filePath = module.FullPath;
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"   [!] File non trovato: {filePath}");
                continue;
            }

            var lines = File.ReadAllLines(filePath, ansiEncoding);
            var originalContent = string.Join(Environment.NewLine, lines);

            // Applica le sostituzioni ordinate (già ordinate da fine a inizio in BuildReplaces)
            // Raggruppa per riga per efficienza
            var replacesByLine = module.Replaces
                .GroupBy(r => r.LineNumber)
                .OrderByDescending(g => g.Key);

            int replacesApplied = 0;

            foreach (var lineGroup in replacesByLine)
            {
                int lineNumber = lineGroup.Key;
                if (lineNumber <= 0 || lineNumber > lines.Length)
                    continue;

                var line = lines[lineNumber - 1]; // Array è 0-based

                // Applica tutte le sostituzioni su questa riga (già ordinate per StartChar desc)
                var replacesForLine = lineGroup.OrderByDescending(r => r.StartChar).ToList();

                foreach (var replace in replacesForLine)
                {
                    // Verifica che la sostituzione sia ancora valida (potrebbe essere cambiata da sostituzioni precedenti)
                    if (replace.StartChar < 0 || replace.EndChar > line.Length || replace.StartChar >= replace.EndChar)
                        continue;

                    var currentText = line.Substring(replace.StartChar, replace.EndChar - replace.StartChar);

                    // Verifica che il testo corrente corrisponda ancora (case-insensitive)
                    if (!string.Equals(currentText, replace.OldText, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"   [WARN] Line {lineNumber}: expected '{replace.OldText}' at pos {replace.StartChar}, found '{currentText}'");
                        continue;
                    }

                    // Applica sostituzione
                    line = line.Remove(replace.StartChar, replace.EndChar - replace.StartChar);
                    line = line.Insert(replace.StartChar, replace.NewText);
                    replacesApplied++;
                }

                lines[lineNumber - 1] = line;
            }

            var newContent = string.Join(Environment.NewLine, lines);

            if (newContent != originalContent)
            {
                // Backup del file originale
                string relativePath = Path.GetRelativePath(backupBaseDir, filePath);
                var backupFilePath = Path.Combine(backupDir, relativePath);
                var backupFileDir = Path.GetDirectoryName(backupFilePath)!;

                if (!Directory.Exists(backupFileDir))
                    Directory.CreateDirectory(backupFileDir);

                File.Copy(filePath, backupFilePath, overwrite: true);
                filesBackedUp++;

                // Scrivi file modificato
                File.WriteAllText(filePath, newContent, ansiEncoding);
                filesProcessed++;
                totalReplaces += replacesApplied;

                Console.WriteLineColor($"   [OK] {replacesApplied} sostituzioni applicate", ConsoleColor.Green);
            }
            else
            {
                Console.WriteLineColor("   [i] Nessuna modifica (contenuto identico)", ConsoleColor.Cyan);
            }
        }

        Console.WriteLine();
        Console.WriteLineColor("===========================================", ConsoleColor.Green);
        Console.WriteLineColor("[OK] Refactoring completato!", ConsoleColor.Green);
        Console.WriteLineColor($"     File modificati:   {filesProcessed}", ConsoleColor.Green);
        Console.WriteLineColor($"     File backuppati:   {filesBackedUp}", ConsoleColor.Green);
        Console.WriteLineColor($"     Sostituzioni totali: {totalReplaces}", ConsoleColor.Green);
        Console.WriteLineColor($"     Cartella backup:   {backupDir}", ConsoleColor.Green);
        Console.WriteLineColor("===========================================", ConsoleColor.Green);
    }
}
