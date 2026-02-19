using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  // ---------------------------------------------------------
  // RISOLUZIONE ACCESSI AI CAMPI
  // ---------------------------------------------------------

  private static void ResolveFieldAccesses(
      VbModule mod,
      VbProcedure proc,
      string[] fileLines,
      Dictionary<string, VbTypeDef> typeIndex,
      Dictionary<string, string> env)
  {
    // Controlli di sicurezza per evitare IndexOutOfRangeException
    if (proc.StartLine <= 0)
    {
      Console.WriteLine($"[WARN] Procedure {proc.Name} has invalid StartLine: {proc.StartLine}, using LineNumber: {proc.LineNumber}");
      proc.StartLine = proc.LineNumber;
    }

    if (proc.EndLine <= 0)
    {
      Console.WriteLine($"[WARN] Procedure {proc.Name} has invalid EndLine: {proc.EndLine}, scanning until end of file");
      proc.EndLine = fileLines.Length;
    }

    var startIndex = Math.Max(0, proc.StartLine - 1);
    var endIndex = Math.Min(fileLines.Length, proc.EndLine);

    if (startIndex >= fileLines.Length)
    {
      Console.WriteLine($"[WARN] Procedure {proc.Name} StartLine {proc.StartLine} is beyond file length {fileLines.Length}");
      return;
    }

    for (int i = startIndex; i < endIndex; i++)
    {
      var raw = fileLines[i].Trim();

      // Rimuovi commenti
      var noComment = raw;
      var idx = noComment.IndexOf("'");
      if (idx >= 0)
        noComment = noComment.Substring(0, idx).Trim();

      foreach (Match m in ReFieldAccess.Matches(noComment))
      {
        var varName = m.Groups[1].Value;
        var fieldName = m.Groups[2].Value;

        // Estrai il nome base della variabile rimuovendo l'accesso array
        // Es: "m_QueuePolling(i)" -> "m_QueuePolling"
        var baseVarName = varName;
        var parenIndex = varName.IndexOf('(');
        if (parenIndex >= 0)
          baseVarName = varName.Substring(0, parenIndex);

        if (env.TryGetValue(baseVarName, out var typeName))
        {
          if (typeIndex.TryGetValue(typeName, out var typeDef))
          {
            var field = typeDef.Fields.FirstOrDefault(f =>
                !string.IsNullOrEmpty(f.Name) &&
                string.Equals(f.Name, fieldName, StringComparison.OrdinalIgnoreCase));

            if (field != null)
            {
              field.Used = true;
              field.References.AddLineNumber(mod.Name, proc.Name, i + 1);
            }
          }
        }
      }
    }
  }

  // ---------------------------------------------------------
  // RISOLUZIONE ACCESSI AI CONTROLLI
  // ---------------------------------------------------------

  private static void ResolveControlAccesses(
      VbModule mod,
      VbProcedure proc,
      string[] fileLines)
  {
    // Indicizza i controlli del modulo per nome (ora unificati, nessun duplicato)
    var controlIndex = mod.Controls.ToDictionary(
        c => c.Name,
        c => c,
        StringComparer.OrdinalIgnoreCase);

    // Controlli di sicurezza per evitare IndexOutOfRangeException
    if (proc.StartLine <= 0)
      proc.StartLine = proc.LineNumber;

    if (proc.EndLine <= 0)
      proc.EndLine = fileLines.Length;

    var startIndex = Math.Max(0, proc.StartLine - 1);
    var endIndex = Math.Min(fileLines.Length, proc.EndLine);

    if (startIndex >= fileLines.Length)
      return;

    for (int i = startIndex; i < endIndex; i++)
    {
      var raw = fileLines[i].Trim();

      // Fine procedura - controlla SOLO i terminatori specifici della procedura
      // (Non usa "End " generico per evitare match con End If, End With, ecc.)
      if (i > proc.LineNumber - 1 && IsProcedureEndLine(raw, proc.Kind))
        break;

      // Rimuovi commenti
      var noComment = raw;
      var idx = noComment.IndexOf("'");
      if (idx >= 0)
        noComment = noComment.Substring(0, idx).Trim();

      // Pattern: controlName.Property o controlName.Method() oppure controlName(index).Property
      // ANCHE: ModuleName.controlName(index).Property per referenze cross-module
      foreach (Match m in Regex.Matches(noComment, @"(\w+)(?:\([^\)]*\))?\.(\w+)"))
      {
        var controlName = m.Groups[1].Value;

        // Verifica se è un controllo del form corrente
        if (controlIndex.TryGetValue(controlName, out var control))
        {
          MarkControlAsUsed(control, mod.Name, proc.Name, i + 1);
        }
      }

      // Pattern avanzato per referenze cross-module: ModuleName.ControlName(index).Property
      foreach (Match m in Regex.Matches(noComment, @"(\w+)\.(\w+)(?:\([^\)]*\))?\.(\w+)"))
      {
        var moduleName = m.Groups[1].Value;
        var controlName = m.Groups[2].Value;

        // Cerca il modulo nel progetto
        var targetModule = mod.Owner?.Modules?.FirstOrDefault(module =>
            string.Equals(module.Name, moduleName, StringComparison.OrdinalIgnoreCase));

        if (targetModule != null)
        {
          // Cerca TUTTI i controlli con lo stesso nome (array di controlli)
          var controls = targetModule.Controls.Where(c =>
              string.Equals(c.Name, controlName, StringComparison.OrdinalIgnoreCase));

          foreach (var control in controls)
          {
            MarkControlAsUsed(control, mod.Name, proc.Name, i + 1);
          }
        }
      }
    }
  }

  // ---------------------------------------------------------
  // RISOLUZIONE REFERENCE PARAMETRI E VARIABILI LOCALI
  // ---------------------------------------------------------

  private static void ResolveParameterAndLocalVariableReferences(
      VbModule mod,
      VbProcedure proc,
      string[] fileLines)
  {
    // Indicizza i parametri per nome
    var parameterIndex = proc.Parameters.ToDictionary(
        p => p.Name,
        p => p,
        StringComparer.OrdinalIgnoreCase);

    // Indicizza le variabili locali per nome
    var localVariableIndex = proc.LocalVariables.ToDictionary(
        v => v.Name,
        v => v,
        StringComparer.OrdinalIgnoreCase);

    // Controlli di sicurezza per evitare IndexOutOfRangeException
    if (proc.StartLine <= 0)
      proc.StartLine = proc.LineNumber;

    if (proc.EndLine <= 0)
      proc.EndLine = fileLines.Length;

    var startIndex = Math.Max(0, proc.StartLine - 1);
    var endIndex = Math.Min(fileLines.Length, proc.EndLine);

    if (startIndex >= fileLines.Length)
      return;

    for (int i = startIndex; i < endIndex; i++)
    {
      var raw = fileLines[i].Trim();
      int currentLineNumber = i + 1;

      // Fine procedura - controlla SOLO i terminatori specifici della procedura
      // (Non usa "End " generico per evitare match con End If, End With, ecc.)
      if (i > proc.LineNumber - 1 && IsProcedureEndLine(raw, proc.Kind))
        break;

      // Rimuovi commenti
      var noComment = raw;
      var idx = noComment.IndexOf("'", StringComparison.Ordinal);
      if (idx >= 0)
        noComment = noComment.Substring(0, idx).Trim();

      // Rimuovi stringhe per evitare di catturare nomi dentro stringhe
      noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

      // Cerca tutti i token word nella riga
      foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\b"))
      {
        var tokenName = m.Groups[1].Value;

        // Controlla se è un parametro
        if (parameterIndex.TryGetValue(tokenName, out var parameter))
        {
          parameter.Used = true;
          parameter.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
        }

        // Controlla se è una variabile locale
        if (localVariableIndex.TryGetValue(tokenName, out var localVar))
        {
          // Esclude la riga di dichiarazione della variabile (usa direttamente LineNumber)
          if (localVar.LineNumber == currentLineNumber)
            continue;

          localVar.Used = true;
          localVar.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
        }
      }
    }
  }

  // ---------------------------------------------------------
  // MARCATURA VALORI ENUM USATI
  // ---------------------------------------------------------

  private static void MarkUsedEnumValues(VbProject project)
  {
    // Indicizza tutti i valori enum per nome
    var allEnumValues = new Dictionary<string, List<VbEnumValue>>(StringComparer.OrdinalIgnoreCase);

    foreach (var mod in project.Modules)
    {
      foreach (var enumDef in mod.Enums)
      {
        foreach (var enumValue in enumDef.Values)
        {
          if (!allEnumValues.ContainsKey(enumValue.Name))
            allEnumValues[enumValue.Name] = new List<VbEnumValue>();

          allEnumValues[enumValue.Name].Add(enumValue);
        }
      }
    }

    // Cerca l'uso dei valori enum in tutti i moduli
    foreach (var mod in project.Modules)
    {
      var fileLines = File.ReadAllLines(mod.FullPath);

      foreach (var proc in mod.Procedures)
      {
        // Scansiona il corpo della procedura
        for (int i = proc.LineNumber - 1; i < fileLines.Length; i++)
        {
          var line = fileLines[i];

          // Fine procedura
          if (line.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
            break;

          // Rimuovi commenti
          var noComment = line;
          var idx = noComment.IndexOf("'");
          if (idx >= 0)
            noComment = noComment.Substring(0, idx);

          // Cerca ogni valore enum nel codice
          foreach (var kvp in allEnumValues)
          {
            var enumValueName = kvp.Key;
            var enumValues = kvp.Value;

            // Usa word boundary per evitare match parziali
            if (Regex.IsMatch(noComment, $@"\b{Regex.Escape(enumValueName)}\b", RegexOptions.IgnoreCase))
            {
              // Marca tutti i valori enum con questo nome (potrebbero esserci duplicati in enum diversi)
              foreach (var enumValue in enumValues)
              {
                enumValue.Used = true;
                enumValue.References.AddLineNumber(mod.Name, proc.Name, i + 1);
              }
            }
          }
        }
      }
    }
  }

  // ---------------------------------------------------------
  // MARCATURA EVENTI USATI (RaiseEvent)
  // ---------------------------------------------------------

  private static void MarkUsedEvents(VbProject project)
  {
    // Indicizza tutti gli eventi per modulo
    var eventsByModule = new Dictionary<string, List<VbEvent>>(StringComparer.OrdinalIgnoreCase);

    foreach (var mod in project.Modules)
    {
      if (mod.Events.Count > 0)
        eventsByModule[mod.Name] = mod.Events;
    }

    // Cerca RaiseEvent in tutti i moduli
    foreach (var mod in project.Modules)
    {
      var fileLines = File.ReadAllLines(mod.FullPath);

      foreach (var proc in mod.Procedures)
      {
        // Scansiona il corpo della procedura
        for (int i = proc.LineNumber - 1; i < fileLines.Length; i++)
        {
          var line = fileLines[i];

          // Fine procedura
          if (line.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
            break;

          // Rimuovi commenti
          var noComment = line;
          var idx = noComment.IndexOf("'");
          if (idx >= 0)
            noComment = noComment.Substring(0, idx);

          // Pattern: RaiseEvent EventName o RaiseEvent EventName(params)
          var raiseEventMatch = Regex.Match(noComment, @"RaiseEvent\s+(\w+)", RegexOptions.IgnoreCase);
          if (raiseEventMatch.Success)
          {
            var eventName = raiseEventMatch.Groups[1].Value;

            // Cerca l'evento nel modulo corrente (gli eventi sono sempre locali al modulo/classe)
            if (eventsByModule.TryGetValue(mod.Name, out var events))
            {
              var evt = events.FirstOrDefault(e =>
                  e.Name.Equals(eventName, StringComparison.OrdinalIgnoreCase));

              if (evt != null)
              {
                evt.Used = true;
                evt.References.AddLineNumber(mod.Name, proc.Name, i + 1);
              }
            }
          }
        }
      }
    }
  }
}
