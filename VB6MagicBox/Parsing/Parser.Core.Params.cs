using VB6MagicBox.Models;
using System.Text.RegularExpressions;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  // -------------------------
  // PARAMETRI
  // -------------------------

  private static List<VbParameter> ParseParameters(string paramList, int originalLineNumber = 0)
  {
    var result = new List<VbParameter>();
    if (string.IsNullOrWhiteSpace(paramList))
      return result;

    var parts = paramList.Split(',');
    var reParam = new Regex(
        @"^(Optional\s+)?(ByVal|ByRef)?\s*(\w+)([$%&!#@]?)(\([^)]*\))?\s*(As\s+([\w\.\(\)]+))?",
        RegexOptions.IgnoreCase);

    foreach (var p in parts)
    {
      var s = p.Trim();
      if (string.IsNullOrEmpty(s))
        continue;

      var m = reParam.Match(s);
      if (!m.Success)
        continue;

      result.Add(new VbParameter
      {
        Name = m.Groups[3].Value,
        Passing = string.IsNullOrEmpty(m.Groups[2].Value) ? "ByRef" : m.Groups[2].Value,
        Type = m.Groups[7].Value,
        Used = false,
        LineNumber = originalLineNumber
      });
    }

    return result;
  }

  /// <summary>
  /// Aggiunge automaticamente References per i parametri delle Declare Function/Sub
  /// che si estendono su più righe con il carattere di continuazione "_"
  /// </summary>
  private static void AddParameterReferencesForMultilineDeclaration(
      VbProcedure procedure,
      string moduleName,
      string[] originalLines,
      int startLineNumber,
      int[] lineMapping,
      int collapsedIndex)
  {
    if (procedure.Parameters == null || procedure.Parameters.Count == 0)
      return;

    // Trova tutte le righe originali che costituivano questa dichiarazione collapsed
    var originalStartIndex = startLineNumber - 1; // Convert to 0-based
    var originalEndIndex = originalStartIndex;

    // Trova l'ultima riga della dichiarazione (seguendo i "_")
    while (originalEndIndex < originalLines.Length - 1)
    {
      var line = originalLines[originalEndIndex].TrimEnd();
      if (!line.EndsWith("_"))
        break;
      originalEndIndex++;
    }

    // Per ogni parametro, cerca in quale riga originale si trova
    foreach (var param in procedure.Parameters)
    {
      for (int lineIdx = originalStartIndex; lineIdx <= originalEndIndex; lineIdx++)
      {
        var originalLine = originalLines[lineIdx];

        // Cerca il nome del parametro in questa riga (word boundary per evitare match parziali)
        var paramPattern = $@"\b{Regex.Escape(param.Name)}\b";
        if (Regex.IsMatch(originalLine, paramPattern, RegexOptions.IgnoreCase))
        {
          // Trovato! Aggiungi una Reference a questa riga specifica
          // ma solo se non l'ho già segnato

          if (!param.References.Any(r => r.Module == moduleName && r.Procedure == procedure.Name))
          {
            param.References.Add(new VbReference
            {
              Module = moduleName, // Sarà impostato dal chiamante se necessario
              Procedure = procedure.Name,
              LineNumbers = new List<int> { lineIdx + 1 } // Convert back to 1-based
            });
          }

          // Un parametro può apparire solo una volta, quindi esci dal loop
          break;
        }
      }
    }
  }
}
