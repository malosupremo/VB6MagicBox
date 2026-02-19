using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  public static partial class NamingConvention
  {
    private static readonly HashSet<string> HungarianPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        "int", "str", "lng", "dbl", "sng", "cur", "bol", "byt", "chr", "dat", "obj", "arr", "udt"
    };

    private static string GetBaseNameFromHungarian(string name, string variableType = "")
    {
      if (string.IsNullOrEmpty(name)) return name;

      // Exclude common non-type prefixes that should be preserved
      var preservePrefixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "msg",
            "plc"
        };

      if (name.Length >= 3)
      {
        var prefix = name.Substring(0, 3);
        if (preservePrefixes.Contains(prefix))
          return name;
      }

      // Caso speciale: mCamelCase (es: mLogger) -> Logger
      // Riconosce pattern m + maiuscola diretta (convenzione member variable)
      if (name.Length > 1 &&
          name.StartsWith("m", StringComparison.OrdinalIgnoreCase) &&
          char.IsUpper(name[1]))
      {
        return name.Substring(1); // Rimuovi la 'm' iniziale
      }

      // Caso speciale: m_Nome (già conforme alla convenzione) -> mantieni così com'è
      if (name.StartsWith("m_", StringComparison.OrdinalIgnoreCase))
      {
        return name; // Già nel formato corretto m_Nome
      }

      // Caso speciale: s_Nome (già conforme alla convenzione) -> mantieni così com'è
      if (name.StartsWith("s_", StringComparison.OrdinalIgnoreCase))
      {
        return name; // Già nel formato corretto s_Nome
      }

      // For non-module locals: only strip known 3-letter prefixes when name does NOT start with 'm'
      // (ma ora gestiamo già i casi mCamelCase e m_ sopra)
      if (name.StartsWith("m", StringComparison.OrdinalIgnoreCase))
        return name;

      var match = Regex.Match(name, @"^([a-z]{3})([A-Z].*)");
      if (match.Success)
      {
        var prefix = match.Groups[1].Value;

        // Caso speciale per 'cur': rimuovi solo se il tipo è Currency
        if (prefix.Equals("cur", StringComparison.OrdinalIgnoreCase))
        {
          // Rimuovi 'cur' solo se il tipo è Currency, altrimenti mantieni (potrebbe essere 'current')
          if (!string.IsNullOrEmpty(variableType) &&
              variableType.Equals("Currency", StringComparison.OrdinalIgnoreCase))
          {
            return match.Groups[2].Value; // Rimuovi prefisso cur per variabili Currency
          }
          else
          {
            return name; // Mantieni nome completo (probabilmente 'current')
          }
        }

        if (HungarianPrefixes.Contains(prefix))
          return match.Groups[2].Value;
      }

      return name;
    }

    private static bool IsPrivate(string visibility)
    {
      return string.IsNullOrEmpty(visibility) ||
             visibility.Equals("Private", StringComparison.OrdinalIgnoreCase) ||
             visibility.Equals("Dim", StringComparison.OrdinalIgnoreCase);
    }

    public static string ToPascalCase(string s)
    {
      if (string.IsNullOrEmpty(s)) return s;

      var valid = Regex.Replace(s, @"[^\w]", ""); // Remove invalid chars
      if (string.IsNullOrEmpty(valid)) return valid;

      var parts = valid.Split('_', StringSplitOptions.RemoveEmptyEntries);
      var result = string.Empty;

      foreach (var part in parts)
      {
        if (part.Length > 0)
        {
          // Se tutte le lettere del segmento sono maiuscole (es. "MAX", "EXEC1", "SQM242HND")
          // applica PascalCase a ogni sequenza di lettere, lasciando le cifre invariate
          if (part.Any(char.IsLetter) && part.Where(char.IsLetter).All(char.IsUpper))
          {
            result += Regex.Replace(part, @"[A-Za-z]+", m =>
                char.ToUpper(m.Value[0]) + m.Value.Substring(1).ToLower());
          }
          else
          {
            // Per parti con case misto, preserva la casing esistente ma assicura maiuscola iniziale
            result += char.ToUpper(part[0]) + part.Substring(1);
          }
        }
      }
      return result;
    }

    public static string ToCamelCase(string s)
    {
      var pascal = ToPascalCase(s);
      if (string.IsNullOrEmpty(pascal)) return pascal;
      return char.ToLower(pascal[0]) + pascal.Substring(1);
    }

    public static string ToUpperSnakeCase(string s)
    {
      if (string.IsNullOrEmpty(s)) return s;

      // Se il nome è già in UPPER_SNAKE_CASE (tutte maiuscole + underscore), restituiscilo così com'è
      // Questo preserva nomi come RIC_AL_PDxI1_LOWVOLTAGE
      if (Regex.IsMatch(s, @"^[A-Z0-9_]+$"))
        return s;

      // Altrimenti converti da camelCase/PascalCase a UPPER_SNAKE_CASE
      // Use regex to handle transitions from lower to upper case, and acronyms
      var s1 = Regex.Replace(s, @"([a-z])([A-Z])", "$1_$2");
      var s2 = Regex.Replace(s1, @"([A-Z])([A-Z][a-z])", "$1_$2");

      var res = s2.ToUpper().Replace("__", "_");
      return Regex.Replace(res, @"_+", "_").Trim('_');
    }

    private static bool IsReservedWord(string name)
    {
      if (string.IsNullOrWhiteSpace(name)) return false;
      // Controlla sia keyword C# che keyword/built-in VB6 (es. Loop, Len, Mid, Left…)
      return IsCSharpKeyword(name) || VbKeywords.Contains(name);
    }

    /// <summary>
    /// Controlla solo le keyword C#. Usato per i campi dei Type, dove l'accesso
    /// avviene sempre via punto (myVar.Len) e i built-in VB6 non creano conflitti.
    /// </summary>
    private static bool IsCSharpKeyword(string name)
    {
      if (string.IsNullOrWhiteSpace(name)) return false;
      return CSharpKeywords.Contains(name.ToLowerInvariant());
    }

    private static string ToPascalCaseFromScreamingSnake(string name)
    {
      if (string.IsNullOrEmpty(name)) return name;

      // If the name doesn't contain underscores and has lowercase letters,
      // it's already in PascalCase/camelCase format - preserve it
      if (!name.Contains('_') && name.Any(char.IsLower))
      {
        return ToPascalCase(name);
      }

      var parts = name.Split('_', StringSplitOptions.RemoveEmptyEntries);
      var result = new System.Text.StringBuilder();

      foreach (var part in parts)
      {
        if (part.Length > 0)
        {
          // Capitalize first letter, lowercase the rest
          result.Append(char.ToUpperInvariant(part[0]));
          if (part.Length > 1)
          {
            result.Append(part.Substring(1).ToLowerInvariant());
          }
        }
      }

      return result.ToString();
    }

    /// <summary>
    /// Risolve i conflitti di naming aggiungendo un suffisso numerico progressivo.
    /// Es: se "result" è già usato, prova "result2", "result3", ecc.
    /// </summary>
    private static string ResolveNameConflict(string proposedName, HashSet<string> usedNames)
    {
      if (string.IsNullOrEmpty(proposedName)) return proposedName;

      var finalName = proposedName;
      int counter = 2;

      while (usedNames.Contains(finalName))
      {
        finalName = proposedName + counter;
        counter++;
      }

      return finalName;
    }

    /// <summary>
    /// Verifica se una stringa è già in formato PascalCase.
    /// Una stringa è PascalCase se inizia con una lettera maiuscola
    /// e non contiene underscore o caratteri speciali.
    /// </summary>
    private static bool IsPascalCase(string name)
    {
      if (string.IsNullOrEmpty(name)) return false;

      // Deve iniziare con lettera maiuscola
      if (!char.IsUpper(name[0])) return false;

      // Non deve contenere underscore
      if (name.Contains('_')) return false;

      // Deve contenere solo lettere e numeri
      return name.All(c => char.IsLetterOrDigit(c));
    }
  }
}
