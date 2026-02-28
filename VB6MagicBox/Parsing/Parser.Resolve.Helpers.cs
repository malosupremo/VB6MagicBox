using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  /// <summary>
  /// Returns true when <paramref name="line"/> is the closing statement for a
  /// procedure of the given <paramref name="procKind"/> (Sub / Function / Property).
  /// </summary>
  private static bool IsProcedureEndLine(string line, string procKind)
  {
    if (string.IsNullOrEmpty(procKind))
      return false;

    string expectedEnd;
    if (procKind.Equals("Sub", StringComparison.OrdinalIgnoreCase))
      expectedEnd = "End Sub";
    else if (procKind.Equals("Function", StringComparison.OrdinalIgnoreCase))
      expectedEnd = "End Function";
    else if (procKind.StartsWith("Property", StringComparison.OrdinalIgnoreCase))
      expectedEnd = "End Property";
    else
      return false;

    var trimmed = line.TrimStart();
    return trimmed.Equals(expectedEnd, StringComparison.OrdinalIgnoreCase) ||
           trimmed.StartsWith(expectedEnd + " ", StringComparison.OrdinalIgnoreCase);
  }

  /// <summary>
  /// Marca un controllo come usato e aggiunge reference con line numbers
  /// </summary>
  private static void MarkControlAsUsed(VbControl control, string moduleName, string procedureName, int lineNumber, int occurrenceIndex = -1, int startChar = -1)
  {
    control.Used = true;
    control.References.AddLineNumber(moduleName, procedureName, lineNumber, occurrenceIndex, startChar, owner: control);
  }

  private static string StripInlineComment(string line)
  {
    if (string.IsNullOrEmpty(line))
      return line;

    bool inString = false;
    for (int i = 0; i < line.Length; i++)
    {
      var ch = line[i];
      if (ch == '"')
      {
        if (!inString)
          inString = true;
        else if (i + 1 < line.Length && line[i + 1] == '"')
          i++;
        else
          inString = false;
      }
      else if (!inString && ch == '\'')
      {
        return line.Substring(0, i);
      }
    }

    return line;
  }

  private static string MaskStringLiterals(string line)
  {
    if (string.IsNullOrEmpty(line))
      return line;

    var chars = line.ToCharArray();
    bool inString = false;
    for (int i = 0; i < chars.Length; i++)
    {
      if (chars[i] == '"')
      {
        if (!inString)
        {
          inString = true;
        }
        else if (i + 1 < chars.Length && chars[i + 1] == '"')
        {
          chars[i + 1] = ' ';
          i++;
        }
        else
        {
          inString = false;
        }
      }
      else if (inString)
      {
        chars[i] = ' ';
      }
    }

    return new string(chars);
  }
}
