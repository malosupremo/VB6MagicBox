using VB6MagicBox.Models;
using System.Text.RegularExpressions;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  /// <summary>
  /// Ricostruisce le righe che utilizzano il carattere di continuazione "_"
  /// Restituisce un array di righe "collapsed" e un mapping dalla riga collapsed alla riga originale
  /// </summary>
  private static (string[] collapsedLines, int[] lineMapping) CollapseLineContinuations(string[] lines)
  {
    var collapsedLines = new List<string>();
    var lineMapping = new List<int>(); // Mapping da riga collapsed a riga originale
    
    for (int i = 0; i < lines.Length; i++)
    {
      var currentLine = lines[i];
      var originalLineNumber = i + 1;
      
      // Se la riga termina con "_", concatena le righe successive
      while (i < lines.Length && currentLine.TrimEnd().EndsWith("_"))
      {
        // Rimuovi il "_" e gli spazi finali
        var withoutContinuation = currentLine.TrimEnd();
        if (withoutContinuation.EndsWith("_"))
          withoutContinuation = withoutContinuation.Substring(0, withoutContinuation.Length - 1);
        
        currentLine = withoutContinuation;
        
        // Aggiungi la riga successiva (se esiste)
        if (i + 1 < lines.Length)
        {
          i++; // Vai alla riga successiva
          var nextLine = lines[i].TrimStart(); // Rimuovi indentazione
          currentLine += " " + nextLine; // Unisci con uno spazio
        }
      }
      
      collapsedLines.Add(currentLine);
      lineMapping.Add(originalLineNumber);
    }
    
    return (collapsedLines.ToArray(), lineMapping.ToArray());
  }

  private static readonly bool DEBUG_CALLS = false;  // Cambia a false per disattivare debug
  private static readonly string DEBUG_METHOD = "Alarm_Hnd";  // Stampa debug solo per questa funzione

  // -------------------------
  // REGEX PRINCIPALI
  // -------------------------

  private static readonly Regex ReFunction =
      new(@"^(Public|Private|Friend)?\s*(Static)?\s*Function\s+(\w+)\s*\((.*?)\)\s*(As\s+([\w\.\(\)]+))?",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReSub =
      new(@"^(Public|Private|Friend)?\s*(Static)?\s*Sub\s+(\w+)\s*\((.*?)\)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReDeclareFunction =
      new(@"^Public\s+Declare\s+Function\s+(\w+)\s+Lib\s+""([^""]+)""\s*\((.*?)\)\s*(As\s+([\w\.\(\)]+))?",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReDeclareSub =
      new(@"^Public\s+Declare\s+Sub\s+(\w+)\s+Lib\s+""([^""]+)""\s*\((.*?)\)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReConstant =
      new(@"^(Public|Private|Friend|Global)?\s*Const\s+(\w+)\s*(?:As\s+([\w\.\(\)]+))?\s*=\s*(.+)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReProperty =
      new(@"^(Public|Private|Friend)?\s*(Static)?\s*Property\s+(Get|Let|Set)\s+(\w+)\s*\((.*?)\)\s*(As\s+([\w\.\(\)]+))?",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReEvent =
      new(@"^(Public|Private|Friend)?\s*Event\s+(\w+)\s*\((.*?)\)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReGlobalVar =
      new(@"^(Public|Private|Global|Friend|Dim)\s+(WithEvents\s+)?(\w+)(\([^)]*\))?\s+As\s+([\w\.\(\)]+)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReLocalVar =
      new(@"^\s*(Dim|Static)\s+(\w+)(\([^)]*\))?\s+As\s+([\w\.\(\)]+)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReMemberVar =
      new(@"^(Public|Private|Friend|Dim)\s+(\w+)(\([^)]*\))?\s+As\s+([\w\.\(\)]+)",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReTypeStart =
      new(@"^(Public|Private|Friend)?\s*Type\s+(\w+)", RegexOptions.IgnoreCase);



  private static readonly Regex ReEnumStart =
      new(@"^(Public|Private|Friend)?\s*Enum\s+(\w+)", RegexOptions.IgnoreCase);


  private static readonly Regex ReField =
      new(@"^\s*(\w+)(\([^)]*\))?\s+As\s+([\w\.\(\)]+)", RegexOptions.IgnoreCase);

  private static readonly Regex ReCall =
      new(@"(?:(\w+)\s*\.\s*)?(\w+)\s*\(",
          RegexOptions.IgnoreCase);

  private static readonly Regex ReBareCall =
      new(@"^\s*(\w+)\s*$", RegexOptions.IgnoreCase);

  private static readonly Regex ReCallWithoutParens =
      new(@"^\s*(\w+)\s+", RegexOptions.IgnoreCase);

  private static readonly Regex ReFieldAccess =
      new(@"(\w+)\.(\w+)", RegexOptions.IgnoreCase);

  private static readonly Regex ReFormControlBegin =
      new(@"^Begin\s+(\S+)\s+(\w+)", RegexOptions.IgnoreCase);

  private static readonly HashSet<string> VbKeywords = new(StringComparer.OrdinalIgnoreCase)
      {
          // Keyword struttura
          "If","Then","Else","ElseIf","End","For","Next","Do","Loop","While","Wend",
          "Select","Case","Function","Sub","Property","Get","Let","Set","Call",
          "With","EndIf","End Sub","End Function","End Property","Type","Enum",
          "Private","Public","Friend","Global","Dim","Option","As","ByVal","ByRef",
          "Not","And","Or","Xor","Mod","New","On","Error","Resume","Goto",
          "IIf","Exit","ExitFunction","ExitSub","ExitFor","ExitDo","ExitWhile",
          // Funzioni native VB6
          "Abs","Array","Asc","AscB","AscW","Atn","CBoolean","CBool","CByte","CCur","CDate",
          "CDbl","Choose","Chr","ChrB","ChrW","CInt","CLng","Close","Const","Cos","CreateObject",
          "CSng","CStr","CVar","CVErr","Date","DateAdd","DateDiff","DatePart","DateSerial",
          "DateValue","Day","DDB","Declare","Defbool","DefByte","DefCur","DefDate","DefDbl",
          "DefInt","DefLng","DefObj","DefSng","DefStr","DefVar","DeleteSetting","Environ",
          "Eof","Error","ErrorNumber","Exp","FileAttr","FileDateTime","FileLen","Filter",
          "Fix","Format","FormatCurrency","FormatDateTime","FormatNumber","FormatPercent",
          "FreeFile","FV","GetAllSettings","GetAttr","GetObject","GetSetting","Goto","Hex",
          "Hour","IIf","Input","InputBox","InStr","InStrRev","Int","IPmt","IRR","IsArray",
          "IsDate","IsEmpty","IsError","IsMissing","IsNull","IsNumeric","IsObject","Join",
          "Kill","LBound","LCase","Left","Len","LenB","Line","LoadPicture","Loc","Lock",
          "Log","LTrim","Mid","MidB","Minute","MirrorPalette","Month","MonthName","MsgBox",
          "Name","Now","Nper","NPV","Oct","Open","Partition","Pmt","PPmt","PV","Rate",
          "Rename","Replace","Reset","RGB","Right","RmDir","Rnd","RTrim","SavePicture",
          "SaveSetting","Second","Seek","SetAttr","Shell","Sin","SLN","Space","Spc","Split",
          "Sqr","Str","StrComp","StrConv","String","StrReverse","SYD","Tab","Tan","Time",
          "Timer","TimeSerial","TimeValue","Trim","TypeName","UBound","UCase","Unlock",
          "Val","VarType","Weekday","WeekdayName","While","Width","Year", "String",
          "Attribute","Lib","Alias","Default","Global","Preserve","ResumeNext","Variant", "Me", "Event"
      };

  // -------------------------
  // DEBUG HELPER
  // -------------------------

  private static void DebugLog(string procName, int originalLineNumber, string callName, string reason, string regexName)
  {
    if (!DEBUG_CALLS)
      return;

    if (procName.Contains(DEBUG_METHOD, StringComparison.OrdinalIgnoreCase) || callName.Contains(DEBUG_METHOD, StringComparison.OrdinalIgnoreCase))
    {
      Console.WriteLine($"[DEBUG] Proc: {procName} | Line: {originalLineNumber} | Call: {callName} | Regex: {regexName} | {reason}");
    }
  }

  // -------------------------
  // ENTRY POINT PARSING
  // -------------------------

  public static VbProject ParseProjectFromVbp(string vbpPath)
  {
    var project = new VbProject { ProjectFile = vbpPath };
    var baseDir = Path.GetDirectoryName(vbpPath)!;

    var files = ParseVbpFile(vbpPath);

    foreach (var f in files)
    {
      var fullPath = Path.Combine(baseDir, f.Path);
      if (!File.Exists(fullPath))
        continue;

      var mod = ParseModule(fullPath, f.Kind);
      // set FullPath and relative Path
      mod.FullPath = fullPath;
      // Compute relative path with leading separator, relative to baseDir
      // Prefer simple substring when module is inside the vbp directory tree
      var baseDirFull = Path.GetFullPath(baseDir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
      var fullPathFull = Path.GetFullPath(fullPath);

      string rel;
      if (fullPathFull.StartsWith(baseDirFull + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
          fullPathFull.Equals(baseDirFull, StringComparison.OrdinalIgnoreCase))
      {
        rel = fullPathFull.Substring(baseDirFull.Length).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
      }
      else
      {
        rel = Path.GetRelativePath(baseDirFull, fullPathFull).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        if (!rel.StartsWith(Path.DirectorySeparatorChar) && !rel.StartsWith(Path.AltDirectorySeparatorChar))
          rel = Path.DirectorySeparatorChar + rel;
      }

      // Ensure leading separator
      if (!rel.StartsWith(Path.DirectorySeparatorChar) && !rel.StartsWith(Path.AltDirectorySeparatorChar))
        rel = Path.DirectorySeparatorChar + rel;

      mod.Path = rel;
      project.Modules.Add(mod);
    }

    return project;
  }

  // -------------------------
  // PARSING VBP
  // -------------------------

  private class VbpEntry
  {
    public string Kind { get; set; }
    public string Path { get; set; }
  }

  private static List<VbpEntry> ParseVbpFile(string vbpPath)
  {
    var list = new List<VbpEntry>();
    var lines = File.ReadAllLines(vbpPath);
    var baseDir = Path.GetDirectoryName(vbpPath)!;

    foreach (var raw in lines)
    {
      var line = raw.Trim();

      if (line.StartsWith("Form=", StringComparison.OrdinalIgnoreCase))
      {
        var parts = line.Substring("Form=".Length).Split(';');
        string path = null;

        if (parts.Length == 2)
          path = parts[1].Trim();
        else if (parts.Length == 1)
        {
          // Formato senza separatore: Form=Form\Restart.frm oppure Form=frmMain
          var name = parts[0].Trim();

          // Se ha già l'estensione, usalo direttamente
          if (name.EndsWith(".frm", StringComparison.OrdinalIgnoreCase))
            path = name;
          else
          {
            // Altrimenti aggiungere l'estensione
            var potentialPath = Path.Combine(baseDir, name + ".frm");
            if (File.Exists(potentialPath))
              path = name + ".frm";
          }
        }

        if (!string.IsNullOrEmpty(path))
          list.Add(new VbpEntry { Kind = "frm", Path = path });
      }
      else if (line.StartsWith("Module=", StringComparison.OrdinalIgnoreCase))
      {
        var parts = line.Substring("Module=".Length).Split(';');
        string path = null;

        if (parts.Length == 2)
          path = parts[1].Trim();
        else if (parts.Length == 1)
        {
          var name = parts[0].Trim();

          if (name.EndsWith(".bas", StringComparison.OrdinalIgnoreCase))
            path = name;
          else
          {
            var potentialPath = Path.Combine(baseDir, name + ".bas");
            if (File.Exists(potentialPath))
              path = name + ".bas";
          }
        }

        if (!string.IsNullOrEmpty(path))
          list.Add(new VbpEntry { Kind = "bas", Path = path });
      }
      else if (line.StartsWith("Class=", StringComparison.OrdinalIgnoreCase))
      {
        var parts = line.Substring("Class=".Length).Split(';');
        string path = null;

        if (parts.Length == 2)
          path = parts[1].Trim();
        else if (parts.Length == 1)
        {
          var name = parts[0].Trim();

          if (name.EndsWith(".cls", StringComparison.OrdinalIgnoreCase))
            path = name;
          else
          {
            var potentialPath = Path.Combine(baseDir, name + ".cls");
            if (File.Exists(potentialPath))
              path = name + ".cls";
          }
        }

        if (!string.IsNullOrEmpty(path))
          list.Add(new VbpEntry { Kind = "cls", Path = path });
      }
    }

    return list;
  }

  // -------------------------
  // PARSING MODULO
  // -------------------------

  public static VbModule ParseModule(string path, string kind)
  {
    var mod = new VbModule
    {
      Name = Path.GetFileNameWithoutExtension(path),
      ConventionalName = string.Empty,
      FullPath = string.Empty,
      Path = path,
      Kind = kind
    };

    // Set FullPath for I/O; Path may be overridden by caller to be relative
    mod.FullPath = path;

    var originalLines = File.ReadAllLines(path);
    
    // Gestisci le righe con continuazione "_"
    var (collapsedLines, lineMapping) = CollapseLineContinuations(originalLines);

    VbProcedure currentProc = null;
    VbTypeDef currentType = null;
    VbEnumDef currentEnum = null;
    bool insideControl = false;
    VbControl currentControl = null;
    string currentWith = null;

    int lineIndex = 0;

    foreach (var raw in collapsedLines)
    {
      var originalLineNumber = lineMapping[lineIndex];
      lineIndex++;
      var line = raw.Trim();

      // Rimuovi commenti
      var noComment = line;
      var commentIndex = noComment.IndexOf("'");
      if (commentIndex >= 0)
        noComment = noComment.Substring(0, commentIndex).Trim();

      // -------------------------
      // FORM CONTROLS
      // -------------------------
      if (kind.Equals("frm", StringComparison.OrdinalIgnoreCase))
      {
        var mc = ReFormControlBegin.Match(noComment);
        if (mc.Success)
        {
          insideControl = true;
          var controlName = mc.Groups[2].Value;
          var controlType = mc.Groups[1].Value;
          
          // Controlla se esiste già un controllo con lo stesso nome (array di controlli)
          currentControl = mod.Controls.FirstOrDefault(c => 
              c.Name.Equals(controlName, StringComparison.OrdinalIgnoreCase));
          
          if (currentControl != null)
          {
            // Esiste già - marca come array
            currentControl.IsArray = true;
          }
          else
          {
            // Nuovo controllo
            currentControl = new VbControl
            {
              ControlType = controlType,
              Name = controlName,
              LineNumber = originalLineNumber
            };
            mod.Controls.Add(currentControl);
          }
          continue;
        }

        if (insideControl && noComment.StartsWith("End", StringComparison.OrdinalIgnoreCase))
        {
          insideControl = false;
          currentControl = null;
          continue;
        }

        if (insideControl && currentControl != null)
        {
          var propMatch = Regex.Match(noComment, @"^\s*(\w+)\s*=\s*(.+)$");
          if (propMatch.Success && !propMatch.Groups[1].Value.Equals("Begin", StringComparison.OrdinalIgnoreCase))
          {
            var propName = propMatch.Groups[1].Value;
            var propValue = propMatch.Groups[2].Value.Trim();
            currentControl.Properties[propName] = propValue;
          }
          continue;
        }
      }

      // -------------------------
      // WITH
      // -------------------------
      if (noComment.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
      {
        currentWith = noComment.Substring(5).Trim();
        continue;
      }
      if (noComment.StartsWith("End With", StringComparison.OrdinalIgnoreCase))
      {
        currentWith = null;
        continue;
      }

      // -------------------------
      // TYPE (UDT)
      // -------------------------
      var mt = ReTypeStart.Match(noComment);
      bool isRealType =
          mt.Success &&
          !noComment.Contains("=") &&
          !raw.StartsWith(" ") &&
          !raw.StartsWith("\t") &&
          !noComment.Contains(" As ");

      if (isRealType)
      {
        currentType = new VbTypeDef
        {
          Name = mt.Groups[2].Value,  // Gruppo 2: nome del Type (dopo Public/Private opzionale)
          LineNumber = originalLineNumber
        };
        mod.Types.Add(currentType);
        continue;
      }

      if (currentType != null)
      {
        if (noComment.StartsWith("End Type", StringComparison.OrdinalIgnoreCase))
        {
          currentType = null;
        }
        else
        {
          var mf = ReField.Match(noComment);
          if (mf.Success)
          {
            var fieldName = mf.Groups[1].Value;
            var arrayPart = mf.Groups[2].Value; // può essere vuoto o "(dimensione)"
            var fieldType = mf.Groups[3].Value;
            
            currentType.Fields.Add(new VbField
            {
              Name = fieldName,
              Type = fieldType,
              IsArray = !string.IsNullOrEmpty(arrayPart),
              LineNumber = originalLineNumber
            });
          }
        }
        continue;
      }

      // -------------------------
      // ENUM
      // -------------------------
      var me = ReEnumStart.Match(noComment);
      if (me.Success && !noComment.Contains("="))
      {
        currentEnum = new VbEnumDef
        {
          Name = me.Groups[2].Value,  // Gruppo 2: nome dell'Enum (dopo Public/Private opzionale)
          LineNumber = originalLineNumber
        };
        mod.Enums.Add(currentEnum);
        continue;
      }

      if (currentEnum != null)
      {
        if (noComment.StartsWith("End Enum", StringComparison.OrdinalIgnoreCase))
        {
          currentEnum = null;
        }
        else
        {
          // Estrai solo il nome dell'enum value (prima di '=' se presente)
          // Es: "RIC_CONTROL_EMPTY = 0" ? "RIC_CONTROL_EMPTY"
          var enumValueName = noComment;
          var equalIndex = noComment.IndexOf('=');
          if (equalIndex >= 0)
          {
            enumValueName = noComment.Substring(0, equalIndex).Trim();
          }
          
          // Salta righe vuote o solo spazi
          if (!string.IsNullOrWhiteSpace(enumValueName))
          {
            currentEnum.Values.Add(new VbEnumValue { Name = enumValueName, LineNumber = originalLineNumber });
          }
        }
        continue;
      }

      // -------------------------
      // FUNCTION
      // -------------------------
      var mf2 = ReFunction.Match(noComment);
      if (mf2.Success)
      {
        var isStatic = Regex.IsMatch(noComment, @"\bStatic\b", RegexOptions.IgnoreCase);
        
        currentProc = new VbProcedure
        {
          Visibility = string.IsNullOrEmpty(mf2.Groups[1].Value) ? "Public" : mf2.Groups[1].Value,
          Name = mf2.Groups[3].Value,
          ConventionalName = string.Empty,
          Kind = "Function",
          IsStatic = isStatic,
          Scope = "Module",
          Parameters = ParseParameters(mf2.Groups[4].Value, originalLineNumber),
          ReturnType = mf2.Groups[6].Value,
          LineNumber = originalLineNumber
        };
        mod.Procedures.Add(currentProc);
        
        // Traccia event handler per controlli (es. Command1_Click)
        foreach (var control in mod.Controls)
        {
          var prefix = $"{control.Name}_";
          if (currentProc.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
          {
            // Questa procedura è un event handler di questo controllo
            var existingRef = control.References.FirstOrDefault(r =>
              string.Equals(r.Module, mod.Name, StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r.Procedure, currentProc.Name, StringComparison.OrdinalIgnoreCase));
            
            if (existingRef != null)
            {
              // Aggiungi il line number se non presente
              if (!existingRef.LineNumbers.Contains(originalLineNumber))
                existingRef.LineNumbers.Add(originalLineNumber);
            }
            else
            {
              // Crea nuova Reference con line number
              control.References.Add(new VbReference
              {
                Module = mod.Name,
                Procedure = currentProc.Name,
                LineNumbers = new List<int> { originalLineNumber }
              });
            }
            break; // Un controllo può avere un solo event handler con questo nome
          }
        }
        
        continue;
      }

      // -------------------------
      // SUB
      // -------------------------
      var ms = ReSub.Match(noComment);
      if (ms.Success)
      {
        var isStatic = Regex.IsMatch(noComment, @"\bStatic\b", RegexOptions.IgnoreCase);
        
        currentProc = new VbProcedure
        {
          Visibility = string.IsNullOrEmpty(ms.Groups[1].Value) ? "Public" : ms.Groups[1].Value,
          Name = ms.Groups[3].Value,
          ConventionalName = string.Empty,
          Kind = "Sub",
          IsStatic = isStatic,
          Scope = "Module",
          Parameters = ParseParameters(ms.Groups[4].Value, originalLineNumber),
          LineNumber = originalLineNumber
        };
        mod.Procedures.Add(currentProc);
        
        // Traccia event handler per controlli (es. Command1_Click)
        foreach (var control in mod.Controls)
        {
          var prefix = $"{control.Name}_";
          if (currentProc.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
          {
            // Questa procedura è un event handler di questo controllo
            var existingRef = control.References.FirstOrDefault(r =>
              string.Equals(r.Module, mod.Name, StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r.Procedure, currentProc.Name, StringComparison.OrdinalIgnoreCase));
            
            if (existingRef != null)
            {
              // Aggiungi il line number se non presente
              if (!existingRef.LineNumbers.Contains(originalLineNumber))
                existingRef.LineNumbers.Add(originalLineNumber);
            }
            else
            {
              // Crea nuova Reference con line number
              control.References.Add(new VbReference
              {
                Module = mod.Name,
                Procedure = currentProc.Name,
                LineNumbers = new List<int> { originalLineNumber }
              });
            }
            break; // Un controllo può avere un solo event handler con questo nome
          }
        }
        
        continue;
      }

      // -------------------------
      // PROPERTY
      // -------------------------
      var mp = ReProperty.Match(noComment);
      if (mp.Success)
      {
        var isStatic = Regex.IsMatch(noComment, @"\bStatic\b", RegexOptions.IgnoreCase);
        
        currentProc = new VbProcedure
        {
          Visibility = string.IsNullOrEmpty(mp.Groups[1].Value) ? "Public" : mp.Groups[1].Value,
          Kind = $"Property{mp.Groups[3].Value}",
          Name = mp.Groups[4].Value,
          ConventionalName = string.Empty,
          IsStatic = isStatic,
          Scope = "Module",
          Parameters = ParseParameters(mp.Groups[5].Value, originalLineNumber),
          ReturnType = mp.Groups[7].Value,
          LineNumber = originalLineNumber
        };
        mod.Procedures.Add(currentProc);
        
        // Traccia event handler per controlli (es. Command1_Click)
        foreach (var control in mod.Controls)
        {
          var prefix = $"{control.Name}_";
          if (currentProc.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
          {
            // Questa procedura è un event handler di questo controllo
            var existingRef = control.References.FirstOrDefault(r =>
              string.Equals(r.Module, mod.Name, StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r.Procedure, currentProc.Name, StringComparison.OrdinalIgnoreCase));
            
            if (existingRef != null)
            {
              // Aggiungi il line number se non presente
              if (!existingRef.LineNumbers.Contains(originalLineNumber))
                existingRef.LineNumbers.Add(originalLineNumber);
            }
            else
            {
              // Crea nuova Reference con line number
              control.References.Add(new VbReference
              {
                Module = mod.Name,
                Procedure = currentProc.Name,
                LineNumbers = new List<int> { originalLineNumber }
              });
            }
            break; // Un controllo può avere un solo event handler con questo nome
          }
        }
        
        continue;
      }

      // -------------------------
      // EVENT (solo classi)
      // -------------------------
      if (kind.Equals("cls", StringComparison.OrdinalIgnoreCase))
      {
        var mev = ReEvent.Match(noComment);
        if (mev.Success)
        {
          mod.Events.Add(new VbEvent
          {
            Name = mev.Groups[2].Value,
            Visibility = string.IsNullOrEmpty(mev.Groups[1].Value) ? "Public" : mev.Groups[1].Value,
            Scope = "Module",
            Parameters = ParseParameters(mev.Groups[3].Value, originalLineNumber),
            LineNumber = originalLineNumber
          });
          continue;
        }
      }

      // -------------------------
      // DECLARE FUNCTION (External APIs)
      // -------------------------
      var mdf = ReDeclareFunction.Match(noComment);
      if (mdf.Success)
      {
        mod.Procedures.Add(new VbProcedure
        {
          Visibility = "Public",
          Name = mdf.Groups[1].Value,
          ConventionalName = string.Empty,
          Kind = "ExternalFunction",
          IsStatic = false,
          Scope = "Module",
          Parameters = ParseParameters(mdf.Groups[3].Value, originalLineNumber),
          ReturnType = mdf.Groups[5].Value,
          LineNumber = originalLineNumber
        });
        continue;
      }

      // -------------------------
      // DECLARE SUB (External APIs)
      // -------------------------
      var mds = ReDeclareSub.Match(noComment);
      if (mds.Success)
      {
        mod.Procedures.Add(new VbProcedure
        {
          Visibility = "Public",
          Name = mds.Groups[1].Value,
          ConventionalName = string.Empty,
          Kind = "ExternalFunction",
          IsStatic = false,
          Scope = "Module",
          Parameters = ParseParameters(mds.Groups[3].Value, originalLineNumber),
          LineNumber = originalLineNumber
        });
        continue;
      }

      // -------------------------
      // CONST (Module level)
      // -------------------------
      var mconst = ReConstant.Match(noComment);
      if (mconst.Success && currentProc == null && currentType == null)
      {
        var scopeValue = mconst.Groups[1].Value;
        var constName = mconst.Groups[2].Value;
        var constType = mconst.Groups[3].Value;
        var constValue = mconst.Groups[4].Value.Trim();
        
        var isStatic = Regex.IsMatch(noComment, @"\bStatic\b", RegexOptions.IgnoreCase);
        var normalizedScope = scopeValue.Equals("Private", StringComparison.OrdinalIgnoreCase) ? "Module" :
                              scopeValue.Equals("Public", StringComparison.OrdinalIgnoreCase) ? "Project" :
                              string.IsNullOrEmpty(scopeValue) ? "Module" : scopeValue;

        mod.Constants.Add(new VbConstant
        {
          Name = constName,
          Type = constType,
          Value = constValue.Replace("\"", ""),  // Rimuove solo le virgolette, mantiene espressioni
          Scope = normalizedScope,
          Visibility = string.IsNullOrEmpty(scopeValue) ? "Private" : scopeValue,
          Level = "Global",
          LineNumber = originalLineNumber
        });
        continue;
      }

      // -------------------------
      // VARIABILI GLOBALI
      // -------------------------
      var mg = ReGlobalVar.Match(noComment);
      if (mg.Success && currentProc == null && currentType == null)
      {
        var scopeValue = mg.Groups[1].Value;
        var isStatic = Regex.IsMatch(noComment, @"\bStatic\b", RegexOptions.IgnoreCase);
        var normalizedScope = (scopeValue.Equals("Private", StringComparison.OrdinalIgnoreCase) ||
                               scopeValue.Equals("Dim", StringComparison.OrdinalIgnoreCase)) ? "Module" :
                              scopeValue.Equals("Public", StringComparison.OrdinalIgnoreCase) ? "Project" :
                              scopeValue;

        // mg.Groups: 1=Scope, 2=WithEvents, 3=Name, 4=Array, 5=Type
        mod.GlobalVariables.Add(new VbVariable
        {
          Name = mg.Groups[3].Value,
          ConventionalName = string.Empty,
          Type = mg.Groups[5].Value,
          IsStatic = isStatic,
          IsArray = !string.IsNullOrEmpty(mg.Groups[4].Value),
          IsWithEvents = !string.IsNullOrEmpty(mg.Groups[2].Value),
          Scope = normalizedScope,
          Visibility = scopeValue.Equals("Dim", StringComparison.OrdinalIgnoreCase) ? "Private" : scopeValue,
          Level = "Global",
          LineNumber = originalLineNumber
        });
        continue;
      }

      // -------------------------
      // VARIABILI MEMBRO (solo classi)
      // -------------------------
      if (kind.Equals("cls", StringComparison.OrdinalIgnoreCase))
      {
        var mm = ReMemberVar.Match(noComment);
        if (mm.Success && currentProc == null && currentType == null)
        {
          var scopeValue = mm.Groups[1].Value;
          var isStatic = Regex.IsMatch(noComment, @"\bStatic\b", RegexOptions.IgnoreCase);
          var normalizedScope = (scopeValue.Equals("Private", StringComparison.OrdinalIgnoreCase) ||
                                 scopeValue.Equals("Dim", StringComparison.OrdinalIgnoreCase)) ? "Module" :
                                scopeValue.Equals("Public", StringComparison.OrdinalIgnoreCase) ? "Project" :
                                scopeValue;

          // mm.Groups: 1=Scope, 2=Name, 3=Array, 4=Type
          mod.GlobalVariables.Add(new VbVariable
          {
            Name = mm.Groups[2].Value,
            ConventionalName = string.Empty,
            Type = mm.Groups[4].Value,
            IsStatic = isStatic,
            IsArray = !string.IsNullOrEmpty(mm.Groups[3].Value),
            Scope = normalizedScope,
            Visibility = scopeValue.Equals("Dim", StringComparison.OrdinalIgnoreCase) ? "Private" : scopeValue,
            Level = "Member",
            LineNumber = originalLineNumber
          });
          continue;
        }
      }

      // -------------------------
      // END PROCEDURE
      // -------------------------
      if (currentProc != null)
      {
        var trimmedNoComment = noComment.TrimStart();
        // Only treat as end of procedure when matching the exact terminators
        // based on possible proc kinds: Sub, Function, Property
        if (trimmedNoComment.StartsWith("End Sub", StringComparison.OrdinalIgnoreCase)
            || trimmedNoComment.StartsWith("End Function", StringComparison.OrdinalIgnoreCase)
            || trimmedNoComment.StartsWith("End Property", StringComparison.OrdinalIgnoreCase))
        {
          currentProc = null;
          continue;
        }
      }

      // -------------------------
      // DENTRO PROCEDURA
      // -------------------------
      if (currentProc != null)
      {
        // costanti locali
        var mc = ReConstant.Match(noComment);
        if (mc.Success && currentProc != null && currentType == null)
        {
          var scopeValue = mc.Groups[1].Value;
          var constName = mc.Groups[2].Value;
          var constType = mc.Groups[3].Value;
          var constValue = mc.Groups[4].Value.Trim();
          
          var normalizedScope = scopeValue.Equals("Private", StringComparison.OrdinalIgnoreCase) ? "Module" :
                                scopeValue.Equals("Public", StringComparison.OrdinalIgnoreCase) ? "Project" :
                                string.IsNullOrEmpty(scopeValue) ? "Module" : scopeValue;

          currentProc.Constants.Add(new VbConstant
          {
            Name = constName,
            Type = constType,
            Value = constValue.Replace("\"", ""),
            Scope = "Procedure",
            Visibility = string.IsNullOrEmpty(scopeValue) ? "Private" : scopeValue,
            Level = "Local",
            LineNumber = originalLineNumber
          });
          continue;
        }

        // variabili locali
        var ml = ReLocalVar.Match(noComment);
        if (ml.Success)
        {
          var declarationType = ml.Groups[1].Value; // Dim or Static
          var isStatic = declarationType.Equals("Static", StringComparison.OrdinalIgnoreCase);
          
          // ml.Groups: 1=Scope(Dim/Static), 2=Name, 3=Array, 4=Type
          currentProc.LocalVariables.Add(new VbVariable
          {
            Name = ml.Groups[2].Value,
            ConventionalName = string.Empty,
            Type = ml.Groups[4].Value,
            IsStatic = isStatic,
            IsArray = !string.IsNullOrEmpty(ml.Groups[3].Value),
            Scope = "Procedure",
            Visibility = "Private",
            Level = "Local",
            LineNumber = originalLineNumber
          });
        }

        // chiamate con parentesi
        foreach (Match callMatch in ReCall.Matches(noComment))
        {
          var objName = callMatch.Groups[1].Success ? callMatch.Groups[1].Value : null;
          var methodName = callMatch.Groups[2].Value;

          if (VbKeywords.Contains(methodName))
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: VB Keyword", "ReCall");
            continue;
          }

          // Esclude parametri e variabili locali
          if (currentProc.Parameters.Any(p => p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase)))
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Parameter", "ReCall");
            continue;
          }

          if (currentProc.LocalVariables.Any(v => v.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase)))
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Local Variable", "ReCall");
            continue;
          }

          // Esclude autoreferenza (ricorsione)
          if (methodName.Equals(currentProc.Name, StringComparison.OrdinalIgnoreCase))
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Self-reference (recursion)", "ReCall");
            continue;
          }

          var newCall = new VbCall();
          if (currentWith != null && noComment.TrimStart().StartsWith("."))
          {
            newCall = new VbCall
            {
              Raw = $"{currentWith}.{methodName}",
              ObjectName = currentWith,
              MethodName = methodName,
              LineNumber = originalLineNumber
            };
          }
          else
          {
            newCall = new VbCall
            {
              Raw = objName != null ? $"{objName}.{methodName}" : methodName,
              ObjectName = objName,
              MethodName = methodName,
              LineNumber = originalLineNumber
            };
          }

          // Evita duplicati
          if (!currentProc.Calls.Any(c => c.Raw.Equals(newCall.Raw, StringComparison.OrdinalIgnoreCase) && c.LineNumber == originalLineNumber))
          {
            DebugLog(currentProc.Name, originalLineNumber, newCall.Raw, "ADDED", "ReCall");
            currentProc.Calls.Add(newCall);
          }
          else
          {
            DebugLog(currentProc.Name, originalLineNumber, newCall.Raw, "SKIPPED: Duplicate", "ReCall");
          }
        }

        // chiamate nude: Exec_Tick
        var mbare = ReBareCall.Match(noComment);
        if (mbare.Success)
        {
          var name = mbare.Groups[1].Value;

          DebugLog(currentProc.Name, originalLineNumber, name, "Bare call matched", "ReBareCall");

          var isKeyword = VbKeywords.Contains(name);
          var isParam = currentProc.Parameters.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
          var isLocalVar = currentProc.LocalVariables.Any(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
          var isGlobalVar = mod.GlobalVariables.Any(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
          var isType = mod.Types.Any(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
          var isSelfRef = name.Equals(currentProc.Name, StringComparison.OrdinalIgnoreCase);

          if (isKeyword)
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: VB Keyword", "ReBareCall");
          else if (isParam)
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: Parameter", "ReBareCall");
          else if (isLocalVar)
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: Local Variable", "ReBareCall");
          else if (isGlobalVar)
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: Global Variable", "ReBareCall");
          else if (isType)
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: Type", "ReBareCall");
          else if (isSelfRef)
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: Self-reference (return value assignment)", "ReBareCall");
          else if (!currentProc.Calls.Any(c => c.MethodName.Equals(name, StringComparison.OrdinalIgnoreCase) && c.LineNumber == originalLineNumber))
          {
            DebugLog(currentProc.Name, originalLineNumber, name, "ADDED", "ReBareCall");
            currentProc.Calls.Add(new VbCall
            {
              Raw = name,
              MethodName = name,
              LineNumber = originalLineNumber
            });
          }
          else
          {
            DebugLog(currentProc.Name, originalLineNumber, name, "SKIPPED: Duplicate", "ReBareCall");
          }
        }

        // chiamate senza parentesi
        foreach (Match callMatch in ReCallWithoutParens.Matches(noComment))
        {
          var methodName = callMatch.Groups[1].Value;

          if (VbKeywords.Contains(methodName))
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: VB Keyword", "ReCallWithoutParens");
            continue;
          }

          var isParam = currentProc.Parameters.Any(p => p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));
          var isLocalVar = currentProc.LocalVariables.Any(v => v.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));
          var isGlobalVar = mod.GlobalVariables.Any(v => v.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));
          var isSelfRef = methodName.Equals(currentProc.Name, StringComparison.OrdinalIgnoreCase);

          // Esclude parametri, variabili locali, globali e autoreferenza
          if (isParam)
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Parameter", "ReCallWithoutParens");
            continue;
          }
          if (isLocalVar)
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Local Variable", "ReCallWithoutParens");
            continue;
          }
          if (isGlobalVar)
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Global Variable", "ReCallWithoutParens");
            continue;
          }
          if (isSelfRef)
          {
            DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Self-reference (recursion)", "ReCallWithoutParens");
            continue;
          }

          var newCall = new VbCall
          {
            Raw = methodName,
            MethodName = methodName,
            LineNumber = originalLineNumber
          };

          // Evita duplicati
          if (!currentProc.Calls.Any(c => c.Raw.Equals(newCall.Raw, StringComparison.OrdinalIgnoreCase) && c.LineNumber == originalLineNumber))
          {
            DebugLog(currentProc.Name, originalLineNumber, newCall.Raw, "ADDED", "ReCallWithoutParens");
            currentProc.Calls.Add(newCall);
          }
          else
          {
            DebugLog(currentProc.Name, originalLineNumber, newCall.Raw, "SKIPPED: Duplicate", "ReCallWithoutParens");
          }
        }

        continue;
      }
    }

    return mod;
  }

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
        @"^(Optional\s+)?(ByVal|ByRef)?\s*(\w+)\s*(As\s+([\w\.\(\)]+))?",
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
        Type = m.Groups[5].Value,
        Used = false,
        LineNumber = originalLineNumber
      });
    }

    return result;
  }
}
