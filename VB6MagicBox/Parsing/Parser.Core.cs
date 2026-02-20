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
        new(@"^(Public|Private|Friend)?\s*(Static)?\s*Function\s+(\w+)\s*\((.*)\)\s*(As\s+([\w\.\(\)]+))?",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReSub =
        new(@"^(Public|Private|Friend)?\s*(Static)?\s*Sub\s+(\w+)\s*\((.*)\)",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReDeclareFunction =
        new(@"^(Public|Private|Friend)?\s*Declare\s+Function\s+(\w+)\s+Lib\s+""([^""]+)""(?:\s+Alias\s+""[^""]+"")?\s*\((.*)\)\s*(As\s+([\w\.\(\)]+))?",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReDeclareSub =
        new(@"^(Public|Private|Friend)?\s*Declare\s+Sub\s+(\w+)\s+Lib\s+""([^""]+)""(?:\s+Alias\s+""[^""]+"")?\s*\((.*)\)",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReConstant =
        new(@"^(Public|Private|Friend|Global)?\s*Const\s+(\w+)\s*(?:As\s+([\w\.\(\)]+))?\s*=\s*(.+)",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReProperty =
        new(@"^(Public|Private|Friend)?\s*(Static)?\s*Property\s+(Get|Let|Set)\s+(\w+)\s*\((.*)\)\s*(As\s+([\w\.\(\)]+))?",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReEvent =
        new(@"^(Public|Private|Friend)?\s*Event\s+(\w+)\s*\((.*?)\)",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReGlobalVar =
        new(@"^(Public|Private|Global|Friend|Dim)\s+(WithEvents\s+)?(\w+)(\([^)]*\))?\s+As\s+(?:New\s+)?([\w\.\(\)]+)",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReLocalVar =
        new(@"^\s*(Dim|Static)\s+(\w+)(\([^)]*\))?\s+As\s+(?:New\s+)?([\w\.\(\)]+)",
            RegexOptions.IgnoreCase);

    private static readonly Regex ReMemberVar =
        new(@"^(Public|Private|Friend|Dim)\s+(\w+)(\([^)]*\))?\s+As\s+(?:New\s+)?([\w\.\(\)]+)",
            RegexOptions.IgnoreCase);

    // Fallback: variabile globale/membro senza "As Tipo" — usato da TypeAnnotator
    // Gruppi: 1=keyword, 2=WithEvents?, 3=nome, 4=suffisso tipo ($%&!#@), 5=dimensioni array
    private static readonly Regex ReGlobalVarNoType = new(
        @"^(Public|Private|Global|Friend|Dim)\s+(WithEvents\s+)?(\w+)([$%&!#@]?)(\([^)]*\))?\s*$",
        RegexOptions.IgnoreCase);

    // Fallback: variabile locale senza "As Tipo" — usato da TypeAnnotator
    // Gruppi: 1=keyword (Dim/Static), 2=nome, 3=suffisso tipo, 4=dimensioni array
    private static readonly Regex ReLocalVarNoType = new(
        @"^\s*(Dim|Static)\s+(\w+)([$%&!#@]?)(\([^)]*\))?\s*$",
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
        new(@"([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)", RegexOptions.IgnoreCase);

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
}
