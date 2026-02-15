using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  public static class NamingConvention
  {
    // Mappa dei prefissi standard per controlli VB6 comuni
    private static readonly Dictionary<string, string> ControlPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        // Controlli standard VB6
        { "TextBox", "txt" },
        { "CommandButton", "cmd" },
        { "Label", "lbl" },
        { "Frame", "fra" },
        { "CheckBox", "chk" },
        { "OptionButton", "opt" },
        { "ListBox", "lst" },
        { "ComboBox", "cbo" },
        { "Timer", "tmr" },
        { "PictureBox", "pic" },
        { "Image", "img" },
        { "Shape", "shp" },
        { "Line", "lin" },
        { "HScrollBar", "hsb" },
        { "VScrollBar", "vsb" },
        { "DirListBox", "dir" },
        { "DriveListBox", "drv" },
        { "FileListBox", "fil" },
        { "Data", "dat" },
        { "OLE", "ole" },
        { "CommonDialog", "dlg" },
        { "Menu", "mnu" },
        
        // Controlli ActiveX comuni
        { "MSFlexGrid", "flx" },
        { "MSHFlexGrid", "flx" },
        { "DataGrid", "grd" },
        { "TreeView", "tvw" },
        { "ListView", "lvw" },
        { "ProgressBar", "prg" },
        { "Slider", "sld" },
        { "TabStrip", "tab" },
        { "ToolBar", "tlb" },
        { "StatusBar", "stb" },
        { "ImageList", "iml" },
        { "RichTextBox", "rtf" },
        { "MonthView", "mvw" },
        { "DateTimePicker", "dtp" },
        { "UpDown", "upd" },
        { "Animation", "ani" },
        { "MSComm", "msc" },
        { "Winsock", "wsk" },
        { "WebBrowser", "web" },
        { "CoolBar", "clb" },
        { "FlatScrollBar", "fsb" }
    };

    private static string GetControlPrefix(string controlType)
    {
      if (string.IsNullOrEmpty(controlType)) return "ctl";

      // Pulisci prima i prefissi namespace (VB., MSComCtl2., etc.)
      var cleanType = controlType;
      if (cleanType.StartsWith("VB.", StringComparison.OrdinalIgnoreCase))
        cleanType = cleanType.Substring(3);
      else if (cleanType.StartsWith("MSComCtl2.", StringComparison.OrdinalIgnoreCase))
        cleanType = cleanType.Substring(10);
      else if (cleanType.StartsWith("MSComctlLib.", StringComparison.OrdinalIgnoreCase))
        cleanType = cleanType.Substring(12);
      else if (cleanType.StartsWith("Threed.SS", StringComparison.OrdinalIgnoreCase))
        cleanType = cleanType.Substring(9);

      // Cerca nella mappa dei prefissi standard
      if (ControlPrefixes.TryGetValue(cleanType, out var prefix))
        return prefix;

      // Regola generica per controlli custom: iniziale + rimozione vocali (tranne iniziale)
      if (cleanType.Length == 0) return "ctl";
      
      var result = new System.Text.StringBuilder();
      
      // Aggiungi sempre la prima lettera (anche se è vocale)
      result.Append(char.ToLowerInvariant(cleanType[0]));
      
      // Per le lettere successive, rimuovi le vocali
      for (int i = 1; i < cleanType.Length && result.Length < 3; i++)
      {
        var ch = cleanType[i];
        if (char.IsLetter(ch) && !"aeiouAEIOU".Contains(ch))
        {
          result.Append(char.ToLowerInvariant(ch));
        }
      }
      
      // Se non ci sono abbastanza consonanti, aggiungi lettere rimanenti
      if (result.Length < 3)
      {
        for (int i = 1; i < cleanType.Length && result.Length < 3; i++)
        {
          var ch = cleanType[i];
          if (char.IsLetter(ch))
          {
            var lowerCh = char.ToLowerInvariant(ch);
            if (!result.ToString().Contains(lowerCh)) // Evita duplicati
            {
              result.Append(lowerCh);
            }
          }
        }
      }
      
      // Se ancora troppo corto, usa le prime 3 lettere
      if (result.Length < 3 && cleanType.Length >= 3)
      {
        return cleanType.Substring(0, 3).ToLowerInvariant();
      }
      
      return result.Length > 0 ? result.ToString() : cleanType.ToLowerInvariant();
    }

    /// <summary>
    /// Converte PascalCase intelligente in SCREAMING_SNAKE_CASE preservando acronimi
    /// Esempi:
    /// - ItemUAObjListener ? ITEM_UA_OBJ_LISTENER
    /// - MaxUnsignedLongAnd1 ? MAX_UNSIGNED_LONG_AND1
    /// - ItemUAObjWriterSim_Real ? ITEM_UA_OBJ_WRITER_SIM_REAL
    /// - Alg_FirstStep ? ALG_FIRST_STEP
    /// - RP_CentralInitial ? RP_CENTRAL_INITIAL
    /// </summary>
    private static string ToScreamingSnakeCase(string name)
    {
      if (string.IsNullOrEmpty(name)) return name;

      // Split per underscore per gestire nomi con _ già presenti
      var parts = name.Split('_');
      var processedParts = new List<string>();

      foreach (var part in parts)
      {
        if (string.IsNullOrEmpty(part))
        {
          // Preserva underscore vuoti
          processedParts.Add(part);
          continue;
        }

        // Applica la logica di word break a ogni parte
        var result = new System.Text.StringBuilder();

        for (int i = 0; i < part.Length; i++)
        {
          var current = part[i];

          // Aggiungi underscore se:
          // 1. Non è l'inizio
          // 2. Il carattere corrente è maiuscolo
          // 3. E il precedente è minuscolo (parola nuova: itemUA)
          // 4. OPPURE il prossimo è minuscolo (fine acronimo: UAObj -> UA_Obj)
          if (i > 0 && char.IsUpper(current))
          {
            var prev = part[i - 1];
            var isNextLower = i + 1 < part.Length && char.IsLower(part[i + 1]);

            // Caso 1: transition lowercase -> uppercase (itemUA)
            if (char.IsLower(prev))
            {
              result.Append('_');
            }
            // Caso 2: fine di acronimo (UAObj -> UA_Obj)
            else if (char.IsUpper(prev) && isNextLower)
            {
              result.Append('_');
            }
          }

          result.Append(char.ToUpperInvariant(current));
          result.Replace("P_DX_I", "PDXI"); // Special case per preservare PDxI come acronimo
        }

        processedParts.Add(result.ToString());
      }

      // Rejoin le parti con underscore e tutto uppercase
      return string.Join("_", processedParts);
    }


    private static string ApplyControlNaming(string controlName, string controlType)
    {
      if (string.IsNullOrEmpty(controlName)) return controlName;

      var expectedPrefix = GetControlPrefix(controlType);

      // Caso speciale: TextBox con prefissi non standard (tb, tx) ? normalizza a txt
      var cleanType = controlType;
      if (cleanType.StartsWith("VB.", StringComparison.OrdinalIgnoreCase))
        cleanType = cleanType.Substring(3);

      if (cleanType.Equals("TextBox", StringComparison.OrdinalIgnoreCase))
      {
        // Controlla se ha prefisso "tb" o "tx" (ma non "txt")
        if (controlName.Length > 2 &&
            (controlName.StartsWith("tb", StringComparison.OrdinalIgnoreCase) ||
             controlName.StartsWith("tx", StringComparison.OrdinalIgnoreCase)) &&
            !controlName.StartsWith("txt", StringComparison.OrdinalIgnoreCase))
        {
          var prefixLen = 2;
          if (controlName.Length > prefixLen && char.IsUpper(controlName[prefixLen]))
          {
            // tbUsername o txEmail ? txtUsername, txtEmail
            var baseName = controlName.Substring(prefixLen);
            return "txt" + baseName;
          }
        }
      }

      // Verifica se il nome inizia già con il prefisso corretto (case insensitive)
      // IMPORTANTE: deve esserci una lettera MAIUSCOLA dopo il prefisso (camelCase)
      if (controlName.Length > expectedPrefix.Length &&
          controlName.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase) &&
          char.IsUpper(controlName[expectedPrefix.Length]))
      {
        // Ha già il prefisso corretto, applica solo il formato camelCase
        var baseName = controlName.Substring(expectedPrefix.Length);
        return expectedPrefix + ToPascalCase(baseName);
      }

      // Controlla se ha un altro prefisso a 3 lettere (es: txt, cmd, lbl)
      if (controlName.Length > 3 &&
          char.IsLower(controlName[0]) &&
          char.IsLower(controlName[1]) &&
          char.IsLower(controlName[2]) &&
          char.IsUpper(controlName[3]))
      {
        // Ha un prefisso diverso, sostituiscilo
        var baseName = controlName.Substring(3);
        return expectedPrefix + baseName;
      }

      // Non ha prefisso, aggiungilo
      return expectedPrefix + ToPascalCase(controlName);
    }

    private static bool IsHungarianM(string name)
    {
      if (string.IsNullOrEmpty(name)) return false;
      return Regex.IsMatch(name, @"^m[a-z]{3}[A-Z]");
    }

    private static string NormalizeHungarianPrefix(string name, string enforcedPrefix = "")
    {
      if (string.IsNullOrEmpty(name)) return name;

      var match = Regex.Match(name, @"^m([a-z]{3})([A-Z].*)");
      if (match.Success)
      {
        var baseName = match.Groups[2].Value; // PollingDisableRequest relative to mudtPollingDisableRequest
                                              // Use enforcedPrefix if provided, otherwise assume PascalCase of baseName.
                                              // BUT if enforcedPrefix is empty, and we stripped 'm'+type, we just return baseName (which is PascalCase).
        return enforcedPrefix + baseName;
      }

      // Fallback if no hungarian prefix found
      if (string.IsNullOrEmpty(enforcedPrefix))
        return ToPascalCase(name);

      return enforcedPrefix + ToPascalCase(name);
    }

    private static readonly HashSet<string> CSharpKeywords = new(StringComparer.Ordinal)
    {
        "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char", "checked", "class", "const", "continue",
        "decimal", "default", "delegate", "do", "double", "else", "enum", "event", "explicit", "extern", "false", "finally",
        "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in", "int", "interface", "internal", "is", "lock",
        "long", "namespace", "new", "null", "object", "operator", "out", "override", "params", "private", "protected",
        "public", "readonly", "ref", "return", "sbyte", "sealed", "short", "sizeof", "stackalloc", "static", "string",
        "struct", "switch", "this", "throw", "true", "try", "typeof", "uint", "ulong", "unchecked", "unsafe", "ushort",
        "using", "virtual", "void", "volatile", "while", "With" // Added With as distinct check or just handle via list
    };

    private static string ToPascalCaseType(string name)
    {
      if (string.IsNullOrEmpty(name)) return name;

      // Rimuovi il suffisso _T o T finale se presente
      var baseName = name;
      if (baseName.EndsWith("_T", StringComparison.OrdinalIgnoreCase))
      {
        baseName = baseName.Substring(0, baseName.Length - 2);
      }
      else if (baseName.EndsWith("T", StringComparison.OrdinalIgnoreCase) &&
               baseName.Length > 1 &&
               baseName[baseName.Length - 2] == '_')
      {
        baseName = baseName.Substring(0, baseName.Length - 1);
      }

      var parts = baseName.Split('_', StringSplitOptions.RemoveEmptyEntries);
      var result = new System.Text.StringBuilder();

      foreach (var part in parts)
      {
        // Normalize case for each part
        if (part.All(char.IsLetter) && part.ToUpperInvariant() == part)
        {
          // All uppercase acronym -> convert to PascalCase
          var normalized = part.ToLowerInvariant();
          if (normalized.Length > 0)
          {
            result.Append(char.ToUpper(normalized[0]));
            if (normalized.Length > 1)
              result.Append(normalized.Substring(1));
          }
        }
        else
        {
          // Mixed case - preserve and capitalize first letter
          if (part.Length > 0)
          {
            result.Append(char.ToUpper(part[0]));
            if (part.Length > 1)
              result.Append(part.Substring(1));
          }
        }
      }

      // Riaggiunge il suffisso _T
      return result.ToString() + "_T";
    }

    private static bool IsClassType(VbProject project, string typeName)
    {
      if (string.IsNullOrWhiteSpace(typeName)) return false;

      // Normalize type name segments (handles namespaces like PDxI.clsPDxI)
      var segments = typeName.Split('.', StringSplitOptions.RemoveEmptyEntries)
                             .Select(s => s.Trim())
                             .ToList();

      foreach (var mod in project.Modules.Where(m => m.IsClass))
      {
        var moduleName = mod.Name; // already without extension
        var moduleNameNoCls = moduleName.StartsWith("cls", StringComparison.OrdinalIgnoreCase)
            ? moduleName.Substring(3)
            : moduleName;

        if (segments.Any(s => s.Equals(moduleName, StringComparison.OrdinalIgnoreCase) ||
                              s.Equals(moduleNameNoCls, StringComparison.OrdinalIgnoreCase)))
          return true;
      }

      // Fallback: if type name includes a cls prefix anywhere (e.g., clsIonGun), treat as class
      return Regex.IsMatch(typeName, @"\bcls\w+", RegexOptions.IgnoreCase);
    }

    public static void Apply(VbProject project)
    {
      foreach (var mod in project.Modules)
      {
        // Trackers per conflict resolution - scope separati
        var globalNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var procedureNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Module/Class naming
        var rawName = Path.GetFileNameWithoutExtension(mod.Name); // Rimuove .cls/.bas/.frm extension, mantiene nome file

        // Per i Form, aggiungi prefisso "frm" come per i controlli
        if (mod.Kind.Equals("frm", StringComparison.OrdinalIgnoreCase))
        {
          if (!rawName.StartsWith("frm", StringComparison.OrdinalIgnoreCase))
            rawName = "frm" + rawName;
        }

        var pascalName = ToPascalCase(rawName);

        // Se il nome risultante è una keyword C# (es. "With"), aggiungi "Class"
        if (CSharpKeywords.Contains(pascalName.ToLower())) // Check lowercase against lowercase keywords usually
        {
          pascalName += "Class";
        }

        mod.ConventionalName = pascalName;

        // Global Variables - con conflict resolution
        foreach (var v in mod.GlobalVariables)
        {
          var baseName = GetBaseNameFromHungarian(v.Name);

          // Controlla se il nome termina con _numero (array di oggetti es: objFM489_0)
          string arraySuffix = null;
          var arrayMatch = Regex.Match(baseName, @"^(.+)(_\d+)$");
          if (arrayMatch.Success)
          {
            baseName = arrayMatch.Groups[1].Value;  // Nome senza il suffisso
            arraySuffix = arrayMatch.Groups[2].Value;  // Es: "_0", "_1"
          }

          string conventionalName;
          if (v.IsStatic)
          {
            // Per le variabili static: controlla se inizia già con "s_"
            if (baseName.StartsWith("s_", StringComparison.OrdinalIgnoreCase))
            {
              // Già nel formato corretto s_Nome, mantieni così com'è
              conventionalName = baseName + (arraySuffix ?? "");
            }
            else
            {
              // Per le variabili static: se il nome inizia con 's' minuscolo, rimuovilo prima di applicare ToPascalCase
              // Es: sQualcosa -> s_Qualcosa (NON s_SQualcosa)
              var nameForPascal = baseName;
              if (baseName.Length > 1 && baseName.StartsWith("s", StringComparison.OrdinalIgnoreCase) && char.IsUpper(baseName[1]))
              {
                // Ha prefisso 's' seguito da maiuscola: sQualcosa -> Qualcosa
                nameForPascal = baseName.Substring(1);
              }
              conventionalName = "s_" + ToPascalCase(nameForPascal) + (arraySuffix ?? "");
            }
          }
          else if (IsPrivate(v.Visibility))
          {
            // Module Private: controlla se inizia già con "m_"
            if (baseName.StartsWith("m_", StringComparison.OrdinalIgnoreCase))
            {
              // Già nel formato corretto m_Nome, mantieni così com'è
              conventionalName = baseName + (arraySuffix ?? "");
            }
            else
            {
              // Module Private use m_
              conventionalName = NormalizeHungarianPrefix(baseName, "m_") + (arraySuffix ?? "");
            }
          }
          else
          {
            // Public globals: comportamento diverso per Form/Class vs Module
            bool isFormOrClass = mod.IsClass; // Form (.frm) o Class (.cls)

            // Se il nome inizia con "gobj", sempre g_ naming (forzato globale)
            if (v.Name.StartsWith("gobj", StringComparison.OrdinalIgnoreCase))
            {
              var raw = "g_" + v.Name.Substring(4);
              var tail = raw.Substring(2);

              // Re-applica il check per array suffix sul tail
              var tailArrayMatch = Regex.Match(tail, @"^(.+)(_\d+)$");
              if (tailArrayMatch.Success)
              {
                tail = tailArrayMatch.Groups[1].Value;
                arraySuffix = tailArrayMatch.Groups[2].Value;
              }

              conventionalName = "g_" + ToPascalCase(tail) + (arraySuffix ?? "");
            }
            else if (IsClassType(project, v.Type))
            {
              // Oggetto custom (non tipo nativo)
              if (isFormOrClass)
              {
                // FORM/CLASS: oggetti pubblici ? objName (camelCase)
                // Es: UAServerObj ? objUAServerObj
                // Es: objFM489 ? objFM489 (già OK)
                var tail = baseName;

                // Rimuovi prefisso obj se presente (per evitare objObjName)
                if (tail.StartsWith("obj", StringComparison.OrdinalIgnoreCase) && tail.Length > 3)
                {
                  tail = tail.Substring(3);
                }

                // Applica PascalCase al tail e poi prefisso obj
                conventionalName = "obj" + ToPascalCase(tail) + (arraySuffix ?? "");
              }
              else
              {
                // MODULE: oggetti pubblici ? g_Name (PascalCase)
                var raw = baseName;
                if (!raw.StartsWith("g_", StringComparison.OrdinalIgnoreCase))
                  raw = "g_" + raw;

                var tail = raw.Substring(2);
                conventionalName = "g_" + ToPascalCase(tail) + (arraySuffix ?? "");
              }
            }
            else
            {
              // Non-class public globals: keep PascalCase (strip hungarian prefix)
              conventionalName = ToPascalCase(baseName) + (arraySuffix ?? "");
            }
          }

          // Conflict resolution per variabili globali
          v.ConventionalName = ResolveNameConflict(conventionalName, globalNamesUsed);
          globalNamesUsed.Add(v.ConventionalName);

          if (IsReservedWord(v.ConventionalName))
          {
            // Se è reserved word, torna al nome originale e rimuovi da globalNamesUsed
            globalNamesUsed.Remove(v.ConventionalName);
            v.ConventionalName = v.Name;
            globalNamesUsed.Add(v.ConventionalName);
          }
        }

        // Constants - con conflict resolution
        foreach (var c in mod.Constants)
        {
          // Constants: intelligentemente converte PascalCase ? SCREAMING_SNAKE_CASE
          // ItemUAObjListener ? ITEM_UA_OBJ_LISTENER
          // RIC_AL_PDxI1_LOWVOLTAGE ? RIC_AL_PDXI1_LOWVOLTAGE (preserva underscore)
          var conventionalName = ToScreamingSnakeCase(c.Name);
          c.ConventionalName = ResolveNameConflict(conventionalName, globalNamesUsed);
          globalNamesUsed.Add(c.ConventionalName);

          if (IsReservedWord(c.ConventionalName))
          {
            globalNamesUsed.Remove(c.ConventionalName);
            c.ConventionalName = c.Name;
            globalNamesUsed.Add(c.ConventionalName);
          }
        }

        // Enums - con conflict resolution (usando globalNamesUsed perché sono a livello modulo)
        foreach (var e in mod.Enums)
        {
          // Convert SCREAMING_SNAKE_CASE to PascalCase
          var conventionalName = ToPascalCaseFromScreamingSnake(e.Name);
          e.ConventionalName = ResolveNameConflict(conventionalName, globalNamesUsed);
          globalNamesUsed.Add(e.ConventionalName);

          if (IsReservedWord(e.ConventionalName))
          {
            globalNamesUsed.Remove(e.ConventionalName);
            e.ConventionalName = e.Name;
            globalNamesUsed.Add(e.ConventionalName);
          }

          // Per i valori enum, conflict resolution interno all'enum
          var enumValueNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
          foreach (var val in e.Values)
          {
            var valConventionalName = ToPascalCaseFromScreamingSnake(val.Name);
            val.ConventionalName = ResolveNameConflict(valConventionalName, enumValueNamesUsed);
            enumValueNamesUsed.Add(val.ConventionalName);

            if (IsReservedWord(val.ConventionalName))
            {
              enumValueNamesUsed.Remove(val.ConventionalName);
              val.ConventionalName = val.Name;
              enumValueNamesUsed.Add(val.ConventionalName);
            }
          }
        }

        // Types - con conflict resolution (usando globalNamesUsed perché sono a livello modulo)
        foreach (var t in mod.Types)
        {
          var conventionalName = ToPascalCaseType(t.Name);
          t.ConventionalName = ResolveNameConflict(conventionalName, globalNamesUsed);
          globalNamesUsed.Add(t.ConventionalName);

          if (IsReservedWord(t.ConventionalName))
          {
            globalNamesUsed.Remove(t.ConventionalName);
            t.ConventionalName = t.Name;
            globalNamesUsed.Add(t.ConventionalName);
          }

          // Per i campi del tipo, conflict resolution interno al tipo
          var fieldNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
          foreach (var f in t.Fields)
          {
            var fieldConventionalName = ToPascalCase(f.Name);

            // Caso speciale: se il campo si chiama "Type", è una keyword riservata
            if (fieldConventionalName.Equals("Type", StringComparison.OrdinalIgnoreCase))
            {
              fieldConventionalName = "TypeValue";
            }
            else if (IsReservedWord(fieldConventionalName))
            {
              // Per altre reserved words, aggiungi suffisso "Value"
              fieldConventionalName = fieldConventionalName + "Value";
            }

            f.ConventionalName = ResolveNameConflict(fieldConventionalName, fieldNamesUsed);
            fieldNamesUsed.Add(f.ConventionalName);
          }
        }

        // Events - con conflict resolution (usando globalNamesUsed perché sono a livello modulo)
        foreach (var ev in mod.Events)
        {
          var conventionalName = ToPascalCase(ev.Name);
          ev.ConventionalName = ResolveNameConflict(conventionalName, globalNamesUsed);
          globalNamesUsed.Add(ev.ConventionalName);

          if (IsReservedWord(ev.ConventionalName))
          {
            globalNamesUsed.Remove(ev.ConventionalName);
            ev.ConventionalName = ev.Name;
            globalNamesUsed.Add(ev.ConventionalName);
          }

          // Per i parametri dell'evento, conflict resolution interno all'evento
          var eventParamNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
          foreach (var p in ev.Parameters)
          {
            var paramConventionalName = ToCamelCase(p.Name);
            p.ConventionalName = ResolveNameConflict(paramConventionalName, eventParamNamesUsed);
            eventParamNamesUsed.Add(p.ConventionalName);

            if (IsReservedWord(p.ConventionalName))
            {
              eventParamNamesUsed.Remove(p.ConventionalName);
              p.ConventionalName = p.Name;
              eventParamNamesUsed.Add(p.ConventionalName);
            }
          }
        }

        // Controls - RAGGRUPPAMENTO per nome originale (gestione array di controlli)
        var controlGroups = mod.Controls.GroupBy(c => c.Name, StringComparer.OrdinalIgnoreCase);
        
        foreach (var group in controlGroups)
        {
          // Calcola ConventionalName una volta sola per il gruppo
          var conventionalName = ApplyControlNaming(group.Key, group.First().ControlType);
          
          if (IsReservedWord(conventionalName))
            conventionalName = group.Key; // Torna al nome originale se reserved

          // Applica lo stesso ConventionalName a tutti i controlli del gruppo
          foreach (var control in group)
          {
            control.ConventionalName = conventionalName;
          }
        }

        // Procedures - con conflict resolution
        foreach (var proc in mod.Procedures)
        {
          string conventionalName;

          // Event Handling Detection
          // Pattern: ObjectName_EventName
          // Check for standard module events: Class_Initialize, etc.
          if (proc.Name.Equals("Class_Initialize", StringComparison.OrdinalIgnoreCase) ||
              proc.Name.Equals("Class_Terminate", StringComparison.OrdinalIgnoreCase) ||
              (mod.Kind.Equals("frm", StringComparison.OrdinalIgnoreCase) &&
               (proc.Name.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) ||
                proc.Name.StartsWith("UserControl_", StringComparison.OrdinalIgnoreCase)))) // Assuming Form_Unload, etc.
          {
            // Standard event handlers: keep name (Case Insensitive? Standardize to PascalCase with underscore?)
            // User said "rimangono uguali" (keep same). But maybe enforce correct casing if user typed "form_load"?
            // Let's assume we keep exact source casing OR standarize to "Form_Load".
            // "Class_Initialize" usually standard.
            // Let's just keep original name if it matches standard patterns, or minimal normalization.
            // User said "Class_Initialize or Form_Unload remain same".
            conventionalName = proc.Name;
          }
          else if (proc.Name.Contains("_"))
          {
            // Possible event handler: Object_Event
            // Need to find if prefix is a known object (Control or WithEvents Variable)
            // Split by LAST underscore? VB6 events usually specific.
            // But User example: objATH3204_0_ReadStatusCompleted -> ObjATH32040_ReadStatusCompleted
            // Here underscore is separator.
            // Try to match prefix against Controls or WithEvents Variables.

            bool isEvent = false;
            conventionalName = proc.Name;

            // Check Controls
            foreach (var ctrl in mod.Controls)
            {
              if (proc.Name.StartsWith(ctrl.Name + "_", StringComparison.OrdinalIgnoreCase))
              {
                // It is a control event - mark control as used
                ctrl.Used = true;
                ctrl.References.Add(new VbReference
                {
                  Module = mod.Name,
                  Procedure = proc.Name
                });

                var eventPart = proc.Name.Substring(ctrl.Name.Length + 1);
                conventionalName = ctrl.ConventionalName + "_" + ToPascalCase(eventPart); // Use ConventionalName of Control!
                isEvent = true;
                break;
              }
            }

            if (!isEvent)
            {
              // Check Global Variables (WithEvents)
              foreach (var v in mod.GlobalVariables)
              {
                // Note: v.Name could be "objATH3204_0". 
                // v.ConventionalName would be "ObjATH32040" (if my previous logic holds).

                // Check if proc name starts with v.Name + "_"
                if (proc.Name.StartsWith(v.Name + "_", StringComparison.OrdinalIgnoreCase))
                {
                  var eventPart = proc.Name.Substring(v.Name.Length + 1);
                  // Use ConventionalName of Variable (which strips prefix/underscore etc) + "_" + Event
                  // But for Event part, we apply PascalCase.
                  conventionalName = v.ConventionalName + "_" + ToPascalCase(eventPart);
                  isEvent = true;
                  break;
                }
              }
            }

            if (!isEvent)
            {
              // Regular procedure with underscore?
              // User said for procedures: PascalCase.
              // If it's not an event, convert to PascalCase (removing underscores).
              conventionalName = ToPascalCase(proc.Name);
            }
          }
          else
          {
            conventionalName = ToPascalCase(proc.Name);
          }

          // SPECIALE: Property Get/Let/Set con lo stesso nome devono mantenere lo stesso ConventionalName
          // Verifica se questa è una Property e se esiste già una Property Get/Let/Set con lo stesso nome
          if (proc.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase))
          {
            // Estrai il nome base della proprietà (senza Get/Let/Set)
            var basePropName = proc.Name;
            
            // Cerca se esiste già una Property (Get/Let/Set) con lo stesso nome base
            var existingProperty = mod.Procedures.FirstOrDefault(p => 
                p != proc &&
                p.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase) &&
                p.Name.Equals(basePropName, StringComparison.OrdinalIgnoreCase));
            
            if (existingProperty != null && !string.IsNullOrEmpty(existingProperty.ConventionalName))
            {
              // Usa lo stesso ConventionalName della Property esistente
              proc.ConventionalName = existingProperty.ConventionalName;
            }
            else
            {
              // Prima Property con questo nome, applica conflict resolution normale
              proc.ConventionalName = ResolveNameConflict(conventionalName, procedureNamesUsed);
              procedureNamesUsed.Add(proc.ConventionalName);
            }
          }
          else
          {
            // Procedure normale, applica conflict resolution
            proc.ConventionalName = ResolveNameConflict(conventionalName, procedureNamesUsed);
            procedureNamesUsed.Add(proc.ConventionalName);
          }

          if (IsReservedWord(proc.ConventionalName))
          {
            procedureNamesUsed.Remove(proc.ConventionalName);
            proc.ConventionalName = proc.Name;
            procedureNamesUsed.Add(proc.ConventionalName);
          }

          // Per ogni procedura, gestione conflict resolution di parametri + variabili locali
          var localScopeNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

          // Parametri
          foreach (var p in proc.Parameters)
          {
            var paramConventionalName = ToCamelCase(p.Name);
            p.ConventionalName = ResolveNameConflict(paramConventionalName, localScopeNamesUsed);
            localScopeNamesUsed.Add(p.ConventionalName);

            if (IsReservedWord(p.ConventionalName))
            {
              localScopeNamesUsed.Remove(p.ConventionalName);
              p.ConventionalName = p.Name;
              localScopeNamesUsed.Add(p.ConventionalName);
            }
          }

          // Variabili locali
          foreach (var v in proc.LocalVariables)
          {
            var baseName = GetBaseNameFromHungarian(v.Name);
            string localConventionalName;

            if (v.IsStatic)
            {
              // Per le variabili static: controlla se inizia già con "s_"
              if (baseName.StartsWith("s_", StringComparison.OrdinalIgnoreCase))
              {
                // Già nel formato corretto s_Nome, mantieni così com'è
                localConventionalName = baseName;
              }
              else
              {
                // Per le variabili static: se il nome inizia con 's' minuscolo, rimuovilo prima di applicare ToPascalCase
                // Es: sQualcosa -> s_Qualcosa (NON s_SQualcosa)
                var nameForPascal = baseName;
                if (baseName.Length > 1 && baseName.StartsWith("s", StringComparison.OrdinalIgnoreCase) && char.IsUpper(baseName[1]))
                {
                  // Ha prefisso 's' seguito da maiuscola: sQualcosa -> Qualcosa
                  nameForPascal = baseName.Substring(1);
                }
                localConventionalName = "s_" + ToPascalCase(nameForPascal);
              }
            }
            else
              localConventionalName = ToCamelCase(baseName);

            v.ConventionalName = ResolveNameConflict(localConventionalName, localScopeNamesUsed);
            localScopeNamesUsed.Add(v.ConventionalName);

            if (IsReservedWord(v.ConventionalName))
            {
              localScopeNamesUsed.Remove(v.ConventionalName);
              v.ConventionalName = v.Name;
              localScopeNamesUsed.Add(v.ConventionalName);
            }
          }

          // Costanti locali
          foreach (var c in proc.Constants)
          {
            var constantConventionalName = ToScreamingSnakeCase(c.Name);
            c.ConventionalName = ResolveNameConflict(constantConventionalName, localScopeNamesUsed);
            localScopeNamesUsed.Add(c.ConventionalName);

            if (IsReservedWord(c.ConventionalName))
            {
              localScopeNamesUsed.Remove(c.ConventionalName);
              c.ConventionalName = c.Name;
              localScopeNamesUsed.Add(c.ConventionalName);
            }
          }
        }
      }
    }

    private static readonly HashSet<string> HungarianPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        "int", "str", "lng", "dbl", "sng", "cur", "bol", "byt", "chr", "dat", "obj", "arr", "udt"
    };

    private static string GetBaseNameFromHungarian(string name)
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
          // Se la parte è tutta maiuscola (come "MAX" o "TIME"), normalizzala
          if (part.All(char.IsUpper) && part.All(char.IsLetter))
          {
            // Converti a PascalCase: prima lettera maiuscola, resto minuscolo
            result += char.ToUpper(part[0]) + part.Substring(1).ToLower();
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
  }
}
