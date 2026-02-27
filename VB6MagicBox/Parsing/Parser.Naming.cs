using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    public static partial class NamingConvention
    {
        // Mappa dei prefissi standard per controlli VB6 comuni
        private static readonly Dictionary<string, string> ControlPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        // Controlli standard VB6
        { "TextBox", "txt" },
        { "CommandButton", "cmd" },
        { "Command", "cmd" },
        { "Label", "lbl" },
        { "Frame", "fra" },
        { "CheckBox", "chk" },
        { "Check", "chk" },
        { "OptionButton", "opt" },
        { "Option", "opt" },
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
            else if (cleanType.StartsWith("MSFlexGridLib.MSFlex", StringComparison.OrdinalIgnoreCase))   
                cleanType = cleanType.Substring(20);

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
                        // tbSQM o txEmail ? txtSQM, txtEmail (preserva capitalizzazione)
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
                // Ha già il prefisso corretto, preserva la capitalizzazione esistente
                var baseName = controlName.Substring(expectedPrefix.Length);
                return expectedPrefix + baseName;
            }

            // Controlla se ha un altro prefisso a 3 lettere (es: txt, cmd, lbl)
            if (controlName.Length > 3 &&
                char.IsLower(controlName[0]) &&
                char.IsLower(controlName[1]) &&
                char.IsLower(controlName[2]) &&
                char.IsUpper(controlName[3]))
            {
                // Ha un prefisso diverso, sostituiscilo ma preserva capitalizzazione
                var baseName = controlName.Substring(3);
                return expectedPrefix + baseName;
            }

            // Non ha prefisso, aggiungilo ma preserva capitalizzazione
            // Solo la prima lettera deve essere maiuscola per camelCase
            var nameToFix = controlName;
            if (nameToFix.Length > 0 && char.IsLower(nameToFix[0]))
            {
                nameToFix = char.ToUpper(nameToFix[0]) + nameToFix.Substring(1);
            }
            return expectedPrefix + nameToFix;
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
            else if (baseName.EndsWith("T", StringComparison.OrdinalIgnoreCase) &&
                     baseName.Length > 1 &&
                     (char.IsLower(baseName[baseName.Length - 2]) || char.IsDigit(baseName[baseName.Length - 2])))
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

            // REGOLA 1: Se il tipo contiene un punto (namespace), probabilmente è una classe esterna
            // Esempi: Sweeper.NIDAQMX, OpticalMonitor.Device, TimeGetWrapper.TimeGetWrap
            if (segments.Count > 1)
            {
                return true; // Considera qualsiasi tipo con namespace come classe
            }

            // REGOLA 2: Cerca nei moduli locali del progetto
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

            // REGOLA 3: Fallback - se il nome include "cls" da qualche parte
            return Regex.IsMatch(typeName, @"\bcls\w+", RegexOptions.IgnoreCase);
        }

    }
}
