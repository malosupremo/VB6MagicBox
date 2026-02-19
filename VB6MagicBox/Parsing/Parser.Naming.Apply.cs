using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  public static partial class NamingConvention
  {
    public static void Apply(VbProject project)
    {
      foreach (var mod in project.Modules)
      {
        // Trackers per conflict resolution - scope separati
        var globalNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var procedureNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Module/Class naming
        var rawName = Path.GetFileNameWithoutExtension(mod.Name); // Rimuove .cls/.bas/.frm extension, mantiene nome file

        // Per Form e Classi, estrai il nome base (senza prefisso) prima di applicare ToPascalCase
        // così SQM242.frm ? FrmSqm242 e EXEC1.bas ? Exec1
        string pascalName;
        if (mod.Kind.Equals("frm", StringComparison.OrdinalIgnoreCase))
        {
          var baseName = rawName.StartsWith("frm", StringComparison.OrdinalIgnoreCase)
              ? rawName.Substring(3)
              : rawName;
          pascalName = "Frm" + ToPascalCase(baseName);
        }
        else if (mod.Kind.Equals("cls", StringComparison.OrdinalIgnoreCase))
        {
          var baseName = rawName.StartsWith("cls", StringComparison.OrdinalIgnoreCase)
              ? rawName.Substring(3)
              : rawName;
          pascalName = "Cls" + ToPascalCase(baseName);
        }
        else
        {
          pascalName = ToPascalCase(rawName);
        }

        // Se il nome risultante è una keyword C# (es. "With"), aggiungi "Class"
        if (CSharpKeywords.Contains(pascalName.ToLower())) // Check lowercase against lowercase keywords usually
        {
          pascalName += "Class";
        }

        mod.ConventionalName = pascalName;

        // Global Variables - con conflict resolution
        foreach (var v in mod.GlobalVariables)
        {
          var baseName = GetBaseNameFromHungarian(v.Name, v.Type);

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
            // Public globals: comportamento diverso per Module vs Form/Class
            bool isModule = mod.Kind.Equals("bas", StringComparison.OrdinalIgnoreCase);

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
              if (isModule)
              {
                // MODULE: oggetti pubblici ? g_Name (globali reali)
                if (baseName.StartsWith("g_", StringComparison.OrdinalIgnoreCase))
                {
                  var tail = baseName.Substring(2);
                  // Se il tail è già in PascalCase, mantieni il nome originale
                  if (IsPascalCase(tail))
                  {
                    conventionalName = baseName + (arraySuffix ?? "");
                  }
                  else
                  {
                    // Il tail non è PascalCase, applicalo
                    conventionalName = "g_" + ToPascalCase(tail) + (arraySuffix ?? "");
                  }
                }
                else
                {
                  // Non inizia con g_, aggiungilo
                  conventionalName = "g_" + ToPascalCase(baseName) + (arraySuffix ?? "");
                }
              }
              else if (mod.IsClass)
              {
                // CLASSE: membri pubblici ? PascalCase puro (come properties e procedure)
                // Es: ServerObj       ? ServerObj
                // Es: objUAServerObj  ? UAServerObj  (strip prefisso obj ungherese)
                conventionalName = ToPascalCase(baseName) + (arraySuffix ?? "");
              }
              else
              {
                // FORM: oggetti pubblici ? objName
                // Es: UAServerObj ? objUAServerObj
                // Es: objFM489   ? objFM489 (già OK)
                var tail = baseName;

                // Rimuovi prefisso obj se presente (per evitare objObjName)
                if (tail.StartsWith("obj", StringComparison.OrdinalIgnoreCase) && tail.Length > 3)
                  tail = tail.Substring(3);

                conventionalName = "obj" + ToPascalCase(tail) + (arraySuffix ?? "");
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
            else if (IsCSharpKeyword(fieldConventionalName))
            {
              // Per keyword C# nei campi, aggiungi suffisso "Value"
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
            var baseName = GetBaseNameFromHungarian(v.Name, v.Type);
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
              v.ConventionalName = v.ConventionalName + "Value";
              localScopeNamesUsed.Add(v.ConventionalName);
            }
          }

          // Costanti locali
          foreach (var c in proc.Constants)
          {
            var constantConventionalName = ToScreamingSnakeCase(c.Name);
            c.ConventionalName = ResolveNameConflict(constantConventionalName, localScopeNamesUsed);
            localScopeNamesUsed.Add(c.ConventionalName);

          }
        }

        // Properties - con conflict resolution (separata dalle procedure)
        foreach (var prop in mod.Properties)
        {
          string conventionalName;

          // Event Handling Detection (simile alle procedure)
          if (prop.Name.Equals("Class_Initialize", StringComparison.OrdinalIgnoreCase) ||
              prop.Name.Equals("Class_Terminate", StringComparison.OrdinalIgnoreCase) ||
              (mod.Kind.Equals("frm", StringComparison.OrdinalIgnoreCase) &&
               (prop.Name.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) ||
                prop.Name.StartsWith("UserControl_", StringComparison.OrdinalIgnoreCase))))
          {
            conventionalName = prop.Name;
          }
          else if (prop.Name.Contains("_"))
          {
            // Possibile event handler: Object_Event
            var prefixName = prop.Name.Split('_')[0];
            var isControlEventHandler = mod.Controls.Any(ctrl =>
              string.Equals(ctrl.Name, prefixName, StringComparison.OrdinalIgnoreCase));

            if (isControlEventHandler)
            {
              // È un event handler di controllo: ObjectName_EventName
              var parts = prop.Name.Split('_', 2);
              if (parts.Length == 2)
              {
                var objectName = ToPascalCase(parts[0]);
                var eventName = ToPascalCase(parts[1]);
                conventionalName = $"{objectName}_{eventName}";
              }
              else
              {
                conventionalName = ToPascalCase(prop.Name);
              }
            }
            else
            {
              conventionalName = ToPascalCase(prop.Name);
            }
          }
          else
          {
            conventionalName = ToPascalCase(prop.Name);
          }

          // SPECIALE: Property Get/Let/Set con lo stesso nome devono mantenere lo stesso ConventionalName
          // Cerca se esiste già una Property (Get/Let/Set) con lo stesso nome base
          var existingProperty = mod.Properties.FirstOrDefault(p =>
              p != prop &&
              p.Name.Equals(prop.Name, StringComparison.OrdinalIgnoreCase));

          if (existingProperty != null && !string.IsNullOrEmpty(existingProperty.ConventionalName))
          {
            // Usa lo stesso ConventionalName della Property esistente
            prop.ConventionalName = existingProperty.ConventionalName;
          }
          else
          {
            // Prima Property con questo nome, applica conflict resolution normale
            prop.ConventionalName = ResolveNameConflict(conventionalName, procedureNamesUsed);
            procedureNamesUsed.Add(prop.ConventionalName);
          }

          if (IsReservedWord(prop.ConventionalName))
          {
            procedureNamesUsed.Remove(prop.ConventionalName);
            prop.ConventionalName = prop.Name;
            procedureNamesUsed.Add(prop.ConventionalName);
          }

          // Per ogni proprietà, gestione conflict resolution di parametri
          var localScopeNamesUsed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

          // Parametri
          foreach (var param in prop.Parameters)
          {
            var paramConventionalName = ToCamelCase(param.Name);
            param.ConventionalName = ResolveNameConflict(paramConventionalName, localScopeNamesUsed);
            localScopeNamesUsed.Add(param.ConventionalName);

            if (IsReservedWord(param.ConventionalName))
            {
              localScopeNamesUsed.Remove(param.ConventionalName);
              param.ConventionalName = param.Name;
              localScopeNamesUsed.Add(param.ConventionalName);
            }
          }
        }
      }
    }
  }
}
