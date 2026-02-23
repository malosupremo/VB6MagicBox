using VB6MagicBox.Models;
using System.Text.RegularExpressions;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  // -------------------------
  // REGEX MODULO
  // -------------------------

  private static readonly Regex ReVbName =
      new(@"^Attribute\s+VB_Name\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase);

  private static readonly Regex ReFormBegin =
      new(@"^Begin\s+VB\.Form\s+(\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

  // -------------------------
  // PARSING MODULO
  // -------------------------

  public static VbModule ParseModule(string path, string kind)
  {
    var mod = new VbModule
    {
      Name = Path.GetFileNameWithoutExtension(path),
      ConventionalName = string.Empty,
      Path = path,
      Kind = kind
    };

    // Set FullPath for I/O; Path may be overridden by caller to be relative
    mod.FullPath = path;

    var originalLines = File.ReadAllLines(path);

    // Gestisci le righe con continuazione "_"
    var (collapsedLines, lineMapping) = CollapseLineContinuations(originalLines);

    VbProcedure currentProc = null;
    VbProperty currentProperty = null;
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
      // ATTRIBUTE VB_NAME (per classi e moduli)
      // -------------------------
      var mvbName = ReVbName.Match(noComment);
      if (mvbName.Success)
      {
        mod.Name = mvbName.Groups[1].Value;

        //introduco la primissima reference
        mod.References.Add(new VbReference()
        {
          Module = mod.Name,
          Procedure = string.Empty,
          LineNumbers = new List<int> { originalLineNumber }
        });
        continue;
      }

      // -------------------------
      // BEGIN VB.FORM (per form)
      // -------------------------
      if (kind.Equals("frm", StringComparison.OrdinalIgnoreCase))
      {
        var mFormBegin = ReFormBegin.Match(noComment);
        if (mFormBegin.Success)
        {
          mod.Name = mFormBegin.Groups[1].Value;

          //introduco la primissima reference
          mod.References.Add(new VbReference()
          {
            Module = mod.Name,
            Procedure = string.Empty,
            LineNumbers = new List<int> { originalLineNumber }
          });
          continue;
        }
      }

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

          // Crea sempre un nuovo controllo temporaneo per raccogliere i dati
          currentControl = new VbControl
          {
            ControlType = controlType,
            Name = controlName,
            LineNumber = originalLineNumber
          };

          mod.Controls.Add(currentControl);
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
      // IMPLEMENTS (class interfaces)
      // -------------------------
      var mi = ReImplements.Match(noComment);
      if (mi.Success)
      {
        var interfaceName = mi.Groups[1].Value;
        if (!string.IsNullOrEmpty(interfaceName) &&
            !mod.ImplementsInterfaces.Any(i => i.Equals(interfaceName, StringComparison.OrdinalIgnoreCase)))
        {
          mod.ImplementsInterfaces.Add(interfaceName);
        }
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
          Kind = "Function",
          IsStatic = isStatic,
          Scope = "Module",
          Parameters = ParseParameters(mf2.Groups[4].Value, originalLineNumber),
          ReturnType = NormalizeTypeName(mf2.Groups[6].Value),
          LineNumber = originalLineNumber,
          StartLine = originalLineNumber
        };
        FixParameterLineNumbersForMultilineSignature(currentProc, originalLines, originalLineNumber);
        FixParameterTypeLineNumbersForMultilineSignature(currentProc, originalLines, originalLineNumber);
        FixReturnTypeLineNumberForMultilineSignature(currentProc, originalLines, originalLineNumber);
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
          Kind = "Sub",
          IsStatic = isStatic,
          Scope = "Module",
          Parameters = ParseParameters(ms.Groups[4].Value, originalLineNumber),
          LineNumber = originalLineNumber,
          StartLine = originalLineNumber
        };
        FixParameterLineNumbersForMultilineSignature(currentProc, originalLines, originalLineNumber);
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

        // Crea SOLO Property (NON aggiungere a Procedures per evitare duplicazioni)
        currentProperty = new VbProperty
        {
          Visibility = string.IsNullOrEmpty(mp.Groups[1].Value) ? "Public" : mp.Groups[1].Value,
          Kind = mp.Groups[3].Value, // Get, Let, Set
          Name = mp.Groups[4].Value,
          Scope = "Module",
          Parameters = ParseParameters(mp.Groups[5].Value, originalLineNumber),
          ReturnType = NormalizeTypeName(mp.Groups[7].Value),
          LineNumber = originalLineNumber,
          StartLine = originalLineNumber
        };
        FixParameterLineNumbersForMultilineSignature(currentProperty, originalLines, originalLineNumber);
        FixParameterTypeLineNumbersForMultilineSignature(currentProperty, originalLines, originalLineNumber);
        FixReturnTypeLineNumberForMultilineSignature(currentProperty, originalLines, originalLineNumber);
        mod.Properties.Add(currentProperty);

        // Imposta currentProc SOLO per tracciare la fine della Property (NON aggiungere a mod.Procedures)
        // Questo è un oggetto temporaneo usato solo per il parsing interno
        currentProc = new VbProcedure
        {
          Visibility = string.IsNullOrEmpty(mp.Groups[1].Value) ? "Public" : mp.Groups[1].Value,
          Kind = $"Property{mp.Groups[3].Value}",
          Name = mp.Groups[4].Value,
          IsStatic = isStatic,
          Scope = "Module",
          Parameters = ParseParameters(mp.Groups[5].Value, originalLineNumber),
          ReturnType = NormalizeTypeName(mp.Groups[7].Value),
          LineNumber = originalLineNumber,
          StartLine = originalLineNumber
        };
        // NON aggiungere: mod.Procedures.Add(currentProc); ? Questo causava le duplicazioni!

        // Traccia event handler per controlli (es. Command1_Click)
        foreach (var control in mod.Controls)
        {
          var prefix = $"{control.Name}_";
          if (currentProperty.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
          {
            // Questa proprietà è un event handler di questo controllo
            var existingRef = control.References.FirstOrDefault(r =>
              string.Equals(r.Module, mod.Name, StringComparison.OrdinalIgnoreCase) &&
              string.Equals(r.Procedure, currentProperty.Name, StringComparison.OrdinalIgnoreCase));

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
                Procedure = currentProperty.Name,
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
        var declareProc = new VbProcedure
        {
          Visibility = string.IsNullOrEmpty(mdf.Groups[1].Value) ? "Public" : mdf.Groups[1].Value,
          Name = mdf.Groups[2].Value,
          Kind = "ExternalFunction",
          IsStatic = false,
          Scope = "Module",
          Parameters = ParseParameters(mdf.Groups[4].Value, originalLineNumber),
          ReturnType = NormalizeTypeName(mdf.Groups[6].Value),
          LineNumber = originalLineNumber,
          StartLine = originalLineNumber,
          EndLine = originalLineNumber
        };
        FixReturnTypeLineNumberForMultilineSignature(declareProc, originalLines, originalLineNumber);
        FixParameterTypeLineNumbersForMultilineSignature(declareProc, originalLines, originalLineNumber);

        // AGGIUNGE REFERENCES AUTOMATICHE per i parametri su righe multiple
        AddParameterReferencesForMultilineDeclaration(declareProc, mod.Name, originalLines, originalLineNumber, lineMapping, lineIndex);

        mod.Procedures.Add(declareProc);
        continue;
      }

      // -------------------------
      // DECLARE SUB (External APIs)
      // -------------------------
      var mds = ReDeclareSub.Match(noComment);
      if (mds.Success)
      {
        var declareProc = new VbProcedure
        {
          Visibility = string.IsNullOrEmpty(mds.Groups[1].Value) ? "Public" : mds.Groups[1].Value,
          Name = mds.Groups[2].Value,
          Kind = "ExternalFunction",
          IsStatic = false,
          Scope = "Module",
          Parameters = ParseParameters(mds.Groups[4].Value, originalLineNumber),
          LineNumber = originalLineNumber,
          StartLine = originalLineNumber,
          EndLine = originalLineNumber
        };

        // AGGIUNGE REFERENCES AUTOMATICHE per i parametri su righe multiple
        AddParameterReferencesForMultilineDeclaration(declareProc, mod.Name, originalLines, originalLineNumber, lineMapping, lineIndex);
        FixParameterTypeLineNumbersForMultilineSignature(declareProc, originalLines, originalLineNumber);

        mod.Procedures.Add(declareProc);
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
      // VARIABILI GLOBALI/MEMBRO SENZA TIPO (fallback per TypeAnnotator)
      // -------------------------
      var mgnt = ReGlobalVarNoType.Match(noComment);
      if (mgnt.Success && currentProc == null && currentType == null)
      {
        var scopeValueNt = mgnt.Groups[1].Value;
        var normalizedScopeNt = (scopeValueNt.Equals("Private", StringComparison.OrdinalIgnoreCase) ||
                                  scopeValueNt.Equals("Dim", StringComparison.OrdinalIgnoreCase)) ? "Module" :
                                  scopeValueNt.Equals("Public", StringComparison.OrdinalIgnoreCase) ? "Project" :
                                  scopeValueNt;
        mod.GlobalVariables.Add(new VbVariable
        {
          Name = mgnt.Groups[3].Value,
          Type = "",
          IsArray = !string.IsNullOrEmpty(mgnt.Groups[5].Value),
          IsWithEvents = !string.IsNullOrEmpty(mgnt.Groups[2].Value),
          Scope = normalizedScopeNt,
          Visibility = scopeValueNt.Equals("Dim", StringComparison.OrdinalIgnoreCase) ? "Private" : scopeValueNt,
          Level = kind.Equals("cls", StringComparison.OrdinalIgnoreCase) ? "Member" : "Global",
          LineNumber = originalLineNumber
        });
        continue;
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
          currentProc.EndLine = originalLineNumber;

          // Se stiamo chiudendo una Property, aggiorna anche currentProperty
          if (trimmedNoComment.StartsWith("End Property", StringComparison.OrdinalIgnoreCase)
              && currentProperty != null)
          {
            currentProperty.EndLine = originalLineNumber;
            currentProperty = null;
          }

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
            Type = ml.Groups[4].Value,
            IsStatic = isStatic,
            IsArray = !string.IsNullOrEmpty(ml.Groups[3].Value),
            Scope = "Procedure",
            Visibility = "Private",
            Level = "Local",
            LineNumber = originalLineNumber
          });
        }
        else
        {
          // variabili locali senza tipo esplicito (fallback per TypeAnnotator)
          var mlnt = ReLocalVarNoType.Match(noComment);
          if (mlnt.Success)
          {
            currentProc.LocalVariables.Add(new VbVariable
            {
              Name = mlnt.Groups[2].Value,
              Type = "",
              IsStatic = mlnt.Groups[1].Value.Equals("Static", StringComparison.OrdinalIgnoreCase),
              IsArray = !string.IsNullOrEmpty(mlnt.Groups[4].Value),
              Scope = "Procedure",
              Visibility = "Private",
              Level = "Local",
              LineNumber = originalLineNumber
            });
          }
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
                    {
                        //MAO, una funzione può richiamare se stessa ricorsivamente anche senza parentesi, ad esempio un avanzamento immediato di tick
                        DebugLog(currentProc.Name, originalLineNumber, name, "ADDED: Self-reference (return value assignment)", "ReBareCall");
                        currentProc.Calls.Add(new VbCall
                        {
                            Raw = name,
                            MethodName = name,
                            LineNumber = originalLineNumber
                        });
                    }
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
          //if (isSelfRef)
          //{
          //  DebugLog(currentProc.Name, originalLineNumber, methodName, "SKIPPED: Self-reference (recursion)", "ReCallWithoutParens");
          //  continue;
          //}

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

    // -------------------------
    // POST-PROCESSING: RAGGRUPPA CONTROLLI ARRAY
    // -------------------------
    if (kind.Equals("frm", StringComparison.OrdinalIgnoreCase))
    {
      // Raggruppa i controlli per nome
      var controlGroups = mod.Controls
          .GroupBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
          .ToList();

      // Sostituisci la lista controlli con quella raggruppata
      mod.Controls.Clear();

      foreach (var group in controlGroups)
      {
        var controlList = group.OrderBy(c => c.LineNumber).ToList();
        var primaryControl = controlList.First();

        // Configura il controllo principale
        primaryControl.IsArray = controlList.Count > 1;
        primaryControl.LineNumber = controlList.First().LineNumber; // Prima riga
        primaryControl.LineNumbers = controlList.Select(c => c.LineNumber).ToList(); // Tutte le righe

        // Aggiungi solo il controllo principale (no duplicati)
        mod.Controls.Add(primaryControl);
      }
    }

    return mod;
  }
}
