using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  // ---------------------------------------------------------
  // RISOLUZIONE TIPI, CHIAMATE, CAMPI
  // ---------------------------------------------------------

  public static void ResolveTypesAndCalls(VbProject project)
  {
    // Indicizzazione procedure per nome (ESCLUSE le proprietà)
    var procIndex = new Dictionary<string, List<(string Module, VbProcedure Proc)>>(
        StringComparer.OrdinalIgnoreCase);

    // Indicizzazione proprietà per nome (SEPARATA dalle procedure)
    var propIndex = new Dictionary<string, List<(string Module, VbProperty Prop)>>(
        StringComparer.OrdinalIgnoreCase);

    foreach (var mod in project.Modules)
    {
      // Indicizza solo le procedure normali (NON le proprietà)
      foreach (var proc in mod.Procedures.Where(p => !p.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase)))
      {
        if (!procIndex.TryGetValue(proc.Name, out var list))
        {
          list = new List<(string, VbProcedure)>();
          procIndex[proc.Name] = list;
        }
        list.Add((mod.Name, proc));
      }

      // Indicizza le proprietà separatamente
      foreach (var prop in mod.Properties)
      {
        if (!propIndex.TryGetValue(prop.Name, out var propList))
        {
          propList = new List<(string, VbProperty)>();
          propIndex[prop.Name] = propList;
        }
        propList.Add((mod.Name, prop));
      }
    }

    // Indicizzazione tipi
    var typeIndex = project.Modules
        .SelectMany(m => m.Types.Select(t => new { Module = m, Type = t }))
        .GroupBy(x => x.Type.Name, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(g => g.Key, g => g.First().Type, StringComparer.OrdinalIgnoreCase);

    // Scansione a livello di modulo per rilevare riferimenti a procedure definite in altri moduli
    // (es. costrutti nel form fuori da qualsiasi procedura: "If Not Is_Ready_To_Start Then")
    foreach (var mod in project.Modules)
    {
      var fileLines = File.ReadAllLines(mod.FullPath);
      foreach (var rawLine in fileLines)
      {
        var noCommentLine = rawLine;
        var idxc = noCommentLine.IndexOf("'");
        if (idxc >= 0)
          noCommentLine = noCommentLine.Substring(0, idxc);

        foreach (Match wm in Regex.Matches(noCommentLine, @"\b([A-Za-z_]\w*)\b"))
        {
          var token = wm.Groups[1].Value;

          if (VbKeywords.Contains(token))
            continue;

          // Ignore tokens that are global variables or types in the same module
          if (mod.GlobalVariables.Any(v => string.Equals(v.Name, token, StringComparison.OrdinalIgnoreCase)))
            continue;
          if (mod.Types.Any(t => string.Equals(t.Name, token, StringComparison.OrdinalIgnoreCase)))
            continue;

          // Le proprietà NON vengono considerate come chiamate nude a livello di modulo
          // Solo le procedure normali possono essere chiamate così
          if (procIndex.TryGetValue(token, out var targets) && targets.Count > 0)
          {
            foreach (var t in targets)
            {
              // mark only procedures defined in other modules (usage from this module)
              if (!string.Equals(t.Module, mod.Name, StringComparison.OrdinalIgnoreCase) && t.Proc != null)
                t.Proc.Used = true;
            }
          }
        }
      }
    }

    // Risoluzione chiamate e campi
    int moduleIndex = 0;
    int totalModules = project.Modules.Count;
    
    foreach (var mod in project.Modules)
    {
      moduleIndex++;
      
      // Progress inline per il parsing
      Console.Write($"\r      [{moduleIndex}/{totalModules}] {Path.GetFileName(mod.FullPath)}...".PadRight(Console.WindowWidth - 1));
      
      var fileLines = File.ReadAllLines(mod.FullPath);          

      // Pre-scan: Traccia i tipi globali attraverso assegnamenti Set a livello di modulo
      var globalTypeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      foreach (var line in fileLines)
      {
        var noComment = line;
        var idx = noComment.IndexOf("'");
        if (idx >= 0)
          noComment = noComment.Substring(0, idx);

        // Pattern: Set varName = New ClassName
        var matchSetNew = Regex.Match(noComment, @"Set\s+(\w+)\s*=\s*New\s+(\w+)", RegexOptions.IgnoreCase);
        if (matchSetNew.Success)
        {
          var varName = matchSetNew.Groups[1].Value;
          var className = matchSetNew.Groups[2].Value;
          globalTypeMap[varName] = className;
        }

        // Pattern: Set varName = otherVar (type aliasing) - include object.property access
        var matchSetAlias = Regex.Match(noComment, @"Set\s+(\w+)\s*=\s+(\w+(?:\.\w+)?)\b", RegexOptions.IgnoreCase);
        if (matchSetAlias.Success)
        {
          var varName = matchSetAlias.Groups[1].Value;
          var sourceVar = matchSetAlias.Groups[2].Value;
          
          // Se è object.property, estrai il tipo da globalTypeMap
          if (sourceVar.Contains('.'))
          {
            var parts = sourceVar.Split('.');
            var objName = parts[0];
            var propName = parts[1];
            
            // Se conosciamo il tipo di objName, usalo
            if (globalTypeMap.TryGetValue(objName, out var objType))
            {
              globalTypeMap[varName] = objType;
            }
          }
          else if (globalTypeMap.TryGetValue(sourceVar, out var sourceType) && !string.IsNullOrEmpty(sourceType))
          {
            globalTypeMap[varName] = sourceType;
          }
        }
      }

      // Indicizzazione classi per nome
      var classIndex = new Dictionary<string, VbModule>(StringComparer.OrdinalIgnoreCase);
      foreach (var classModule in project.Modules.Where(m => m.IsClass))
      {
        var fileName = Path.GetFileNameWithoutExtension(classModule.Name);
        classIndex[fileName] = classModule;
        
        // Aggiungi anche il nome senza namespace (ultimo token dopo il punto)
        if (fileName.Contains('.'))
        {
          var shortName = fileName.Split('.').Last();
          if (!classIndex.ContainsKey(shortName))
            classIndex[shortName] = classModule;
        }
        
        // Aggiungi anche basato su procedure names e tipi definiti nel modulo
        // (Nel caso il file contenga una classe ma non sia chiaro dal nome)
        // Se il modulo ha procedure pubbliche, potrebbe essere una classe
        if (classModule.Procedures.Any(p => p.Scope.Equals("Public", StringComparison.OrdinalIgnoreCase)))
        {
          // Prova ad aggiungere varianti del nome
          var variants = new[] { fileName };
          foreach (var variant in variants)
          {
            if (!classIndex.ContainsKey(variant))
              classIndex[variant] = classModule;
          }
        }
      }

      // Indice di tutti i moduli per VB_Name, per rilevare accessi diretti
      // Es: FrmRestart.Show, SHARESTRUCT.MY_CONST (senza variabile dichiarata)
      var moduleByName = project.Modules
          .Where(m => !string.IsNullOrEmpty(m.Name))
          .ToDictionary(m => m.Name, m => m, StringComparer.OrdinalIgnoreCase);

      foreach (var proc in mod.Procedures)
      {
        // Ambiente variabili à tipo
        var env = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Debug disabled - keep variable but set to null to avoid any logging
        List<string> debugInfo = null;

        // Carica i tipi globali
        foreach (var kvp in globalTypeMap)
          env[kvp.Key] = kvp.Value;

        // Carica GlobalVariables da TUTTI i moduli del progetto (non solo da mod)
        foreach (var anyMod in project.Modules)
        {
          foreach (var v in anyMod.GlobalVariables)
            if (!string.IsNullOrEmpty(v.Name) && !string.IsNullOrEmpty(v.Type))
            {
              // Non sovrascrivere se esiste già (priorità al modulo corrente)
              if (!env.ContainsKey(v.Name))
              {
                env[v.Name] = v.Type;
                if (debugInfo != null && (v.Name.StartsWith("gobj") || v.Name.Contains("Plasma") || v.Name.Contains("Qpc") || v.Name.Contains("Plc")))
                {
                  debugInfo.Add($"Loaded: {v.Name} -> {v.Type}");
                }
              }
            }
        }

        foreach (var p in proc.Parameters)
          if (!string.IsNullOrEmpty(p.Name) && !string.IsNullOrEmpty(p.Type))
            env[p.Name] = p.Type;

        foreach (var lv in proc.LocalVariables)
          if (!string.IsNullOrEmpty(lv.Name) && !string.IsNullOrEmpty(lv.Type))
            env[lv.Name] = lv.Type;

        // Traccia i tipi attraverso assegnamenti Set locali alla procedura PRIMA dei pass
        for (int typeTrackLine = proc.LineNumber - 1; typeTrackLine < fileLines.Length; typeTrackLine++)
        {
          var rawSetLine = fileLines[typeTrackLine];
          var noCommentSetLine = rawSetLine;
          var idxSet = noCommentSetLine.IndexOf("'", StringComparison.Ordinal);
          if (idxSet >= 0)
            noCommentSetLine = noCommentSetLine.Substring(0, idxSet);

          if (typeTrackLine > proc.LineNumber - 1 && noCommentSetLine.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
            break;

          // Pattern: Set varName = New ClassName
          var matchSetNew = Regex.Match(noCommentSetLine, @"Set\s+(\w+)\s*=\s*New\s+(\w+)", RegexOptions.IgnoreCase);
          if (matchSetNew.Success)
          {
            var varName = matchSetNew.Groups[1].Value;
            var className = matchSetNew.Groups[2].Value;
            env[varName] = className;
          }

          // Pattern: Set varName = otherVar (type aliasing) - include object.property access
          var matchSetAlias = Regex.Match(noCommentSetLine, @"Set\s+(\w+)\s*=\s+(\w+(?:\.\w+)?)\b", RegexOptions.IgnoreCase);
          if (matchSetAlias.Success)
          {
            var varName = matchSetAlias.Groups[1].Value;
            var sourceVar = matchSetAlias.Groups[2].Value;
            
            // Se è object.property, estrai il tipo da env (già popolato con global vars)
            if (sourceVar.Contains("."))
            {
              var parts = sourceVar.Split('.');
              var objName = parts[0];
              
              // Se conosciamo il tipo di objName, usalo
              if (env.TryGetValue(objName, out var objType) && !string.IsNullOrEmpty(objType))
              {
                env[varName] = objType;
              }
            }
            else if (env.TryGetValue(sourceVar, out var sourceType) && !string.IsNullOrEmpty(sourceType))
            {
              env[varName] = sourceType;
            }
          }
        }

        // Marcatura tipi usati nel corpo della procedura
        foreach (var line in fileLines.Skip(proc.LineNumber - 1))
        {
          // Fine procedura
          if (line.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
            break;

          foreach (var p in proc.Parameters)
          {
            if (p == null || string.IsNullOrEmpty(p.Name) || p.Used)
              continue;

            if (Regex.IsMatch(line, $@"\b{Regex.Escape(p.Name)}\b", RegexOptions.IgnoreCase))
              p.Used = true;
          }
        }

        // Risoluzione chiamate
        foreach (var call in proc.Calls)
        {
          // Se è object.method
          if (!string.IsNullOrEmpty(call.ObjectName))
          {
            if (env.TryGetValue(call.ObjectName, out var objType))
              call.ResolvedType = objType;
          }

          var bareName = call.MethodName ?? call.Raw;

          // Prova a risolvere come procedure
          if (procIndex.TryGetValue(bareName, out var targets) && targets.Count > 0)
          {
            if (!string.IsNullOrEmpty(call.ResolvedType))
            {
              // Cerca modulo classe corrispondente
              var match = targets.FirstOrDefault(t =>
                  Path.GetFileNameWithoutExtension(t.Module)
                      .Equals(call.ResolvedType, StringComparison.OrdinalIgnoreCase));

              if (match.Proc != null)
              {
                call.ResolvedModule = match.Module;
                call.ResolvedProcedure = match.Proc.Name;
                call.ResolvedKind = match.Proc.Kind;
              }
              else
              {
                call.ResolvedModule = targets[0].Module;
                call.ResolvedProcedure = targets[0].Proc.Name;
                call.ResolvedKind = targets[0].Proc.Kind;
              }
            }
            else
            {
              call.ResolvedModule = targets[0].Module;
              call.ResolvedProcedure = targets[0].Proc.Name;
              call.ResolvedKind = targets[0].Proc.Kind;
            }
          }
          else
          {
            // Prova a risolvere come variabile globale nel modulo
            var globalVar = mod.GlobalVariables.FirstOrDefault(v =>
                v.Name.Equals(bareName, StringComparison.OrdinalIgnoreCase));

            if (globalVar != null)
            {
              call.ResolvedModule = mod.Name;
              call.ResolvedProcedure = globalVar.Name;
              call.ResolvedKind = "Variable";
            }
          }
        }

        // Risoluzione accessi ai campi: var.field
        ResolveFieldAccesses(mod, proc, fileLines, typeIndex, env);

        // Risoluzione accessi ai controlli: control.Property o control.Method()
        ResolveControlAccesses(mod, proc, fileLines);

        // Risoluzione reference per parametri e variabili locali
        ResolveParameterAndLocalVariableReferences(mod, proc, fileLines);

        // Rilevamento delle occorrenze nude di altre procedure nel corpo (es. "Not Is_Caller_Busy" o "If Is_Caller_Busy Then")
        // Controlli di sicurezza per evitare IndexOutOfRangeException
        if (proc.StartLine <= 0)
        {
          proc.StartLine = proc.LineNumber;
        }
        
        if (proc.EndLine <= 0)
        {
          proc.EndLine = fileLines.Length;
        }
        
        // Assicurati che gli indici siano validi
        var startIndex = Math.Max(0, proc.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, proc.EndLine);
        
        if (startIndex >= fileLines.Length)
        {
          continue; // Skip this procedure
        }
        
        for (int li = startIndex; li < endIndex; li++)
        {
          var rawLine = fileLines[li];
          
          if (debugInfo != null && li == proc.LineNumber - 1)
          {
            debugInfo.Add($"\n[START] proc.LineNumber={proc.LineNumber}, fileLines.Length={fileLines.Length}");
          }
          
          var noCommentLine = rawLine;
          var idx = noCommentLine.IndexOf("'");
          if (idx >= 0)
            noCommentLine = noCommentLine.Substring(0, idx);

          // Rimuovi stringhe per evitare di catturare pattern dentro stringhe
          noCommentLine = Regex.Replace(noCommentLine, @"""[^""]*""", "\"\"");

          // Se incontriamo la fine della procedura interrompiamo
          if (li > proc.LineNumber - 1 && IsProcedureEndLine(noCommentLine, proc.Kind))
            break;
          
          if (debugInfo != null && li > proc.LineNumber + 1 && li <= proc.LineNumber + 20)
          {
            debugInfo.Add($"[Line {li}] Checking: {rawLine.Substring(0, Math.Min(60, rawLine.Length))}");
          }

          // PASS 1: Estrai tutte le chiamate with parentheses object.method() o method()
          foreach (Match callMatch in Regex.Matches(noCommentLine, @"(?:(\w+)\.)?(\w+)\s*\("))
          {
            var objName = callMatch.Groups[1].Success ? callMatch.Groups[1].Value : null;
            var methodName = callMatch.Groups[2].Value;

            // Ignora keywords
            if (VbKeywords.Contains(methodName))
              continue;

            // Filtra auto-referenza: se è una procedura nel modulo corrente, non aggiungerla
            if (string.IsNullOrEmpty(objName) && string.Equals(methodName, proc.Name, StringComparison.OrdinalIgnoreCase))
              continue;

            // Skip if already in calls
            if (proc.Calls.Any(c => string.Equals(c.Raw, objName != null ? $"{objName}.{methodName}" : methodName, StringComparison.OrdinalIgnoreCase)))
              continue;

            // Se è object.method, risolvi il tipo dell'oggetto
            if (!string.IsNullOrEmpty(objName))
            {
              if (env.TryGetValue(objName, out var objType) && !string.IsNullOrEmpty(objType))
              {
                // Se il tipo contiene un namespace (es. QuarzPC.clsQuarzPC), prendi solo l'ultima parte
                var classNameToLookup = objType.Contains(".") ? objType.Split('.').Last() : objType;
                
                // Cerca la procedura nella classe
                if (classIndex.TryGetValue(classNameToLookup, out var classModule))
                {
                  // PRIMA cerca nelle proprietà (hanno precedenza negli accessi con punto)
                  var classProp = classModule.Properties.FirstOrDefault(p =>
                      p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));

                  if (classProp != null)
                  {
                    classProp.Used = true;
                    
                    classProp.References.AddLineNumber(mod.Name, proc.Name, li + 1);

                    proc.Calls.Add(new VbCall
                    {
                      Raw = objName != null ? $"{objName}.{methodName}" : methodName,
                      ObjectName = objName,
                      MethodName = methodName,
                      ResolvedType = objType,
                      ResolvedModule = classModule.Name,
                      ResolvedProcedure = classProp.Name,
                      ResolvedKind = $"Property{classProp.Kind}",
                      LineNumber = li + 1
                    });
                  }
                  else
                  {
                    // Se non è una proprietà, cerca nelle procedure normali
                    var classProc = classModule.Procedures.FirstOrDefault(p =>
                        p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));

                    if (classProc != null)
                    {
                      classProc.Used = true;
                      proc.Calls.Add(new VbCall
                      {
                        Raw = objName != null ? $"{objName}.{methodName}" : methodName,
                        ObjectName = objName,
                        MethodName = methodName,
                        ResolvedType = objType,
                        ResolvedModule = classModule.Name,
                        ResolvedProcedure = classProc.Name,
                        ResolvedKind = classProc.Kind,
                        LineNumber = li + 1
                      });
                    }
                  }
                }
              }
            }

            // Altrimenti risolvi come procedura semplice
            if (procIndex.TryGetValue(methodName, out var targets) && targets.Count > 0)
            {
              (string Module, VbProcedure TargetProc) selected;
              if (env.TryGetValue(methodName, out var resolvedType))
              {
                selected = targets.FirstOrDefault(t => Path.GetFileNameWithoutExtension(t.Module).Equals(resolvedType, StringComparison.OrdinalIgnoreCase));
                if (selected.TargetProc == null)
                  selected = targets[0];
              }
              else
              {
                selected = targets[0];
              }

              if (selected.TargetProc != null)
              {
                selected.TargetProc.Used = true;

                proc.Calls.Add(new VbCall
                {
                  Raw = methodName,
                  MethodName = methodName,
                  ResolvedModule = selected.Module,
                  ResolvedProcedure = selected.TargetProc.Name,
                  ResolvedKind = selected.TargetProc.Kind,
                  LineNumber = li + 1
                });
              }
            }
          }

          // PASS 1.5: Estrai object.method SENZA parentesi (es. gobjPlasmaSource.Timer in un'assegnazione)
          foreach (Match methodAccessMatch in Regex.Matches(noCommentLine, @"(\w+)\.(\w+)(?!\s*\()"))
          {
            var objName = methodAccessMatch.Groups[1].Value;
            var methodName = methodAccessMatch.Groups[2].Value;

            if (VbKeywords.Contains(methodName) || VbKeywords.Contains(objName))
              continue;

            // Controlla se già nelle Calls (per evitare duplicati nelle Calls,
            // ma per le proprietà aggiungiamo comunque i LineNumbers alle References)
            var alreadyInCalls = proc.Calls.Any(c => string.Equals(c.Raw, $"{objName}.{methodName}", StringComparison.OrdinalIgnoreCase));

            // Risolvi il tipo dell'oggetto
            if (env.TryGetValue(objName, out var objType) && !string.IsNullOrEmpty(objType))
            {
              // Se il tipo contiene un namespace (es. QuarzPC.clsQuarzPC), prendi solo l'ultima parte
              var classNameToLookup = objType.Contains(".") ? objType.Split('.').Last() : objType;
              
              // Cerca nella classe
              if (classIndex.TryGetValue(classNameToLookup, out var classModule))
              {
                // PRIMA cerca nelle proprietà (hanno precedenza negli accessi con punto)
                var classProp = classModule.Properties.FirstOrDefault(p =>
                    p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));

                if (classProp != null)
                {
                  classProp.Used = true;
                  
                  classProp.References.AddLineNumber(mod.Name, proc.Name, li + 1);

                  if (!alreadyInCalls)
                  {
                    proc.Calls.Add(new VbCall
                    {
                      Raw = $"{objName}.{methodName}",
                      ObjectName = objName,
                      MethodName = methodName,
                      ResolvedType = classNameToLookup,
                      ResolvedModule = classModule.Name,
                      ResolvedProcedure = classProp.Name,
                      ResolvedKind = $"Property{classProp.Kind}",
                      LineNumber = li + 1
                    });
                  }
                }
                else if (!alreadyInCalls)
                {
                  // Se non è una proprietà, cerca nelle procedure normali
                  var classProc = classModule.Procedures.FirstOrDefault(p =>
                      p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));

                  if (classProc != null)
                  {
                    classProc.Used = true;
                    proc.Calls.Add(new VbCall
                    {
                      Raw = $"{objName}.{methodName}",
                      ObjectName = objName,
                      MethodName = methodName,
                      ResolvedType = classNameToLookup,
                      ResolvedModule = classModule.Name,
                      ResolvedProcedure = classProc.Name,
                      ResolvedKind = classProc.Kind,
                      LineNumber = li + 1
                    });
                  }
                }
              }
            }
          }

          // PASS 1.5b: Generico - cerca object.method dove object è in env (è una variabile nota)
          // Pattern: qualsiasi IDENTIFIER.IDENTIFIER OVUNQUE nella riga
          var trimmedLineForMethods = noCommentLine.Trim();
          foreach (Match genericMethodMatch in Regex.Matches(trimmedLineForMethods, @"(\w+)\.(\w+)", RegexOptions.IgnoreCase))
          {
            var objName = genericMethodMatch.Groups[1].Value;
            var methodName = genericMethodMatch.Groups[2].Value;

            // NON escludere keywords per object.method - possono essere metodi custom
            // (es. gobjPlc.Timer è valido anche se Timer è una built-in function)
            if (VbKeywords.Contains(objName))
              continue;

            // Traccia riferimento diretto a modulo noto: FrmRestart.Show, Module.Proc
            // Copre i casi in cui il modulo è usato per nome senza variabile dichiarata
            if (moduleByName.TryGetValue(objName, out var referencedModule) &&
                !string.Equals(referencedModule.Name, mod.Name, StringComparison.OrdinalIgnoreCase))
            {
              referencedModule.Used = true;
              referencedModule.References.AddLineNumber(mod.Name, proc.Name, li + 1);
            }

            // Se objName NON è un oggetto noto in env, non proseguire con la risoluzione del tipo
            var objInEnv = env.TryGetValue(objName, out var objType);
            if (!objInEnv || string.IsNullOrEmpty(objType))
              continue;

            // Controlla se già nelle Calls (per le proprietà aggiungiamo comunque i LineNumbers)
            var alreadyInCalls = proc.Calls.Any(c => string.Equals(c.Raw, $"{objName}.{methodName}", StringComparison.OrdinalIgnoreCase));

            // Se il tipo contiene un namespace (es. QuarzPC.clsQuarzPC), prendi solo l'ultima parte
            var classNameToLookup = objType.Contains(".") ? objType.Split('.').Last() : objType;
            
            // Cerca nella classe
            if (classIndex.TryGetValue(classNameToLookup, out var classModule))
            {
              // PRIMA cerca nelle proprietà (hanno precedenza negli accessi con punto)
              var classProp = classModule.Properties.FirstOrDefault(p =>
                  p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));

              if (classProp != null)
              {
                classProp.Used = true;
                
                classProp.References.AddLineNumber(mod.Name, proc.Name, li + 1);

                if (!alreadyInCalls)
                {
                  proc.Calls.Add(new VbCall
                  {
                    Raw = $"{objName}.{methodName}",
                    ObjectName = objName,
                    MethodName = methodName,
                    ResolvedType = classNameToLookup,
                    ResolvedModule = classModule.Name,
                    ResolvedProcedure = classProp.Name,
                    ResolvedKind = $"Property{classProp.Kind}",
                    LineNumber = li + 1
                  });
                }
              }
              else if (!alreadyInCalls)
              {
                // Se non è una proprietà, cerca nelle procedure normali
                var classProc = classModule.Procedures.FirstOrDefault(p =>
                    p.Name.Equals(methodName, StringComparison.OrdinalIgnoreCase));

                if (classProc != null)
                {
                  classProc.Used = true;
                  proc.Calls.Add(new VbCall
                  {
                    Raw = $"{objName}.{methodName}",
                    ObjectName = objName,
                    MethodName = methodName,
                    ResolvedType = classNameToLookup,
                    ResolvedModule = classModule.Name,
                    ResolvedProcedure = classProc.Name,
                    ResolvedKind = classProc.Kind,
                    LineNumber = li + 1
                  });
                }
              }
            }
          }

          // Print debug info at end of procedure
          if (debugInfo != null && debugInfo.Count > 0)
          {
            // Aggiungi lista di classi disponibili nel classIndex
            if (proc.Name.Equals("CallObjectTimer", StringComparison.OrdinalIgnoreCase))
            {
              debugInfo.Add($"\nAvailable classes in classIndex: {string.Join(", ", classIndex.Keys.OrderBy(k => k))}");
            }
            Console.WriteLine(string.Join("\n", debugInfo));
          }
        }
      }
    }

    Console.WriteLine(); // Vai a capo dopo il progress del parsing

    // Aggiunge References ai tipi per ogni posizione in cui appaiono in "As TypeName"
    ResolveTypeReferences(project, typeIndex);

    // Aggiunge References alle classi per ogni dichiarazione "As [New] ClassName"
    ResolveClassModuleReferences(project);

    // Marcatura tipi usati
    MarkUsedTypes(project);
  }

  // ---------------------------------------------------------
  // COSTRUZIONE DIPENDENZE + MARCATURA USED
  // ---------------------------------------------------------

  /// <summary>
  /// Legge un file con FileShare.Read per evitare blocchi di file
  /// quando il file è aperto da altri processi (es. IDE)
  /// </summary>
  private static string[] ReadAllLinesShared(string filePath)
  {
    try
    {
      using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
      using (var reader = new StreamReader(stream))
      {
        var content = reader.ReadToEnd();
        return content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
      }
    }
    catch (IOException ex)
    {
      Console.WriteLine($"    [WARN] Impossibile leggere {Path.GetFileName(filePath)}: {ex.Message}");
      return Array.Empty<string>();
    }
  }

  public static void BuildDependenciesAndUsage(VbProject project)
  {
    var procByModuleAndName = new Dictionary<(string Module, string Name), VbProcedure>();

    foreach (var mod in project.Modules)
      foreach (var proc in mod.Procedures)
        procByModuleAndName[(mod.Name, proc.Name)] = proc;

    var varByModuleAndName = new Dictionary<(string Module, string Name), VbVariable>();

    foreach (var mod in project.Modules)
      foreach (var variable in mod.GlobalVariables)
        varByModuleAndName[(mod.Name, variable.Name)] = variable;

    int moduleIndex = 0;
    int totalModules = project.Modules.Count;

    foreach (var mod in project.Modules)
    {
      moduleIndex++;
      
      // Estrai il nome del file senza path per il log
      var fileName = Path.GetFileName(mod.FullPath);
      var moduleName = Path.GetFileNameWithoutExtension(mod.Name);
      Console.WriteLine($"\r  [{moduleIndex}/{totalModules}] {fileName} ({moduleName})...".PadRight(Console.WindowWidth - 1));

      int counter = 0; 

      foreach (var proc in mod.Procedures)
      {
        // Progress inline per il parsing
        Console.Write($"\r      [Procedure {counter++}/{mod.Procedures.Count}] {proc.Name}...".PadRight(Console.WindowWidth - 1));

        foreach (var call in proc.Calls.DistinctBy(c => $"{c.Raw}|{c.ResolvedModule}|{c.ResolvedProcedure}|{c.LineNumber}"))
        {
          project.Dependencies.Add(new DependencyEdge
          {
            CallerModule = mod.Name,
            CallerProcedure = proc.Name,
            CalleeRaw = call.Raw,
            CalleeModule = call.ResolvedModule,
            CalleeProcedure = call.ResolvedProcedure
          });

          // Marca procedure chiamate
          if (!string.IsNullOrEmpty(call.ResolvedModule) &&
              !string.IsNullOrEmpty(call.ResolvedProcedure) &&
              procByModuleAndName.TryGetValue((call.ResolvedModule, call.ResolvedProcedure), out var targetProc))
          {
            targetProc.Used = true;
            // Usa il line number dalla call, se non disponibile usa il line number della procedura
            var lineNum = call.LineNumber > 0 ? call.LineNumber : proc.LineNumber;
            targetProc.References.AddLineNumber(mod.Name, proc.Name, lineNum);
          }

          // Marca classi usate
          if (!string.IsNullOrEmpty(call.ResolvedType))
          {
            var clsMod = project.Modules.FirstOrDefault(m =>
                m.IsClass &&
                Path.GetFileNameWithoutExtension(m.Name)
                    .Equals(call.ResolvedType, StringComparison.OrdinalIgnoreCase));

            if (clsMod != null)
              clsMod.Used = true;
          }
        }
      }
      counter = 0;

      // Marca variabili globali usate e traccia references
      // Per variabili Public/Global, cerca in TUTTI i moduli
      // Per variabili Private/Dim, cerca solo nel modulo corrente
      foreach (var v in mod.GlobalVariables)
      {
        // Progress inline per il parsing
        Console.Write($"\r      [Variable {counter++}/{mod.GlobalVariables.Count}] {v.Name}...".PadRight(Console.WindowWidth - 1));

        bool isPublic = string.IsNullOrEmpty(v.Visibility) || 
                       v.Visibility.Equals("Public", StringComparison.OrdinalIgnoreCase) ||
                       v.Visibility.Equals("Global", StringComparison.OrdinalIgnoreCase);

        // Determina in quali moduli cercare
        var modulesToSearch = isPublic 
            ? project.Modules  // Public/Global: cerca ovunque
            : new List<VbModule> { mod };  // Private/Dim: solo nel modulo corrente

        foreach (var searchMod in modulesToSearch)
        {
          var searchLines = ReadAllLinesShared(searchMod.FullPath);
          int lineNum = 0;
          
          foreach (var line in searchLines)
          {
            lineNum++;
            if (line.IndexOf(v.Name, StringComparison.OrdinalIgnoreCase) >= 0)
            {
              v.Used = true;
              // Trova la procedura corretta che contiene questa riga
              var procAtLine = searchMod.GetProcedureAtLine(lineNum);
              if (procAtLine != null)
              {
                // CONTROLLO SHADOW: Se la procedura ha una variabile locale con lo stesso nome,
                // quella locale fa "shadow" della globale, quindi NON aggiungere reference
                var hasLocalWithSameName = procAtLine.LocalVariables.Any(lv => 
                    lv.Name.Equals(v.Name, StringComparison.OrdinalIgnoreCase)) ||
                  procAtLine.Parameters.Any(p => 
                    p.Name.Equals(v.Name, StringComparison.OrdinalIgnoreCase));
                
                if (hasLocalWithSameName)
                {
                  // La variabile locale fa shadow di quella globale, skip
                  continue;
                }

                v.References.AddLineNumber(searchMod.Name, procAtLine.Name, lineNum);
              }
            }
          }
        }
      }

      counter = 0;
      // Marca costanti usate (modulo level) e traccia references
      // Per costanti Public/Global, cerca in TUTTI i moduli
      // Per costanti Private, cerca solo nel modulo corrente
      foreach (var c in mod.Constants)
      {
        // Progress inline per il parsing
        Console.Write($"\r      [Costant {counter++}/{mod.Constants.Count}] {c.Name}...".PadRight(Console.WindowWidth - 1));

        bool isPublic = string.IsNullOrEmpty(c.Visibility) || 
                       c.Visibility.Equals("Public", StringComparison.OrdinalIgnoreCase) ||
                       c.Visibility.Equals("Global", StringComparison.OrdinalIgnoreCase);

        // Determina in quali moduli cercare
        var modulesToSearch = isPublic 
            ? project.Modules  // Public/Global: cerca ovunque
            : new List<VbModule> { mod };  // Private: solo nel modulo corrente

        foreach (var searchMod in modulesToSearch)
        {
          var searchLines = ReadAllLinesShared(searchMod.FullPath);
          int lineNum = 0;
          
          foreach (var line in searchLines)
          {
            lineNum++;
            if (line.IndexOf(c.Name, StringComparison.OrdinalIgnoreCase) >= 0)
            {
              c.Used = true;
              // Trova la procedura corretta che contiene questa riga
              var procAtLine = searchMod.GetProcedureAtLine(lineNum);
              if (procAtLine != null)
              {
                // CONTROLLO SHADOW: Se la procedura ha una costante locale con lo stesso nome,
                // quella locale fa "shadow" della globale, quindi NON aggiungere reference
                var hasLocalWithSameName = procAtLine.Constants.Any(lc => 
                    lc.Name.Equals(c.Name, StringComparison.OrdinalIgnoreCase));
                
                if (hasLocalWithSameName)
                {
                  // La costante locale fa shadow di quella globale, skip
                  continue;
                }

                c.References.AddLineNumber(searchMod.Name, procAtLine.Name, lineNum);
              }
            }
          }
        }
      }

    }

    Console.WriteLine(); // Vai a capo dopo il progress del parsing

    // Marcatura tipi usati
    MarkUsedTypes(project);

    // Propaga Used al modulo: se qualunque membro è usato, il modulo è usato
    foreach (var mod in project.Modules)
    {
      if (!mod.Used)
      {
        mod.Used = mod.Procedures.Any(p => p.Used)
                || mod.Properties.Any(p => p.Used)
                || mod.GlobalVariables.Any(v => v.Used)
                || mod.Constants.Any(c => c.Used)
                || mod.Enums.Any(e => e.Used || e.Values.Any(v => v.Used))
                || mod.Types.Any(t => t.Used)
                || mod.Controls.Any(c => c.Used)
                || mod.Events.Any(e => e.Used);
      }
    }

    // Costruisce ModuleReferences: per ogni modulo raccoglie i moduli che lo referenziano
    // attraverso qualsiasi suo membro (costanti, tipi, enum, procedure, property, controlli, variabili).
    // Non modifica le References esistenti — è un aggregato di sola lettura.
    foreach (var mod in project.Modules)
    {
      var callers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

      void Collect(IEnumerable<VbReference> refs)
      {
        foreach (var r in refs)
          if (!string.IsNullOrEmpty(r.Module) &&
              !string.Equals(r.Module, mod.Name, StringComparison.OrdinalIgnoreCase))
            callers.Add(r.Module);
      }

      foreach (var proc in mod.Procedures)   Collect(proc.References);
      foreach (var prop in mod.Properties)   Collect(prop.References);
      foreach (var v    in mod.GlobalVariables) Collect(v.References);
      foreach (var c    in mod.Constants)    Collect(c.References);
      foreach (var e    in mod.Enums)        { Collect(e.References); foreach (var val in e.Values) Collect(val.References); }
      foreach (var t    in mod.Types)        { Collect(t.References); foreach (var f   in t.Fields) Collect(f.References); }
      foreach (var c    in mod.Controls)     Collect(c.References);
      foreach (var ev   in mod.Events)       Collect(ev.References);
      Collect(mod.References);

      mod.ModuleReferences = callers
          .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
          .ToList();
    }
  }
}
