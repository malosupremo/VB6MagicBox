using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    // ---------------------------------------------------------
    // REGEX COMPILATE PER HOT-PATH (usate nei loop di risoluzione)
    // ---------------------------------------------------------

    private static readonly Regex ReSetNew = 
        new(@"Set\s+(\w+)\s*=\s*New\s+(\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReSetAlias = 
        new(@"Set\s+(\w+)\s*=\s+(\w+(?:\.\w+)?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReWordBoundary = 
        new(@"\b([A-Za-z_]\w*)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReObjectMethod = 
        new(@"(\w+)\.(\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

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

        var enumValueIndex = project.Modules
            .SelectMany(m => m.Enums.SelectMany(e => e.Values))
            .Select(v => v.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

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

                foreach (Match wm in ReWordBoundary.Matches(noCommentLine))
                {
                    var token = wm.Groups[1].Value;

                    if (VbKeywords.Contains(token))
                        continue;

                    // Ignore tokens that are global variables or types in the same module
                    if (mod.GlobalVariables.Any(v => string.Equals(v.Name, token, StringComparison.OrdinalIgnoreCase)))
                        continue;
                    if (mod.Types.Any(t => string.Equals(t.Name, token, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    // Controlla se è una procedura pubblica in altri moduli
                    if (procIndex.TryGetValue(token, out var targets) && targets.Count > 0)
                    {
                        foreach (var t in targets)
                        {
                            // mark only procedures defined in other modules (usage from this module)
                            if (!string.Equals(t.Module, mod.Name, StringComparison.OrdinalIgnoreCase) && t.Proc != null)
                                t.Proc.Used = true;
                        }
                    }

                    // Controlla se è una proprietà pubblica in altri moduli (bare usage)
                    if (propIndex.TryGetValue(token, out var propTargets) && propTargets.Count > 0)
                    {
                        foreach (var pt in propTargets)
                        {
                            // mark only properties defined in other modules
                            if (!string.Equals(pt.Module, mod.Name, StringComparison.OrdinalIgnoreCase) && pt.Prop != null)
                            {
                                pt.Prop.Used = true;
                                // Non aggiungiamo References qui perché non conosciamo la procedura chiamante
                                // Le references bare saranno aggiunte dopo nel contesto delle procedure
                            }
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
                var matchSetNew = ReSetNew.Match(noComment);
                if (matchSetNew.Success)
                {
                    var varName = matchSetNew.Groups[1].Value;
                    var className = matchSetNew.Groups[2].Value;
                    globalTypeMap[varName] = className;
                }

                // Pattern: Set varName = otherVar (type aliasing) - include object.property access
                var matchSetAlias = ReSetAlias.Match(noComment);
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

                if (proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrEmpty(proc.Name) &&
                    !string.IsNullOrEmpty(proc.ReturnType))
                {
                    env[proc.Name] = proc.ReturnType;
                }

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
                    var matchSetNew = ReSetNew.Match(noCommentSetLine);
                    if (matchSetNew.Success)
                    {
                        var varName = matchSetNew.Groups[1].Value;
                        var className = matchSetNew.Groups[2].Value;
                        env[varName] = className;
                    }

                    // Pattern: Set varName = otherVar (type aliasing) - include object.property access
                    var matchSetAlias = ReSetAlias.Match(noCommentSetLine);
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
                ResolveFieldAccesses(mod, proc, fileLines, typeIndex, env, classIndex);

                // Risoluzione accessi ai controlli: control.Property o control.Method()
                ResolveControlAccesses(mod, proc, fileLines);

                // Risoluzione reference per parametri e variabili locali
                ResolveParameterAndLocalVariableReferences(mod, proc, fileLines);
                ResolveFunctionReturnReferences(mod, proc, fileLines);

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

                    // PASS 1.2: Chiamate senza parentesi usate come argomenti (es. RunCmdSim Queue_Pop_Sim)
                    if (li != proc.LineNumber - 1 &&
                        !Regex.IsMatch(noCommentLine, @"^\s*(Dim|Static|Const|ReDim|Set)\b", RegexOptions.IgnoreCase))
                    {
                        foreach (Match tokenMatch in Regex.Matches(noCommentLine, @"\b([A-Za-z_]\w*)\b"))
                        {
                            var tokenName = tokenMatch.Groups[1].Value;

                            if (VbKeywords.Contains(tokenName))
                                continue;

                            if (enumValueIndex.Contains(tokenName))
                                continue;

                            if (env.ContainsKey(tokenName))
                                continue;

                            if (string.Equals(tokenName, proc.Name, StringComparison.OrdinalIgnoreCase))
                                continue;

                            if (tokenMatch.Index > 0 && noCommentLine[tokenMatch.Index - 1] == '.')
                                continue;

                            // Prima controlla nelle procedure pubbliche
                            if (procIndex.TryGetValue(tokenName, out var targets) && targets.Count > 0)
                            {
                                if (proc.Calls.Any(c => string.Equals(c.Raw, tokenName, StringComparison.OrdinalIgnoreCase)))
                                    continue;

                                (string Module, VbProcedure TargetProc) selected;
                                if (env.TryGetValue(tokenName, out var resolvedType))
                                {
                                    selected = targets.FirstOrDefault(t =>
                                        Path.GetFileNameWithoutExtension(t.Module)
                                            .Equals(resolvedType, StringComparison.OrdinalIgnoreCase));
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
                                        Raw = tokenName,
                                        MethodName = tokenName,
                                        ResolvedModule = selected.Module,
                                        ResolvedProcedure = selected.TargetProc.Name,
                                        ResolvedKind = selected.TargetProc.Kind,
                                        LineNumber = li + 1
                                    });
                                }
                                continue;
                            }

                            // Poi controlla nelle proprietà pubbliche (bare usage cross-module)
                            // Es: If ExecSts = DEPOSIT_STS Then
                            if (propIndex.TryGetValue(tokenName, out var propTargets) && propTargets.Count > 0)
                            {
                                if (proc.Calls.Any(c => string.Equals(c.Raw, tokenName, StringComparison.OrdinalIgnoreCase)))
                                    continue;

                                // Cerca proprietà Get in altri moduli (public properties usabili bare)
                                var propTarget = propTargets.FirstOrDefault(pt => 
                                    pt.Prop.Kind.Equals("Get", StringComparison.OrdinalIgnoreCase) &&
                                    !string.Equals(pt.Module, mod.Name, StringComparison.OrdinalIgnoreCase));

                                // Se non trovata in altri moduli, prova nel modulo corrente
                                if (propTarget.Prop == null)
                                {
                                    propTarget = propTargets.FirstOrDefault(pt => 
                                        pt.Prop.Kind.Equals("Get", StringComparison.OrdinalIgnoreCase));
                                }

                                if (propTarget.Prop != null)
                                {
                                    propTarget.Prop.Used = true;
                                    propTarget.Prop.References.AddLineNumber(mod.Name, proc.Name, li + 1);

                                    proc.Calls.Add(new VbCall
                                    {
                                        Raw = tokenName,
                                        MethodName = tokenName,
                                        ResolvedModule = propTarget.Module,
                                        ResolvedProcedure = propTarget.Prop.Name,
                                        ResolvedKind = $"Property{propTarget.Prop.Kind}",
                                        LineNumber = li + 1
                                    });
                                }
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
                    foreach (Match genericMethodMatch in ReObjectMethod.Matches(trimmedLineForMethods))
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

            foreach (var prop in mod.Properties)
            {
                var env = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var kvp in globalTypeMap)
                    env[kvp.Key] = kvp.Value;

                foreach (var anyMod in project.Modules)
                {
                    foreach (var v in anyMod.GlobalVariables)
                        if (!string.IsNullOrEmpty(v.Name) && !string.IsNullOrEmpty(v.Type))
                        {
                            if (!env.ContainsKey(v.Name))
                                env[v.Name] = v.Type;
                        }
                }

                foreach (var p in prop.Parameters)
                    if (!string.IsNullOrEmpty(p.Name) && !string.IsNullOrEmpty(p.Type))
                        env[p.Name] = p.Type;

                ResolveFieldAccesses(mod, prop, fileLines, typeIndex, env, classIndex);
                ResolveParameterReferences(mod, prop, fileLines);
                ResolvePropertyReturnReferences(mod, prop, fileLines);
            }
        }

        Console.WriteLine(); // Vai a capo dopo il progress del parsing

        // Aggiunge References ai tipi per ogni posizione in cui appaiono in "As TypeName"
        ResolveTypeReferences(project, typeIndex);

        // Aggiunge References alle classi per ogni dichiarazione "As [New] ClassName"
        ResolveClassModuleReferences(project);

        // Aggiunge References ai valori enum usati (anche senza prefisso)
        ResolveEnumValueReferences(project);

        // Marcatura tipi usati
        MarkUsedTypes(project);
    }
}


