using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    // =================================================================
    // SINGLE-PASS REFERENCE RESOLUTION
    // =================================================================

    /// <summary>
    /// Resolves ALL references in a single pass: file by file, line by line, token by token.
    /// Replaces the old multi-scan approach (PASS 1, 1.2, 1.5, 1.5b + ResolveFieldAccesses +
    /// ResolveControlAccesses + ResolveParameterAndLocalVariableReferences + ResolveEnumValueReferences
    /// + ResolveTypeReferences + ResolveClassModuleReferences).
    /// </summary>
    public static void ResolveTypesAndCalls(VbProject project, Dictionary<string, string[]> fileCache)
    {
        var gIdx = BuildGlobalIndexes(project);

        int moduleIndex = 0;
        int totalModules = project.Modules.Count;

        foreach (var mod in project.Modules)
        {
            moduleIndex++;
            Console.Write($"\r      [{moduleIndex}/{totalModules}] {Path.GetFileName(mod.FullPath)}...".PadRight(Console.WindowWidth - 1));

            var fileLines = GetFileLines(fileCache, mod);

            // --- Per-module indexes ---
            var controlIndex = mod.Controls.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);

            // Build module-level env (variable → type) from global Set New assignments
            var globalTypeMap = BuildGlobalTypeMap(fileLines);

            // --- Process each procedure ---
            foreach (var proc in mod.Procedures)
            {
                var env = BuildProcEnv(proc, mod, gIdx, globalTypeMap, fileLines);

                ResolveProcedureBody(
                    mod, proc.Name, proc.Kind,
                    proc.StartLine, proc.EndLine, proc.LineNumber,
                    proc.Parameters, proc.LocalVariables, proc.Constants, proc.Calls,
                    proc.References,
                    isFunction: proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase),
                    fileLines, env, gIdx, controlIndex);
            }

            // --- Process each property ---
            foreach (var prop in mod.Properties)
            {
                EnsurePropertyBounds(prop, fileLines);
                var env = BuildPropEnv(prop, mod, gIdx, globalTypeMap);

                ResolveProcedureBody(
                    mod, prop.Name, prop.Kind ?? "Property",
                    prop.StartLine, prop.EndLine, prop.LineNumber,
                    prop.Parameters, localVariables: null, localConstants: null, calls: null,
                    prop.References,
                    isFunction: prop.Kind?.Equals("Get", StringComparison.OrdinalIgnoreCase) == true,
                    fileLines, env, gIdx, controlIndex);
            }

            PrunePropertyReferenceOverlaps(mod, fileLines);
        }

        Console.WriteLine();

        // --- Post-processing: declaration-level references (model iteration, no re-scan) ---
        // These add References for "As TypeName" in type fields, global vars, params, locals
        // and "As [New] ClassName" for class modules — lines outside procedure bodies.
        ResolveTypeReferences(project, gIdx, fileCache);
        ResolveClassModuleReferences(project, fileCache);

        // Mark used types from declarations
        MarkUsedTypes(project, fileCache);
    }

    // -----------------------------------------------------------------
    // ENVIRONMENT BUILDERS
    // -----------------------------------------------------------------

    private static Dictionary<string, string> BuildGlobalTypeMap(string[] fileLines)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in fileLines)
        {
            var noComment = StripInlineComment(line).TrimEnd();

            var matchSetNew = ReSetNew.Match(noComment);
            if (matchSetNew.Success)
            {
                map[matchSetNew.Groups[1].Value] = matchSetNew.Groups[2].Value;
            }

            var matchSetAlias = ReSetAlias.Match(noComment);
            if (matchSetAlias.Success)
            {
                var varName = matchSetAlias.Groups[1].Value;
                var source = matchSetAlias.Groups[2].Value;
                if (source.Contains('.'))
                {
                    var parts = source.Split('.');
                    if (map.TryGetValue(parts[0], out var objType))
                        map[varName] = objType;
                }
                else if (map.TryGetValue(source, out var srcType) && !string.IsNullOrEmpty(srcType))
                {
                    map[varName] = srcType;
                }
            }
        }
        return map;
    }

    private static Dictionary<string, string> BuildProcEnv(
        VbProcedure proc,
        VbModule mod,
        GlobalIndexes gIdx,
        Dictionary<string, string> globalTypeMap,
        string[] fileLines)
    {
        var env = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 1. Global type map
        foreach (var kvp in globalTypeMap)
            env[kvp.Key] = kvp.Value;

        // 2. Global variables from ALL modules
        foreach (var anyMod in mod.Owner?.Modules ?? Enumerable.Empty<VbModule>())
            foreach (var v in anyMod.GlobalVariables)
                if (!string.IsNullOrEmpty(v.Name) && !string.IsNullOrEmpty(v.Type))
                    env.TryAdd(v.Name, v.Type);

        // 3. Parameters
        foreach (var p in proc.Parameters)
            if (!string.IsNullOrEmpty(p.Name) && !string.IsNullOrEmpty(p.Type))
                env[p.Name] = p.Type;

        // 4. Local variables
        foreach (var lv in proc.LocalVariables)
            if (!string.IsNullOrEmpty(lv.Name) && !string.IsNullOrEmpty(lv.Type))
                env[lv.Name] = lv.Type;

        // 5. Function name → return type
        if (proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrEmpty(proc.Name) &&
            !string.IsNullOrEmpty(proc.ReturnType))
        {
            env[proc.Name] = proc.ReturnType;
        }

        // 6. Local Set New / Set Alias inside procedure body
        TrackLocalSetAssignments(proc.StartLine > 0 ? proc.StartLine : proc.LineNumber, proc.EndLine, fileLines, env);

        return env;
    }

    private static Dictionary<string, string> BuildPropEnv(
        VbProperty prop,
        VbModule mod,
        GlobalIndexes gIdx,
        Dictionary<string, string> globalTypeMap)
    {
        var env = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var kvp in globalTypeMap)
            env[kvp.Key] = kvp.Value;

        foreach (var anyMod in mod.Owner?.Modules ?? Enumerable.Empty<VbModule>())
            foreach (var v in anyMod.GlobalVariables)
                if (!string.IsNullOrEmpty(v.Name) && !string.IsNullOrEmpty(v.Type))
                    env.TryAdd(v.Name, v.Type);

        foreach (var p in prop.Parameters)
            if (!string.IsNullOrEmpty(p.Name) && !string.IsNullOrEmpty(p.Type))
                env[p.Name] = p.Type;

        return env;
    }

    private static void TrackLocalSetAssignments(int startLine, int endLine, string[] fileLines, Dictionary<string, string> env)
    {
        if (startLine <= 0) return;
        var start = Math.Max(0, startLine - 1);
        var end = Math.Min(fileLines.Length, endLine > 0 ? endLine : fileLines.Length);
        for (int i = start; i < end; i++)
        {
            var noComment = StripInlineComment(fileLines[i]).TrimEnd();
            if (i > start && noComment.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
                break;

            var m = ReSetNew.Match(noComment);
            if (m.Success)
                env[m.Groups[1].Value] = m.Groups[2].Value;

            var ma = ReSetAlias.Match(noComment);
            if (ma.Success)
            {
                var varName = ma.Groups[1].Value;
                var source = ma.Groups[2].Value;
                if (source.Contains('.'))
                {
                    var parts = source.Split('.');
                    if (env.TryGetValue(parts[0], out var objType) && !string.IsNullOrEmpty(objType))
                        env[varName] = objType;
                }
                else if (env.TryGetValue(source, out var srcType) && !string.IsNullOrEmpty(srcType))
                {
                    env[varName] = srcType;
                }
            }
        }
    }

    // -----------------------------------------------------------------
    // CORE SINGLE-PASS: RESOLVE PROCEDURE/PROPERTY BODY
    // -----------------------------------------------------------------

    private static void ResolveProcedureBody(
        VbModule mod,
        string memberName,
        string memberKind,
        int startLine,
        int endLine,
        int lineNumber,
        List<VbParameter> parameters,
        List<VbVariable>? localVariables,
        List<VbConstant>? localConstants,
        List<VbCall>? calls,
        List<VbReference> memberReferences,
        bool isFunction,
        string[] fileLines,
        Dictionary<string, string> env,
        GlobalIndexes gIdx,
        Dictionary<string, VbControl> controlIndex)
    {
        // Safety
        if (startLine <= 0) startLine = lineNumber;
        if (endLine <= 0) endLine = fileLines.Length;

        var startIdx = Math.Max(0, startLine - 1);
        var endIdx = Math.Min(fileLines.Length, endLine);
        if (startIdx >= fileLines.Length) return;

        // --- Per-member indexes (local scope) ---
        var paramIndex = (parameters ?? [])
            .Where(p => !string.IsNullOrEmpty(p.Name))
            .GroupBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var localVarIndex = (localVariables ?? [])
            .Where(v => !string.IsNullOrEmpty(v.Name))
            .GroupBy(v => v.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var localConstIndex = (localConstants ?? [])
            .Where(c => !string.IsNullOrEmpty(c.Name))
            .GroupBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var globalVarModIndex = mod.GlobalVariables
            .Where(v => !string.IsNullOrEmpty(v.Name))
            .GroupBy(v => v.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        // Shadow set: names that hide globals/procs/props
        var localNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in parameters ?? []) localNames.Add(p.Name);
        foreach (var v in localVariables ?? []) localNames.Add(v.Name);
        foreach (var c in localConstants ?? []) localNames.Add(c.Name ?? "");

        var withStack = new Stack<string>();

        // Track which (line, startChar) we have already recorded to avoid duplicates
        var recorded = new HashSet<(int Line, int StartChar)>();

        for (int i = startIdx; i < endIdx; i++)
        {
            var raw = fileLines[i];
            int currentLine = i + 1;

            // Stop at procedure end
            if (i > lineNumber - 1 && IsProcedureEndLine(raw, memberKind))
                break;

            // --- Strip comment, mask strings ---
            var noComment = StripInlineComment(raw);
            var trimmed = noComment.TrimStart();

            // --- With stack management ---
            var isWithLine = false;
            if (trimmed.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
            {
                isWithLine = true;
                var withExpr = trimmed.Substring(5).Trim();
                if (withExpr.StartsWith(".") && withStack.Count > 0)
                    withExpr = withStack.Peek() + withExpr;

                if (!string.IsNullOrEmpty(withExpr))
                    withStack.Push(withExpr);
                // Fall through to normal chain/token resolution for the With line
            }

            if (trimmed.Equals("End With", StringComparison.OrdinalIgnoreCase))
            {
                if (withStack.Count > 0) withStack.Pop();
                continue;
            }

            // Build effective line for dot-chain resolution (expand With prefix)
            // Skip expansion for With lines: their expression is already literal in the source
            var effectiveLine = noComment;
            if (!isWithLine && withStack.Count > 0 && trimmed.StartsWith(".", StringComparison.Ordinal))
            {
                var suffix = trimmed.Substring(1).TrimStart();
                effectiveLine = withStack.Peek() + "." + suffix;
            }
            if (!isWithLine && withStack.Count > 0)
            {
                var withPrefix = withStack.Peek();
                effectiveLine = ReWithDotReplacement.Replace(effectiveLine,
                    m => withPrefix + "." + m.Groups[1].Value);
            }

            var masked = MaskStringLiterals(noComment);
            var maskedEffective = MaskStringLiterals(effectiveLine);

            // Skip declaration lines for bare-token scanning
            var isDeclLine = Regex.IsMatch(trimmed, @"^\s*(Dim|Static|Const|ReDim|Set\s+\w+\s*=\s*New)\b", RegexOptions.IgnoreCase);

            // =============================================================
            // STEP 1: Resolve dot-chains (field access, control access,
            //         class member access, module-qualified access)
            // =============================================================
            var chainTokensClaimed = new HashSet<int>(); // startChar positions claimed by chains

            foreach (var (chainText, chainIndex) in EnumerateDotChains(maskedEffective))
            {
                ResolveChain(
                    chainText, chainIndex, maskedEffective, raw,
                    currentLine, mod, memberName,
                    env, gIdx, controlIndex, localNames,
                    paramIndex, localVarIndex, globalVarModIndex,
                    calls, recorded, chainTokensClaimed);
            }

            // Also search chains inside parentheses (e.g., CStr(obj.Prop))
            foreach (var (innerText, innerStart) in EnumerateParenContents(maskedEffective))
            {
                foreach (var (chainText, chainIndex) in EnumerateDotChains(innerText))
                {
                    ResolveChain(
                        chainText, innerStart + chainIndex, maskedEffective, raw,
                        currentLine, mod, memberName,
                        env, gIdx, controlIndex, localNames,
                        paramIndex, localVarIndex, globalVarModIndex,
                        calls, recorded, chainTokensClaimed);
                }
            }

            // Try unwrap function chains: CStr(obj.Field) → unwrap to obj.Field
            foreach (var (chainText, chainIndex) in EnumerateDotChains(maskedEffective))
            {
                if (TryUnwrapFunctionChain(chainText, chainIndex, out var unwrapped, out var unwrappedIdx))
                {
                    foreach (var (innerChain, innerIdx) in EnumerateDotChains(unwrapped))
                    {
                        ResolveChain(
                            innerChain, unwrappedIdx + innerIdx, maskedEffective, raw,
                            currentLine, mod, memberName,
                            env, gIdx, controlIndex, localNames,
                            paramIndex, localVarIndex, globalVarModIndex,
                            calls, recorded, chainTokensClaimed);
                    }
                }
            }

            // =============================================================
            // STEP 2: Resolve standalone (bare) tokens
            // =============================================================
            foreach (var (token, tokenIdx) in EnumerateTokens(masked))
            {
                // Already claimed by a dot-chain
                if (chainTokensClaimed.Contains(tokenIdx))
                    continue;

                // Member-access token (after a dot) — skip, already handled by chains
                if (IsMemberAccessToken(masked, tokenIdx))
                    continue;

                // VB keyword
                if (VbKeywords.Contains(token))
                    continue;

                // Skip auto-reference (procedure calling itself bare)
                if (string.Equals(token, memberName, StringComparison.OrdinalIgnoreCase))
                {
                    // But track function return assignments
                    if (isFunction && currentLine != lineNumber)
                    {
                        RecordReference(memberReferences, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, null);
                    }
                    continue;
                }

                // Detect "As <TOKEN>" context for type/class references
                if (IsAsTypeContext(masked, tokenIdx, token))
                {
                    // Type reference
                    if (gIdx.TypeIndex.TryGetValue(token, out var typeDef))
                    {
                        typeDef.Used = true;
                        RecordReference(typeDef.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, typeDef);
                        continue;
                    }
                    // Class reference
                    if (gIdx.ClassIndex.TryGetValue(token, out var classMod) &&
                        !string.Equals(classMod.Name, mod.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        classMod.Used = true;
                        RecordReference(classMod.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, classMod);
                        continue;
                    }
                }

                // --- Priority-based classification (one branch wins) ---

                // 1. Parameter
                if (paramIndex.TryGetValue(token, out var param))
                {
                    param.Used = true;
                    RecordReference(param.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, param);
                    continue;
                }

                // 2. Local variable (skip declaration line)
                if (localVarIndex.TryGetValue(token, out var localVar))
                {
                    if (localVar.LineNumber != currentLine)
                    {
                        localVar.Used = true;
                        RecordReference(localVar.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, localVar);
                    }
                    continue;
                }

                // 3. Local constant
                if (localConstIndex.TryGetValue(token, out var localConst))
                {
                    localConst.Used = true;
                    RecordReference(localConst.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, localConst);
                    continue;
                }

                // 4. Module-level global variable (same module)
                if (globalVarModIndex.TryGetValue(token, out var globalVar))
                {
                    globalVar.Used = true;
                    RecordReference(globalVar.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, globalVar);
                    continue;
                }

                // 5. Global variable from another module
                if (gIdx.GlobalVarIndex.TryGetValue(token, out var gvList))
                {
                    // Prefer first match not in current module
                    var gv = gvList.FirstOrDefault(x =>
                        !string.Equals(x.Module, mod.Name, StringComparison.OrdinalIgnoreCase));
                    if (gv.Variable == null && gvList.Count > 0)
                        gv = gvList[0];
                    if (gv.Variable != null && !localNames.Contains(token))
                    {
                        gv.Variable.Used = true;
                        RecordReference(gv.Variable.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, gv.Variable);
                        continue;
                    }
                }

                // 6. Control (same module, bare usage like `lblTitle = "x"`)
                if (controlIndex.TryGetValue(token, out var control) && !localNames.Contains(token))
                {
                    var startChar = GetTokenStartChar(raw, token, tokenIdx);
                    MarkControlAsUsed(control, mod.Name, memberName, currentLine, startChar);
                    recorded.Add((currentLine, tokenIdx));
                    continue;
                }

                // Skip remaining lookups on declaration lines
                if (isDeclLine)
                    continue;

                // 7. Global constant
                if (gIdx.ConstantIndex.TryGetValue(token, out var constList) && !localNames.Contains(token))
                {
                    var c = constList.FirstOrDefault(x =>
                        !string.Equals(x.Module, mod.Name, StringComparison.OrdinalIgnoreCase));
                    if (c.Constant == null && constList.Count > 0)
                        c = constList[0];
                    if (c.Constant != null)
                    {
                        c.Constant.Used = true;
                        RecordReference(c.Constant.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, c.Constant);
                        continue;
                    }
                }

                // 8. Procedure (Sub/Function)
                if (gIdx.ProcIndex.TryGetValue(token, out var procTargets) && procTargets.Count > 0)
                {
                    var selected = SelectProcTarget(procTargets, env, token);
                    if (selected.Proc != null)
                    {
                        selected.Proc.Used = true;
                        RecordReference(selected.Proc.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, selected.Proc);
                        AddCallIfNew(calls, token, null, selected.Module, selected.Proc.Name, selected.Proc.Kind, null, currentLine);
                    }
                    continue;
                }

                // 9. Property (bare usage cross-module, e.g., If ExecSts = ...)
                if (gIdx.PropIndex.TryGetValue(token, out var propTargets) && propTargets.Count > 0)
                {
                    var pt = propTargets.FirstOrDefault(x =>
                        x.Prop.Kind.Equals("Get", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(x.Module, mod.Name, StringComparison.OrdinalIgnoreCase));
                    if (pt.Prop == null)
                        pt = propTargets.FirstOrDefault(x =>
                            x.Prop.Kind.Equals("Get", StringComparison.OrdinalIgnoreCase));
                    if (pt.Prop != null)
                    {
                        pt.Prop.Used = true;
                        RecordReference(pt.Prop.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, pt.Prop);
                        AddCallIfNew(calls, token, null, pt.Module, pt.Prop.Name, $"Property{pt.Prop.Kind}", null, currentLine);
                    }
                    continue;
                }

                // 10. Enum value (bare)
                if (gIdx.EnumValueIndex.TryGetValue(token, out var enumValues) && !localNames.Contains(token))
                {
                    // Filter by enum type context if available on the line
                    var filtered = FilterEnumValues(enumValues, masked, gIdx);
                    if (filtered.Count == 0 && enumValues.Count > 1)
                        continue; // ambiguous, skip

                    foreach (var ev in filtered)
                    {
                        ev.Used = true;
                        RecordReference(ev.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, ev);
                    }
                    continue;
                }

                // 11. Module name (bare usage, e.g., module name passed as value)
                if (gIdx.ModuleByName.TryGetValue(token, out var refMod))
                {
                    refMod.Used = true;
                    RecordReference(refMod.References, mod.Name, memberName, currentLine, tokenIdx, masked, recorded, refMod);
                    continue;
                }
            }
        }
    }

    // -----------------------------------------------------------------
    // DOT-CHAIN RESOLUTION
    // -----------------------------------------------------------------

    private static void ResolveChain(
        string chainText,
        int chainIndex,
        string scanLine,
        string rawLine,
        int currentLine,
        VbModule mod,
        string memberName,
        Dictionary<string, string> env,
        GlobalIndexes gIdx,
        Dictionary<string, VbControl> controlIndex,
        HashSet<string> localNames,
        Dictionary<string, VbParameter> paramIndex,
        Dictionary<string, VbVariable> localVarIndex,
        Dictionary<string, VbVariable> globalVarModIndex,
        List<VbCall>? calls,
        HashSet<(int Line, int StartChar)> recorded,
        HashSet<int> chainTokensClaimed)
    {
        var parts = SplitChainParts(chainText);
        if (parts.Length < 2) return;

        // Only depth-0 tokens: chain-structural identifiers (base, field names after dots).
        // Tokens inside parenthesized arguments (e.g., UBound, inner variables) are left
        // unclaimed so the bare-token scan at STEP 2 can resolve them independently.
        var tokenPositions = GetDepthZeroTokenPositions(chainText, chainIndex);

        // Claim structural token positions in this chain
        foreach (var (_, pos) in tokenPositions)
            chainTokensClaimed.Add(pos);

        // --- Base variable ---
        var baseVarName = parts[0];
        var parenIdx = baseVarName.IndexOf('(');
        if (parenIdx >= 0) baseVarName = baseVarName.Substring(0, parenIdx);

        var baseTokenPos = tokenPositions.FirstOrDefault();

        // --- Check for qualified enum reference: Enum.Value ---
        if (parts.Length == 2 && gIdx.EnumDefIndex.TryGetValue(baseVarName, out var enumDefs))
        {
            var valueName = parts[1];
            var valueParenIdx = valueName.IndexOf('(');
            if (valueParenIdx >= 0) valueName = valueName.Substring(0, valueParenIdx);

            foreach (var enumDef in enumDefs)
            {
                enumDef.Used = true;
                RecordReference(enumDef.References, mod.Name, memberName, currentLine, baseTokenPos.Item2, scanLine, recorded, enumDef);

                var enumValue = enumDef.Values.FirstOrDefault(v =>
                    v.Name.Equals(valueName, StringComparison.OrdinalIgnoreCase));
                if (enumValue != null)
                {
                    enumValue.Used = true;
                    var valueTokenPos = tokenPositions.ElementAtOrDefault(1);
                    RecordReference(enumValue.References, mod.Name, memberName, currentLine, valueTokenPos.Item2, scanLine, recorded, enumValue);
                }
            }
            return; // Fully handled as Enum.Value
        }

        // Record reference for base variable
        if (!string.IsNullOrEmpty(baseVarName))
        {
            if (paramIndex.TryGetValue(baseVarName, out var paramRef))
            {
                paramRef.Used = true;
                RecordReference(paramRef.References, mod.Name, memberName, currentLine, baseTokenPos.Item2, scanLine, recorded, paramRef);
            }
            else if (localVarIndex.TryGetValue(baseVarName, out var localRef) && localRef.LineNumber != currentLine)
            {
                localRef.Used = true;
                RecordReference(localRef.References, mod.Name, memberName, currentLine, baseTokenPos.Item2, scanLine, recorded, localRef);
            }
            else if (globalVarModIndex.TryGetValue(baseVarName, out var globalRef))
            {
                globalRef.Used = true;
                RecordReference(globalRef.References, mod.Name, memberName, currentLine, baseTokenPos.Item2, scanLine, recorded, globalRef);
            }
            else if (gIdx.GlobalVarIndex.TryGetValue(baseVarName, out var gvList) && !localNames.Contains(baseVarName))
            {
                var gv = gvList.FirstOrDefault(x => !string.Equals(x.Module, mod.Name, StringComparison.OrdinalIgnoreCase));
                if (gv.Variable == null && gvList.Count > 0) gv = gvList[0];
                if (gv.Variable != null)
                {
                    gv.Variable.Used = true;
                    RecordReference(gv.Variable.References, mod.Name, memberName, currentLine, baseTokenPos.Item2, scanLine, recorded, gv.Variable);
                }
            }
            else if (controlIndex.TryGetValue(baseVarName, out var baseControl) && !localNames.Contains(baseVarName))
            {
                var sc = GetTokenStartChar(rawLine, baseVarName, baseTokenPos.Item2);
                MarkControlAsUsed(baseControl, mod.Name, memberName, currentLine, sc);
                recorded.Add((currentLine, baseTokenPos.Item2));
            }
            else if (gIdx.ModuleByName.TryGetValue(baseVarName, out var baseMod))
            {
                // Module-qualified access (e.g., FrmRestart.Show)
                baseMod.Used = true;
                RecordReference(baseMod.References, mod.Name, memberName, currentLine, baseTokenPos.Item2, scanLine, recorded, baseMod);
            }
        }

        // --- Resolve chain type ---
        string? typeName = null;
        int startPartIndex = 1;

        // Try resolve base from env (variable → type)
        if (!env.TryGetValue(baseVarName, out typeName) || string.IsNullOrEmpty(typeName))
        {
            // Try resolve as module-qualified access: Module.GlobalVar / Module.Property
            var moduleMatch = mod.Owner?.Modules?.FirstOrDefault(m =>
                m.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));

            if (moduleMatch != null && parts.Length > 1)
            {
                var member = parts[1];
                var memberParen = member.IndexOf('(');
                if (memberParen >= 0) member = member.Substring(0, memberParen);

                var gVar = moduleMatch.GlobalVariables.FirstOrDefault(v =>
                    v.Name.Equals(member, StringComparison.OrdinalIgnoreCase));

                if (gVar != null && !string.IsNullOrEmpty(gVar.Type))
                {
                    typeName = gVar.Type;
                    // Record reference for the member token
                    gVar.Used = true;
                    var memberTokenPos = tokenPositions.ElementAtOrDefault(1);
                    RecordReference(gVar.References, mod.Name, memberName, currentLine, memberTokenPos.Item2, scanLine, recorded, gVar);
                    startPartIndex = 2;
                }
                else
                {
                    var mProp = moduleMatch.Properties.FirstOrDefault(p =>
                        p.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
                    if (mProp != null && !string.IsNullOrEmpty(mProp.ReturnType))
                    {
                        typeName = mProp.ReturnType;
                        mProp.Used = true;
                        var memberTokenPos = tokenPositions.ElementAtOrDefault(1);
                        RecordReference(mProp.References, mod.Name, memberName, currentLine, memberTokenPos.Item2, scanLine, recorded, mProp);
                        startPartIndex = 2;
                    }
                    else
                    {
                        // Could be Module.Procedure or Module.Constant — record references
                        var mProc = moduleMatch.Procedures.FirstOrDefault(p =>
                            p.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
                        if (mProc != null)
                        {
                            mProc.Used = true;
                            var memberTokenPos = tokenPositions.ElementAtOrDefault(1);
                            RecordReference(mProc.References, mod.Name, memberName, currentLine, memberTokenPos.Item2, scanLine, recorded, mProc);
                            AddCallIfNew(calls, $"{baseVarName}.{member}", baseVarName, moduleMatch.Name, mProc.Name, mProc.Kind, null, currentLine);
                        }
                        else
                        {
                            // Module.Constant
                            var mConst = moduleMatch.Constants.FirstOrDefault(c =>
                                c.Name != null && c.Name.Equals(member, StringComparison.OrdinalIgnoreCase));
                            if (mConst != null)
                            {
                                mConst.Used = true;
                                var memberTokenPos = tokenPositions.ElementAtOrDefault(1);
                                RecordReference(mConst.References, mod.Name, memberName, currentLine, memberTokenPos.Item2, scanLine, recorded, mConst);
                            }
                        }
                        // No further chain traversal for procedures or constants
                        return;
                    }
                }
            }
        }

        if (string.IsNullOrEmpty(typeName))
            return;

        // Check if base type is internal (project-defined)
        var initialType = NormalizeModuleTypeName(typeName);
        bool isInternal = gIdx.TypeIndex.ContainsKey(initialType) || gIdx.ClassIndex.ContainsKey(initialType);
        if (!isInternal)
            return;

        // --- Walk the chain ---
        for (int pi = startPartIndex; pi < parts.Length; pi++)
        {
            var fieldName = parts[pi];
            var fieldParen = fieldName.IndexOf('(');
            if (fieldParen >= 0) fieldName = fieldName.Substring(0, fieldParen);
            if (string.IsNullOrEmpty(fieldName)) break;

            if (string.IsNullOrEmpty(typeName))
            {
                // Unknown type from previous step: fallback search in all types
                var found = FallbackFieldSearch(fieldName, pi, tokenPositions, scanLine, currentLine, mod.Name, memberName, gIdx.TypeIndex, recorded);
                if (!found) break;
                typeName = null; // can't continue chain without type
                break;
            }

            var baseType = NormalizeModuleTypeName(typeName);

            // Try as class member
            if (gIdx.ClassIndex.TryGetValue(baseType, out var classModule))
            {
                var classProp = classModule.Properties.FirstOrDefault(p =>
                    MatchesName(p.Name, p.ConventionalName, fieldName));
                if (classProp != null)
                {
                    classProp.Used = true;
                    var tp = tokenPositions.ElementAtOrDefault(pi);
                    RecordReference(classProp.References, mod.Name, memberName, currentLine, tp.Item2, scanLine, recorded, classProp);
                    AddCallIfNew(calls, $"{baseVarName}.{fieldName}", baseVarName, classModule.Name, classProp.Name, $"Property{classProp.Kind}", baseType, currentLine);
                    typeName = classProp.ReturnType;
                    if (string.IsNullOrEmpty(typeName)) break;
                    continue;
                }

                var classProc = classModule.Procedures.FirstOrDefault(p =>
                    MatchesName(p.Name, p.ConventionalName, fieldName));
                if (classProc != null)
                {
                    classProc.Used = true;
                    var tp = tokenPositions.ElementAtOrDefault(pi);
                    RecordReference(classProc.References, mod.Name, memberName, currentLine, tp.Item2, scanLine, recorded, classProc);
                    AddCallIfNew(calls, $"{baseVarName}.{fieldName}", baseVarName, classModule.Name, classProc.Name, classProc.Kind, baseType, currentLine);
                    typeName = classProc.ReturnType;
                    if (string.IsNullOrEmpty(typeName)) typeName = null;
                    continue;
                }

                // Member not found in class
                typeName = null;
                continue;
            }

            // Try as UDT field
            if (!gIdx.TypeIndex.TryGetValue(baseType, out var typeDef))
            {
                // Fallback: search field in all known types
                FallbackFieldSearch(fieldName, pi, tokenPositions, scanLine, currentLine, mod.Name, memberName, gIdx.TypeIndex, recorded);
                typeName = null;
                break;
            }

            var field = typeDef.Fields.FirstOrDefault(f =>
                !string.IsNullOrEmpty(f.Name) && MatchesName(f.Name, f.ConventionalName, fieldName));
            if (field == null) break;

            var tokenPos = tokenPositions.ElementAtOrDefault(pi);
            field.Used = true;
            RecordReference(field.References, mod.Name, memberName, currentLine, tokenPos.Item2, scanLine, recorded, field);
            typeName = field.Type;
            if (string.IsNullOrEmpty(typeName)) break;
        }
    }

    // -----------------------------------------------------------------
    // HELPERS
    // -----------------------------------------------------------------

    private static void RecordReference(
        List<VbReference> references,
        string module,
        string procedure,
        int lineNumber,
        int startChar,
        string scanLine,
        HashSet<(int Line, int StartChar)> recorded,
        object? owner)
    {
        if (startChar < 0 || !recorded.Add((lineNumber, startChar)))
            return;

        references.AddLineNumber(module, procedure, lineNumber, startChar, owner: owner);
    }

    /// <summary>Extracts the identifier token starting at the given index in the line.</summary>
    private static string ExtractTokenAt(string line, int index)
    {
        if (index < 0 || index >= line.Length) return string.Empty;
        int end = index;
        while (end < line.Length && IsIdentifierChar(line[end])) end++;
        return line.Substring(index, end - index);
    }

    /// <summary>Checks if the token at the given position is preceded by "As " (possibly with New).</summary>
    private static bool IsAsTypeContext(string line, int tokenIndex, string token)
    {
        // Look backward for "As " or "As New "
        var before = line.Substring(0, tokenIndex).TrimEnd();
        return before.EndsWith("As", StringComparison.OrdinalIgnoreCase) ||
               before.EndsWith("As New", StringComparison.OrdinalIgnoreCase);
    }

    private static (string Module, VbProcedure Proc) SelectProcTarget(
        List<(string Module, VbProcedure Proc)> targets,
        Dictionary<string, string> env,
        string token)
    {
        if (env.TryGetValue(token, out var resolvedType) && !string.IsNullOrEmpty(resolvedType))
        {
            var match = targets.FirstOrDefault(t =>
                Path.GetFileNameWithoutExtension(t.Module)
                    .Equals(resolvedType, StringComparison.OrdinalIgnoreCase));
            if (match.Proc != null) return match;
        }
        return targets[0];
    }

    private static void AddCallIfNew(
        List<VbCall>? calls,
        string raw,
        string? objectName,
        string resolvedModule,
        string resolvedProcedure,
        string? resolvedKind,
        string? resolvedType,
        int lineNumber)
    {
        if (calls == null) return;
        var exists = calls.Any(c => string.Equals(c.Raw, raw, StringComparison.OrdinalIgnoreCase));
        if (!exists)
        {
            calls.Add(new VbCall
            {
                Raw = raw,
                ObjectName = objectName,
                MethodName = resolvedProcedure,
                ResolvedModule = resolvedModule,
                ResolvedProcedure = resolvedProcedure,
                ResolvedKind = resolvedKind,
                ResolvedType = resolvedType,
                LineNumber = lineNumber
            });
        }
    }

    private static List<VbEnumValue> FilterEnumValues(
        List<VbEnumValue> values,
        string line,
        GlobalIndexes gIdx)
    {
        if (values.Count <= 1) return values;

        // Try to narrow by "As EnumType" on the same line
        var enumTypeDefsOnLine = new HashSet<VbEnumDef>();
        foreach (Match m in Regex.Matches(line, @"\bAs\s+([A-Za-z_]\w*)", RegexOptions.IgnoreCase))
        {
            var typeToken = m.Groups[1].Value;
            if (gIdx.EnumDefIndex.TryGetValue(typeToken, out var defs))
                foreach (var d in defs) enumTypeDefsOnLine.Add(d);
        }

        if (enumTypeDefsOnLine.Count > 0)
        {
            var filtered = values.Where(v =>
                gIdx.EnumValueOwners.TryGetValue(v, out var owner) &&
                enumTypeDefsOnLine.Contains(owner)).ToList();
            if (filtered.Count > 0) return filtered;
        }

        // Try to narrow by any enum name present on the line
        var enumDefsOnLine = new HashSet<VbEnumDef>();
        foreach (var (token, _) in EnumerateTokens(line))
        {
            if (gIdx.EnumDefIndex.TryGetValue(token, out var defs))
                foreach (var d in defs) enumDefsOnLine.Add(d);
        }

        if (enumDefsOnLine.Count > 0)
        {
            var filtered = values.Where(v =>
                gIdx.EnumValueOwners.TryGetValue(v, out var owner) &&
                enumDefsOnLine.Contains(owner)).ToList();
            if (filtered.Count > 0) return filtered;
        }

        return values;
    }

    private static bool FallbackFieldSearch(
        string fieldName,
        int partIndex,
        List<(string Value, int Index)> tokenPositions,
        string scanLine,
        int currentLine,
        string moduleName,
        string memberName,
        Dictionary<string, VbTypeDef> typeIndex,
        HashSet<(int Line, int StartChar)> recorded)
    {
        foreach (var (_, typeDef) in typeIndex)
        {
            var field = typeDef.Fields.FirstOrDefault(f =>
                !string.IsNullOrEmpty(f.Name) &&
                MatchesName(f.Name, f.ConventionalName, fieldName));
            if (field != null)
            {
                var tp = tokenPositions.ElementAtOrDefault(partIndex);
                field.Used = true;
                RecordReference(field.References, moduleName, memberName, currentLine, tp.Item2, scanLine, recorded, field);
                return true;
            }
        }
        return false;
    }
}
