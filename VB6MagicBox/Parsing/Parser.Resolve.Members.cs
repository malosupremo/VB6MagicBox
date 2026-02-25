using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    // ---------------------------------------------------------
    // REGEX COMPILATE PER HOT-PATH (ResolveFieldAccesses, ResolveControlAccesses, etc.)
    // ---------------------------------------------------------

    private static readonly Regex ReWithDotReplacement = 
        new(@"(?<![\w)])\.(\s*[A-Za-z_]\w*(?:\([^)]*\))?)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReChainPattern = 
        new(@"([A-Za-z_]\w*(?:\([^)]*\))?)(?:\s*\.\s*[A-Za-z_]\w*(?:\([^)]*\))?)+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReTokens = 
        new(@"\b[A-Za-z_]\w*\b", RegexOptions.Compiled);

    private static readonly Regex ReControlAccess = 
        new(@"(\w+)(?:\([^\)]*\))?\.(\w+)", RegexOptions.Compiled);

    private static readonly Regex ReControlAccessCrossModule = 
        new(@"(\w+)\.(\w+)(?:\([^\)]*\))?\.(\w+)", RegexOptions.Compiled);

    // ---------------------------------------------------------
    // RISOLUZIONE ACCESSI AI CAMPI
    // ---------------------------------------------------------

    private static void ResolveFieldAccesses(
        VbModule mod,
        VbProcedure proc,
        string[] fileLines,
        Dictionary<string, VbTypeDef> typeIndex,
        Dictionary<string, string> env,
        Dictionary<string, VbModule> classIndex)
    {
        // Controlli di sicurezza per evitare IndexOutOfRangeException
        if (proc.StartLine <= 0)
        {
            Console.WriteLine($"[WARN] Procedure {proc.Name} has invalid StartLine: {proc.StartLine}, using LineNumber: {proc.LineNumber}");
            proc.StartLine = proc.LineNumber;
        }

        if (proc.EndLine <= 0)
        {
            Console.WriteLine($"[WARN] Procedure {proc.Name} has invalid EndLine: {proc.EndLine}, scanning until end of file");
            proc.EndLine = fileLines.Length;
        }

        var startIndex = Math.Max(0, proc.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, proc.EndLine);

        if (startIndex >= fileLines.Length)
        {
            Console.WriteLine($"[WARN] Procedure {proc.Name} StartLine {proc.StartLine} is beyond file length {fileLines.Length}");
            return;
        }

        var withStack = new Stack<string>();

        for (int i = startIndex; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();

            // Rimuovi commenti (ignorando apostrofi dentro stringhe)
            var noComment = StripInlineComment(raw).Trim();

            var trimmedNoComment = noComment.TrimStart();
            if (trimmedNoComment.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
            {
                var withExpr = trimmedNoComment.Substring(5).Trim();
                if (withExpr.StartsWith(".") && withStack.Count > 0)
                    withExpr = withStack.Peek() + withExpr;
                if (!string.IsNullOrEmpty(withExpr))
                    withStack.Push(withExpr);
                continue;
            }

            if (trimmedNoComment.Equals("End With", StringComparison.OrdinalIgnoreCase))
            {
                if (withStack.Count > 0)
                    withStack.Pop();
                continue;
            }

            if (withStack.Count > 0 && trimmedNoComment.StartsWith(".", StringComparison.Ordinal))
            {
                var suffix = trimmedNoComment.Substring(1).TrimStart();
                noComment = withStack.Peek() + "." + suffix;
            }

            if (withStack.Count > 0)
            {
                var withPrefix = withStack.Peek();
                noComment = ReWithDotReplacement.Replace(noComment,
                    m => withPrefix + "." + m.Groups[1].Value);
            }

            var scanLine = MaskStringLiterals(noComment);
            if (scanLine.IndexOf('.', StringComparison.Ordinal) < 0)
                continue;
            var chainMatches = new List<(string Text, int Index)>();

            foreach (Match m in ReChainPattern.Matches(scanLine))
                chainMatches.Add((m.Value, m.Index));

            foreach (var (innerText, innerStart) in EnumerateParenContents(scanLine))
            {
                foreach (Match m in ReChainPattern.Matches(innerText))
                    chainMatches.Add((m.Value, innerStart + m.Index));
            }

            foreach (var (chainText, chainIndex) in chainMatches)
            {
                if (TryUnwrapFunctionChain(chainText, chainIndex, out var unwrappedChain, out var unwrappedIndex))
                { }

                var effectiveChain = unwrappedChain ?? chainText;
                var effectiveIndex = unwrappedChain != null ? unwrappedIndex : chainIndex;

                var parts = SplitChainParts(effectiveChain);

                if (parts.Length < 2)
                    continue;

                var tokenMatches = ReTokens.Matches(effectiveChain);
                var tokenPositions = tokenMatches.Select(m => (m.Value, effectiveIndex + m.Index)).ToList();

                var baseVarName = parts[0];
                var parenIndex = baseVarName.IndexOf('(');
                if (parenIndex >= 0)
                    baseVarName = baseVarName.Substring(0, parenIndex);

                var baseTokenPosition = tokenPositions.FirstOrDefault();
                if (!string.IsNullOrEmpty(baseVarName))
                {
                    var baseOccIdx = GetOccurrenceIndex(scanLine, baseVarName, baseTokenPosition.Item2, i + 1);

                    var paramRef = proc.Parameters.FirstOrDefault(p =>
                        p.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));
                    if (paramRef != null)
                    {
                        paramRef.Used = true;
                        paramRef.References.AddLineNumber(mod.Name, proc.Name, i + 1, baseOccIdx);
                    }

                    var localRef = proc.LocalVariables.FirstOrDefault(v =>
                        v.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));
                    if (localRef != null && localRef.LineNumber != i + 1)
                    {
                        localRef.Used = true;
                        localRef.References.AddLineNumber(mod.Name, proc.Name, i + 1, baseOccIdx);
                    }

                    var globalRef = mod.GlobalVariables.FirstOrDefault(v =>
                        v.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));
                    if (globalRef != null)
                    {
                        globalRef.Used = true;
                        globalRef.References.AddLineNumber(mod.Name, proc.Name, i + 1, baseOccIdx);
                    }
                }

                string typeName = null;
                var startPartIndex = 1;

                if (!env.TryGetValue(baseVarName, out typeName) || string.IsNullOrEmpty(typeName))
                {
                    var moduleMatch = mod.Owner?.Modules?.FirstOrDefault(m =>
                        m.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));

                    if (moduleMatch != null && parts.Length > 1)
                    {
                        var memberName = parts[1];
                        var memberParenIndex = memberName.IndexOf('(');
                        if (memberParenIndex >= 0)
                            memberName = memberName.Substring(0, memberParenIndex);

                        var globalVar = moduleMatch.GlobalVariables.FirstOrDefault(v =>
                            v.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase));

                        if (globalVar != null && !string.IsNullOrEmpty(globalVar.Type))
                        {
                            typeName = globalVar.Type;
                            startPartIndex = 2;
                        }
                        else
                        {
                            var prop = moduleMatch.Properties.FirstOrDefault(p =>
                                p.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase));
                            if (prop != null && !string.IsNullOrEmpty(prop.ReturnType))
                            {
                                typeName = prop.ReturnType;
                                startPartIndex = 2;
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(typeName))
                    continue;

                // Verifica se il tipo base � nel progetto (interno) o esterno
                // Se � esterno, NON tracciare References per nessun membro della catena
                var initialTypeName = typeName;
                if (initialTypeName.Contains('('))
                    initialTypeName = initialTypeName.Substring(0, initialTypeName.IndexOf('('));
                if (initialTypeName.Contains('.'))
                    initialTypeName = initialTypeName.Split('.').Last();

                bool isInternalType = typeIndex.ContainsKey(initialTypeName) || classIndex.ContainsKey(initialTypeName);

                // Se il tipo � esterno (non nel progetto), interrompi la catena
                // Es: gobjFM489.ActualState.Program.Frequency_Long dove gobjFM489 � tipo esterno
                if (!isInternalType)
                    continue;

                for (int partIndex = startPartIndex; partIndex < parts.Length; partIndex++)
                {
                    var fieldName = parts[partIndex];
                    var fieldParenIndex = fieldName.IndexOf('(');
                    if (fieldParenIndex >= 0)
                        fieldName = fieldName.Substring(0, fieldParenIndex);

                    if (string.IsNullOrEmpty(fieldName))
                        break;

                    // Se typeName � null (tipo di ritorno sconosciuto dal passo precedente,
                    // es. una Function di classe il cui ReturnType non � risolto),
                    // cerca il campo corrente in tutti i tipi interni noti.
                    if (string.IsNullOrEmpty(typeName))
                    {
                        var chainFallbackFound = false;
                        foreach (var (anyTypeName, anyTypeDef) in typeIndex)
                        {
                            var anyField = anyTypeDef.Fields.FirstOrDefault(f =>
                                !string.IsNullOrEmpty(f.Name) &&
                                string.Equals(f.Name, fieldName, StringComparison.OrdinalIgnoreCase));
                            if (anyField != null)
                            {
                                var tp = tokenPositions.Skip(partIndex).FirstOrDefault(t =>
                                    t.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                                var oi = GetOccurrenceIndex(scanLine, fieldName, tp.Item2, i + 1);
                                
                                anyField.Used = true;
                                anyField.References.AddLineNumber(mod.Name, proc.Name, i + 1, oi);
                                typeName = anyField.Type;
                                chainFallbackFound = true;
                                break;
                            }
                        }
                        if (!chainFallbackFound || string.IsNullOrEmpty(typeName))
                            break;
                        continue;
                    }

                    var baseTypeName = typeName;
                    if (baseTypeName.Contains('('))
                        baseTypeName = baseTypeName.Substring(0, baseTypeName.IndexOf('('));
                    if (baseTypeName.Contains('.'))
                        baseTypeName = baseTypeName.Split('.').Last();

                    if (classIndex.TryGetValue(baseTypeName, out var classModule))
                    {
                        // Cerca prima nelle propriet� della classe
                        var classProp = classModule.Properties.FirstOrDefault(p =>
                            MatchesName(p.Name, p.ConventionalName, fieldName));
                        if (classProp != null)
                        {
                            classProp.Used = true;
                            classProp.References.AddLineNumber(mod.Name, proc.Name, i + 1);
                            typeName = classProp.ReturnType;
                            if (string.IsNullOrEmpty(typeName))
                                break;
                            continue;
                        }

                        // Cerca anche tra le funzioni/procedure (es. Item(i) � una Function che ritorna un tipo)
                        var classFunc = classModule.Procedures.FirstOrDefault(p =>
                            MatchesName(p.Name, p.ConventionalName, fieldName));
                        if (classFunc != null)
                        {
                            typeName = classFunc.ReturnType;
                            if (string.IsNullOrEmpty(typeName))
                                typeName = null; // ReturnType sconosciuto: la prossima iterazione user� il fallback
                            continue;
                        }

                        // Membro non trovato nella classe: tipo sconosciuto per i segmenti successivi
                        typeName = null;
                        continue;
                    }

                    if (!typeIndex.TryGetValue(baseTypeName, out var typeDef))
                    {
                        // Tipo base non � n� classe n� UDT noto: cerca il campo in tutti i tipi interni
                        var fieldFoundInAnyType = false;
                        foreach (var (anyTypeName, anyTypeDef) in typeIndex)
                        {
                            var anyField = anyTypeDef.Fields.FirstOrDefault(f =>
                                !string.IsNullOrEmpty(f.Name) &&
                                MatchesName(f.Name, f.ConventionalName, fieldName));

                            if (anyField != null)
                            {
                                var fallbackTokenPos = tokenPositions.Skip(partIndex).FirstOrDefault(t =>
                                    t.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                                var fallbackOccIdx = GetOccurrenceIndex(noComment, fieldName, fallbackTokenPos.Item2, i + 1);
                                
                                anyField.Used = true;
                                anyField.References.AddLineNumber(mod.Name, proc.Name, i + 1, fallbackOccIdx);
                                fieldFoundInAnyType = true;
                                break;
                            }
                        }
                        if (!fieldFoundInAnyType)
                            break;
                        typeName = null;
                        break;
                    }

                    var field = typeDef.Fields.FirstOrDefault(f =>
                        !string.IsNullOrEmpty(f.Name) &&
                        MatchesName(f.Name, f.ConventionalName, fieldName));

                    if (field == null)
                        break;

                    var tokenPosition = tokenPositions.Skip(partIndex).FirstOrDefault(t =>
                        t.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                    var occurrenceIndex = GetOccurrenceIndex(scanLine, fieldName, tokenPosition.Item2, i + 1);

                    

                    field.Used = true;
                    field.References.AddLineNumber(mod.Name, proc.Name, i + 1, occurrenceIndex);
                    typeName = field.Type;

                    if (string.IsNullOrEmpty(typeName))
                        break;
                }
            }
        }
    }

    private static void ResolveFieldAccesses(
        VbModule mod,
        VbProperty prop,
        string[] fileLines,
        Dictionary<string, VbTypeDef> typeIndex,
        Dictionary<string, string> env,
        Dictionary<string, VbModule> classIndex)
    {
        if (prop.StartLine <= 0)
            prop.StartLine = prop.LineNumber;

        if (prop.EndLine <= 0)
            prop.EndLine = fileLines.Length;

        var startIndex = Math.Max(0, prop.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, prop.EndLine);

        if (startIndex >= fileLines.Length)
            return;

        var withStack = new Stack<string>();

        for (int i = startIndex; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();

            // Rimuovi commenti (ignorando apostrofi dentro stringhe)
            var noComment = StripInlineComment(raw).Trim();

            var trimmedNoComment = noComment.TrimStart();
            if (trimmedNoComment.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
            {
                var withExpr = trimmedNoComment.Substring(5).Trim();
                if (withExpr.StartsWith(".") && withStack.Count > 0)
                    withExpr = withStack.Peek() + withExpr;
                if (!string.IsNullOrEmpty(withExpr))
                    withStack.Push(withExpr);
                continue;
            }

            if (trimmedNoComment.Equals("End With", StringComparison.OrdinalIgnoreCase))
            {
                if (withStack.Count > 0)
                    withStack.Pop();
                continue;
            }

            if (withStack.Count > 0 && trimmedNoComment.StartsWith(".", StringComparison.Ordinal))
            {
                var suffix = trimmedNoComment.Substring(1).TrimStart();
                noComment = withStack.Peek() + "." + suffix;
            }

            if (withStack.Count > 0)
            {
                var withPrefix = withStack.Peek();
                noComment = ReWithDotReplacement.Replace(noComment,
                    m => withPrefix + "." + m.Groups[1].Value);
            }

            var scanLine = MaskStringLiterals(noComment);
            if (scanLine.IndexOf('.', StringComparison.Ordinal) < 0)
                continue;
            var chainMatches = new List<(string Text, int Index)>();

            foreach (Match m in ReChainPattern.Matches(scanLine))
                chainMatches.Add((m.Value, m.Index));

            foreach (var (innerText, innerStart) in EnumerateParenContents(scanLine))
            {
                foreach (Match m in ReChainPattern.Matches(innerText))
                    chainMatches.Add((m.Value, innerStart + m.Index));
            }

            foreach (var (chainText, chainIndex) in chainMatches)
            {
                if (TryUnwrapFunctionChain(chainText, chainIndex, out var unwrappedChain, out var unwrappedIndex))
                { }

                var effectiveChain = unwrappedChain ?? chainText;
                var effectiveIndex = unwrappedChain != null ? unwrappedIndex : chainIndex;

                var parts = SplitChainParts(effectiveChain);

                if (parts.Length < 2)
                    continue;

                var tokenMatches = ReTokens.Matches(effectiveChain);
                var tokenPositions = tokenMatches.Select(m => (m.Value, effectiveIndex + m.Index)).ToList();

                var baseVarName = parts[0];
                var parenIndex = baseVarName.IndexOf('(');
                if (parenIndex >= 0)
                    baseVarName = baseVarName.Substring(0, parenIndex);

                var baseTokenPosition = tokenPositions.FirstOrDefault();
                if (!string.IsNullOrEmpty(baseVarName))
                {
                    var baseOccIdx = GetOccurrenceIndex(scanLine, baseVarName, baseTokenPosition.Item2, i + 1);

                    var paramRef = prop.Parameters.FirstOrDefault(p =>
                        p.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));
                    if (paramRef != null)
                    {
                        paramRef.Used = true;
                        paramRef.References.AddLineNumber(mod.Name, prop.Name, i + 1, baseOccIdx);
                    }

                    var globalRef = mod.GlobalVariables.FirstOrDefault(v =>
                        v.Name.Equals(baseVarName, StringComparison.OrdinalIgnoreCase));
                    if (globalRef != null)
                    {
                        globalRef.Used = true;
                        globalRef.References.AddLineNumber(mod.Name, prop.Name, i + 1, baseOccIdx);
                    }
                }

                if (!env.TryGetValue(baseVarName, out var typeName) || string.IsNullOrEmpty(typeName))
                    continue;

                // Verifica se il tipo base � nel progetto (interno) o esterno
                // Se � esterno, NON tracciare References per nessun membro della catena
                var initialTypeName = typeName;
                if (initialTypeName.Contains('('))
                    initialTypeName = initialTypeName.Substring(0, initialTypeName.IndexOf('('));
                if (initialTypeName.Contains('.'))
                    initialTypeName = initialTypeName.Split('.').Last();

                bool isInternalType = typeIndex.ContainsKey(initialTypeName) || classIndex.ContainsKey(initialTypeName);

                // Se il tipo � esterno (non nel progetto), interrompi la catena
                if (!isInternalType)
                    continue;

                for (int partIndex = 1; partIndex < parts.Length; partIndex++)
                {
                    var fieldName = parts[partIndex];
                    var fieldParenIndex = fieldName.IndexOf('(');
                    if (fieldParenIndex >= 0)
                        fieldName = fieldName.Substring(0, fieldParenIndex);

                    if (string.IsNullOrEmpty(fieldName))
                        break;

                    // Se typeName � null (tipo di ritorno sconosciuto dal passo precedente),
                    // cerca il campo corrente in tutti i tipi interni noti.
                    if (string.IsNullOrEmpty(typeName))
                    {
                        var chainFallbackFound = false;
                        foreach (var (anyTypeName, anyTypeDef) in typeIndex)
                        {
                            var anyField = anyTypeDef.Fields.FirstOrDefault(f =>
                                !string.IsNullOrEmpty(f.Name) &&
                                string.Equals(f.Name, fieldName, StringComparison.OrdinalIgnoreCase));
                            if (anyField != null)
                            {
                                var tp = tokenPositions.Skip(partIndex).FirstOrDefault(t =>
                                    t.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                                var oi = GetOccurrenceIndex(scanLine, fieldName, tp.Item2, i + 1);
                                anyField.Used = true;
                                anyField.References.AddLineNumber(mod.Name, prop.Name, i + 1, oi);
                                typeName = anyField.Type;
                                chainFallbackFound = true;
                                break;
                            }
                        }
                        if (!chainFallbackFound || string.IsNullOrEmpty(typeName))
                            break;
                        continue;
                    }

                    var baseTypeName = typeName;
                    if (baseTypeName.Contains('('))
                        baseTypeName = baseTypeName.Substring(0, baseTypeName.IndexOf('('));
                    if (baseTypeName.Contains('.'))
                        baseTypeName = baseTypeName.Split('.').Last();

                    if (classIndex.TryGetValue(baseTypeName, out var classModule))
                    {
                        // Cerca prima nelle propriet� della classe
                        var classProp = classModule.Properties.FirstOrDefault(p =>
                            MatchesName(p.Name, p.ConventionalName, fieldName));
                        if (classProp != null)
                        {
                            classProp.Used = true;
                            classProp.References.AddLineNumber(mod.Name, prop.Name, i + 1);
                            typeName = classProp.ReturnType;
                            if (string.IsNullOrEmpty(typeName))
                                break;
                            continue;
                        }

                        // Cerca anche tra le funzioni/procedure (es. Item(i) � una Function che ritorna un tipo)
                        var classFunc = classModule.Procedures.FirstOrDefault(p =>
                            MatchesName(p.Name, p.ConventionalName, fieldName));
                        if (classFunc != null)
                        {
                            typeName = classFunc.ReturnType;
                            if (string.IsNullOrEmpty(typeName))
                                typeName = null;
                            continue;
                        }

                        // Membro non trovato nella classe: tipo sconosciuto per i segmenti successivi
                        typeName = null;
                        continue;
                    }

                    if (!typeIndex.TryGetValue(baseTypeName, out var typeDef))
                    {
                        var fieldFoundInAnyType = false;
                        foreach (var (anyTypeName, anyTypeDef) in typeIndex)
                        {
                            var anyField = anyTypeDef.Fields.FirstOrDefault(f =>
                                !string.IsNullOrEmpty(f.Name) &&
                                MatchesName(f.Name, f.ConventionalName, fieldName));
                            if (anyField != null)
                            {
                                var fallbackTokenPos = tokenPositions.Skip(partIndex).FirstOrDefault(t =>
                                    t.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                                var fallbackOccIdx = GetOccurrenceIndex(noComment, fieldName, fallbackTokenPos.Item2, i + 1);
                                anyField.Used = true;
                                anyField.References.AddLineNumber(mod.Name, prop.Name, i + 1, fallbackOccIdx);
                                fieldFoundInAnyType = true;
                                break;
                            }
                        }
                        if (!fieldFoundInAnyType)
                            break;
                        typeName = null;
                        break;
                    }

                    var field = typeDef.Fields.FirstOrDefault(f =>
                        !string.IsNullOrEmpty(f.Name) &&
                        MatchesName(f.Name, f.ConventionalName, fieldName));

                    if (field == null)
                        break;

                    var tokenPosition = tokenPositions.Skip(partIndex).FirstOrDefault(t =>
                        t.Value.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                    var occurrenceIndex = GetOccurrenceIndex(scanLine, fieldName, tokenPosition.Item2, i + 1);

                    field.Used = true;
                    field.References.AddLineNumber(mod.Name, prop.Name, i + 1, occurrenceIndex);
                    typeName = field.Type;

                    if (string.IsNullOrEmpty(typeName))
                        break;
                }
            }
        }
    }

    // ---------------------------------------------------------
    // RISOLUZIONE ACCESSI AI CONTROLLI
    // ---------------------------------------------------------

    private static void ResolveControlAccesses(
        VbModule mod,
        VbProcedure proc,
        string[] fileLines)
    {
        // Indicizza i controlli del modulo per nome (ora unificati, nessun duplicato)
        var controlIndex = mod.Controls.ToDictionary(
            c => c.Name,
            c => c,
            StringComparer.OrdinalIgnoreCase);

        var moduleByName = mod.Owner?.Modules?.ToDictionary(
            m => m.Name,
            m => m,
            StringComparer.OrdinalIgnoreCase) ?? new Dictionary<string, VbModule>(StringComparer.OrdinalIgnoreCase);

        var withStack = new Stack<string>();

        var shadowedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in proc.Parameters)
            shadowedNames.Add(p.Name);
        foreach (var lv in proc.LocalVariables)
            shadowedNames.Add(lv.Name);

        // Controlli di sicurezza per evitare IndexOutOfRangeException
        if (proc.StartLine <= 0)
            proc.StartLine = proc.LineNumber;

        if (proc.EndLine <= 0)
            proc.EndLine = fileLines.Length;

        var startIndex = Math.Max(0, proc.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, proc.EndLine);

        if (startIndex >= fileLines.Length)
            return;

        for (int i = startIndex; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();

            // Fine procedura - controlla SOLO i terminatori specifici della procedura
            // (Non usa "End " generico per evitare match con End If, End With, ecc.)
            if (i > proc.LineNumber - 1 && IsProcedureEndLine(raw, proc.Kind))
                break;

            // Rimuovi commenti
            var noComment = StripInlineComment(raw);

            var trimmedNoComment = noComment.TrimStart();
            if (trimmedNoComment.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
            {
                var withExpr = trimmedNoComment.Substring(5).Trim();
                if (withExpr.StartsWith(".") && withStack.Count > 0)
                    withExpr = withStack.Peek() + withExpr;
                if (!string.IsNullOrEmpty(withExpr))
                    withStack.Push(withExpr);
                continue;
            }

            if (trimmedNoComment.Equals("End With", StringComparison.OrdinalIgnoreCase))
            {
                if (withStack.Count > 0)
                    withStack.Pop();
                continue;
            }

            if (withStack.Count > 0 && trimmedNoComment.StartsWith(".", StringComparison.Ordinal))
            {
                var suffix = trimmedNoComment.Substring(1).TrimStart();
                noComment = withStack.Peek() + "." + suffix;
            }

            if (withStack.Count > 0)
            {
                var withPrefix = withStack.Peek();
                noComment = ReWithDotReplacement.Replace(noComment,
                    m => withPrefix + "." + m.Groups[1].Value);
            }

            noComment = MaskStringLiterals(noComment);

            foreach (var (moduleName, controlName, moduleIndex, controlNameIndex) in EnumerateControlAccesses(noComment))
            {
                if (string.IsNullOrEmpty(controlName))
                    continue;

                if (!string.IsNullOrEmpty(moduleName))
                {
                    if (TryResolveModule(moduleName, proc, mod, moduleByName, out var targetModule))
                    {
                        var controls = targetModule.Controls.Where(c =>
                            string.Equals(c.Name, controlName, StringComparison.OrdinalIgnoreCase));

                        foreach (var control in controls)
                        {
                            var occurrenceIndex = GetOccurrenceIndex(noComment, controlName, controlNameIndex, i + 1);
                            MarkControlAsUsed(control, mod.Name, proc.Name, i + 1, occurrenceIndex);
                        }

                        continue;
                    }
                }

                if (controlIndex.TryGetValue(moduleName, out var sameModuleControl))
                {
                    var occurrenceIndex = GetOccurrenceIndex(noComment, moduleName, moduleIndex, i + 1);
                    MarkControlAsUsed(sameModuleControl, mod.Name, proc.Name, i + 1, occurrenceIndex);
                }
            }
            foreach (Match match in ReTokens.Matches(noComment))
            {
                if (IsMemberAccessToken(noComment, match.Index))
                    continue;

                var token = match.Value;
                if (shadowedNames.Contains(token))
                    continue;

                if (controlIndex.TryGetValue(token, out var control))
                {
                    var occurrenceIndex = GetOccurrenceIndex(noComment, token, match.Index, i + 1);
                    MarkControlAsUsed(control, mod.Name, proc.Name, i + 1, occurrenceIndex);
                }
            }
        }
    }

    private static bool TryResolveModule(
        string moduleName,
        VbProcedure proc,
        VbModule mod,
        Dictionary<string, VbModule> moduleByName,
        out VbModule targetModule)
    {
        if (moduleByName.TryGetValue(moduleName, out targetModule))
            return true;

        string typeName = null;
        var param = proc.Parameters.FirstOrDefault(p =>
            string.Equals(p.Name, moduleName, StringComparison.OrdinalIgnoreCase));
        if (param != null)
            typeName = param.Type;

        if (string.IsNullOrEmpty(typeName))
        {
            var local = proc.LocalVariables.FirstOrDefault(v =>
                string.Equals(v.Name, moduleName, StringComparison.OrdinalIgnoreCase));
            if (local != null)
                typeName = local.Type;
        }

        if (string.IsNullOrEmpty(typeName))
        {
            var global = mod.GlobalVariables.FirstOrDefault(v =>
                string.Equals(v.Name, moduleName, StringComparison.OrdinalIgnoreCase));
            if (global != null)
                typeName = global.Type;
        }

        if (string.IsNullOrEmpty(typeName))
        {
            targetModule = null;
            return false;
        }

        var normalizedType = NormalizeModuleTypeName(typeName);
        return moduleByName.TryGetValue(normalizedType, out targetModule);
    }

    private static string NormalizeModuleTypeName(string typeName)
    {
        if (string.IsNullOrWhiteSpace(typeName))
            return string.Empty;

        var normalized = typeName;
        var parenIndex = normalized.IndexOf('(');
        if (parenIndex >= 0)
            normalized = normalized.Substring(0, parenIndex);

        if (normalized.Contains('.'))
            normalized = normalized.Split('.').Last();

        return normalized.Trim();
    }

    // ---------------------------------------------------------
    // RISOLUZIONE REFERENCE PARAMETRI E VARIABILI LOCALI
    // ---------------------------------------------------------

    private static void ResolveParameterAndLocalVariableReferences(
        VbModule mod,
        VbProcedure proc,
        string[] fileLines)
    {
        // Indicizza i parametri per nome
        var parameterIndex = proc.Parameters.ToDictionary(
            p => p.Name,
            p => p,
            StringComparer.OrdinalIgnoreCase);

        // Indicizza le variabili locali per nome
        var localVariableIndex = proc.LocalVariables.ToDictionary(
            v => v.Name,
            v => v,
            StringComparer.OrdinalIgnoreCase);

    var globalVariableIndex = mod.GlobalVariables.ToDictionary(
        v => v.Name,
        v => v,
        StringComparer.OrdinalIgnoreCase);

        // Controlli di sicurezza per evitare IndexOutOfRangeException
        if (proc.StartLine <= 0)
            proc.StartLine = proc.LineNumber;

        if (proc.EndLine <= 0)
            proc.EndLine = fileLines.Length;

        var startIndex = Math.Max(0, proc.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, proc.EndLine);

        if (startIndex >= fileLines.Length)
            return;

        bool trackReturn = proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase) &&
                           !string.IsNullOrEmpty(proc.Name);

        for (int i = startIndex; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();
            int currentLineNumber = i + 1;

            // Fine procedura - controlla SOLO i terminatori specifici della procedura
            // (Non usa "End " generico per evitare match con End If, End With, ecc.)
            if (i > proc.LineNumber - 1 && IsProcedureEndLine(raw, proc.Kind))
                break;

            // Rimuovi commenti (ignorando apostrofi dentro stringhe)
            var noComment = StripInlineComment(raw).Trim();

            // Rimuovi stringhe per evitare di catturare nomi dentro stringhe
            noComment = MaskStringLiterals(noComment);

            // Cerca tutti i token word nella riga
            foreach (Match m in ReTokens.Matches(noComment))
            {
                var tokenName = m.Value;

                // Controlla se � un parametro
                if (parameterIndex.TryGetValue(tokenName, out var parameter))
                {
                    parameter.Used = true;
                    parameter.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
                }

                // Controlla se � una variabile locale
                if (localVariableIndex.TryGetValue(tokenName, out var localVar))
                {
                    // Esclude la riga di dichiarazione della variabile (usa direttamente LineNumber)
                    if (localVar.LineNumber == currentLineNumber)
                        continue;

                    localVar.Used = true;
                    localVar.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
                }

                // Controlla se � una variabile globale del modulo (e non � shadowed)
                if (!parameterIndex.ContainsKey(tokenName) && !localVariableIndex.ContainsKey(tokenName))
                {
                  if (globalVariableIndex.TryGetValue(tokenName, out var globalVar))
                  {
                    globalVar.Used = true;
                    globalVar.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
                  }
                }
            }

            if (trackReturn && currentLineNumber != proc.LineNumber &&
                ContainsStandaloneToken(noComment, proc.Name))
            {
                proc.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
            }
        }
    }

    private static void ResolveParameterReferences(
        VbModule mod,
        VbProperty prop,
        string[] fileLines)
    {
        if (prop.Parameters == null || prop.Parameters.Count == 0)
            return;

        var parameterIndex = prop.Parameters.ToDictionary(
            p => p.Name,
            p => p,
            StringComparer.OrdinalIgnoreCase);

        if (prop.StartLine <= 0)
            prop.StartLine = prop.LineNumber;

        if (prop.EndLine <= 0)
            prop.EndLine = fileLines.Length;

        var startIndex = Math.Max(0, prop.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, prop.EndLine);

        if (startIndex >= fileLines.Length)
            return;

        for (int i = startIndex; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();
            int currentLineNumber = i + 1;

            // Rimuovi commenti (ignorando apostrofi dentro stringhe)
            var noComment = StripInlineComment(raw).Trim();

            // Rimuovi stringhe per evitare di catturare nomi dentro stringhe
            noComment = MaskStringLiterals(noComment);

            foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\b"))
            {
                if (IsMemberAccessToken(noComment, m.Index))
                    continue;

                var tokenName = m.Groups[1].Value;

                if (parameterIndex.TryGetValue(tokenName, out var parameter))
                {
                    parameter.Used = true;
                    parameter.References.AddLineNumber(mod.Name, prop.Name, currentLineNumber);
                }
            }
        }
    }

    private static void ResolvePropertyReturnReferences(
        VbModule mod,
        VbProperty prop,
        string[] fileLines)
    {
        if (string.IsNullOrEmpty(prop.Name))
            return;

        if (prop.StartLine <= 0)
            prop.StartLine = prop.LineNumber;

        if (prop.EndLine <= 0)
            prop.EndLine = fileLines.Length;

        var startIndex = Math.Max(0, prop.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, prop.EndLine);

        if (startIndex >= fileLines.Length)
            return;

        for (int i = startIndex + 1; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();
            int currentLineNumber = i + 1;

            var noComment = StripInlineComment(raw).Trim();

            noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

            if (ContainsStandaloneToken(noComment, prop.Name))
            {
                prop.References.AddLineNumber(mod.Name, prop.Name, currentLineNumber);
            }
        }
    }

    private static void ResolveFunctionReturnReferences(
        VbModule mod,
        VbProcedure proc,
        string[] fileLines)
    {
        if (!proc.Kind.Equals("Function", StringComparison.OrdinalIgnoreCase))
            return;

        if (string.IsNullOrEmpty(proc.Name))
            return;

        if (proc.StartLine <= 0)
            proc.StartLine = proc.LineNumber;

        if (proc.EndLine <= 0)
            proc.EndLine = fileLines.Length;

        var startIndex = Math.Max(0, proc.StartLine - 1);
        var endIndex = Math.Min(fileLines.Length, proc.EndLine);

        if (startIndex >= fileLines.Length)
            return;

        for (int i = startIndex + 1; i < endIndex; i++)
        {
            var raw = fileLines[i].Trim();
            var noComment = StripInlineComment(raw).Trim();

            noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

            if (ContainsStandaloneToken(noComment, proc.Name))
            {
                proc.References.AddLineNumber(mod.Name, proc.Name, i + 1);
            }
        }
    }

    private static bool ContainsStandaloneToken(string line, string token)
    {
        if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(token))
            return false;

        int index = 0;
        while ((index = line.IndexOf(token, index, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            if (IsWordBoundary(line, index, token.Length) && !IsMemberAccessToken(line, index))
                return true;

            index += token.Length;
        }

        return false;
    }

    private static IEnumerable<(string Text, int Index)> EnumerateParenContents(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        var starts = new Stack<int>();
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '(')
            {
                starts.Push(i + 1);
                continue;
            }

            if (line[i] != ')' || starts.Count == 0)
                continue;

            var start = starts.Pop();
            if (start > i)
                continue;

            yield return (line.Substring(start, i - start), start);
        }
    }

    private static IEnumerable<(string Token, int Index)> EnumerateTokens(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        for (int i = 0; i < line.Length; i++)
        {
            if (!IsIdentifierStart(line[i]))
                continue;

            if (i > 0 && IsIdentifierChar(line[i - 1]))
                continue;

            int start = i;
            i++;
            while (i < line.Length && IsIdentifierChar(line[i]))
                i++;

            yield return (line.Substring(start, i - start), start);
            i--;
        }
    }

    private static IEnumerable<(string Left, string Right)> EnumerateQualifiedTokens(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        for (int i = 0; i < line.Length; i++)
        {
            if (!IsIdentifierStart(line[i]))
                continue;

            if (i > 0 && IsIdentifierChar(line[i - 1]))
                continue;

            int leftStart = i;
            i++;
            while (i < line.Length && IsIdentifierChar(line[i]))
                i++;

            var left = line.Substring(leftStart, i - leftStart);

            int dotIndex = i;
            while (dotIndex < line.Length && char.IsWhiteSpace(line[dotIndex]))
                dotIndex++;

            if (dotIndex >= line.Length || line[dotIndex] != '.')
            {
                i--;
                continue;
            }

            dotIndex++;
            while (dotIndex < line.Length && char.IsWhiteSpace(line[dotIndex]))
                dotIndex++;

            if (dotIndex >= line.Length || !IsIdentifierStart(line[dotIndex]))
            {
                i--;
                continue;
            }

            int rightStart = dotIndex;
            dotIndex++;
            while (dotIndex < line.Length && IsIdentifierChar(line[dotIndex]))
                dotIndex++;

            var right = line.Substring(rightStart, dotIndex - rightStart);
            yield return (left, right);
            i = dotIndex - 1;
        }
    }

    private static bool IsIdentifierStart(char value)
        => char.IsLetter(value) || value == '_';

    private static IEnumerable<(string ModuleName, string ControlName, int ModuleIndex, int ControlIndex)> EnumerateControlAccesses(string line)
    {
        if (string.IsNullOrEmpty(line))
            yield break;

        int i = 0;
        while (i < line.Length)
        {
            if (!IsIdentifierStart(line[i]) || (i > 0 && IsIdentifierChar(line[i - 1])))
            {
                i++;
                continue;
            }

            int firstStart = i;
            i++;
            while (i < line.Length && IsIdentifierChar(line[i]))
                i++;

            var first = line.Substring(firstStart, i - firstStart);
            int index = SkipWhitespace(line, i);
            index = SkipOptionalParentheses(line, index);
            index = SkipWhitespace(line, index);

            if (index >= line.Length || line[index] != '.')
                continue;

            index++;
            index = SkipWhitespace(line, index);

            if (index >= line.Length || !IsIdentifierStart(line[index]))
            {
                i = index;
                continue;
            }

            int secondStart = index;
            index++;
            while (index < line.Length && IsIdentifierChar(line[index]))
                index++;

            var second = line.Substring(secondStart, index - secondStart);
            int afterSecond = SkipWhitespace(line, index);
            int afterSecondParen = SkipOptionalParentheses(line, afterSecond);
            afterSecondParen = SkipWhitespace(line, afterSecondParen);

            if (afterSecondParen < line.Length && line[afterSecondParen] == '.')
            {
                yield return (first, second, firstStart, secondStart);
                i = afterSecondParen + 1;
            }
            else
            {
                yield return (first, second, firstStart, secondStart);
                i = index;
            }
        }
    }

    private static int SkipWhitespace(string line, int index)
    {
        while (index < line.Length && char.IsWhiteSpace(line[index]))
            index++;

        return index;
    }

    private static int SkipOptionalParentheses(string line, int index)
    {
        if (index >= line.Length || line[index] != '(')
            return index;

        int depth = 0;
        for (int i = index; i < line.Length; i++)
        {
            if (line[i] == '(')
                depth++;
            else if (line[i] == ')')
            {
                depth--;
                if (depth == 0)
                    return i + 1;
            }
        }

        return line.Length;
    }

    private static bool ContainsToken(string line, string token)
    {
        if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(token))
            return false;

        int index = 0;
        while ((index = line.IndexOf(token, index, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            if (IsWordBoundary(line, index, token.Length))
                return true;

            index += token.Length;
        }

        return false;
    }

    private static bool IsWordBoundary(string line, int index, int length)
    {
        bool startOk = index == 0 || !IsIdentifierChar(line[index - 1]);
        int endIndex = index + length;
        bool endOk = endIndex >= line.Length || !IsIdentifierChar(line[endIndex]);
        return startOk && endOk;
    }

    private static bool IsIdentifierChar(char value)
        => char.IsLetterOrDigit(value) || value == '_';

    private static bool IsMemberAccessToken(string line, int tokenIndex)
    {
        if (tokenIndex <= 0)
            return false;

        var index = tokenIndex - 1;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        return index >= 0 && line[index] == '.';
    }

    private static bool TryGetRaiseEventName(string line, out string eventName)
    {
        eventName = null;
        if (string.IsNullOrWhiteSpace(line))
            return false;

        var keywordIndex = line.IndexOf("RaiseEvent", StringComparison.OrdinalIgnoreCase);
        if (keywordIndex < 0)
            return false;

        int index = keywordIndex + "RaiseEvent".Length;
        while (index < line.Length && char.IsWhiteSpace(line[index]))
            index++;

        if (index >= line.Length || !IsIdentifierChar(line[index]))
            return false;

        int start = index;
        while (index < line.Length && IsIdentifierChar(line[index]))
            index++;

        eventName = line.Substring(start, index - start);
        return !string.IsNullOrEmpty(eventName);
    }

    // ---------------------------------------------------------
    // MARCATURA VALORI ENUM USATI
    // ---------------------------------------------------------

    private static void MarkUsedEnumValues(VbProject project, Dictionary<string, string[]> fileCache = null)
    {
        // Indicizza tutti i valori enum per nome
        var allEnumValues = new Dictionary<string, List<VbEnumValue>>(StringComparer.OrdinalIgnoreCase);

        foreach (var mod in project.Modules)
        {
            foreach (var enumDef in mod.Enums)
            {
                foreach (var enumValue in enumDef.Values)
                {
                    if (!allEnumValues.ContainsKey(enumValue.Name))
                        allEnumValues[enumValue.Name] = new List<VbEnumValue>();

                    allEnumValues[enumValue.Name].Add(enumValue);
                }
            }
        }

        // Cerca l'uso dei valori enum in tutti i moduli
        foreach (var mod in project.Modules)
        {
            var fileLines = GetFileLines(fileCache, mod);

            foreach (var proc in mod.Procedures)
            {
                // Scansiona il corpo della procedura
                for (int i = proc.LineNumber - 1; i < fileLines.Length; i++)
                {
                    var line = fileLines[i];

                    // Fine procedura
                    if (line.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
                        break;

                    // Rimuovi commenti
                    var noComment = StripInlineComment(line);
                    noComment = MaskStringLiterals(noComment);

                    // Cerca ogni valore enum nel codice
                    foreach (var kvp in allEnumValues)
                    {
                        var enumValueName = kvp.Key;
                        var enumValues = kvp.Value;

                        // Usa word boundary per evitare match parziali
                        if (ContainsToken(noComment, enumValueName))
                        {
                            // Marca tutti i valori enum con questo nome (potrebbero esserci duplicati in enum diversi)
                            foreach (var enumValue in enumValues)
                            {
                                enumValue.Used = true;
                                enumValue.References.AddLineNumber(mod.Name, proc.Name, i + 1);
                            }
                        }
                    }
                }
            }
        }
    }

    // ---------------------------------------------------------
    // MARCATURA EVENTI USATI (RaiseEvent)
    // ---------------------------------------------------------

    private static void MarkUsedEvents(VbProject project, Dictionary<string, string[]> fileCache = null)
    {
        // Indicizza tutti gli eventi per modulo
        var eventsByModule = new Dictionary<string, List<VbEvent>>(StringComparer.OrdinalIgnoreCase);

        foreach (var mod in project.Modules)
        {
            if (mod.Events.Count > 0)
                eventsByModule[mod.Name] = mod.Events;
        }

        // Cerca RaiseEvent in tutti i moduli
        foreach (var mod in project.Modules)
        {
            var fileLines = GetFileLines(fileCache, mod);

            foreach (var proc in mod.Procedures)
            {
                // Scansiona il corpo della procedura
                for (int i = proc.LineNumber - 1; i < fileLines.Length; i++)
                {
                    var line = fileLines[i];

                    // Fine procedura
                    if (line.TrimStart().StartsWith("End ", StringComparison.OrdinalIgnoreCase))
                        break;

                    // Rimuovi commenti
                    var noComment = StripInlineComment(line);
                    noComment = MaskStringLiterals(noComment);

                    // Pattern: RaiseEvent EventName o RaiseEvent EventName(params)
                    if (TryGetRaiseEventName(noComment, out var eventName))
                    {
                        // Cerca l'evento nel modulo corrente (gli eventi sono sempre locali al modulo/classe)
                        if (eventsByModule.TryGetValue(mod.Name, out var events))
                        {
                            var evt = events.FirstOrDefault(e =>
                                e.Name.Equals(eventName, StringComparison.OrdinalIgnoreCase));

                            if (evt != null)
                            {
                                evt.Used = true;
                                evt.References.AddLineNumber(mod.Name, proc.Name, i + 1);
                            }
                        }
                    }
                }
            }
        }
    }

    private static void ResolveEnumValueReferences(VbProject project, Dictionary<string, string[]> fileCache)
    {
        var enumValueIndex = project.Modules
            .SelectMany(m => m.Enums.SelectMany(e => e.Values))
            .GroupBy(v => v.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var enumDefIndex = project.Modules
            .SelectMany(m => m.Enums)
            .GroupBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        foreach (var mod in project.Modules)
        {
            var fileLines = GetFileLines(fileCache, mod);

            foreach (var proc in mod.Procedures)
            {
                if (proc.StartLine <= 0)
                    proc.StartLine = proc.LineNumber;
                if (proc.EndLine <= 0)
                    proc.EndLine = fileLines.Length;

                var startIndex = Math.Max(0, proc.StartLine - 1);
                var endIndex = Math.Min(fileLines.Length, proc.EndLine);

                if (startIndex >= fileLines.Length)
                    continue;

                var shadowedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in proc.Parameters)
                    shadowedNames.Add(p.Name);
                foreach (var lv in proc.LocalVariables)
                    shadowedNames.Add(lv.Name);
                foreach (var c in proc.Constants)
                    shadowedNames.Add(c.Name);

                for (int i = startIndex; i < endIndex; i++)
                {
                    var raw = fileLines[i];
                    var noComment = raw;
                    var idx = noComment.IndexOf("'", StringComparison.Ordinal);
                    if (idx >= 0)
                        noComment = noComment.Substring(0, idx);

                    noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

                    foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\b"))
                    {
                        var token = m.Groups[1].Value;
                        if (shadowedNames.Contains(token))
                            continue;

                        if (!enumValueIndex.TryGetValue(token, out var enumValues))
                            continue;

                        foreach (var enumValue in enumValues)
                        {
                            enumValue.Used = true;
                            enumValue.References.AddLineNumber(mod.Name, proc.Name, i + 1);
                        }
                    }

                    foreach (var (enumName, valueName) in EnumerateQualifiedTokens(noComment))
                    {
                        if (enumDefIndex.TryGetValue(enumName, out var enumDefs))
                        {
                          foreach (var enumDef in enumDefs)
                          {
                            enumDef.Used = true;
                            enumDef.References.AddLineNumber(mod.Name, proc.Name, i + 1);

                            var value = enumDef.Values.FirstOrDefault(v =>
                                v.Name.Equals(valueName, StringComparison.OrdinalIgnoreCase));
                            if (value != null)
                            {
                              value.Used = true;
                              value.References.AddLineNumber(mod.Name, proc.Name, i + 1);
                            }
                          }
                        }
                    }
                }
            }

            foreach (var prop in mod.Properties)
            {
                if (prop.StartLine <= 0)
                    prop.StartLine = prop.LineNumber;
                if (prop.EndLine <= 0)
                    prop.EndLine = fileLines.Length;

                var startIndex = Math.Max(0, prop.StartLine - 1);
                var endIndex = Math.Min(fileLines.Length, prop.EndLine);

                if (startIndex >= fileLines.Length)
                    continue;

                var shadowedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in prop.Parameters)
                    shadowedNames.Add(p.Name);

                for (int i = startIndex; i < endIndex; i++)
                {
                    var raw = fileLines[i];
                    var noComment = raw;
                    var idx = noComment.IndexOf("'", StringComparison.Ordinal);
                    if (idx >= 0)
                        noComment = noComment.Substring(0, idx);

                    noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

                    foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\b"))
                    {
                        var token = m.Groups[1].Value;
                        if (shadowedNames.Contains(token))
                            continue;

                        if (!enumValueIndex.TryGetValue(token, out var enumValues))
                            continue;

                        foreach (var enumValue in enumValues)
                        {
                          enumValue.Used = true;
                          enumValue.References.AddLineNumber(mod.Name, prop.Name, i + 1);
                        }
                    }

                    foreach (var (enumName, valueName) in EnumerateQualifiedTokens(noComment))
                    {
                        if (enumDefIndex.TryGetValue(enumName, out var enumDefs))
                        {
                          foreach (var enumDef in enumDefs)
                          {
                            enumDef.Used = true;
                            enumDef.References.AddLineNumber(mod.Name, prop.Name, i + 1);

                            var value = enumDef.Values.FirstOrDefault(v =>
                                v.Name.Equals(valueName, StringComparison.OrdinalIgnoreCase));
                            if (value != null)
                            {
                              value.Used = true;
                              value.References.AddLineNumber(mod.Name, prop.Name, i + 1);
                            }
                          }
                        }
                    }
                }
            }
        }
    }

  private static int GetOccurrenceIndex(string line, string token, int tokenIndex, int currentLineNumber = 0)
  {
        bool isDebug = false;

    if (tokenIndex < 0)
      return -1;

    var matches = Regex.Matches(line, $@"\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase);

    if (isDebug)
    {
      Console.WriteLine($"[DEBUG GetOccurrenceIndex] Line {currentLineNumber}, Token={token}, TokenIndex={tokenIndex}");
      Console.WriteLine($"[DEBUG]   Line: {line}");
      Console.WriteLine($"[DEBUG]   Total matches: {matches.Count}");
      for (int j = 0; j < matches.Count; j++)
        Console.WriteLine($"[DEBUG]     Match {j+1} at index {matches[j].Index}: '{matches[j].Value}'");
    }

    for (int i = 0; i < matches.Count; i++)
    {
      if (matches[i].Index == tokenIndex)
      {
        if (isDebug)
          Console.WriteLine($"[DEBUG]   ? Returning occurrence {i+1}");
        return i + 1; // 1-based occurrence index
      }
    }

    if (isDebug)
      Console.WriteLine($"[DEBUG]   ? Token not found at specified index, returning -1");

    return -1;
  }

  private static bool TryUnwrapFunctionChain(string chainText, int chainIndex, out string unwrappedChain, out int unwrappedIndex)
  {
    unwrappedChain = null;
    unwrappedIndex = chainIndex;

    if (string.IsNullOrEmpty(chainText))
      return false;

    var parenIndex = chainText.IndexOf('(');
    if (parenIndex <= 0)
      return false;

    int depth = 0;
    int closeIndex = -1;
    for (int i = parenIndex; i < chainText.Length; i++)
    {
      if (chainText[i] == '(')
        depth++;
      else if (chainText[i] == ')')
      {
        depth--;
        if (depth == 0)
        {
          closeIndex = i;
          break;
        }
      }
    }

    if (closeIndex >= 0 && closeIndex + 1 < chainText.Length && chainText[closeIndex + 1] == '.')
      return false;

    var prefix = chainText.Substring(0, parenIndex).Trim();
    if (prefix.Contains('.'))
      return false;

    unwrappedChain = chainText.Substring(parenIndex + 1).TrimStart();
    unwrappedIndex = chainIndex + parenIndex + 1;
    return !string.IsNullOrEmpty(unwrappedChain);
  }

  private static bool MatchesName(string name, string conventionalName, string token)
  {
    return string.Equals(name, token, StringComparison.OrdinalIgnoreCase) ||
           string.Equals(conventionalName, token, StringComparison.OrdinalIgnoreCase);
  }

  private static string[] SplitChainParts(string chainText)
  {
    if (string.IsNullOrEmpty(chainText))
      return Array.Empty<string>();

    var parts = new List<string>();
    int depth = 0;
    int start = 0;

    for (int i = 0; i < chainText.Length; i++)
    {
      if (chainText[i] == '(')
        depth++;
      else if (chainText[i] == ')')
        depth = Math.Max(0, depth - 1);
      else if (chainText[i] == '.' && depth == 0)
      {
        var part = chainText.Substring(start, i - start).Trim();
        if (part.Length > 0)
          parts.Add(part);
        start = i + 1;
      }
    }

    if (start <= chainText.Length)
    {
      var tail = chainText.Substring(start).Trim();
      if (tail.Length > 0)
        parts.Add(tail);
    }

    return parts.ToArray();
  }
}
