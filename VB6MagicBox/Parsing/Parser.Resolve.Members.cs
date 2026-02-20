using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
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

            // Rimuovi commenti
            var noComment = raw;
            var idx = noComment.IndexOf("'");
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            var trimmedNoComment = noComment.TrimStart();
            if (trimmedNoComment.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
            {
                var withExpr = trimmedNoComment.Substring(5).Trim();
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
                noComment = Regex.Replace(noComment,
                    @"(?<!\w)\.(\s*[A-Za-z_]\w*(?:\([^)]*\))?)",
                    m => withPrefix + "." + m.Groups[1].Value,
                    RegexOptions.IgnoreCase);
            }

            var chainPattern = @"([A-Za-z_]\w*(?:\([^)]*\))?)(?:\s*\.\s*[A-ZaZ_]\w*(?:\([^)]*\))?)+";
            var chainTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Match m in Regex.Matches(noComment, chainPattern, RegexOptions.IgnoreCase))
                chainTexts.Add(m.Value);

            foreach (Match inner in Regex.Matches(noComment, @"\(([^)]*)\)"))
            {
                var innerText = inner.Groups[1].Value;
                foreach (Match m in Regex.Matches(innerText, chainPattern, RegexOptions.IgnoreCase))
                    chainTexts.Add(m.Value);
            }

            foreach (var chainText in chainTexts)
            {
                var parts = chainText
                    .Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

                if (parts.Length < 2)
                    continue;

                var baseVarName = parts[0];
                var parenIndex = baseVarName.IndexOf('(');
                if (parenIndex >= 0)
                    baseVarName = baseVarName.Substring(0, parenIndex);

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

                for (int partIndex = startPartIndex; partIndex < parts.Length; partIndex++)
                {
                    var fieldName = parts[partIndex];
                    var fieldParenIndex = fieldName.IndexOf('(');
                    if (fieldParenIndex >= 0)
                        fieldName = fieldName.Substring(0, fieldParenIndex);

                    if (string.IsNullOrEmpty(fieldName))
                        break;

                    var baseTypeName = typeName;
                    if (baseTypeName.Contains('('))
                        baseTypeName = baseTypeName.Substring(0, baseTypeName.IndexOf('('));
                    if (baseTypeName.Contains('.'))
                        baseTypeName = baseTypeName.Split('.').Last();

                    if (classIndex.TryGetValue(baseTypeName, out var classModule))
                    {
                        var classProp = classModule.Properties.FirstOrDefault(p =>
                            p.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                        if (classProp != null)
                        {
                          classProp.Used = true;
                          classProp.References.AddLineNumber(mod.Name, proc.Name, i + 1);
                          typeName = classProp.ReturnType;
                          if (string.IsNullOrEmpty(typeName))
                            break;
                          continue;
                        }
                    }

                    if (!typeIndex.TryGetValue(baseTypeName, out var typeDef))
                        break;

                    var field = typeDef.Fields.FirstOrDefault(f =>
                        !string.IsNullOrEmpty(f.Name) &&
                        string.Equals(f.Name, fieldName, StringComparison.OrdinalIgnoreCase));

                    if (field == null)
                        break;

                    field.Used = true;
                    field.References.AddLineNumber(mod.Name, proc.Name, i + 1);
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

            // Rimuovi commenti
            var noComment = raw;
            var idx = noComment.IndexOf("'");
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            var trimmedNoComment = noComment.TrimStart();
            if (trimmedNoComment.StartsWith("With ", StringComparison.OrdinalIgnoreCase))
            {
                var withExpr = trimmedNoComment.Substring(5).Trim();
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
                noComment = Regex.Replace(noComment,
                    @"(?<!\w)\.(\s*[A-Za-z_]\w*(?:\([^)]*\))?)",
                    m => withPrefix + "." + m.Groups[1].Value,
                    RegexOptions.IgnoreCase);
            }

            var chainPattern = @"([A-Za-z_]\w*(?:\([^)]*\))?)(?:\s*\.\s*[A-ZaZ_]\w*(?:\([^)]*\))?)+";
            var chainTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Match m in Regex.Matches(noComment, chainPattern, RegexOptions.IgnoreCase))
                chainTexts.Add(m.Value);

            foreach (Match inner in Regex.Matches(noComment, @"\(([^)]*)\)"))
            {
                var innerText = inner.Groups[1].Value;
                foreach (Match m in Regex.Matches(innerText, chainPattern, RegexOptions.IgnoreCase))
                    chainTexts.Add(m.Value);
            }

            foreach (var chainText in chainTexts)
            {
                var parts = chainText
                    .Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

                if (parts.Length < 2)
                    continue;

                var baseVarName = parts[0];
                var parenIndex = baseVarName.IndexOf('(');
                if (parenIndex >= 0)
                    baseVarName = baseVarName.Substring(0, parenIndex);

                if (!env.TryGetValue(baseVarName, out var typeName) || string.IsNullOrEmpty(typeName))
                    continue;

                for (int partIndex = 1; partIndex < parts.Length; partIndex++)
                {
                    var fieldName = parts[partIndex];
                    var fieldParenIndex = fieldName.IndexOf('(');
                    if (fieldParenIndex >= 0)
                        fieldName = fieldName.Substring(0, fieldParenIndex);

                    if (string.IsNullOrEmpty(fieldName))
                        break;

                    var baseTypeName = typeName;
                    if (baseTypeName.Contains('('))
                        baseTypeName = baseTypeName.Substring(0, baseTypeName.IndexOf('('));
                    if (baseTypeName.Contains('.'))
                        baseTypeName = baseTypeName.Split('.').Last();

                    if (classIndex.TryGetValue(baseTypeName, out var classModule))
                    {
                        var classProp = classModule.Properties.FirstOrDefault(p =>
                            p.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                        if (classProp != null)
                        {
                          classProp.Used = true;
                          classProp.References.AddLineNumber(mod.Name, prop.Name, i + 1);
                          typeName = classProp.ReturnType;
                          if (string.IsNullOrEmpty(typeName))
                            break;
                          continue;
                        }
                    }

                    if (!typeIndex.TryGetValue(baseTypeName, out var typeDef))
                        break;

                    var field = typeDef.Fields.FirstOrDefault(f =>
                        !string.IsNullOrEmpty(f.Name) &&
                        string.Equals(f.Name, fieldName, StringComparison.OrdinalIgnoreCase));

                    if (field == null)
                        break;

                    field.Used = true;
                    field.References.AddLineNumber(mod.Name, prop.Name, i + 1);
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
            var noComment = raw;
            var idx = noComment.IndexOf("'");
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            // Pattern: controlName.Property o controlName.Method() oppure controlName(index).Property
            // ANCHE: ModuleName.controlName(index).Property per referenze cross-module
            foreach (Match m in Regex.Matches(noComment, @"(\w+)(?:\([^\)]*\))?\.(\w+)"))
            {
                var controlName = m.Groups[1].Value;

                // Verifica se è un controllo del form corrente
                if (controlIndex.TryGetValue(controlName, out var control))
                {
                    MarkControlAsUsed(control, mod.Name, proc.Name, i + 1);
                }
            }

            // Pattern avanzato per referenze cross-module: ModuleName.ControlName(index).Property
            foreach (Match m in Regex.Matches(noComment, @"(\w+)\.(\w+)(?:\([^\)]*\))?\.(\w+)"))
            {
                var moduleName = m.Groups[1].Value;
                var controlName = m.Groups[2].Value;

                // Cerca il modulo nel progetto
                var targetModule = mod.Owner?.Modules?.FirstOrDefault(module =>
                    string.Equals(module.Name, moduleName, StringComparison.OrdinalIgnoreCase));

                if (targetModule != null)
                {
                    // Cerca TUTTI i controlli con lo stesso nome (array di controlli)
                    var controls = targetModule.Controls.Where(c =>
                        string.Equals(c.Name, controlName, StringComparison.OrdinalIgnoreCase));

                    foreach (var control in controls)
                    {
                        MarkControlAsUsed(control, mod.Name, proc.Name, i + 1);
                    }
                }
            }
        }
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
            int currentLineNumber = i + 1;

            // Fine procedura - controlla SOLO i terminatori specifici della procedura
            // (Non usa "End " generico per evitare match con End If, End With, ecc.)
            if (i > proc.LineNumber - 1 && IsProcedureEndLine(raw, proc.Kind))
                break;

            // Rimuovi commenti
            var noComment = raw;
            var idx = noComment.IndexOf("'", StringComparison.Ordinal);
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            // Rimuovi stringhe per evitare di catturare nomi dentro stringhe
            noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

            // Cerca tutti i token word nella riga
            foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\b"))
            {
                var tokenName = m.Groups[1].Value;

                // Controlla se è un parametro
                if (parameterIndex.TryGetValue(tokenName, out var parameter))
                {
                    parameter.Used = true;
                    parameter.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
                }

                // Controlla se è una variabile locale
                if (localVariableIndex.TryGetValue(tokenName, out var localVar))
                {
                    // Esclude la riga di dichiarazione della variabile (usa direttamente LineNumber)
                    if (localVar.LineNumber == currentLineNumber)
                        continue;

                    localVar.Used = true;
                    localVar.References.AddLineNumber(mod.Name, proc.Name, currentLineNumber);
                }
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

            // Rimuovi commenti
            var noComment = raw;
            var idx = noComment.IndexOf("'", StringComparison.Ordinal);
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            // Rimuovi stringhe per evitare di catturare nomi dentro stringhe
            noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

            foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\b"))
            {
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

            var noComment = raw;
            var idx = noComment.IndexOf("'", StringComparison.Ordinal);
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

            if (Regex.IsMatch(noComment, $@"\b{Regex.Escape(prop.Name)}\b", RegexOptions.IgnoreCase))
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
            var noComment = raw;
            var idx = noComment.IndexOf("'", StringComparison.Ordinal);
            if (idx >= 0)
                noComment = noComment.Substring(0, idx).Trim();

            noComment = Regex.Replace(noComment, @"""[^""]*""", "\"\"");

            if (Regex.IsMatch(noComment, $@"\b{Regex.Escape(proc.Name)}\b", RegexOptions.IgnoreCase))
            {
                proc.References.AddLineNumber(mod.Name, proc.Name, i + 1);
            }
        }
    }

    // ---------------------------------------------------------
    // MARCATURA VALORI ENUM USATI
    // ---------------------------------------------------------

    private static void MarkUsedEnumValues(VbProject project)
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
            var fileLines = File.ReadAllLines(mod.FullPath);

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
                    var noComment = line;
                    var idx = noComment.IndexOf("'");
                    if (idx >= 0)
                        noComment = noComment.Substring(0, idx);

                    // Cerca ogni valore enum nel codice
                    foreach (var kvp in allEnumValues)
                    {
                        var enumValueName = kvp.Key;
                        var enumValues = kvp.Value;

                        // Usa word boundary per evitare match parziali
                        if (Regex.IsMatch(noComment, $@"\b{Regex.Escape(enumValueName)}\b", RegexOptions.IgnoreCase))
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

    private static void MarkUsedEvents(VbProject project)
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
            var fileLines = File.ReadAllLines(mod.FullPath);

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
                    var noComment = line;
                    var idx = noComment.IndexOf("'");
                    if (idx >= 0)
                        noComment = noComment.Substring(0, idx);

                    // Pattern: RaiseEvent EventName o RaiseEvent EventName(params)
                    var raiseEventMatch = Regex.Match(noComment, @"RaiseEvent\s+(\w+)", RegexOptions.IgnoreCase);
                    if (raiseEventMatch.Success)
                    {
                        var eventName = raiseEventMatch.Groups[1].Value;

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

    private static void ResolveEnumValueReferences(VbProject project)
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
            var fileLines = File.ReadAllLines(mod.FullPath);

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

                    foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\s*\.\s*([A-Za-z_]\w*)\b"))
                    {
                        var enumName = m.Groups[1].Value;
                        var valueName = m.Groups[2].Value;

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

                    foreach (Match m in Regex.Matches(noComment, @"\b([A-Za-z_]\w*)\s*\.\s*([A-Za-z_]\w*)\b"))
                    {
                        var enumName = m.Groups[1].Value;
                        var valueName = m.Groups[2].Value;

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
}
