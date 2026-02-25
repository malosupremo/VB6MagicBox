using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    /// <summary>
    /// Costruisce la lista di sostituzioni (Replaces) per ogni modulo basandosi su:
    /// - References già risolte (LineNumbers)
    /// - ConventionalName vs Name (per determinare cosa rinominare)
    /// 
    /// Questo metodo analizza TUTTI i simboli di TUTTI i moduli e per ogni riferimento
    /// trova la posizione esatta del token nel codice sorgente, costruendo una LineReplace.
    /// 
    /// VANTAGGI:
    /// - Fase 2 (Refactoring) diventa triviale: basta applicare le sostituzioni pre-calcolate
    /// - Nessun re-parsing dei file
    /// - Nessuna ambiguità su cosa sostituire
    /// - Export in .linereplace.json per verifica manuale
    /// </summary>
    public static void BuildReplaces(VbProject project, Dictionary<string, string[]> fileCache)
    {
        Console.WriteLine();
        ConsoleX.WriteLineColor("===========================================", ConsoleColor.Yellow);
        ConsoleX.WriteLineColor("  2: Costruzione sostituzioni (Replaces)", ConsoleColor.Yellow);
        ConsoleX.WriteLineColor("===========================================", ConsoleColor.Yellow);
        Console.WriteLine();

        int totalReplaces = 0;

        // Cache delle righe dei file per evitare letture multiple
        fileCache ??= new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var module in project.Modules)
        {
            if (fileCache.ContainsKey(module.FullPath!))
                continue;

            if (File.Exists(module.FullPath))
                fileCache[module.FullPath] = File.ReadAllLines(module.FullPath);
        }

        // STEP 1: Per ogni modulo, raccogli TUTTI i simboli
        // Dobbiamo processare ANCHE simboli già convenzionali se hanno References
        // (es. bare property cross-module dove il nome è già corretto ma serve tracciare la posizione)
        var enumValueOwners = new Dictionary<VbEnumValue, VbEnumDef>();
        var enumValueConflicts = new HashSet<VbEnumValue>();

        var enumValueGroups = project.Modules
            .SelectMany(m => m.Enums.SelectMany(e => e.Values.Select(v => new { Enum = e, Value = v })))
            .GroupBy(x => x.Value.ConventionalName, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Select(x => x.Enum.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
            .ToList();

        foreach (var group in enumValueGroups)
        {
            foreach (var item in group)
            {
                enumValueOwners[item.Value] = item.Enum;
                enumValueConflicts.Add(item.Value);
            }
        }

        var allSymbols = new List<(string? oldName, string? newName, string category, object source, string? definingModule, VbEnumDef? enumOwner, bool qualifyEnumValueRefs)>();

        foreach (var module in project.Modules)
        {
            if (module.IsSharedExternal)
                continue;

            if (!module.IsConventional)
                allSymbols.Add((module.Name, module.ConventionalName, "Module", module, module.Name, null, false));

            foreach (var v in module.GlobalVariables)
            {
                if (!v.IsConventional || v.References.Count > 0)
                    allSymbols.Add((v.Name, v.ConventionalName, "GlobalVariable", v, module.Name, null, false));
            }

            foreach (var c in module.Constants)
            {
                if (!c.IsConventional || c.References.Count > 0)
                    allSymbols.Add((c.Name, c.ConventionalName, "Constant", c, module.Name, null, false));
            }

            foreach (var t in module.Types)
            {
                if (!t.IsConventional || t.References.Count > 0)
                    allSymbols.Add((t.Name, t.ConventionalName, "Type", t, module.Name, null, false));
                foreach (var f in t.Fields)
                {
                    if (!f.IsConventional || f.References.Count > 0)
                        allSymbols.Add((f.Name, f.ConventionalName, "Field", f, module.Name, null, false));
                }
            }

            foreach (var e in module.Enums)
            {
                if (!e.IsConventional || e.References.Count > 0)
                    allSymbols.Add((e.Name, e.ConventionalName, "Enum", e, module.Name, null, false));
                foreach (var v in e.Values)
                {
                    if (!v.IsConventional || v.References.Count > 0)
                        allSymbols.Add((v.Name, v.ConventionalName, "EnumValue", v, module.Name, e, enumValueConflicts.Contains(v)));
                }
            }

            foreach (var c in module.Controls)
            {
                if (!c.IsConventional || c.References.Count > 0)
                    allSymbols.Add((c.Name, c.ConventionalName, "Control", c, module.Name, null, false));
            }

            foreach (var p in module.Procedures)
            {
                if (!p.IsConventional || p.References.Count > 0)
                    allSymbols.Add((p.Name, p.ConventionalName, "Procedure", p, module.Name, null, false));
                foreach (var param in p.Parameters)
                {
                    if (!param.IsConventional || param.References.Count > 0)
                        allSymbols.Add((param.Name, param.ConventionalName, "Parameter", param, module.Name, null, false));
                }
                foreach (var lv in p.LocalVariables)
                {
                    if (!lv.IsConventional || lv.References.Count > 0)
                        allSymbols.Add((lv.Name, lv.ConventionalName, "LocalVariable", lv, module.Name, null, false));
                }
            }

            foreach (var prop in module.Properties)
            {
                if (!prop.IsConventional || prop.References.Count > 0)
                    allSymbols.Add((prop.Name, prop.ConventionalName, "Property", prop, module.Name, null, false));
                foreach (var param in prop.Parameters)
                {
                    if (!param.IsConventional || param.References.Count > 0)
                        allSymbols.Add((param.Name, param.ConventionalName, "PropertyParameter", param, module.Name, null, false));
                }
            }
        }

        // STEP 2: Per ogni simbolo, elabora le sue References e costruisci i LineReplace
        int symbolIndex = 0;
        var consoleLock = new object();
        var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Parallel.ForEach(allSymbols, parallelOptions, item =>
        {
            var (oldName, newName, category, source, definingModule, enumOwner, qualifyEnumValueRefs) = item;
            var currentIndex = Interlocked.Increment(ref symbolIndex);
            if (oldName == newName && !qualifyEnumValueRefs && !string.Equals(category, "Type", StringComparison.OrdinalIgnoreCase))
                return;

            bool isDebugSymbol = oldName?.Equals("Dsp_h", StringComparison.OrdinalIgnoreCase) == true
                          || oldName?.Equals("Msg_h", StringComparison.OrdinalIgnoreCase) == true;

            lock (consoleLock)
            {
                Console.Write($"\r   Processando simboli: [{currentIndex}/{allSymbols.Count}] {category}: {oldName}...".PadRight(Console.WindowWidth - 1));
            }

            // Trova il modulo che definisce il simbolo
            var ownerModule = project.Modules.FirstOrDefault(m =>
            string.Equals(m.Name, definingModule, StringComparison.OrdinalIgnoreCase));

            if (ownerModule == null)
                return;

            // DICHIARAZIONE: Aggiungi replace per la dichiarazione del simbolo (solo nel modulo che lo definisce)
            if (oldName != newName)
                AddDeclarationReplace(ownerModule, source, oldName, newName, category, fileCache);

            string? referenceNewName = newName;
            if (qualifyEnumValueRefs && enumOwner != null)
                referenceNewName = $"{enumOwner.ConventionalName}.{newName}";

            // REFERENCES: Aggiungi replace per tutti i riferimenti
            AddReferencesReplaces(project, source, oldName!, newName!, category, fileCache, isDebugSymbol, referenceNewName);

            // ATTRIBUTI VB6: Gestione speciale per "Attribute VB_Name" e "Attribute VarName."
            if (oldName != newName)
                AddAttributeReplaces(ownerModule, source, oldName!, newName!, category, fileCache);
        });

        // STEP 3: Conta i replace totali
        foreach (var module in project.Modules)
        {
            totalReplaces += module.Replaces.Count;

            // Ordina i Replaces per applicazione sicura (da fine a inizio)
            module.Replaces = module.Replaces
                .OrderByDescending(r => r.LineNumber)
                .ThenByDescending(r => r.StartChar)
                .ToList();
        }

        Console.WriteLine();
        ConsoleX.WriteLineColor($"   [OK] {totalReplaces} sostituzioni preparate per {project.Modules.Count} moduli", ConsoleColor.Green);
    }

    /// <summary>
    /// Aggiunge un Replace per la dichiarazione del simbolo
    /// </summary>
    private static void AddDeclarationReplace(
        VbModule module,
        object source,
        string? oldName,
        string? newName,
        string category,
        Dictionary<string, string[]> fileCache)
    {
        var lineNumberProp = source?.GetType().GetProperty("LineNumber");
        if (lineNumberProp?.GetValue(source) is not int lineNum || lineNum <= 0)
            return;

        if (!fileCache.TryGetValue(module.FullPath!, out var lines))
            return;

        if (lineNum > lines.Length)
            return;

        var line = lines[lineNum - 1]; // LineNumber è 1-based
        var (codePart, _) = SplitCodeAndComment(line);

        // Per le costanti, usa AddReplaceFromLine con skipStringLiterals
        if (source is VbConstant)
        {
            module.Replaces.AddReplaceFromLine(codePart, lineNum, oldName!, newName!, category + "_Declaration", -1, skipStringLiterals: true);
            return;
        }

        // Controlli (anche array): sostituisci il nome in tutte le righe Begin
        if (source is VbControl control && control.LineNumbers.Count > 0)
        {
            foreach (var controlLine in control.LineNumbers)
            {
                if (controlLine <= 0 || controlLine > lines.Length)
                    continue;

                var controlLineText = lines[controlLine - 1];
                var (controlCodePart, _) = SplitCodeAndComment(controlLineText);
                var controlPattern = $@"(?<=^.*Begin\s+\S+\s+){Regex.Escape(oldName!)}\b";
                var controlMatches = Regex.Matches(controlCodePart, controlPattern, RegexOptions.IgnoreCase);

                foreach (Match match in controlMatches)
                {
                    module.Replaces.AddReplace(
                        controlLine,
                        match.Index,
                        match.Index + match.Length,
                        match.Value,
                        newName!,
                        category + "_Declaration");
                }
            }

            return;
        }

        // Per altri simboli, trova tutte le occorrenze del nome nella dichiarazione
        var pattern = $@"\b{Regex.Escape(oldName!)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        foreach (Match match in matches)
        {
            module.Replaces.AddReplace(
                lineNum,
                match.Index,
                match.Index + match.Length,
                match.Value,
                newName!,
                category + "_Declaration");
        }
    }

    /// <summary>
    /// Aggiunge Replaces per tutti i riferimenti del simbolo
    /// </summary>
    private static void AddReferencesReplaces(
        VbProject project,
        object source,
        string oldName,
        string newName,
        string category,
        Dictionary<string, string[]> fileCache,
        bool isDebugSymbol = false,
        string? referenceNewNameOverride = null)
    {
        var referencesProp = source?.GetType().GetProperty("References");
        if (referencesProp?.GetValue(source) is not System.Collections.IEnumerable references)
            return;

        foreach (var reference in references)
        {
            var moduleProp = reference?.GetType().GetProperty("Module");
            var refModuleName = moduleProp?.GetValue(reference) as string;

            if (string.IsNullOrEmpty(refModuleName))
                continue;

            var refModule = project.Modules.FirstOrDefault(m =>
                string.Equals(m.Name, refModuleName, StringComparison.OrdinalIgnoreCase));

            if (refModule == null)
                continue;

            if (refModule.IsSharedExternal)
                continue;

            if (!fileCache.TryGetValue(refModule.FullPath!, out var lines))
                continue;

            var lineNumbersProp = reference?.GetType().GetProperty("LineNumbers");
            var occurrenceIndexesProp = reference?.GetType().GetProperty("OccurrenceIndexes");

            if (lineNumbersProp?.GetValue(reference) is not System.Collections.Generic.List<int> refLineNumbers)
                continue;

            var occurrenceIndexes = occurrenceIndexesProp?.GetValue(reference) as System.Collections.Generic.List<int>;

            for (int idx = 0; idx < refLineNumbers.Count; idx++)
            {
                var lineNum = refLineNumbers[idx];
                if (lineNum <= 0 || lineNum > lines.Length)
                    continue;

                var line = lines[lineNum - 1];
                var (codePart, _) = SplitCodeAndComment(line);

                var occIndex = (occurrenceIndexes != null && idx < occurrenceIndexes.Count) ? occurrenceIndexes[idx] : -1;

                var sourceModule = GetDefiningModule(project, source!);
                var sourceModuleInfo = project.Modules.FirstOrDefault(m =>
                    string.Equals(m.Name, sourceModule, StringComparison.OrdinalIgnoreCase));
                var sourceModuleReferenceName = sourceModuleInfo?.ConventionalName ?? sourceModule;

                int replacesBefore = isDebugSymbol && lineNum == 3146 ? refModule.Replaces.Count : 0;

                var referenceNewName = string.IsNullOrEmpty(referenceNewNameOverride) ? newName : referenceNewNameOverride;
                var stringRanges = GetStringLiteralRanges(codePart);
                var trimmedCodePart = codePart.TrimStart();
                var allowStringReplace = trimmedCodePart.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase) ||
                                         trimmedCodePart.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

                if (category == "EnumValue" && referenceNewName.Contains('.', StringComparison.Ordinal))
                {
                    AddEnumValueReferenceReplaces(refModule, codePart, lineNum, oldName, newName, referenceNewName, category, occIndex, stringRanges);
                    continue;
                }

                if (source is VbProperty)
                {
                    var shouldQualify = ShouldQualifyPropertyReference(refModule, lineNum, referenceNewName);
                    var propertyQualifier = GetPropertyQualifier(sourceModule, sourceModuleReferenceName, refModule, refModuleName, shouldQualify);
                    AddPropertyReferenceReplaces(refModule, codePart, lineNum, oldName, referenceNewName, category, occIndex, stringRanges, allowStringReplace, propertyQualifier);
                }
                else if (source is VbVariable variable && IsGlobalVariable(project, sourceModule, variable))
                {
                    var shouldQualify = ShouldQualifyModuleMemberReference(refModule, lineNum, referenceNewName);
                    var qualifier = shouldQualify ? sourceModuleReferenceName : null;
                    AddModuleMemberReferenceReplaces(refModule, codePart, lineNum, oldName, referenceNewName, category, occIndex, stringRanges, allowStringReplace, qualifier, sourceModule, sourceModuleReferenceName);
                }
                else if (source is VbControl && Regex.IsMatch(codePart.TrimStart(), @"^Begin\s+\S+\s+", RegexOptions.IgnoreCase))
                {
                    var pattern = $@"(?<=^.*Begin\s+\S+\s+){Regex.Escape(oldName)}\b";
                    var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

                    foreach (Match match in matches)
                    {
                        if (!allowStringReplace && IsInsideStringLiteral(stringRanges, match.Index))
                            continue;

                        refModule.Replaces.AddReplace(
                            lineNum,
                            match.Index,
                            match.Index + match.Length,
                            match.Value,
                            referenceNewName,
                            category + "_Reference");
                    }
                }
                else if (string.Equals(category, "Constant", StringComparison.OrdinalIgnoreCase))
                {
                    AddConstantReferenceReplaces(refModule, codePart, lineNum, oldName, referenceNewName, category, occIndex, stringRanges, allowStringReplace);
                }
                else
                {
                    var effectiveOldName = oldName;
                    if (string.Equals(category, "Type", StringComparison.OrdinalIgnoreCase))
                    {
                        var alternateOldName = GetTypeAlternateName(oldName);
                        if (!string.IsNullOrEmpty(alternateOldName) &&
                            !Regex.IsMatch(codePart, $@"\b{Regex.Escape(oldName)}\b", RegexOptions.IgnoreCase) &&
                            Regex.IsMatch(codePart, $@"\b{Regex.Escape(alternateOldName)}\b", RegexOptions.IgnoreCase))
                        {
                            effectiveOldName = alternateOldName;
                        }
                    }

                    refModule.Replaces.AddReplaceFromLine(codePart, lineNum, effectiveOldName, referenceNewName, category + "_Reference", occIndex, skipStringLiterals: !allowStringReplace);
                }

                //if (isDebugSymbol && lineNum == 3146)
                //{
                //  var added = refModule.Replaces.Count - replacesBefore;
                //  var lastReplace = added > 0 ? refModule.Replaces.Last() : null;
                //  Console.WriteLine($"\n[DBG] '{oldName}'@{refModuleName}:{lineNum} occIdx={occIndex} → {added} replace(s)" +
                //      (lastReplace != null ? $" char {lastReplace.StartChar}-{lastReplace.EndChar} '{lastReplace.OldText}'→'{lastReplace.NewText}'" : " NONE"));
                //}
            }
        }
    }

    private static void AddEnumValueReferenceReplaces(
        VbModule refModule,
        string codePart,
        int lineNum,
        string oldName,
        string newName,
        string qualifiedName,
        string category,
        int occIndex,
        List<(int start, int end)> stringRanges)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (occIndex > 0 && occIndex <= matches.Count)
            targetMatches = new[] { matches[occIndex - 1] };

        foreach (var match in targetMatches)
        {
            if (IsInsideStringLiteral(stringRanges, match.Index))
                continue;

            var replacement = IsQualifiedEnumReference(codePart, match.Index) ? newName : qualifiedName;
            if (string.Equals(match.Value, replacement, StringComparison.OrdinalIgnoreCase))
                continue;

            refModule.Replaces.AddReplace(
                lineNum,
                match.Index,
                match.Index + match.Length,
                match.Value,
                replacement,
                category + "_Reference");
        }
    }

    private static bool IsQualifiedEnumReference(string line, int tokenIndex)
    {
        var index = tokenIndex - 1;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        if (index < 0 || line[index] != '.')
            return false;

        index--;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        if (index < 0)
            return false;

        var end = index;
        while (index >= 0 && (char.IsLetterOrDigit(line[index]) || line[index] == '_'))
            index--;

        return end > index;
    }

    private static void AddConstantReferenceReplaces(
        VbModule refModule,
        string codePart,
        int lineNum,
        string oldName,
        string newName,
        string category,
        int occIndex,
        List<(int start, int end)> stringRanges,
        bool allowStringReplace)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (occIndex > 0 && occIndex <= matches.Count)
            targetMatches = new[] { matches[occIndex - 1] };

        foreach (var match in targetMatches)
        {
            if (!allowStringReplace && IsInsideStringLiteral(stringRanges, match.Index))
                continue;

            if (IsConstantMemberAccessToken(codePart, match.Index))
                continue;

            if (IsTypeFieldDeclaration(codePart, match.Index, oldName))
                continue;

            refModule.Replaces.AddReplace(
                lineNum,
                match.Index,
                match.Index + match.Length,
                match.Value,
                newName,
                category + "_Reference");
        }
    }

    private static bool IsConstantMemberAccessToken(string line, int tokenIndex)
    {
        if (tokenIndex <= 0)
            return false;

        var index = tokenIndex - 1;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        return index >= 0 && line[index] == '.';
    }

    private static bool IsTypeFieldDeclaration(string line, int tokenIndex, string token)
    {
        if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(token))
            return false;

        var trimmed = line.TrimStart();
        if (trimmed.StartsWith("Public ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Private ", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed.Substring(trimmed.IndexOf(' ') + 1).TrimStart();
        }

        var tokenMatch = Regex.Match(trimmed, $@"^\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase);
        if (!tokenMatch.Success)
            return false;

        var asIndex = trimmed.IndexOf(" As ", StringComparison.OrdinalIgnoreCase);
        if (asIndex < 0)
            return false;

        return tokenIndex <= line.IndexOf(" As ", StringComparison.OrdinalIgnoreCase);
    }

    private static List<(int start, int end)> GetStringLiteralRanges(string line)
    {
        var ranges = new List<(int start, int end)>();
        if (string.IsNullOrEmpty(line))
            return ranges;

        bool inString = false;
        int stringStart = -1;

        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"')
            {
                if (!inString)
                {
                    inString = true;
                    stringStart = i;
                }
                else if (i + 1 < line.Length && line[i + 1] == '"')
                {
                    i++;
                }
                else
                {
                    inString = false;
                    if (stringStart >= 0)
                        ranges.Add((stringStart, i + 1));
                }
            }
        }

        return ranges;
    }

    private static bool IsInsideStringLiteral(List<(int start, int end)> ranges, int index)
    {
        return ranges.Any(r => index >= r.start && index < r.end);
    }

    /// <summary>
    /// Aggiunge Replaces per attributi VB6 speciali (Attribute VB_Name, Attribute VarName.)
    /// </summary>
    private static void AddAttributeReplaces(
        VbModule module,
        object source,
        string oldName,
        string newName,
        string category,
        Dictionary<string, string[]> fileCache)
    {
        if (!fileCache.TryGetValue(module.Name!, out var lines))
            return;

        // VB_Name per moduli/classi/form
        if (source is VbModule && (module.IsClass || module.IsForm))
        {
            for (int i = 0; i < Math.Min(20, lines.Length); i++)
            {
                var line = lines[i];
                var vbNameMatch = Regex.Match(line, @"Attribute\s+VB_Name\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase);

                if (vbNameMatch.Success && vbNameMatch.Groups[1].Value.Equals(oldName, StringComparison.OrdinalIgnoreCase))
                {
                    // La sostituzione è dentro le virgolette
                    var nameGroup = vbNameMatch.Groups[1];
                    module.Replaces.AddReplace(
                        i + 1,
                        nameGroup.Index,
                        nameGroup.Index + nameGroup.Length,
                        nameGroup.Value,
                        newName,
                        category + "_AttributeVBName");
                }
            }
        }

        // Attribute VarName.VB_VarXXX (righe dopo dichiarazione variabili globali)
        var lineNumberProp = source?.GetType().GetProperty("LineNumber");
        if (lineNumberProp?.GetValue(source) is int declarationLineNum && declarationLineNum > 0 && declarationLineNum < lines.Length)
        {
            var nextLine = lines[declarationLineNum]; // declarationLineNum è 1-based, +1 per riga successiva, -1 per array
            var trimmedNextLine = nextLine.TrimStart();

            if (trimmedNextLine.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                var attributeMatch = Regex.Match(trimmedNextLine, @"^Attribute\s+(\w+)\.", RegexOptions.IgnoreCase);
                if (attributeMatch.Success && attributeMatch.Groups[1].Value.Equals(oldName, StringComparison.OrdinalIgnoreCase))
                {
                    var nameGroup = attributeMatch.Groups[1];
                    var absoluteIndex = nextLine.IndexOf(trimmedNextLine) + nameGroup.Index;

                    module.Replaces.AddReplace(
                        declarationLineNum + 1,
                        absoluteIndex,
                        absoluteIndex + nameGroup.Length,
                        nameGroup.Value,
                        newName,
                        category + "_AttributeVar");
                }
            }
        }
    }

    /// <summary>
    /// Helper: estrae il nome del modulo che definisce il simbolo
    /// </summary>
    private static bool ShouldQualifyPropertyReference(VbModule refModule, int lineNum, string? referenceName)
    {
        if (string.IsNullOrWhiteSpace(referenceName))
            return false;

        var proc = refModule.GetProcedureAtLine(lineNum);
        if (proc != null)
            return HasShadowing(proc.Parameters, proc.LocalVariables, referenceName);

        var prop = refModule.Properties.FirstOrDefault(p => p.ContainsLine(lineNum));
        if (prop != null)
            return HasShadowing(prop.Parameters, null, referenceName);

        return false;
    }

    private static bool ShouldQualifyModuleMemberReference(VbModule refModule, int lineNum, string? referenceName)
    {
        if (string.IsNullOrWhiteSpace(referenceName))
            return false;

        var proc = refModule.GetProcedureAtLine(lineNum);
        if (proc != null)
            return HasShadowing(proc.Parameters, proc.LocalVariables, referenceName);

        var prop = refModule.Properties.FirstOrDefault(p => p.ContainsLine(lineNum));
        if (prop != null)
            return HasShadowing(prop.Parameters, null, referenceName);

        return false;
    }

    private static bool IsGlobalVariable(VbProject project, string sourceModule, VbVariable variable)
    {
        if (variable == null)
            return false;

        var module = project.Modules.FirstOrDefault(m =>
            string.Equals(m.Name, sourceModule, StringComparison.OrdinalIgnoreCase));

        return module != null && module.GlobalVariables.Contains(variable);
    }

    private static string? GetPropertyQualifier(string sourceModule, string? sourceModuleReferenceName, VbModule refModule, string refModuleName, bool shouldQualify)
    {
        if (!shouldQualify)
            return null;

        return sourceModuleReferenceName;
    }

    private static void AddPropertyReferenceReplaces(
        VbModule refModule,
        string codePart,
        int lineNum,
        string oldName,
        string newName,
        string category,
        int occIndex,
        List<(int start, int end)> stringRanges,
        bool allowStringReplace,
        string? qualifier)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (occIndex > 0 && occIndex <= matches.Count)
            targetMatches = new[] { matches[occIndex - 1] };

        foreach (var match in targetMatches)
        {
            if (!allowStringReplace && IsInsideStringLiteral(stringRanges, match.Index))
                continue;

            var replacement = newName;
            if (!IsMemberAccessToken(codePart, match.Index) && !string.IsNullOrWhiteSpace(qualifier))
                replacement = $"{qualifier}.{newName}";

            if (string.Equals(match.Value, replacement, StringComparison.OrdinalIgnoreCase))
                continue;

            refModule.Replaces.AddReplace(
                lineNum,
                match.Index,
                match.Index + match.Length,
                match.Value,
                replacement,
                category + "_Reference");
        }
    }

    private static void AddModuleMemberReferenceReplaces(
        VbModule refModule,
        string codePart,
        int lineNum,
        string oldName,
        string newName,
        string category,
        int occIndex,
        List<(int start, int end)> stringRanges,
        bool allowStringReplace,
        string? qualifier,
        string? sourceModuleName,
        string? sourceModuleReferenceName)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (occIndex > 0 && occIndex <= matches.Count)
            targetMatches = new[] { matches[occIndex - 1] };

        foreach (var match in targetMatches)
        {
            if (!allowStringReplace && IsInsideStringLiteral(stringRanges, match.Index))
                continue;

            if (IsMemberAccessToken(codePart, match.Index))
            {
                var memberQualifier = GetMemberAccessQualifier(codePart, match.Index);
                if (!string.IsNullOrWhiteSpace(memberQualifier) &&
                    !string.Equals(memberQualifier, sourceModuleName, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(memberQualifier, sourceModuleReferenceName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            var replacement = newName;
            if (!IsMemberAccessToken(codePart, match.Index) && !string.IsNullOrWhiteSpace(qualifier))
                replacement = $"{qualifier}.{newName}";

            if (string.Equals(match.Value, replacement, StringComparison.OrdinalIgnoreCase))
                continue;

            refModule.Replaces.AddReplace(
                lineNum,
                match.Index,
                match.Index + match.Length,
                match.Value,
                replacement,
                category + "_Reference");
        }
    }

    private static string? GetMemberAccessQualifier(string line, int tokenIndex)
    {
        if (tokenIndex <= 0)
            return null;

        var index = tokenIndex - 1;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        if (index < 0 || line[index] != '.')
            return null;

        index--;
        while (index >= 0 && char.IsWhiteSpace(line[index]))
            index--;

        if (index < 0 || !IsIdentifierChar(line[index]))
            return null;

        var end = index + 1;
        while (index >= 0 && IsIdentifierChar(line[index]))
            index--;

        var start = index + 1;
        return line.Substring(start, end - start);
    }

    private static bool HasShadowing(IEnumerable<VbParameter> parameters, IEnumerable<VbVariable>? locals, string referenceName)
    {
        var comparer = StringComparer.OrdinalIgnoreCase;
        foreach (var parameter in parameters)
        {
            var name = GetEffectiveName(parameter);
            if (!string.IsNullOrEmpty(name) && comparer.Equals(name, referenceName))
                return true;
        }

        if (locals != null)
        {
            foreach (var local in locals)
            {
                var name = GetEffectiveName(local);
                if (!string.IsNullOrEmpty(name) && comparer.Equals(name, referenceName))
                    return true;
            }
        }

        return false;
    }

    private static string? GetEffectiveName(VbParameter parameter)
        => string.IsNullOrWhiteSpace(parameter.ConventionalName) ? parameter.Name : parameter.ConventionalName;

    private static string? GetEffectiveName(VbVariable variable)
        => string.IsNullOrWhiteSpace(variable.ConventionalName) ? variable.Name : variable.ConventionalName;

    private static string GetDefiningModule(VbProject project, object source)
    {
        if (source is VbModule mod)
            return mod.Name;

        // Per altri simboli, cerca nel progetto
        foreach (var module in project.Modules)
        {
            if (module.GlobalVariables.Contains(source))
                return module.Name;
            if (module.Constants.Contains(source))
                return module.Name;
            if (module.Types.Any(t => t == source || t.Fields.Contains(source)))
                return module.Name;
            if (module.Enums.Any(e => e == source || e.Values.Contains(source)))
                return module.Name;
            if (module.Controls.Contains(source))
                return module.Name;
            if (module.Procedures.Any(p => p == source || p.Parameters.Contains(source) || p.LocalVariables.Contains(source)))
                return module.Name;
            if (module.Properties.Any(p => p == source || p.Parameters.Contains(source)))
                return module.Name;
        }

        return string.Empty;
    }

    /// <summary>
    /// Helper: separa codice da commento (gestisce stringhe correttamente)
    /// </summary>
    private static (string code, string comment) SplitCodeAndComment(string line)
    {
        bool inString = false;
        for (int i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (!inString)
                    inString = true;
                else if (i + 1 < line.Length && line[i + 1] == '"')
                    i++; // escaped double quote
                else
                    inString = false;
            }
            else if (!inString && ch == '\'')
                return (line[..i].TrimEnd(), line[i..]);
        }
        return (line, string.Empty);
    }

    private static string GetTypeAlternateName(string typeName)
    {
        if (string.IsNullOrWhiteSpace(typeName))
            return string.Empty;

        if (typeName.EndsWith("_T", StringComparison.OrdinalIgnoreCase))
            return typeName.Substring(0, typeName.Length - 2);

        return typeName + "_T";
    }
}
