using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    private sealed record ReferenceReplaceEntry(
        VbModule RefModule,
        int LineNumber,
        int OccurrenceIndex,
        int StartChar,
        string OldName,
        string NewName,
        string Category,
        object Source,
        string? ReferenceNewNameOverride);
    private sealed class StartCharCheckEntry
    {
        public string? Module { get; set; }
        public string? Procedure { get; set; }
        public int LineNumber { get; set; }
        public int StartChar { get; set; }
        public int OccurrenceIndex { get; set; }
        public string? OldName { get; set; }
        public string? NewName { get; set; }
        public string? Category { get; set; }
    }

    public static void ExportStartCharChecks(string outputPath)
    {
        var entries = StartCharChecks?.Values
            .OrderBy(e => e.Module, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.Procedure, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.LineNumber)
            .ThenBy(e => e.StartChar)
            .ToList() ?? new List<StartCharCheckEntry>();

        using var writer = new StreamWriter(outputPath, false);
        writer.WriteLine("Module;Procedure;LineNumber;StartChar;OccurrenceIndex;OldName;NewName;Category");
        foreach (var entry in entries)
        {
            writer.WriteLine($"{entry.Module};{entry.Procedure};{entry.LineNumber};{entry.StartChar};{entry.OccurrenceIndex};{entry.OldName};{entry.NewName};{entry.Category}");
        }
    }

    private static ConcurrentDictionary<string, StartCharCheckEntry> StartCharChecks = new(StringComparer.OrdinalIgnoreCase);
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
        StartCharChecks = new ConcurrentDictionary<string, StartCharCheckEntry>(StringComparer.OrdinalIgnoreCase);

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

        // STEP 2: Raccogli tutte le references in una lista per elaborazione per riga
        var replaceEntries = new ConcurrentBag<ReferenceReplaceEntry>();
        var replaceEntryKeys = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        int symbolIndex = 0;
        var consoleLock = new object();
        var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Parallel.ForEach(allSymbols, parallelOptions, item =>
        {
            var (oldName, newName, category, source, definingModule, enumOwner, qualifyEnumValueRefs) = item;
            var currentIndex = Interlocked.Increment(ref symbolIndex);
            if (oldName == newName && !qualifyEnumValueRefs && !string.Equals(category, "Type", StringComparison.OrdinalIgnoreCase))
                return;

            lock (consoleLock)
            {
                Console.Write($"\r   Processando simboli: [{currentIndex + 1}/{allSymbols.Count}] {category}: {oldName}...".PadRight(Console.WindowWidth - 1));
            }

            var ownerModule = project.Modules.FirstOrDefault(m =>
                string.Equals(m.Name, definingModule, StringComparison.OrdinalIgnoreCase));

            if (ownerModule == null)
                return;

            if (oldName != newName)
                AddDeclarationReplace(ownerModule, source, oldName, newName, category, fileCache);

            string? referenceNewName = newName;
            if (qualifyEnumValueRefs && enumOwner != null)
                referenceNewName = $"{enumOwner.ConventionalName}.{newName}";

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

                if (refModule == null || refModule.IsSharedExternal)
                    continue;

                if (!fileCache.TryGetValue(refModule.FullPath!, out var lines))
                    continue;

                var lineNumbersProp = reference?.GetType().GetProperty("LineNumbers");
                var occurrenceIndexesProp = reference?.GetType().GetProperty("OccurrenceIndexes");
                var startCharsProp = reference?.GetType().GetProperty("StartChars");

                if (lineNumbersProp?.GetValue(reference) is not System.Collections.Generic.List<int> refLineNumbers)
                    continue;

                var occurrenceIndexes = occurrenceIndexesProp?.GetValue(reference) as System.Collections.Generic.List<int>;
                var startChars = startCharsProp?.GetValue(reference) as System.Collections.Generic.List<int>;

                for (int idx = 0; idx < refLineNumbers.Count; idx++)
                {
                    var lineNum = refLineNumbers[idx];
                    if (lineNum <= 0 || lineNum > lines.Length)
                        continue;

                    var occIndex = (occurrenceIndexes != null && idx < occurrenceIndexes.Count) ? occurrenceIndexes[idx] : -1;
                    var startChar = (startChars != null && idx < startChars.Count) ? startChars[idx] : -1;

                    if (occIndex < 0 && startChar < 0)
                        continue;

                    var entryKey = $"{refModule.Name}|{lineNum}|{startChar}|{occIndex}|{oldName}|{newName}|{category}";
                    if (replaceEntryKeys.TryAdd(entryKey, 0))
                    {
                        replaceEntries.Add(new ReferenceReplaceEntry(
                            refModule,
                            lineNum,
                            occIndex,
                            startChar,
                            oldName!,
                            newName!,
                            category,
                            source!,
                            referenceNewName));
                    }
                }
            }

            if (oldName != newName)
                AddAttributeReplaces(ownerModule, source, oldName!, newName!, category, fileCache);
        });

        Console.WriteLine();
        ConsoleX.WriteLineColor("   [i] Raccolta references completata. Inizio applicazione replaces...", ConsoleColor.Cyan);

        var lineCache = new Dictionary<(string Module, int Line), (string CodePart, List<(int start, int end)> StringRanges, bool AllowStringReplace)>();
        var sourceModuleCache = new Dictionary<object, string>(ReferenceEqualityComparer.Instance);

        var moduleOrder = project.Modules
            .Select((m, index) => new { m.Name, index })
            .ToDictionary(x => x.Name, x => x.index, StringComparer.OrdinalIgnoreCase);

        var replaceEntryList = replaceEntries.ToList();
        replaceEntryList.Sort((a, b) =>
        {
            var moduleCompare = moduleOrder[a.RefModule.Name].CompareTo(moduleOrder[b.RefModule.Name]);
            if (moduleCompare != 0)
                return moduleCompare;

            var lineCompare = a.LineNumber.CompareTo(b.LineNumber);
            if (lineCompare != 0)
                return lineCompare;

            return b.StartChar.CompareTo(a.StartChar);
        });

        string? currentModuleName = null;
        int processedEntries = 0;
        int totalEntries = replaceEntryList.Count;

        foreach (var entry in replaceEntryList)
        {
            if (!string.Equals(currentModuleName, entry.RefModule.Name, StringComparison.OrdinalIgnoreCase))
            {
                currentModuleName = entry.RefModule.Name;
                Console.Write($"\r   [i] Applying replaces ({processedEntries + 1}/{totalEntries}): {currentModuleName}...".PadRight(Console.WindowWidth - 1));
            }

            if (!fileCache.TryGetValue(entry.RefModule.FullPath!, out var lines))
                continue;

            if (entry.LineNumber <= 0 || entry.LineNumber > lines.Length)
                continue;

            var cacheKey = (entry.RefModule.Name, entry.LineNumber);
            if (!lineCache.TryGetValue(cacheKey, out var cacheEntry))
            {
                var line = lines[entry.LineNumber - 1];
                var (codePart, _) = SplitCodeAndComment(line);
                var stringRanges = GetStringLiteralRanges(codePart);
                var trimmedCodePart = codePart.TrimStart();
                var allowStringReplace = trimmedCodePart.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase) ||
                                         trimmedCodePart.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);
                cacheEntry = (codePart, stringRanges, allowStringReplace);
                lineCache[cacheKey] = cacheEntry;
            }

            ApplyReferenceReplace(project, entry, cacheEntry.CodePart, cacheEntry.StringRanges, cacheEntry.AllowStringReplace, sourceModuleCache);
            processedEntries++;
        }

        Console.WriteLine();
        ConsoleX.WriteLineColor($"   [OK] Ordinamento in corso...", ConsoleColor.Green);

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
        ConsoleX.WriteLineColor($"   [OK] {totalReplaces} sostituzioni ordinate e preparate per {project.Modules.Count} moduli", ConsoleColor.Green);
    }

    private static void ApplyReferenceReplace(
        VbProject project,
        ReferenceReplaceEntry entry,
        string codePart,
        List<(int start, int end)> stringRanges,
        bool allowStringReplace,
        Dictionary<object, string> sourceModuleCache)
    {
        if (!sourceModuleCache.TryGetValue(entry.Source, out var sourceModule))
        {
            sourceModule = GetDefiningModule(project, entry.Source);
            sourceModuleCache[entry.Source] = sourceModule;
        }
        var sourceModuleInfo = project.Modules.FirstOrDefault(m =>
            string.Equals(m.Name, sourceModule, StringComparison.OrdinalIgnoreCase));
        var sourceModuleReferenceName = sourceModuleInfo?.ConventionalName ?? sourceModule;

        var referenceNewName = string.IsNullOrEmpty(entry.ReferenceNewNameOverride)
            ? entry.NewName
            : entry.ReferenceNewNameOverride;

        if (entry.Category == "EnumValue" && referenceNewName.Contains('.', StringComparison.Ordinal))
        {
            AddEnumValueReferenceReplaces(entry.RefModule, codePart, entry.LineNumber, entry.OldName, entry.NewName, referenceNewName,
                entry.Category, entry.OccurrenceIndex, entry.StartChar, stringRanges);
            return;
        }

        if (entry.Source is VbProperty)
        {
            var shouldQualify = ShouldQualifyPropertyReference(entry.RefModule, entry.LineNumber, referenceNewName);
            var propertyQualifier = GetPropertyQualifier(sourceModule, sourceModuleReferenceName, entry.RefModule, entry.RefModule.Name, shouldQualify);
            AddPropertyReferenceReplaces(entry.RefModule, codePart, entry.LineNumber, entry.OldName, referenceNewName,
                entry.Category, entry.OccurrenceIndex, entry.StartChar, stringRanges, allowStringReplace, propertyQualifier);
            return;
        }

        if (entry.Source is VbVariable variable && IsGlobalVariable(project, sourceModule, variable))
        {
            var shouldQualify = ShouldQualifyModuleMemberReference(entry.RefModule, entry.LineNumber, referenceNewName);
            var qualifier = shouldQualify ? sourceModuleReferenceName : null;
            if (entry.StartChar >= 0 && entry.StartChar < codePart.Length)
            {
                var length = Math.Min(entry.OldName.Length, codePart.Length - entry.StartChar);
                var foundText = codePart.Substring(entry.StartChar, length);
                if (string.Equals(foundText, entry.OldName, StringComparison.OrdinalIgnoreCase))
                {
                    if (allowStringReplace || !IsInsideStringLiteral(stringRanges, entry.StartChar))
                    {
                        var replacement = referenceNewName;
                        if (!IsMemberAccessToken(codePart, entry.StartChar) && !string.IsNullOrWhiteSpace(qualifier))
                            replacement = $"{qualifier}.{referenceNewName}";

                        entry.RefModule.Replaces.AddReplace(
                            entry.LineNumber,
                            entry.StartChar,
                            entry.StartChar + entry.OldName.Length,
                            foundText,
                            replacement,
                            entry.Category + "_Reference");
                        return;
                    }
                }
            }

            AddModuleMemberReferenceReplaces(entry.RefModule, codePart, entry.LineNumber, entry.OldName, referenceNewName,
                entry.Category, entry.OccurrenceIndex, entry.StartChar, stringRanges, allowStringReplace, qualifier, sourceModule, sourceModuleReferenceName);
            return;
        }

        if (entry.Source is VbControl && Regex.IsMatch(codePart.TrimStart(), @"^Begin\s+\S+\s+", RegexOptions.IgnoreCase))
        {
            var pattern = $@"(?<=^.*Begin\s+\S+\s+){Regex.Escape(entry.OldName)}\b";
            var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                if (!allowStringReplace && IsInsideStringLiteral(stringRanges, match.Index))
                    continue;

                entry.RefModule.Replaces.AddReplace(
                    entry.LineNumber,
                    match.Index,
                    match.Index + match.Length,
                    match.Value,
                    referenceNewName,
                    entry.Category + "_Reference");
            }
            return;
        }

        if (string.Equals(entry.Category, "Constant", StringComparison.OrdinalIgnoreCase))
        {
            AddConstantReferenceReplaces(entry.RefModule, codePart, entry.LineNumber, entry.OldName, referenceNewName, entry.Category,
                entry.OccurrenceIndex, entry.StartChar, stringRanges, allowStringReplace);
            return;
        }

        var effectiveOldName = entry.OldName;
        if (string.Equals(entry.Category, "Type", StringComparison.OrdinalIgnoreCase))
        {
            var alternateOldName = GetTypeAlternateName(entry.OldName);
            if (!string.IsNullOrEmpty(alternateOldName) &&
                !Regex.IsMatch(codePart, $@"\b{Regex.Escape(entry.OldName)}\b", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(codePart, $@"\b{Regex.Escape(alternateOldName)}\b", RegexOptions.IgnoreCase))
            {
                effectiveOldName = alternateOldName;
            }
        }

        if (entry.StartChar >= 0 && entry.StartChar < codePart.Length)
        {
            var length = Math.Min(effectiveOldName.Length, codePart.Length - entry.StartChar);
            var foundText = codePart.Substring(entry.StartChar, length);
            if (string.Equals(foundText, effectiveOldName, StringComparison.OrdinalIgnoreCase))
            {
                if (allowStringReplace || !IsInsideStringLiteral(stringRanges, entry.StartChar))
                {
                    entry.RefModule.Replaces.AddReplace(
                        entry.LineNumber,
                        entry.StartChar,
                        entry.StartChar + effectiveOldName.Length,
                        foundText,
                        referenceNewName,
                        entry.Category + "_Reference");
                }
                return;
            }
        }

        entry.RefModule.Replaces.AddReplaceFromLine(codePart, entry.LineNumber, effectiveOldName, referenceNewName,
            entry.Category + "_Reference", entry.OccurrenceIndex, skipStringLiterals: !allowStringReplace);
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
            var procedureProp = reference?.GetType().GetProperty("Procedure");
            var occurrenceIndexesProp = reference?.GetType().GetProperty("OccurrenceIndexes");
            var startCharsProp = reference?.GetType().GetProperty("StartChars");

            if (lineNumbersProp?.GetValue(reference) is not System.Collections.Generic.List<int> refLineNumbers)
                continue;

            var occurrenceIndexes = occurrenceIndexesProp?.GetValue(reference) as System.Collections.Generic.List<int>;
            var startChars = startCharsProp?.GetValue(reference) as System.Collections.Generic.List<int>;

            for (int idx = 0; idx < refLineNumbers.Count; idx++)
            {
                var lineNum = refLineNumbers[idx];
                if (lineNum <= 0 || lineNum > lines.Length)
                    continue;

                var line = lines[lineNum - 1];
                var (codePart, _) = SplitCodeAndComment(line);

                var occIndex = (occurrenceIndexes != null && idx < occurrenceIndexes.Count) ? occurrenceIndexes[idx] : -1;
                var startChar = (startChars != null && idx < startChars.Count) ? startChars[idx] : -1;

                if (occIndex < 0 && startChar < 0)
                    continue;

                var sourceModule = GetDefiningModule(project, source!);
                var sourceModuleInfo = project.Modules.FirstOrDefault(m =>
                    string.Equals(m.Name, sourceModule, StringComparison.OrdinalIgnoreCase));
                var sourceModuleReferenceName = sourceModuleInfo?.ConventionalName ?? sourceModule;

                int replacesBefore = refModule.Replaces.Count;

                var referenceNewName = string.IsNullOrEmpty(referenceNewNameOverride) ? newName : referenceNewNameOverride;
                var stringRanges = GetStringLiteralRanges(codePart);
                var trimmedCodePart = codePart.TrimStart();
                var allowStringReplace = trimmedCodePart.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase) ||
                                         trimmedCodePart.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

                if (category == "EnumValue" && referenceNewName.Contains('.', StringComparison.Ordinal))
                {
                    AddEnumValueReferenceReplaces(refModule, codePart, lineNum, oldName, newName, referenceNewName, category, occIndex, startChar, stringRanges);
                    continue;
                }

                if (source is VbProperty)
                {
                    var shouldQualify = ShouldQualifyPropertyReference(refModule, lineNum, referenceNewName);
                    var propertyQualifier = GetPropertyQualifier(sourceModule, sourceModuleReferenceName, refModule, refModuleName, shouldQualify);
                    AddPropertyReferenceReplaces(refModule, codePart, lineNum, oldName, referenceNewName, category, occIndex, startChar, stringRanges, allowStringReplace, propertyQualifier);
                }
                else if (source is VbVariable variable && IsGlobalVariable(project, sourceModule, variable))
                {
                    var shouldQualify = ShouldQualifyModuleMemberReference(refModule, lineNum, referenceNewName);
                    var qualifier = shouldQualify ? sourceModuleReferenceName : null;
                    if (startChar >= 0 && startChar < codePart.Length)
                    {
                        var length = Math.Min(oldName.Length, codePart.Length - startChar);
                        var foundText = codePart.Substring(startChar, length);
                        if (string.Equals(foundText, oldName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (allowStringReplace || !IsInsideStringLiteral(stringRanges, startChar))
                            {
                                var replacement = referenceNewName;
                                if (!IsMemberAccessToken(codePart, startChar) && !string.IsNullOrWhiteSpace(qualifier))
                                    replacement = $"{qualifier}.{referenceNewName}";

                                refModule.Replaces.AddReplace(
                                    lineNum,
                                    startChar,
                                    startChar + oldName.Length,
                                    foundText,
                                    replacement,
                                    category + "_Reference");
                                continue;
                            }
                        }
                    }

                    AddModuleMemberReferenceReplaces(refModule, codePart, lineNum, oldName, referenceNewName, category, occIndex, startChar, stringRanges, allowStringReplace, qualifier, sourceModule, sourceModuleReferenceName);
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
                    AddConstantReferenceReplaces(refModule, codePart, lineNum, oldName, referenceNewName, category, occIndex, startChar, stringRanges, allowStringReplace);
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

                    if (startChar >= 0 && startChar < codePart.Length)
                    {
                        var length = Math.Min(effectiveOldName.Length, codePart.Length - startChar);
                        var foundText = codePart.Substring(startChar, length);
                        if (string.Equals(foundText, effectiveOldName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (allowStringReplace || !IsInsideStringLiteral(stringRanges, startChar))
                            {
                                refModule.Replaces.AddReplace(
                                    lineNum,
                                    startChar,
                                    startChar + effectiveOldName.Length,
                                    foundText,
                                    referenceNewName,
                                    category + "_Reference");
                            }
                            continue;
                        }
                    }

                    refModule.Replaces.AddReplaceFromLine(codePart, lineNum, effectiveOldName, referenceNewName, category + "_Reference", occIndex, skipStringLiterals: !allowStringReplace);
                }

                if (startChar >= 0 && refModule.Replaces.Count == replacesBefore)
                {
                    var refProcedureName = procedureProp?.GetValue(reference) as string;
                    var key = $"{refModule.Name}|{refProcedureName}|{lineNum}|{startChar}|{oldName}|{category}";
                    StartCharChecks.TryAdd(key, new StartCharCheckEntry
                    {
                        Module = refModule.Name,
                        Procedure = refProcedureName,
                        LineNumber = lineNum,
                        StartChar = startChar,
                        OccurrenceIndex = occIndex,
                        OldName = oldName,
                        NewName = referenceNewName,
                        Category = category
                    });
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
        int startChar,
        List<(int start, int end)> stringRanges)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (startChar >= 0)
        {
            var preciseMatch = matches.Cast<Match>().FirstOrDefault(m => m.Index == startChar);
            if (preciseMatch != null)
                targetMatches = new[] { preciseMatch };
        }
        else if (occIndex > 0 && occIndex <= matches.Count)
        {
            targetMatches = new[] { matches[occIndex - 1] };
        }

        foreach (var match in targetMatches)
        {
            if (IsInsideStringLiteral(stringRanges, match.Index))
                continue;

            var replacement = IsQualifiedEnumReference(codePart, match.Index) ? newName : qualifiedName;
            if (!string.Equals(replacement, newName, StringComparison.OrdinalIgnoreCase))
            {
                var ownerName = GetEnumOwnerFromQualifiedName(qualifiedName);
                if (!string.IsNullOrWhiteSpace(ownerName) &&
                    !ContainsTokenInLine(codePart, ownerName))
                {
                    replacement = newName;
                }
            }
            if (string.Equals(match.Value, replacement, StringComparison.Ordinal))
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

    private static string? GetEnumOwnerFromQualifiedName(string qualifiedName)
    {
        if (string.IsNullOrWhiteSpace(qualifiedName))
            return null;

        var dotIndex = qualifiedName.IndexOf('.');
        if (dotIndex <= 0)
            return null;

        return qualifiedName.Substring(0, dotIndex);
    }

    private static bool ContainsTokenInLine(string line, string token)
    {
        if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(token))
            return false;

        return Regex.IsMatch(line, $@"\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase);
    }

    private static void AddConstantReferenceReplaces(
        VbModule refModule,
        string codePart,
        int lineNum,
        string oldName,
        string newName,
        string category,
        int occIndex,
        int startChar,
        List<(int start, int end)> stringRanges,
        bool allowStringReplace)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (startChar >= 0)
        {
            var preciseMatch = matches.Cast<Match>().FirstOrDefault(m => m.Index == startChar);
            if (preciseMatch != null)
                targetMatches = new[] { preciseMatch };
        }
        else if (occIndex > 0 && occIndex <= matches.Count)
        {
            targetMatches = new[] { matches[occIndex - 1] };
        }

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
        int startChar,
        List<(int start, int end)> stringRanges,
        bool allowStringReplace,
        string? qualifier)
    {
        var pattern = $@"\b{Regex.Escape(oldName)}\b";
        var matches = Regex.Matches(codePart, pattern, RegexOptions.IgnoreCase);

        if (matches.Count == 0)
            return;

        IEnumerable<Match> targetMatches = matches.Cast<Match>();
        if (startChar >= 0)
        {
            var preciseMatch = matches.Cast<Match>().FirstOrDefault(m => m.Index == startChar);
            if (preciseMatch != null)
                targetMatches = new[] { preciseMatch };
        }
        else if (occIndex > 0 && occIndex <= matches.Count)
        {
            targetMatches = new[] { matches[occIndex - 1] };
        }

        foreach (var match in targetMatches)
        {
            if (!allowStringReplace && IsInsideStringLiteral(stringRanges, match.Index))
                continue;

            var replacement = newName;
            if (!IsMemberAccessToken(codePart, match.Index) && !string.IsNullOrWhiteSpace(qualifier))
                replacement = $"{qualifier}.{newName}";

            if (string.Equals(match.Value, replacement, StringComparison.Ordinal))
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
        int startChar,
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
        if (startChar >= 0)
        {
            var preciseMatch = matches.Cast<Match>().FirstOrDefault(m => m.Index == startChar);
            if (preciseMatch != null)
                targetMatches = new[] { preciseMatch };
        }
        else if (occIndex > 0 && occIndex <= matches.Count)
        {
            targetMatches = new[] { matches[occIndex - 1] };
        }

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

            if (string.Equals(match.Value, replacement, StringComparison.Ordinal))
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

    private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
    {
        public static readonly ReferenceEqualityComparer Instance = new();

        public new bool Equals(object? x, object? y)
            => ReferenceEquals(x, y);

        public int GetHashCode(object obj)
            => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
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
