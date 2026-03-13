using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;


public static partial class VbParser
{
    // ---------------------------------------------------------
    // ORDINAMENTO COMPLETO
    // ---------------------------------------------------------

    public static void SortProject(VbProject project)
    {
        // Moduli - ordinati per path (non per nome)
        // Applica le convenzioni di naming prima dell'export
        NamingConvention.Apply(project);

        project.Modules = project.Modules
            .OrderBy(m => m.Path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var mod in project.Modules)
        {
            // Procedure - ordina per nome (non per numero di riga)
            mod.Procedures = mod.Procedures
                .OrderBy(p => p.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Proprietà - ordina per nome
            mod.Properties = mod.Properties
                .OrderBy(p => p.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Variabili globali - ordina per nome
            mod.GlobalVariables = mod.GlobalVariables
                .OrderBy(v => v.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Costanti - ordina per nome
            mod.Constants = mod.Constants
                .OrderBy(c => c.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Tipi - ordina per nome
            mod.Types = mod.Types
                .OrderBy(t => t.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Enum - ordina per nome
            mod.Enums = mod.Enums
                .OrderBy(e => e.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Controlli form - ordina per nome
            mod.Controls = mod.Controls
                .OrderBy(c => c.ConventionalName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Campi dei Type - NON ordinare, mantengono l'ordine originale (come nel file)
            // foreach (var t in mod.Types)
            // {
            //   t.Fields rimangono nell'ordine di dichiarazione
            // }

            // Valori enum - rimangono ordinati alfabeticamente
            foreach (var e in mod.Enums)
            {
                e.Values = e.Values
                    .OrderBy(v => v.ConventionalName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // Variabili locali + costanti locali - ordinati per nome
            // Calls - non ordinare per mantenerli in ordine di apparizione senza ripetizioni
            // Parameters - NON ordinare! Mantengono l'ordine originale (è l'ordine dei parametri!)
            foreach (var p in mod.Procedures)
            {
                p.LocalVariables = p.LocalVariables
                    .OrderBy(v => v.ConventionalName, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                p.Constants = p.Constants
                    .OrderBy(c => c.ConventionalName, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // p.Parameters rimangono nell'ordine originale (cruciale!)
                // p.Calls rimangono nell'ordine di apparizione
            }
        }
    }

    // ---------------------------------------------------------
    // ESPORTAZIONE JSON
    // ---------------------------------------------------------

    public static void ExportJson(VbProject project, string outputPath)
    {
        // Filtra le calls non valide: bare calls senza MethodName risolto
        foreach (var mod in project.Modules)
        {
            foreach (var proc in mod.Procedures)
            {
                // Mantieni solo calls che hanno:
                // - ObjectName (es. gobjPlc.Timer) oppure
                // - MethodName con risoluzione (ResolvedModule/ResolvedProcedure/ResolvedKind non nulli)
                proc.Calls = proc.Calls
                    .Where(c => !string.IsNullOrEmpty(c.ObjectName) ||
                                (!string.IsNullOrEmpty(c.ResolvedModule) &&
                                 !string.IsNullOrEmpty(c.ResolvedProcedure)))
                    .ToList();

                // Deduplicare e ordinare References nelle procedure (Module + Procedure)
                proc.References = proc.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // Deduplicare e ordinare References nelle variabili globali (Module + Procedure)
            foreach (var variable in mod.GlobalVariables)
            {
                variable.References = variable.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // Deduplicare e ordinare References nelle costanti (Module + Procedure)
            foreach (var c in mod.Constants)
            {
                c.References = c.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // Deduplicare e ordinare References nei campi dei tipi (Module + Procedure)
            foreach (var type in mod.Types)
            {
                // References del Type stesso
                type.References = type.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var field in type.Fields)
                {
                    field.References = field.References
                        .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                        .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }

            // Deduplicare e ordinare References negli Enum (Module + Procedure)
            foreach (var enumDef in mod.Enums)
            {
                // References dell'Enum stesso
                enumDef.References = enumDef.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // References dei singoli valori enum
                foreach (var enumValue in enumDef.Values)
                {
                    enumValue.References = enumValue.References
                        .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                        .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }

            // Deduplicare e ordinare References nei controlli (Module + Procedure)
            foreach (var control in mod.Controls)
            {
                control.References = control.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // Deduplicare e ordinare References negli Eventi (Module + Procedure)
            foreach (var evt in mod.Events)
            {
                evt.References = evt.References
                    .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                    .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            // Deduplicare e ordinare References nelle procedure (per i parametri e variabili locali)
            foreach (var proc in mod.Procedures)
            {
                // Deduplicare References nei parametri
                foreach (var param in proc.Parameters)
                {
                    param.References = param.References
                        .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                        .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }

                // Deduplicare References nelle variabili locali
                foreach (var localVar in proc.LocalVariables)
                {
                    localVar.References = localVar.References
                        .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                        .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }

                // Deduplicare References nelle costanti locali
                foreach (var localConst in proc.Constants)
                {
                    localConst.References = localConst.References
                        .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                        .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }

            // Deduplicare e ordinare References del Modulo stesso (per le classi)
            mod.References = mod.References
                .DistinctBy(r => new { r.Module, r.Procedure, r.LineNumbers })
                .OrderBy(r => r.Module, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Procedure, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        var json = JsonSerializer.Serialize(project, options);
        File.WriteAllText(outputPath, json);
    }

    // ---------------------------------------------------------
    // ESPORTAZIONE CSV SHADOWS (locali che nascondono oggetti esterni)
    // ---------------------------------------------------------

    public static void ExportShadowsCsv(VbProject project, string outputPath)
    {
        var csvLines = new List<string>
    {
      "Module,LineNumber,Procedure,LocalKind,LocalName,LocalType,ShadowedKind,ShadowedName,ShadowedType,ShadowedModule"
    };

        var shadowEntries = new List<(string Module, int LineNumber, string Procedure, string LocalKind, string LocalName, string? LocalType, string ShadowKind, string ShadowName, string? ShadowType, string ShadowModule)>();

        static bool IsPrivateVisibility(string? visibility)
        {
            return string.Equals(visibility, "Private", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(visibility, "Dim", StringComparison.OrdinalIgnoreCase);
        }

        static IEnumerable<string> GetNameCandidates(string? name, string? conventionalName)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(name))
                names.Add(name);
            if (!string.IsNullOrWhiteSpace(conventionalName))
                names.Add(conventionalName);
            return names;
        }

        var externalSymbols = new List<(string Module, string Kind, string Name, string? Type)>();

        foreach (var mod in project.Modules)
        {
            foreach (var v in mod.GlobalVariables.Where(v => !IsPrivateVisibility(v.Visibility)))
            {
                foreach (var name in GetNameCandidates(v.Name, v.ConventionalName))
                    externalSymbols.Add((mod.Name, "GlobalVariable", name, v.Type));
            }

            foreach (var c in mod.Constants.Where(c => !IsPrivateVisibility(c.Visibility)))
            {
                foreach (var name in GetNameCandidates(c.Name, c.ConventionalName))
                    externalSymbols.Add((mod.Name, "Constant", name, c.Type));
            }

            if (!mod.IsClass)
            {
                foreach (var p in mod.Properties.Where(p => !IsPrivateVisibility(p.Visibility)))
                {
                    foreach (var name in GetNameCandidates(p.Name, p.ConventionalName))
                        externalSymbols.Add((mod.Name, $"Property{p.Kind}", name, p.ReturnType));
                }
            }

            foreach (var ctrl in mod.Controls)
            {
                foreach (var name in GetNameCandidates(ctrl.Name, ctrl.ConventionalName))
                    externalSymbols.Add((mod.Name, "Control", name, ctrl.ControlType));
            }
        }

        var rowKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mod in project.Modules)
        {
            foreach (var proc in mod.Procedures)
            {
                var locals = new List<(string Kind, string Name, int LineNumber, string? Type)>
                {
                    // Parameters
                };

                foreach (var param in proc.Parameters)
                {
                    foreach (var name in GetNameCandidates(param.Name, param.ConventionalName))
                        locals.Add(("Parameter", name, param.LineNumber > 0 ? param.LineNumber : proc.LineNumber, param.Type));
                }

                foreach (var localVar in proc.LocalVariables)
                {
                    foreach (var name in GetNameCandidates(localVar.Name, localVar.ConventionalName))
                        locals.Add(("LocalVariable", name, localVar.LineNumber > 0 ? localVar.LineNumber : proc.LineNumber, localVar.Type));
                }

                foreach (var localConst in proc.Constants)
                {
                    foreach (var name in GetNameCandidates(localConst.Name, localConst.ConventionalName))
                        locals.Add(("LocalConstant", name, localConst.LineNumber > 0 ? localConst.LineNumber : proc.LineNumber, localConst.Type));
                }

                foreach (var (localKind, localName, localLineNumber, localType) in locals)
                {
                    foreach (var (shadowModule, shadowKind, shadowName, shadowType) in externalSymbols)
                    {
                        if (string.Equals(shadowModule, mod.Name, StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (!string.Equals(localName, shadowName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var key = $"{mod.Name}|{localLineNumber}|{proc.Name}|{localKind}|{localName}|{localType}|{shadowKind}|{shadowName}|{shadowType}|{shadowModule}";
                        if (!rowKeys.Add(key))
                            continue;

                        shadowEntries.Add((
                          mod.Name,
                          localLineNumber,
                          proc.Name,
                          localKind,
                          localName,
                          localType,
                          shadowKind,
                          shadowName,
                          shadowType,
                          shadowModule));
                    }
                }
            }
        }

        foreach (var entry in shadowEntries
            .OrderBy(e => e.Module, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.LineNumber)
            .ThenBy(e => e.Procedure, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.LocalKind, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.LocalName, StringComparer.OrdinalIgnoreCase))
        {
            csvLines.Add(
              $"\"{EscapeCsv(entry.Module)}\",{entry.LineNumber},\"{EscapeCsv(entry.Procedure)}\",\"{EscapeCsv(entry.LocalKind)}\",\"{EscapeCsv(entry.LocalName)}\",\"{EscapeCsv(entry.LocalType ?? string.Empty)}\",\"{EscapeCsv(entry.ShadowKind)}\",\"{EscapeCsv(entry.ShadowName)}\",\"{EscapeCsv(entry.ShadowType ?? string.Empty)}\",\"{EscapeCsv(entry.ShadowModule)}\"");
        }

        File.WriteAllText(outputPath, string.Join(Environment.NewLine, csvLines));
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        // Escape delle virgolette doppie nel CSV (raddoppiarle)
        return value.Replace("\"", "\"\"");
    }

    // ---------------------------------------------------------
    // ESPORTAZIONE LINE REPLACE JSON
    // ---------------------------------------------------------

    /// <summary>
    /// Esporta il file .linereplace.json contenente tutte le sostituzioni
    /// precise (riga + posizione carattere) da applicare per il refactoring.
    /// Utile per verifica manuale e debugging.
    /// </summary>
    public static void ExportLineReplaceJson(VbProject project, string outputPath)
    {
        var replaceData = new
        {
            ProjectFile = project.ProjectFile,
            TotalReplaces = project.Modules.Sum(m => m.Replaces.Count),
            Modules = project.Modules
              .Where(m => m.Replaces.Count > 0)
              .Select(m => new
              {
                  ModuleName = m.Name,
                  FilePath = m.Path,
                  ReplaceCount = m.Replaces.Count,
                  Replaces = m.Replaces
                      .OrderBy(r => r.LineNumber)
                      .ThenBy(r => r.StartChar)
                      .ToList()
              })
              .ToList()
        };

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        var json = JsonSerializer.Serialize(replaceData, options);
        File.WriteAllText(outputPath, json);
    }
    // ---------------------------------------------------------
    // ESPORTAZIONE CSV TODO ENUM PREFIX
    // ---------------------------------------------------------

    public static void ExportEnumPrefixTodoCsv(VbProject project, string outputPath)
    {
        var csvLines = new List<string>
    {
      "Module,Path,LineNumber,OldText,NewText,Category"
    };

        foreach (var mod in project.Modules)
        {
            foreach (var replace in mod.Replaces)
            {
                if (!string.Equals(replace.Category, "EnumValue_Reference", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (string.IsNullOrEmpty(replace.NewText) || !replace.NewText.Contains('.'))
                    continue;

                if (!string.IsNullOrEmpty(replace.OldText) && replace.OldText.Contains('.'))
                    continue;

                csvLines.Add(
                  $"\"{EscapeCsv(mod.Name)}\"," +
                  $"\"{EscapeCsv(mod.Path ?? string.Empty)}\"," +
                  $"{replace.LineNumber}," +
                  $"\"{EscapeCsv(replace.OldText ?? string.Empty)}\"," +
                  $"\"{EscapeCsv(replace.NewText ?? string.Empty)}\"," +
                  $"\"{EscapeCsv(replace.Category ?? string.Empty)}\"");
            }
        }

        File.WriteAllText(outputPath, string.Join(Environment.NewLine, csvLines));
    }

    // ---------------------------------------------------------
    // ESPORTAZIONE MERMAID
    // ---------------------------------------------------------

    public static void ExportMermaid(VbProject project, string outputPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("```mermaid");
        sb.AppendLine("graph TD");

        // Lookup Name → ConventionalName per visualizzare i nomi convenzionali nel grafo
        var conventionalName = project.Modules
            .ToDictionary(m => m.Name, m => m.ConventionalName ?? m.Name, StringComparer.OrdinalIgnoreCase);

        // Costruisce gli archi dal nuovo ModuleReferences: ogni entry è un caller del modulo callee
        var edges = project.Modules
            .Where(mod => mod.ModuleReferences.Count > 0)
            .SelectMany(mod => mod.ModuleReferences
                .Select(caller => new
                {
                    Caller = Sanitize(conventionalName.GetValueOrDefault(caller, caller)),
                    Callee = Sanitize(mod.ConventionalName ?? mod.Name)
                }))
            .Distinct()
            .OrderBy(e => e.Caller, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.Callee, StringComparer.OrdinalIgnoreCase);

        foreach (var edge in edges)
        {
            sb.AppendLine($"    {edge.Caller} --> {edge.Callee}");
        }

        sb.AppendLine("```");

        File.WriteAllText(outputPath, sb.ToString());
    }

    private static string Sanitize(string s)
    {
        return s.Replace('.', '_').Replace(' ', '_');
    }
}
