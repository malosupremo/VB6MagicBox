using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
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
            ConsoleX.WriteLineColor($"    [WARN] Impossibile leggere {Path.GetFileName(filePath)}: {ex.Message}", ConsoleColor.Yellow);
            return Array.Empty<string>();
        }
    }

    private static string[] GetFileLines(Dictionary<string, string[]> fileCache, VbModule mod)
    {
        if (fileCache != null && fileCache.TryGetValue(mod.FullPath, out var lines))
            return lines;

        return ReadAllLinesShared(mod.FullPath);
    }

    /// <summary>
    /// Costruisce il grafo delle dipendenze (Dependencies) e propaga i flag Used.
    /// Le References sono già state risolte dal SinglePass resolver.
    /// </summary>
    public static void BuildDependenciesAndUsage(VbProject project, Dictionary<string, string[]> fileCache)
    {
        var procByModuleAndName = new Dictionary<(string Module, string Name), VbProcedure>();
        foreach (var mod in project.Modules)
            foreach (var proc in mod.Procedures)
                procByModuleAndName[(mod.Name, proc.Name)] = proc;

        int moduleIndex = 0;
        int totalModules = project.Modules.Count;

        foreach (var mod in project.Modules)
        {
            moduleIndex++;
            var fileName = Path.GetFileName(mod.FullPath);
            var moduleName = Path.GetFileNameWithoutExtension(mod.Name);
            Console.Write($"\r  [{moduleIndex}/{totalModules}] {fileName} ({moduleName})...".PadRight(Console.WindowWidth - 1));

            // Costruisci Dependencies da Calls (già popolate dal SinglePass)
            foreach (var proc in mod.Procedures)
            {
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

                    // Marca classi usate (per chiamate object.method dove object è una classe)
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
        }

        Console.WriteLine();

        // Marca tipi usati (da dichiarazioni "As TypeName")
        MarkUsedTypes(project, fileCache);

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

            foreach (var proc in mod.Procedures) Collect(proc.References);
            foreach (var prop in mod.Properties) Collect(prop.References);
            foreach (var v in mod.GlobalVariables) Collect(v.References);
            foreach (var c in mod.Constants) Collect(c.References);
            foreach (var e in mod.Enums) { Collect(e.References); foreach (var val in e.Values) Collect(val.References); }
            foreach (var t in mod.Types) { Collect(t.References); foreach (var f in t.Fields) Collect(f.References); }
            foreach (var c in mod.Controls) Collect(c.References);
            foreach (var ev in mod.Events) Collect(ev.References);
            Collect(mod.References);

            mod.ModuleReferences = callers
                .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}
