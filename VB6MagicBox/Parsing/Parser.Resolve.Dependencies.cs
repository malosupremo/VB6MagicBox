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
              else
              {
                var propAtLine = searchMod.Properties.FirstOrDefault(p => p.ContainsLine(lineNum));
                if (propAtLine != null)
                {
                  var hasParamWithSameName = propAtLine.Parameters.Any(p =>
                      p.Name.Equals(v.Name, StringComparison.OrdinalIgnoreCase));

                  if (hasParamWithSameName)
                    continue;

                  v.References.AddLineNumber(searchMod.Name, propAtLine.Name, lineNum);
                }
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
              else
              {
                var propAtLine = searchMod.Properties.FirstOrDefault(p => p.ContainsLine(lineNum));
                if (propAtLine != null)
                {
                  var hasParamWithSameName = propAtLine.Parameters.Any(p =>
                      p.Name.Equals(c.Name, StringComparison.OrdinalIgnoreCase));

                  if (hasParamWithSameName)
                    continue;

                  c.References.AddLineNumber(searchMod.Name, propAtLine.Name, lineNum);
                }
              }
            }
          }
        }
      }

      counter = 0;
      // Marca proprietà usate (modulo level) e traccia references
      foreach (var prop in mod.Properties)
      {
        // Progress inline per il parsing
        Console.Write($"\r      [Property {counter++}/{mod.Properties.Count}] {prop.Name}...".PadRight(Console.WindowWidth - 1));

        bool isPublic = string.IsNullOrEmpty(prop.Visibility) ||
                       prop.Visibility.Equals("Public", StringComparison.OrdinalIgnoreCase) ||
                       prop.Visibility.Equals("Global", StringComparison.OrdinalIgnoreCase);

        var modulesToSearch = isPublic
            ? project.Modules
            : new List<VbModule> { mod };

        foreach (var searchMod in modulesToSearch)
        {
          var searchLines = ReadAllLinesShared(searchMod.FullPath);
          int lineNum = 0;

          foreach (var line in searchLines)
          {
            lineNum++;
            if (line.IndexOf(prop.Name, StringComparison.OrdinalIgnoreCase) >= 0)
            {
              prop.Used = true;
              var procAtLine = searchMod.GetProcedureAtLine(lineNum);
              if (procAtLine != null)
              {
                var hasLocalWithSameName = procAtLine.LocalVariables.Any(lv =>
                    lv.Name.Equals(prop.Name, StringComparison.OrdinalIgnoreCase)) ||
                  procAtLine.Parameters.Any(p =>
                    p.Name.Equals(prop.Name, StringComparison.OrdinalIgnoreCase));

                if (hasLocalWithSameName)
                  continue;

                prop.References.AddLineNumber(searchMod.Name, procAtLine.Name, lineNum);
              }
              else
              {
                var propAtLine = searchMod.Properties.FirstOrDefault(p => p.ContainsLine(lineNum));
                if (propAtLine != null)
                {
                  var hasParamWithSameName = propAtLine.Parameters.Any(p =>
                      p.Name.Equals(prop.Name, StringComparison.OrdinalIgnoreCase));

                  if (hasParamWithSameName)
                    continue;

                  prop.References.AddLineNumber(searchMod.Name, propAtLine.Name, lineNum);
                }
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
              else
              {
                var propAtLine = searchMod.Properties.FirstOrDefault(p => p.ContainsLine(lineNum));
                if (propAtLine != null)
                {
                  var hasParamWithSameName = propAtLine.Parameters.Any(p =>
                      p.Name.Equals(c.Name, StringComparison.OrdinalIgnoreCase));

                  if (hasParamWithSameName)
                    continue;

                  c.References.AddLineNumber(searchMod.Name, propAtLine.Name, lineNum);
                }
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
