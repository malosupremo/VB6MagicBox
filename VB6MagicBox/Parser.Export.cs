using System;
using System.IO;
using System.Linq;
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
  // ESPORTAZIONE JSON RENAME (solo elementi da rinominare)
  // ---------------------------------------------------------

  public static void ExportRenameJson(VbProject project, string outputPath)
  {
    var renameData = new
    {
      ProjectFile = project.ProjectFile,
      Modules = project.Modules
          .Where(m => !m.IsConventional || 
                      m.GlobalVariables.Any(v => !v.IsConventional) ||
                      m.Constants.Any(c => !c.IsConventional) ||
                      m.Types.Any(t => !t.IsConventional || t.Fields.Any(f => !f.IsConventional)) ||
                      m.Enums.Any(e => !e.IsConventional || e.Values.Any(v => !v.IsConventional)) ||
                      m.Controls.Any(c => !c.IsConventional) ||
                      m.Properties.Any(p => !p.IsConventional || p.Parameters.Any(param => !param.IsConventional)) ||
                      m.Procedures.Any(p => !p.IsConventional || p.Parameters.Any(param => !param.IsConventional) || p.LocalVariables.Any(lv => !lv.IsConventional)))
          .Select(m => new
          {
              Module = new { m.Name, m.ConventionalName },
              IsConventional = m.IsConventional,
              GlobalVariables = m.GlobalVariables
                  .Where(v => !v.IsConventional)
                  .Select(v => new { v.Name, v.ConventionalName })
                  .ToList(),
              Constants = m.Constants
                  .Where(c => !c.IsConventional)
                  .Select(c => new { c.Name, c.ConventionalName })
                  .ToList(),
              Types = m.Types
                  .Where(t => !t.IsConventional || t.Fields.Any(f => !f.IsConventional))
                  .Select(t => new
                  {
                      Type = new { t.Name, t.ConventionalName },
                      IsConventional = t.IsConventional,
                      Fields = t.Fields
                          .Where(f => !f.IsConventional)
                          .Select(f => new { f.Name, f.ConventionalName })
                          .ToList()
                  })
                  .ToList(),
              Enums = m.Enums
                  .Where(e => !e.IsConventional || e.Values.Any(v => !v.IsConventional))
                  .Select(e => new
                  {
                      Enum = new { e.Name, e.ConventionalName },
                      IsConventional = e.IsConventional,
                      Values = e.Values
                          .Where(v => !v.IsConventional)
                          .Select(v => new { v.Name, v.ConventionalName })
                          .ToList()
                  })
                  .ToList(),
              Controls = m.Controls
                  .Where(c => !c.IsConventional)
                  .Select(c => new { c.Name, c.ConventionalName })
                  .ToList(),
              Properties = m.Properties
                  .Where(p => !p.IsConventional || p.Parameters.Any(param => !param.IsConventional))
                  .Select(p => new
                  {
                      Property = new { p.Name, p.ConventionalName, p.Kind },
                      IsConventional = p.IsConventional,
                      Parameters = p.Parameters
                          .Where(param => !param.IsConventional)
                          .Select(param => new { param.Name, param.ConventionalName })
                          .ToList()
                  })
                  .ToList(),
              Procedures = m.Procedures
                  .Where(p => !p.IsConventional || p.Parameters.Any(param => !param.IsConventional) || p.LocalVariables.Any(lv => !lv.IsConventional))
                  .Select(p => new
                  {
                      Procedure = new { p.Name, p.ConventionalName },
                      IsConventional = p.IsConventional,
                      Parameters = p.Parameters
                          .Where(param => !param.IsConventional)
                          .Select(param => new { param.Name, param.ConventionalName })
                          .ToList(),
                      LocalVariables = p.LocalVariables
                          .Where(lv => !lv.IsConventional)
                          .Select(lv => new { lv.Name, lv.ConventionalName })
                          .ToList()
                  })
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

    var json = JsonSerializer.Serialize(renameData, options);
    File.WriteAllText(outputPath, json);
  }

  // ---------------------------------------------------------
  // ESPORTAZIONE CSV RENAME (file piatto per elementi che cambiano nome)
  // ---------------------------------------------------------

  public static void ExportRenameCsv(VbProject project, string outputPath)
  {
    var csvLines = new List<string>
    {
      "Name,ConventionalName,Type,Visibility,Module"  // Header CSV
    };

    foreach (var mod in project.Modules)
    {
      // Modulo stesso (se deve essere rinominato)
      if (!mod.IsConventional)
      {
        csvLines.Add($"\"{EscapeCsv(mod.Name)}\",\"{EscapeCsv(mod.ConventionalName)}\",\"Module\",\"Public\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Variabili globali che cambiano nome
      foreach (var variable in mod.GlobalVariables.Where(v => !v.IsConventional))
      {
        var visibility = string.IsNullOrEmpty(variable.Visibility) ? "Public" : variable.Visibility;
        csvLines.Add($"\"{EscapeCsv(variable.Name)}\",\"{EscapeCsv(variable.ConventionalName)}\",\"Variable\",\"{visibility}\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Costanti che cambiano nome
      foreach (var constant in mod.Constants.Where(c => !c.IsConventional))
      {
        var visibility = string.IsNullOrEmpty(constant.Visibility) ? "Public" : constant.Visibility;
        csvLines.Add($"\"{EscapeCsv(constant.Name)}\",\"{EscapeCsv(constant.ConventionalName)}\",\"Constant\",\"{visibility}\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Tipi che cambiano nome (VbTypeDef non ha Visibility, default "Public")
      foreach (var type in mod.Types.Where(t => !t.IsConventional))
      {
        csvLines.Add($"\"{EscapeCsv(type.Name)}\",\"{EscapeCsv(type.ConventionalName)}\",\"Type\",\"Public\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Campi dei tipi che cambiano nome (VbField non ha Visibility, default "Public")
      foreach (var type in mod.Types)
      {
        foreach (var field in type.Fields.Where(f => !f.IsConventional))
        {
          csvLines.Add($"\"{EscapeCsv(field.Name)}\",\"{EscapeCsv(field.ConventionalName)}\",\"Field\",\"Public\",\"{EscapeCsv(mod.Name)}\"");
        }
      }

      // Enum che cambiano nome (VbEnumDef non ha Visibility, default "Public")
      foreach (var enumDef in mod.Enums.Where(e => !e.IsConventional))
      {
        csvLines.Add($"\"{EscapeCsv(enumDef.Name)}\",\"{EscapeCsv(enumDef.ConventionalName)}\",\"Enum\",\"Public\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Valori enum che cambiano nome (VbEnumValue non ha Visibility, default "Public")
      foreach (var enumDef in mod.Enums)
      {
        foreach (var enumValue in enumDef.Values.Where(v => !v.IsConventional))
        {
          csvLines.Add($"\"{EscapeCsv(enumValue.Name)}\",\"{EscapeCsv(enumValue.ConventionalName)}\",\"EnumValue\",\"Public\",\"{EscapeCsv(mod.Name)}\"");
        }
      }

      // Controlli che cambiano nome (VbControl non ha Visibility, default "Public")
      foreach (var control in mod.Controls.Where(c => !c.IsConventional))
      {
        csvLines.Add($"\"{EscapeCsv(control.Name)}\",\"{EscapeCsv(control.ConventionalName)}\",\"Control\",\"Public\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Proprietà che cambiano nome
      foreach (var property in mod.Properties.Where(p => !p.IsConventional))
      {
        var visibility = string.IsNullOrEmpty(property.Visibility) ? "Public" : property.Visibility;
        csvLines.Add($"\"{EscapeCsv(property.Name)}\",\"{EscapeCsv(property.ConventionalName)}\",\"Property{property.Kind}\",\"{visibility}\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Parametri delle proprietà che cambiano nome
      foreach (var property in mod.Properties)
      {
        foreach (var parameter in property.Parameters.Where(p => !p.IsConventional))
        {
          csvLines.Add($"\"{EscapeCsv(parameter.Name)}\",\"{EscapeCsv(parameter.ConventionalName)}\",\"PropertyParameter\",\"Local\",\"{EscapeCsv(mod.Name)}\"");
        }
      }

      // Procedure che cambiano nome
      foreach (var procedure in mod.Procedures.Where(p => !p.IsConventional))
      {
        var visibility = string.IsNullOrEmpty(procedure.Visibility) ? "Public" : procedure.Visibility;
        csvLines.Add($"\"{EscapeCsv(procedure.Name)}\",\"{EscapeCsv(procedure.ConventionalName)}\",\"{procedure.Kind}\",\"{visibility}\",\"{EscapeCsv(mod.Name)}\"");
      }

      // Parametri delle procedure che cambiano nome (VbParameter non ha Visibility, default "Local")
      foreach (var procedure in mod.Procedures)
      {
        foreach (var parameter in procedure.Parameters.Where(p => !p.IsConventional))
        {
          csvLines.Add($"\"{EscapeCsv(parameter.Name)}\",\"{EscapeCsv(parameter.ConventionalName)}\",\"Parameter\",\"Local\",\"{EscapeCsv(mod.Name)}\"");
        }
      }

      // Variabili locali delle procedure che cambiano nome (VbVariable può non avere Visibility)
      foreach (var procedure in mod.Procedures)
      {
        foreach (var localVar in procedure.LocalVariables.Where(v => !v.IsConventional))
        {
          csvLines.Add($"\"{EscapeCsv(localVar.Name)}\",\"{EscapeCsv(localVar.ConventionalName)}\",\"LocalVariable\",\"Local\",\"{EscapeCsv(mod.Name)}\"");
        }
      }

      // Costanti locali delle procedure che cambiano nome (VbConstant può non avere Visibility)
      foreach (var procedure in mod.Procedures)
      {
        foreach (var localConst in procedure.Constants.Where(c => !c.IsConventional))
        {
          csvLines.Add($"\"{EscapeCsv(localConst.Name)}\",\"{EscapeCsv(localConst.ConventionalName)}\",\"LocalConstant\",\"Local\",\"{EscapeCsv(mod.Name)}\"");
        }
      }
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
  // ESPORTAZIONE MERMAID
  // ---------------------------------------------------------

  public static void ExportMermaid(VbProject project, string outputPath)
  {
    var sb = new StringBuilder();
    sb.AppendLine("graph TD");

    // Raggruppa dipendenze per modulo (caller -> callee)
    var moduleDependencies = project.Dependencies
        .Where(d => !string.IsNullOrEmpty(d.CallerModule) && !string.IsNullOrEmpty(d.CalleeModule))
        .Select(d => new { Caller = Sanitize(d.CallerModule), Callee = Sanitize(d.CalleeModule) })
        .Distinct()
        .OrderBy(d => d.Caller)
        .ThenBy(d => d.Callee);

    foreach (var dep in moduleDependencies)
    {
      sb.AppendLine($"    {dep.Caller} --> {dep.Callee}");
    }

    File.WriteAllText(outputPath, sb.ToString());
  }

  private static string Sanitize(string s)
  {
    return s.Replace('.', '_').Replace(' ', '_');
  }
}
