using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
  // -------------------------
  // ENTRY POINT PARSING
  // -------------------------

  public static VbProject ParseProjectFromVbp(string vbpPath)
  {
    var project = new VbProject { ProjectFile = vbpPath };
    var baseDir = Path.GetDirectoryName(vbpPath)!;

    var files = ParseVbpFile(vbpPath);

    int fileIndex = 0;
    int totalFiles = files.Count;

    foreach (var f in files)
    {
      fileIndex++;
      Console.Write($"\r   Parsing moduli: [{fileIndex}/{totalFiles}] {Path.GetFileName(f.Path)}...".PadRight(Console.WindowWidth - 1));

      var fullPath = Path.Combine(baseDir, f.Path);
      if (!File.Exists(fullPath))
        continue;

      var mod = ParseModule(fullPath, f.Kind);
      mod.IsSharedExternal = IsSharedExternalPath(f.Path);
      // set FullPath and relative Path
      mod.FullPath = fullPath;
      // Compute relative path with leading separator, relative to baseDir
      // Prefer simple substring when module is inside the vbp directory tree
      var baseDirFull = Path.GetFullPath(baseDir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
      var fullPathFull = Path.GetFullPath(fullPath);

      string rel;
      if (fullPathFull.StartsWith(baseDirFull + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
          fullPathFull.Equals(baseDirFull, StringComparison.OrdinalIgnoreCase))
      {
        rel = fullPathFull.Substring(baseDirFull.Length).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
      }
      else
      {
        rel = Path.GetRelativePath(baseDirFull, fullPathFull).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        if (!rel.StartsWith(Path.DirectorySeparatorChar) && !rel.StartsWith(Path.AltDirectorySeparatorChar))
          rel = Path.DirectorySeparatorChar + rel;
      }

      // Ensure leading separator
      if (!rel.StartsWith(Path.DirectorySeparatorChar) && !rel.StartsWith(Path.AltDirectorySeparatorChar))
        rel = Path.DirectorySeparatorChar + rel;

      mod.Path = rel;
      mod.Owner = project;  // Imposta il riferimento al progetto
      project.Modules.Add(mod);
    }

    if (totalFiles > 0)
      Console.WriteLine();

    return project;
  }

  private static Dictionary<string, string[]> BuildFileCache(VbProject project)
  {
    var cache = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

    foreach (var mod in project.Modules)
    {
      if (string.IsNullOrWhiteSpace(mod.FullPath) || !File.Exists(mod.FullPath))
        continue;

      if (!cache.ContainsKey(mod.FullPath))
        cache[mod.FullPath] = ReadAllLinesShared(mod.FullPath);
    }

    return cache;
  }

  // -------------------------
  // PARSING VBP
  // -------------------------

  private class VbpEntry
  {
    public string Kind { get; set; }
    public string Path { get; set; }
  }

  private static bool IsSharedExternalPath(string path)
  {
    if (string.IsNullOrWhiteSpace(path))
      return false;

    var normalized = path.Replace('/', '\\');
    return normalized.Contains("..\\..\\", StringComparison.Ordinal);
  }

  private static List<VbpEntry> ParseVbpFile(string vbpPath)
  {
    var list = new List<VbpEntry>();
    var lines = File.ReadAllLines(vbpPath);
    var baseDir = Path.GetDirectoryName(vbpPath)!;

    foreach (var raw in lines)
    {
      var line = raw.Trim();

      if (line.StartsWith("Form=", StringComparison.OrdinalIgnoreCase))
      {
        var parts = line.Substring("Form=".Length).Split(';');
        string path = null;

        if (parts.Length == 2)
          path = parts[1].Trim();
        else if (parts.Length == 1)
        {
          // Formato senza separatore: Form=Form\Restart.frm oppure Form=frmMain
          var name = parts[0].Trim();

          // Se ha già l'estensione, usalo direttamente
          if (name.EndsWith(".frm", StringComparison.OrdinalIgnoreCase))
            path = name;
          else
          {
            // Altrimenti aggiungere l'estensione
            var potentialPath = Path.Combine(baseDir, name + ".frm");
            if (File.Exists(potentialPath))
              path = name + ".frm";
          }
        }

        if (!string.IsNullOrEmpty(path))
          list.Add(new VbpEntry { Kind = "frm", Path = path });
      }
      else if (line.StartsWith("Module=", StringComparison.OrdinalIgnoreCase))
      {
        var parts = line.Substring("Module=".Length).Split(';');
        string path = null;

        if (parts.Length == 2)
          path = parts[1].Trim();
        else if (parts.Length == 1)
        {
          var name = parts[0].Trim();

          if (name.EndsWith(".bas", StringComparison.OrdinalIgnoreCase))
            path = name;
          else
          {
            var potentialPath = Path.Combine(baseDir, name + ".bas");
            if (File.Exists(potentialPath))
              path = name + ".bas";
          }
        }

        if (!string.IsNullOrEmpty(path))
          list.Add(new VbpEntry { Kind = "bas", Path = path });
      }
      else if (line.StartsWith("Class=", StringComparison.OrdinalIgnoreCase))
      {
        var parts = line.Substring("Class=".Length).Split(';');
        string path = null;

        if (parts.Length == 2)
          path = parts[1].Trim();
        else if (parts.Length == 1)
        {
          var name = parts[0].Trim();

          if (name.EndsWith(".cls", StringComparison.OrdinalIgnoreCase))
            path = name;
          else
          {
            var potentialPath = Path.Combine(baseDir, name + ".cls");
            if (File.Exists(potentialPath))
              path = name + ".cls";
          }
        }

        if (!string.IsNullOrEmpty(path))
          list.Add(new VbpEntry { Kind = "cls", Path = path });
      }
    }

    return list;
  }
}
