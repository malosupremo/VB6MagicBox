using System.Text.RegularExpressions;
using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    // ---------------------------------------------------------
    // REGEX COMPILATE PER HOT-PATH (usate nei loop di risoluzione)
    // ---------------------------------------------------------

    private static readonly Regex ReSetNew =
        new(@"Set\s+(\w+)\s*=\s*New\s+(\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReSetAlias =
        new(@"Set\s+(\w+)\s*=\s+(\w+(?:\.\w+)?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReWordBoundary =
        new(@"\b([A-Za-z_]\w*)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReObjectMethod =
        new(@"(\w+)\.(\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
}