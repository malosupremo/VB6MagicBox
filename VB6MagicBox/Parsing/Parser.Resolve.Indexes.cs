using VB6MagicBox.Models;

namespace VB6MagicBox.Parsing;

public static partial class VbParser
{
    /// <summary>
    /// Contains all pre-built global indexes needed for the single-pass reference resolution.
    /// Built once before scanning, reused for every module/procedure.
    /// </summary>
    internal sealed class GlobalIndexes
    {
        /// <summary>Procedure (Sub/Function/Declare, NOT Property) by name → list of (Module, Proc).</summary>
        public Dictionary<string, List<(string Module, VbProcedure Proc)>> ProcIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Property (Get/Let/Set) by name → list of (Module, Prop).</summary>
        public Dictionary<string, List<(string Module, VbProperty Prop)>> PropIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>UDT (Type) by name.</summary>
        public Dictionary<string, VbTypeDef> TypeIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Enum by name → list (allows homonyms across modules, though rare).</summary>
        public Dictionary<string, List<VbEnumDef>> EnumDefIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Enum value name → list of VbEnumValue (multiple enums can have same-named values).</summary>
        public Dictionary<string, List<VbEnumValue>> EnumValueIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Reverse lookup: VbEnumValue → owning VbEnumDef.</summary>
        public Dictionary<VbEnumValue, VbEnumDef> EnumValueOwners { get; } = new();

        /// <summary>Class modules by name (file name without extension).</summary>
        public Dictionary<string, VbModule> ClassIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>All modules by VB_Name.</summary>
        public Dictionary<string, VbModule> ModuleByName { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Global constants (module-level) by name → list of (Module, Constant).</summary>
        public Dictionary<string, List<(string Module, VbConstant Constant)>> ConstantIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Global variables by name → list of (Module, Variable).</summary>
        public Dictionary<string, List<(string Module, VbVariable Variable)>> GlobalVarIndex { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Set of all enum value names (fast existence check).</summary>
        public HashSet<string> EnumValueNames { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Builds all global indexes from the parsed project in a single pass over the model.
    /// Must be called AFTER Parser.Core has populated all declarations.
    /// </summary>
    internal static GlobalIndexes BuildGlobalIndexes(VbProject project)
    {
        var idx = new GlobalIndexes();

        foreach (var mod in project.Modules)
        {
            // --- Module by VB_Name ---
            if (!string.IsNullOrEmpty(mod.Name))
                idx.ModuleByName.TryAdd(mod.Name, mod);

            // --- Class index ---
            if (mod.IsClass)
            {
                var className = Path.GetFileNameWithoutExtension(mod.Name);
                idx.ClassIndex.TryAdd(className, mod);

                // Add short name if namespaced (e.g., "PDxI.clsPDxI" → "clsPDxI")
                if (className.Contains('.'))
                {
                    var shortName = className.Split('.').Last();
                    idx.ClassIndex.TryAdd(shortName, mod);
                }
            }

            // --- Procedures (Sub/Function/Declare, NOT Property) ---
            foreach (var proc in mod.Procedures.Where(
                         p => !p.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase)))
            {
                if (!idx.ProcIndex.TryGetValue(proc.Name, out var list))
                {
                    list = new List<(string, VbProcedure)>();
                    idx.ProcIndex[proc.Name] = list;
                }
                list.Add((mod.Name, proc));
            }

            // --- Properties ---
            foreach (var prop in mod.Properties)
            {
                if (!idx.PropIndex.TryGetValue(prop.Name, out var propList))
                {
                    propList = new List<(string, VbProperty)>();
                    idx.PropIndex[prop.Name] = propList;
                }
                propList.Add((mod.Name, prop));
            }

            // --- Types (UDT) ---
            foreach (var typeDef in mod.Types)
            {
                if (!string.IsNullOrEmpty(typeDef.Name))
                    idx.TypeIndex.TryAdd(typeDef.Name, typeDef);
            }

            // --- Enums + Values ---
            foreach (var enumDef in mod.Enums)
            {
                if (!string.IsNullOrEmpty(enumDef.Name))
                {
                    if (!idx.EnumDefIndex.TryGetValue(enumDef.Name, out var enumList))
                    {
                        enumList = new List<VbEnumDef>();
                        idx.EnumDefIndex[enumDef.Name] = enumList;
                    }
                    enumList.Add(enumDef);
                }

                foreach (var val in enumDef.Values)
                {
                    if (string.IsNullOrEmpty(val.Name))
                        continue;

                    idx.EnumValueNames.Add(val.Name);
                    idx.EnumValueOwners[val] = enumDef;

                    if (!idx.EnumValueIndex.TryGetValue(val.Name, out var valList))
                    {
                        valList = new List<VbEnumValue>();
                        idx.EnumValueIndex[val.Name] = valList;
                    }
                    valList.Add(val);
                }
            }

            // --- Global constants ---
            foreach (var c in mod.Constants)
            {
                if (string.IsNullOrEmpty(c.Name))
                    continue;

                if (!idx.ConstantIndex.TryGetValue(c.Name, out var cList))
                {
                    cList = new List<(string, VbConstant)>();
                    idx.ConstantIndex[c.Name] = cList;
                }
                cList.Add((mod.Name, c));
            }

            // --- Global variables ---
            foreach (var v in mod.GlobalVariables)
            {
                if (string.IsNullOrEmpty(v.Name))
                    continue;

                if (!idx.GlobalVarIndex.TryGetValue(v.Name, out var vList))
                {
                    vList = new List<(string, VbVariable)>();
                    idx.GlobalVarIndex[v.Name] = vList;
                }
                vList.Add((mod.Name, v));
            }
        }

        return idx;
    }
}
