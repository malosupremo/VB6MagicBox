# VB6MagicBox - Technical Summary to start a new chat with GitHub Copilot

> This file is the **single source of truth** for starting new Copilot sessions.
> It must be updated after every code modification.

## Purpose
VB6 parser/refactoring tool (.NET 10 console app). Pipeline: parse VB6 `.vbp` project → resolve references (types/calls/fields/controls) → build dependencies/usage → apply naming conventions → **build precise replaces** → export outputs → apply refactoring → annotate types → reorder variables → harmonize spacing.

## Project Structure
```
VB6MagicBox/
├── Program.cs                           — CLI entry: interactive menu (6 options) + command-line mode
├── Refactoring.cs                       — Phase 2: apply pre-calculated renames (right-to-left)
├── TypeAnnotator.cs                     — Phase 3: add missing types, visibility, Call/Step cleanup
├── CodeFormatter.cs                     — Phase 4+5: ReorderLocalVariables + HarmonizeSpacing
├── ConsoleX.cs                          — Console color helpers (WriteLineSuccess/Error/Warning/Color)
├── Models/
│   ├── Models.cs                        — VbProject, VbModule, VbControl, VbVariable, VbConstant, VbTypeDef, VbEnumDef, LineReplace
│   ├── Models.Members.cs                — VbProcedure, VbProperty, VbParameter, VbCall, VbReference, VbEnumValue, VbField
│   └── Models.Data.cs                   — VbReferenceListExtensions (AddLineNumber dedup, AddReplace dedup, ResetReferenceDebugEntries)
├── Parsing/
│   ├── Parser.cs                        — Facade: ParseAndResolve (5-step pipeline), ParseResolveAndExport
│   ├── Parser.Core.cs                   — VbKeywords HashSet, regex definitions (ReFieldAccess, ReSetNew, ReSetAlias, etc.)
│   ├── Parser.Core.Project.cs           — ParseProjectFromVbp: reads .vbp file, iterates modules
│   ├── Parser.Core.Module.cs            — ParseModuleFile: parses procedures/properties/controls/types/enums/globals/locals
│   ├── Parser.Core.Params.cs            — ParseParameters: extracts parameter list from procedure signatures
│   ├── Parser.Resolve.cs                — Legacy file (minimal), BuildFileCache, GetFileLines
│   ├── Parser.Resolve.SinglePass.cs     — ★ Core: ResolveTypesAndCalls, ResolveProcedureBody (STEP 1+2), ResolveChain
│   ├── Parser.Resolve.Indexes.cs        — GlobalIndexes class + BuildGlobalIndexes
│   ├── Parser.Resolve.Members.cs        — EnumerateDotChains, EnumerateTokens, FindTokenInRawLine, GetTokenStartChar, MaskStringLiterals, StripInlineComment
│   ├── Parser.Resolve.Helpers.cs        — MarkControlAsUsed, RecordReference, utility helpers
│   ├── Parser.Resolve.Types.cs          — ResolveTypeReferences, ResolveClassModuleReferences, MarkUsedTypes
│   ├── Parser.Resolve.Dependencies.cs   — BuildDependenciesAndUsage
│   ├── Parser.Replaces.cs               — ★ BuildReplaces: pre-calculates all LineReplace entries per module
│   ├── Parser.Naming.cs                 — Naming convention detection logic
│   ├── Parser.Naming.Apply.cs           — NamingConvention.Apply: assigns ConventionalName to all symbols
│   ├── Parser.Naming.Utilities.cs       — Naming helper functions
│   └── Parser.Export.cs                 — ExportJson, ExportRenameJson/Csv, ExportMermaid, ExportLineReplaceJson, etc.
```

## Architecture — Full Pipeline
**Phase 1 (Analysis)**: `VbParser.ParseAndResolve` — `Parser.cs`
1. **Step 1/5 — Parsing**: `ParseProjectFromVbp` → `ParseModuleFile` per module (`Parser.Core.Project` + `Parser.Core.Module`)
2. **Step 2/5 — Resolution**: `ResolveTypesAndCalls` — single-pass reference resolution (`Parser.Resolve.SinglePass` + `.Indexes` + `.Members` + `.Helpers` + `.Types`)
3. **Step 3/5 — Dependencies**: `BuildDependenciesAndUsage` — cross-module dependency graph + mark `Used` symbols (`Parser.Resolve.Dependencies`)
4. **Step 4/5 — Naming & Sort**: `SortProject` → `NamingConvention.Apply` — assigns `ConventionalName` to all symbols, sorts alphabetically (`Parser.Naming.Apply` + `Parser.Export`)
5. **Step 5/5 — Build Replaces**: `BuildReplaces` — pre-calculates exact character positions for all renames (`Parser.Replaces`)

**Phase 2 (Refactoring)**: `Refactoring.ApplyRenames` — `Refactoring.cs`
- ⚡ Ultra-fast: uses pre-calculated `VbModule.Replaces[]` (line + char position)
- No regex matching, no re-parsing — just mechanical right-to-left substitution
- Validates each replace (checks OldText still matches at StartChar)
- Backup only modified files (`.backup{timestamp}` folder)
- Uses Windows-1252 (ANSI) encoding for VB6 file I/O

**Phase 3 (Type Annotation)**: `TypeAnnotator.AddMissingTypes` — `TypeAnnotator.cs`
- Runs AFTER refactoring (uses conventional names already applied)
- Adds cleanup: visibility normalization (`Public`/`Private`) in forms/classes, `Call` removal, `For ... Step 1` cleanup

**Phase 4 (Variable Reorder)**: `CodeFormatter.ReorderLocalVariables` — `CodeFormatter.cs`
- Moves `Dim`/`Static`/`Const` declarations to top of each procedure body
- Groups: comments → constants → static → Dim; keeps `Attribute` lines attached

**Phase 5 (Spacing)**: `CodeFormatter.HarmonizeSpacing` — `CodeFormatter.cs`
- Normalizes blank lines per spacing rules (S0, S7, S8, S14, single-line If, property adjacency, etc.)

**Magic Wand (option 6)**: runs all 5 phases sequentially on a single `ParseAndResolve` result.

## Key Models (`Models/`)
- **`VbProject`**: `ProjectFile`, `Modules[]`, `Dependencies[]`
- **`VbModule`**: `Name`, `ConventionalName`, `Kind` (`bas`/`cls`/`frm`), `FullPath`, `IsSharedExternal`, `Used`, `IsClass`/`IsForm` (computed), `ImplementsInterfaces[]`
  - Collections: `Procedures[]`, `Properties[]` (separate), `Types[]`, `Enums[]`, `Constants[]`, `GlobalVariables[]`, `Controls[]`, `References[]`, `ModuleReferences[]`
  - **`Replaces[]`** (`List<LineReplace>`) — all substitutions for this module, ordered by line (desc) + char (desc)
  - Methods: `GetProcedureAtLine(lineNum)` → uses `ContainsLine()` for accurate procedure identification
- **`VbProcedure`**: `Name`, `ConventionalName`, `Kind` (Sub/Function/Declare), `StartLine`/`EndLine`, `LineNumber`, `ReturnType`, `Scope`, `Visibility`, `Used`, `IsStatic`
  - Collections: `Parameters[]`, `LocalVariables[]`, `Constants[]`, `Calls[]`, `References[]`
  - Method: `ContainsLine(lineNumber)` — checks `lineNumber >= StartLine && lineNumber <= EndLine`
- **`VbProperty`**: same structure as procedure, `Kind` (Get/Let/Set), `ReturnType`, `Parameters[]`, `References[]`
- **`VbTypeDef`** / **`VbField`**: UDT definitions with `References[]`
- **`VbEnumDef`** / **`VbEnumValue`**: enum definitions with `References[]`
- **`VbControl`**: `Name`, `ConventionalName`, `ControlType`, `LineNumbers[]` (for control arrays), `References[]`
- **`VbReference`**: tracks `LineNumbers[]` + `StartChars[]` + `OccurrenceIndexes[]` for exact replace positions
- **`LineReplace`**: `LineNumber`, `StartChar`, `EndChar`, `OldText`, `NewText`, `Category` — precise substitution with exact character position

### Key Deduplication Points (`Models.Data.cs`)
- `AddLineNumber` on `VbReference`: deduplicates by (line, startChar) pairs
- `AddReplace` on `VbModule`: deduplicates by (line, startChar) — prevents overlapping substitutions
- `replaceEntryKeys` ConcurrentDictionary in `BuildReplaces`: prevents parallel-generated duplicates
- `recorded` HashSet in `ResolveProcedureBody`: prevents same (line, startChar) being recorded twice

## Parsing (`Parser.Core` + `Parser.Core.Module`)
- Handles `Function`/`Sub`/`Property` and `Declare` (supports `Alias`, optional visibility)
- Collapses line continuations `_` with line mapping; multiline signature parameter line numbers are corrected
- **Multi-declaration Dim**: `Dim a, b, c As Integer` parsed as three separate `VbVariable` entries using `^\s*(Dim|Static)\s+(.*)$` regex gate
- Regexes for array parens use `\([^)]*\)`; `ReFieldAccess` pattern: `([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)`
- `Implements` lines recorded in `VbModule.ImplementsInterfaces`
- Form control parsing: handles indented `Begin` blocks and nested `controlDepth` for container controls

## Resolution (Single-Pass Architecture)

### Overview
`ResolveTypesAndCalls` in `Parser.Resolve.SinglePass.cs` is the **single-pass** resolver.
Supporting files:
- `Parser.Resolve.Indexes.cs` — `GlobalIndexes` class
- `Parser.Resolve.Members.cs` — chain detection, token enumeration, position mapping
- `Parser.Resolve.Helpers.cs` — `MarkControlAsUsed`, `RecordReference`, utility
- `Parser.Resolve.Types.cs` — post-processing: type/class declaration references

### GlobalIndexes (`Parser.Resolve.Indexes.cs`)
Built **once** via `BuildGlobalIndexes(project)`, reused for every module/procedure:
- `ProcIndex` — Procedure (Sub/Function/Declare, NOT Property) by name → list of (Module, Proc)
- `PropIndex` — Property (Get/Let/Set) by name → list of (Module, Prop)
- `TypeIndex` — UDT (Type) by name
- `EnumDefIndex` — Enum by name → list
- `EnumValueIndex` — Enum value name → list of VbEnumValue
- `EnumValueOwners` — Reverse lookup: VbEnumValue → owning VbEnumDef
- `ClassIndex` — Class modules **+ forms** by name (forms added for `Dim obj As FrmXxx`)
- `ModuleByName` — All modules by VB_Name
- `ConstantIndex` — Global constants by name → list of (Module, Constant)
- `GlobalVarIndex` — Global variables by name → list of (Module, Variable)
- `EnumValueNames` — Fast existence check HashSet

### Per-Module / Per-Procedure Indexes
- `controlIndex` = `mod.Controls.ToDictionary(c => c.Name, OrdinalIgnoreCase)` — one per module
- `paramIndex`, `localVarIndex`, `localConstIndex`, `globalVarModIndex` — one per procedure, built at top of `ResolveProcedureBody`
- `localNames` — HashSet of param + local var + local const names (shadow guard)
- `env` — variable→type map built by `BuildProcEnv`/`BuildPropEnv` (global type map, all modules' globals, params, locals, function return, local `Set` assignments tracked in-flight)

### Two-Step Line Processing (`ResolveProcedureBody`)
For each line in the procedure body (original file lines, `startLine` to `endLine`):

**STEP 1 — Dot-chain resolution:**
1. `EnumerateDotChains(maskedEffective)` detects `identifier.identifier` chains (left-to-right)
2. `EnumerateParenContents(maskedEffective)` → inner chains inside `(...)`
3. `TryUnwrapFunctionChain` → unwrap `CStr(obj.Field)` patterns
4. Each chain → `ResolveChain` which claims token positions via `chainTokensClaimed`

**STEP 2 — Bare-token resolution:**
`EnumerateTokens(masked)` scans all identifiers; tokens already in `chainTokensClaimed` or preceded by a dot are skipped. **VB keyword filter** checks local indexes before skipping (VB6 allows shadowing).

### Resolution Priority (base variable in ResolveChain)
1. `paramIndex` — procedure parameters
2. `localVarIndex` — local variables (skip declaration line)
3. `globalVarModIndex` — module-level variables (same module)
4. `controlIndex` — current form's controls
5. `GlobalVarIndex` — global variables from other modules (non-Private filter)
6. `ModuleByName` — module-qualified access

### STEP 2 Bare-Token Priority
1. Parameter
2. Local variable (skip declaration line)
3. Local constant
4. Module-level global variable (same module)
5. Control (same form, bare usage like `lblTitle = "x"`)
6. Global variable from another module (non-Private)
7. Global constant (cross-module non-Private, then same-module)
8. Procedure (Sub/Function) via `SelectProcTarget` (local-first)
9. Property (bare cross-module, e.g., `If ExecSts = ...`)
10. Enum value (bare, with context filtering)
11. Module name

### With Block Expansion (inside `ResolveProcedureBody`)
Lines inside a `With` block undergo two-phase expansion to build the `effectiveLine`:
1. **Block 1** (leading dot): if `trimmed.StartsWith(".")`, replaces the entire line with `withStack.Peek() + "." + suffix` (strips leading dot, concatenates With prefix). This loses original indentation but preserves token order.
2. **Block 2** (remaining dots): `ReWithDotReplacement` regex `(?<![\w)]).` expands any remaining `.Member` references NOT preceded by word chars or `)`. This handles the **second and subsequent** With-context dots on the same line.

Example: raw `.ProgNum.text = .ProgNum.text + "*"` inside `With Me` →
- After Block 1: `Me.ProgNum.text = .ProgNum.text + "*"`
- After Block 2: `Me.ProgNum.text = Me.ProgNum.text + "*"`

### Chain Resolution (`ResolveChain`)
- **Function return value self-reference**: if base matches the containing function name (e.g., `Queue_Pop.Seq`), records self-reference and continues chain walk
- **Enum.Value**: 2-part chains matching EnumDefIndex are handled first
- **Module-qualified**: `Module.GlobalVar`, `Module.Property`, `Module.Procedure`, `Module.Constant`, `Module.Control`
- **Me keyword**: `Me.Member` resolved as self-reference to current class/form module (only for `IsClass` or `Kind=="frm"`)
- **Module-match member priority**: when base resolves to a module (Me, or explicit module name), members are checked in this order — **first match wins**:
  1. `GlobalVariables` (only if variable has a non-empty `Type`) → sets `typeName`, chain continues
  2. `Properties` (only if property has a non-empty `ReturnType`) → sets `typeName`, chain continues
  3. `Procedures` → records reference, **returns** (no further chain walk)
  4. `Constants` → records reference, **returns**
  5. `Controls` → `MarkControlAsUsed`, **returns**
  - ⚠️ A Property/GlobalVariable with the same name as a Control will **shadow** the Control.
- **Chain walk**: uses `env` for type resolution → ClassIndex → Properties → Procedures → Controls → UDT fields; stops at external (non-project) types

### Position Mapping Pipeline
With-expanded lines have different character positions from the raw source. The mapping pipeline ensures references point to the correct raw-line positions:
1. `GetDepthZeroTokenPositions(chainText, chainIndex)` — extracts structural tokens with their **effective-line** positions (not inside parens)
2. `FindTokenInRawLine(rawLine, token, positionHint, claimedPositions)` — maps each effective position → **raw-line** position using `\b` regex proximity matching. Accepts `claimedPositions` (= `chainTokensClaimed`) to skip positions already taken by earlier chains, enforcing **forward-only mapping**
3. `chainTokensClaimed.Add(pos)` — claims raw positions so STEP 2 bare-token scan skips them
4. `GetTokenStartChar(rawLine, token, fallbackIndex)` — final position lookup with exact-match-first strategy

**Forward-only invariant**: since `EnumerateDotChains` yields chains left-to-right and `FindTokenInRawLine` skips claimed positions, each chain maps to distinct raw positions. This prevents the **equidistant tie bug** where two chains mapping the same token would both resolve to the same raw position.

**Duplicate chain guard**: `ResolveChain` checks if the first structural token's natural raw position (without claim filtering) is already in `chainTokensClaimed`. If so, the chain was already resolved (by a previous loop — e.g., `EnumerateDotChains` finds the chain at top level, then `EnumerateParenContents` finds the same chain inside function parentheses). Re-resolving would cause `FindTokenInRawLine` to map homonym tokens to wrong occurrences because correct positions are already claimed.

### Key Helpers (`Parser.Resolve.Members.cs`)
- `SkipOptionalParentheses` — returns index unchanged for unbalanced parens (fix for multi-line calls with `_`)
- `EnumerateParenContents` — extracts balanced `()` content for inner chain resolution
- `TryUnwrapFunctionChain` — unwraps function wrappers; skips array access `var(i).Member`
- `MaskStringLiterals` — replaces string content with spaces, preserving positions
- `StripInlineComment` — removes `'` comments (string-aware: ignores apostrophes inside string literals)
- `PrunePropertyReferenceOverlaps` — removes references overlapping between Property Get/Let/Set variants of the same name
- `ReWithDotReplacement` regex (line 12): `(?<![\w)]).(\s*[A-Za-z_]\w*(?:\([^)]*\))?)`

### Post-Processing (declaration-level references)
- `ResolveTypeReferences` — adds References for `As TypeName` in type fields, global vars, params, locals
- `ResolveClassModuleReferences` — adds References for `As [New] ClassName`
- `MarkUsedTypes` — marks types from declarations
- Module `Used` propagated from any used member

## BuildReplaces (`Parser.Replaces`)
- **Pre-calculates ALL substitutions** during analysis phase (after naming conventions applied)
- Iterates all symbols that can produce renames; for each: finds exact character position in source using `References.LineNumbers` + `StartChars`
- Special handling:
  - Properties in other modules: only `.PropertyName` (dot-prefixed)
  - Property/global variable shadowing: if a local parameter/variable matches the renamed symbol (case-insensitive), bare references are qualified as `ModuleName.Symbol`
  - Controls in `Begin` lines: only control name (skip library/type)
  - **SSTab container references**: `Tab(N).Control(N)= "ControlName"` lines inside SSTab definitions are scanned; the control name inside quotes is replaced (preserving any array index like `(2)`)
  - Attributes: `Attribute VB_Name = "..."` and `Attribute VarName.VB_VarHelpID`
  - Constants: `skipStringLiterals=true` to avoid replacing inside string values
  - **Procedure disambiguation**: conflict filters exclude forms (`!m.IsForm`) and classes; `SelectProcTarget` prefers local module; qualified references use `referenceNewName` (not `qualifiedReferenceName`) to avoid double-prefixing
  - **No blind fallback**: when StartChar is valid but text doesn't match, the replace is skipped (no `AddReplaceFromLine` fallback) to prevent homonym position pollution
- **Progress logging**: `Processando simboli: [i/N] <Category>: <Name>...`
- **File cache**: reads each file once, reuses for all symbols
- Output: `VbModule.Replaces[]` ordered by line (desc) + char (desc) for safe application

## Refactoring (`Refactoring.cs`)
- ⚡ **Ultra-fast**: uses pre-calculated `Replaces[]` from Phase 1
- No regex, no parsing, no matching logic — just mechanical application
- Groups by line, applies substitutions right-to-left
- Validates each replace (checks OldText still matches at StartChar position)
- Backup only modified files; uses Windows-1252 ANSI encoding

## Exports (`Parser.Export` + `Program.cs`)
`ExportProjectFiles` writes:
- `*.symbols.json` — complete analysis (all modules, procedures, references, etc.)
- `*.rename.json` — only non-conventional symbols
- `*.rename.csv` — CSV format
- `*.linereplace.json` — precise line-by-line substitutions (LineNumber, StartChar, EndChar, OldText, NewText, Category)
- `*.dependencies.md` — Mermaid dependency graph
- `*._TODO_shadows.csv` — property/variable shadowing cases
- `*.refdebug.csv` — reference debug info (SymbolKind/SymbolName, caller source location)
- `*.disambiguations.csv` — lines where qualified prefix was actually applied
- `*.refissues.csv` — references still containing `-1` StartChar
- `*._CHECK_startchars.csv` — references with valid StartChar but no replace produced

## TypeAnnotator (`TypeAnnotator.cs`)
- Adds missing types only when suffix infers type. No default `As Object`.
- Collects missing type cases (no suffix) and exports `*.missingTypes.csv` with `Module,Procedure,Name,ConventionalName,Kind`
- Tracks missing return types for `Function` and `Property Get` (kind `FunctionReturn`/`PropertyReturn`)
- Skips `As Variant` defaults for constants; non-inferable constants go to missing types
- **Const expressions with `Or`/`And` or parentheses** are treated as non-inferable and go to missing types
- For forms/classes: adds explicit `Public` on procedures/properties without visibility and `Private` on module-level `Dim`/`Const`
- Cleans VB6 `Call` statements and removes `Step 1` from `For` loops
- **Runs AFTER refactoring** (uses conventional names already applied)

## CodeFormatter (`CodeFormatter.cs`)
- **ReorderLocalVariables**: moves Dim/Static/Const to top of procedure body; keeps `Attribute` lines attached to signatures; groups: comments → constants → static → Dim; left-trims and indents with two spaces
- **HarmonizeSpacing**: normalizes blank lines; rules: S0 fallback; S14 wins for initial declaration groups; S7 comments before block use blank + comments + block; S8 comments inside block use block start + comments + block; single-line If always has blank after unless adjacent to If/End If boundaries; multi-line `If...Then` gets blank before when preceded by non-block statements; `End If` does NOT add blank before `Else`/`ElseIf`; blank lines before `Else`/`ElseIf` are removed; Property Get/Let/Set blocks with same name stay adjacent

## Naming (`Parser.Naming` + `Parser.Naming.Apply` + `Parser.Naming.Utilities`)
- Modules: `Frm`/`Cls` prefix, PascalCase
- Procedures: detect standard events, control events, WithEvents, interface-implemented (`Implements`) event handlers keep name (underscore preserved); double underscores in event part → treated as normal (PascalCase)
- **WithEvents variables**: use `M_` + PascalCase (e.g., `mpicContainer` → `M_Container`, event handler `M_Container_Click`). The `M_` prefix avoids conflicts with properties/controls having the same base name. Checked via `v.IsWithEvents` before other naming rules in `Parser.Naming.Apply.cs`. Strips `m_`, `g_`, `gobj`, and `m`+Hungarian 3-letter prefix patterns, then applies `M_` + `ToPascalCase`
- Properties share `ConventionalName` across Get/Let/Set
- Enum naming: strips `e_` (with underscore) or `e` (before uppercase) prefix, then screaming snake to PascalCase; preserves short acronyms (e.g., `SQM`)
- Control prefixes: `Frame` → `Fra`, `Label` → `Lbl`, `Panel` → `Pnl`; numeric-only suffixes retain the stem (e.g., `Label1` → `LblLabel1`)
- **Label Value disambiguation**: when two `VB.Label` controls collide on the same proposed name and one has a numeric-only Caption (after stripping `<>` placeholders) while the other has text, the numeric-caption label gets suffix `Value` instead of `2` (e.g., `LblStatusValue` vs `LblStatus`). This is order-independent: `labelValueOverrides` dictionary pre-computes overrides before the naming loop, so even if the numeric-caption label is declared first it still gets `Value`
- Global `g`+Hungarian prefix: stripped (e.g., `glngLanguage` → `Language`)
- Type name normalization: trailing single-letter `T` stripped before `_T` (e.g., `ParamT` → `Param_T`)
- Convention names are PascalCase (initial uppercase) so event procedures use uppercase initial

## Known Pitfalls / Rules
- Always use `StartLine/EndLine` bounds when scanning procedures; use `GetProcedureAtLine(lineNum)` (not `FirstOrDefault(p => lineNum >= p.LineNumber)`)
- Regex for arrays must escape parentheses correctly (`\([^)]*\)`)
- `ReFieldAccess` and nested chain regex must allow array indexing
- Constants: use `skipStringLiterals=true` when building replaces to avoid altering constant values in quotes
- `chainTokensClaimed` claims positions BEFORE resolution; if a chain is detected but resolution fails, STEP 2 cannot process those tokens. This is by design to avoid double-counting
- Local form controls always take priority over cross-module global variables (VB6 scoping)
- `SkipOptionalParentheses` must return the original index for unbalanced parens (multi-line calls with `_`)
- `FindTokenInRawLine` returns -1 for synthetic With-prefix tokens; callers must guard with `pos >= 0`
- Procedure conflict disambiguation ignores class modules and forms; class/form members are accessed via object and are not ambiguous
- `LineReplace.StartChar`/`EndChar` refer to positions in the original source before substitutions; only `NewText` changes later
- **Forward-only position mapping**: `FindTokenInRawLine` must receive `chainTokensClaimed` to skip already-claimed raw positions. Without this, equidistant tokens (e.g., two `ProgNum` on the same line inside `With Me`) create a proximity tie and both chains map to the first occurrence, losing the second reference
- **Duplicate chain avoidance**: `EnumerateParenContents` can re-discover chains already found by `EnumerateDotChains` (e.g., `Trim$(obj.Field)` where the main loop scans past `Trim$` and finds `obj.Field`). `ResolveChain` must detect and skip such duplicates — otherwise `FindTokenInRawLine` maps homonym tokens to wrong raw positions because the correct ones are already claimed
- **VB keyword shadowing**: STEP 2 keyword filter must check `paramIndex`, `localVarIndex`, `localConstIndex`, `globalVarModIndex`, and `controlIndex` before skipping a token as a VB keyword. VB6 allows locals/controls to shadow built-in function names (e.g., `Dim Exp As Long`, control named `Rate`)
- **Procedure disambiguation**: `SelectProcTarget` must prefer the current module's procedure over cross-module matches. Conflict resolution must exclude forms (`!m.IsForm`) like classes. When force-qualifying, pass `referenceNewName` (not `qualifiedReferenceName`) to avoid double-prefixing (e.g., `FrmBarcode.FrmBarcode.ShowRecipe`)
- **Function return value in dot-chains**: `ResolveChain` must check if the base variable matches the containing function name (self-reference for return value) to avoid missing renames like `Queue_Pop.Seq`

## Recent Fixes
- **External object members**: member-access tokens (e.g., `obj.Prop`) are excluded from parameter/local reference scans so external member names are not renamed
- **Enum value collisions**: when multiple enum values converge to the same `ConventionalName`, references are qualified as `EnumName.Value` to avoid ambiguity
- **Property/global variable shadowing**: when a public property or module-level variable is renamed and a local parameter/variable uses the same name (case-insensitive), bare references are qualified as `ModuleName.Symbol` to avoid shadowing
- **Enum qualification guard**: if the value is already qualified (`EnumName.Value`), only the value is renamed, preventing `Enum.Enum.Value` or wrong enum prefixes
- **Inline calls after `Then`**: procedure references are recorded even if a call is already present in `Calls`, ensuring inline statements are renamed
- **Screaming snake to PascalCase**: enum value naming preserves only known acronyms, so `RIC_RUN_CMD_CALLER` → `RicRunCmdCaller`
- **String-aware comment stripping**: comment removal ignores apostrophes inside string literals, preventing truncation of lines with text (e.g., `"E' ..."`)
- **StartChar-first replaces**: references now carry `StartChar`, so `BuildReplaces` can apply substitutions without re-scanning the line (fallbacks only when missing)
- **StartChar-targeted replaces**: `BuildReplaces` now uses `StartChar` for all reference categories (module members, properties, constants, enum values) before falling back to occurrence-based matches
- **Case-only renames**: replaces now apply even when only casing changes (case-sensitive compare), so `filterMode` → `FilterMode` is not skipped
- **Enum reference filtering**: member-access tokens (e.g., `Enum.Value`) are excluded from bare enum-value matches; qualified tokens record both enum/value positions
- **Type name normalization**: trailing single-letter `T` is stripped before appending `_T` (e.g., `ParamT` → `Param_T`)
- **Local multi-declaration parsing**: comma-separated `Dim` declarations are split into multiple locals so all occurrences are tracked and renamed
- **Base variable in chains**: base identifiers used in member chains are marked as references, ensuring parameters are renamed even in `obj.Member...` expressions
- **String-literal safe replaces**: reference renames skip string literals, with a special allowance for `Attribute ...` lines to keep VB attributes in sync
- **Function-chain unwrap**: member chains inside function calls (e.g., `CStr(...)`) are unwrapped for resolution; array access (`var(i).Member`) is excluded from unwrap
- **Chain splitting outside parentheses**: dot-chain parsing ignores dots inside parentheses, enabling nested calls like `Evap_Info(...)` without breaking earlier cases
- **Control arrays renaming**: controls use all `LineNumbers` for `Begin` lines so every element in a control array is renamed consistently
- **Global constant shadowing**: global constant references are skipped when a local variable/parameter shadows the constant name
- **Self module references**: module references are recorded even inside their own module, so form/module names passed as values are renamed consistently
- **Event handler guard for double underscores**: procedures/properties with an extra underscore in the event part are treated as normal routines (PascalCase), not control/WithEvents events
- **Function/Property event handler guard**: control event references are only added for `Sub` handlers, avoiding false matches on `Function`/`Property` names that share control prefixes
- **Property global references**: global variables referenced inside `Property` blocks are now tracked with `StartChar` even when the property has no parameters
- **Local declaration ordering**: procedure headers keep `Attribute` lines attached; local declarations are grouped as comments → constants → static → Dim, without extra blank lines or alphabetic sorting
- **Spacing rules tweaks**: no blank after initial file `Attribute` block; pre-procedure comment blocks stay contiguous with a single blank line before them; no blank at the start of a procedure even if the first statement is a comment for a following block
- **Single-line If spacing**: single-line `If` statements always add a blank line after; they add a blank line before unless preceded by a comment
- **Property spacing**: paired `Property Get/Let/Set` blocks with the same name stay adjacent (no blank line inserted before/after the pair)
- **Spacing rules**: blank line inserted after local declaration blocks and before `For/Do` loops when preceded by non-block statements; `End With` inserts a blank line unless followed by another block end
- **Disambiguations CSV**: `*.disambiguations.csv` now only lists lines where a qualified prefix was actually applied; sorted by module and line number
- **Shadows CSV**: includes `LineNumber`, `LocalType`, and `ShadowedType`, sorted by module and line number
- **Conflict handling**: procedure/property/constant disambiguation ignores class modules; class members are accessed via object and are not ambiguous
- **Console output styling**: `[OK]` in green, `[WARN]` in yellow, `[X]` in red, `[i]` in cyan
- **Control StartChar precision**: control references now carry exact `StartChar` from raw lines (indent preserved) and overwrite coarse positions when more accurate data arrives
- **Reference cleanup**: coarse references without `StartChar`/`OccurrenceIndex` are skipped for replaces, and precise references update earlier `-1`/incorrect entries
- **StartChar mismatch report**: `_CHECK_startchars.csv` is emitted when a reference has a valid `StartChar` but no replace is produced (skips `-1`)
- **Reference debug export**: `*.refdebug.csv` now includes `SymbolKind/SymbolName` captured at insertion time (no post-search), plus caller source location for any `-1` entries
- **Reference debug/check outputs**: `*.refdebug.csv` and `_CHECK_startchars.csv` are always emitted with headers, even when empty
- **Reference issues export**: `*.refissues.csv` lists every reference still containing `-1` to drive cleanup
- **Dependencies scan precision**: variables/constants/properties found by `BuildDependenciesAndUsage` now record `OccurrenceIndex` + `StartChar` and owner info
- **Qualified enum references**: `Enum.Value` matches use a pair regex to avoid false hits like `AdvancedFeatures` and capture correct positions
- **Frame control prefix**: `Frame` controls starting with `frX` preserve the `fr` prefix (e.g., `frVariable` → `fraFrVariable`)
- **Global g+Hungarian prefix**: globals like `glngLanguage` strip `g` + type prefix to `Language` (conflicts/reserved handled as usual)
- **Control naming PascalCase**: control conventional names start with uppercase, so event handlers become `LstSel_Click`
- **BuildReplaces optimization**: references are aggregated into a per-line list, sorted with `List.Sort` and module index, then applied with cached line parsing; entries are de-duplicated
- **Single-pass resolver**: replaced old multi-pass approach (PASS 1, 1.2, 1.5, 1.5b + ResolveFieldAccesses + ResolveControlAccesses + ResolveParameterAndLocalVariableReferences + ResolveEnumValueReferences) with a single `ResolveTypesAndCalls` that scans each line once
- **GlobalIndexes**: 11 pre-built dictionaries replace ad-hoc per-pass index building
- **OccurrenceIndexes removal**: dead code cleanup across 8 files
- **Visibility filter**: cross-module resolution now skips `Private` symbols
- **Form control parsing**: handles indented `Begin` blocks and nested `controlDepth` for container controls
- **Module.Control lookup**: module-qualified section of ResolveChain checks `moduleMatch.Controls` for `Module.Control` patterns (e.g., `FrmLoading.progress`)
- **Forms in ClassIndex**: forms are added to ClassIndex (alongside classes) so `Dim obj As FrmXxx` chain walks find form controls/properties
- **Control naming prefixes**: `Frame` → `Fra` (not `Fr`), `Label` → `Lbl` (not `Lb`), `Panel` → `Pnl` (not `Pn`); numeric-only suffixes retain the stem (e.g., `Label1` → `LblLabel1`)
- **SkipOptionalParentheses fix**: returns original index for unbalanced parentheses in multi-line calls (lines with `_` continuation); prevents `EnumerateDotChains` from consuming the entire line
- **Me keyword**: `Me.Control`, `Me.Property`, etc. in class/form modules resolved as self-reference
- **Control priority over cross-module globals**: `controlIndex` check moved before `GlobalVarIndex` in both ResolveChain and STEP 2, matching VB6 scoping semantics. Also fixes the `else if` fall-through bug where `GlobalVarIndex` matching the key but failing the inner filter would block the control check
- **With block expansion**: `ReWithDotReplacement` expands `.Member` references inside `With` blocks; negative lookbehind `(?<![\w)])` prevents false matches on `identifier.member`
- **Array dimension constants**: constants used as array dimensions in `Dim arr(CONST)` are resolved
- **Multi-line Declare params**: parameters in multiline `Declare` statements are resolved with corrected line numbers
- **Space-before-paren**: VB6 call syntax `FuncName (args)` with space before paren is handled
- **Position collision fix**: `FindTokenInRawLine` maps effective-line positions to raw-line positions, preventing collisions between With-expanded and raw coordinates
- **Multi-declaration Dim parsing**: `Dim a, b, c As Integer` now correctly parsed as three separate locals using a broader `^\s*(Dim|Static)\s+(.*)$` regex gate in `Parser.Core.Module.cs`
- **VB keyword shadowing by locals**: STEP 2 keyword filter in `Parser.Resolve.SinglePass.cs` now checks `paramIndex`/`localVarIndex`/`localConstIndex` before skipping a token as a VB keyword, so `Dim Exp As Long` gets renamed
- **VB keyword shadowing by controls/globals**: extended the keyword filter to also check `globalVarModIndex` and `controlIndex`, so controls named after VB functions (e.g., `Rate`) are renamed
- **Function return value in dot-chains**: `ResolveChain` now accepts `isFunction`, `memberReferences`, `lineNumber` params and checks if the base variable is the containing function's return value (e.g., `Queue_Pop.Seq`), recording the self-reference
- **Procedure disambiguation — exclude forms**: conflict filters in `BuildReplaces` now use `!m.IsForm` to exclude forms (like classes), preventing false `FrmBarcode.ShowRecipe` qualification
- **Procedure disambiguation — local-first**: `SelectProcTarget` prefers the current module's procedure over cross-module matches
- **Procedure disambiguation — double prefix fix**: when force-qualifying procedure references, `referenceNewName` is passed instead of `qualifiedReferenceName` to avoid `FrmBarcode.FrmBarcode.ShowRecipe`
- **Forward-only position mapping**: `FindTokenInRawLine` now accepts `claimedPositions` parameter (= `chainTokensClaimed`) to skip raw positions already claimed by earlier chains. Fixes the **equidistant tie bug**: on lines like `.ProgNum.text = .ProgNum.text + "*"` inside `With Me`, the effective-line position of the second `ProgNum` is equidistant (distance 8) from both raw occurrences (pos 13 and 29). Without claimed-position filtering, `OrderBy().First()` returns pos 13 (document order), same as Chain 1 → the second reference is deduplicated away by `AddLineNumber`. With the fix, pos 13 is skipped (already claimed) and pos 29 is correctly selected
- **Homonym position pollution fix**: when `ApplyReferenceReplace`'s generic handler (and specialized handlers like `AddPropertyReferenceReplaces`, `AddModuleMemberReferenceReplaces`, `AddEnumValueReferenceReplaces`, `AddConstantReferenceReplaces`) had a valid `StartChar` that didn't match any regex occurrence, the fallback scanned `\bName\b` for ALL occurrences on the line, claiming positions belonging to OTHER symbols with the same name. E.g., a field `.Deposit.Material` with a slightly wrong StartChar would claim position 4 (the bare control `Material` at the start of the line) via the fallback, blocking the control's rename. Fix: when `startChar >= 0` but no precise match exists at that position, return immediately instead of falling back to a blind regex scan
- **WithEvents naming convention**: variables declared `WithEvents` now use `M_` + PascalCase (e.g., `mpicContainer` → `M_Container`). The `M_` prefix clarifies module scope and avoids conflicts with properties/controls sharing the same base name (e.g., property `Container` vs WithEvents `mpicContainer`). Event handlers become `M_Container_Click`. Implemented as a priority branch in `Parser.Naming.Apply.cs` before the `IsStatic`/`IsPrivate`/`Public` checks
- **Enum `e_` prefix stripping**: enum names starting with `e_` (e.g., `e_POLLING_STATE`) now strip the prefix before PascalCase conversion (→ `PollingState`). The existing `e` + uppercase strip (e.g., `ePollingState` → `PollingState`) is preserved as fallback
- **Attribute VB_VarHelpID fix**: `AddAttributeReplaces` used `module.Name` to look up the file cache (keyed by `module.FullPath`), so it never found any match and all `Attribute VarName.VB_VarHelpID = -1` lines were silently skipped. Also affected `Attribute VB_Name` for module renames. Fixed by using `module.FullPath!`
- **Spacing: End If / Else gap fix**: `End If` handler was adding a blank line before `Else`/`ElseIf` because they weren't in `IsBlockEnd`. Added `IsElseLine` helper; `End If` handler now suppresses blank when next is `Else`/`ElseIf`; blank line handler also removes existing blanks before `Else`/`ElseIf`
- **Spacing: blank before multi-line If**: added blank line before `If...Then` (multi-line) when preceded by regular statements, matching the existing `For`/`Do` logic. Uses `IsMultiLineIfStart` (starts with `If`, ends with `Then` after stripping inline comment)
- **SSTab container control references**: `Tab(N).Control(N)= "ControlName"` lines inside SSTab definitions were not renamed when child controls were renamed, causing the SSTab to lose track of its children. Fix: `AddDeclarationReplace` for controls now scans form files for the `Tab(\d+)\.Control(\d+)\s*=\s*"OldName(\(\d+\))?"` pattern and replaces the control name inside quotes (preserving any array index)
- **Duplicate chain resolution fix**: `EnumerateParenContents` re-discovers dot-chains already resolved
- **Label Value disambiguation**: when two `VB.Label` controls collide on the same `ProposedName`, if one has numeric-only Caption (digits/dots after stripping quotes and `<>` placeholders) and the other has text, the numeric-caption label gets `Value` suffix instead of `2`. Pre-computed in `labelValueOverrides` dictionary before the naming loop so ordering is irrelevant. Helpers: `IsLabelType(controlType)` checks `VB.Label`/`Label`; `HasNumericOnlyCaption(controls)` strips quotes/`<>`/`>` from the first control's Caption and checks all-digit (dots allowed for decimals like `123.456`)
- **Recursive Sub call references**: the STEP 2 auto-reference guard (`token == memberName`) only recorded references for `Function` (return value assignments) but silently skipped `Sub` recursive calls — no reference was generated, so `BuildReplaces` could not rename them. Fix: removed the `isFunction` gate; now both Sub and Function self-references are recorded when `currentLine != lineNumber`. Same fix applied to the dot-chain self-reference in `ResolveChain` (e.g., `Me.MySub`)
- **Control prefix idempotency (FraMES bug)**: `ApplyControlNaming` checked `StartsWith("Frame", OrdinalIgnoreCase)` before the "already correct prefix" check, so `FraMES` (prefix `fra` + base `MES`) was falsely matched as `Frame` + `S` → `FraS`. Fix: moved the "already has correct prefix" check (`StartsWith(expectedPrefix) + uppercase after`) to the top of the method, before the Frame/Panel/Label word checks. Now `FraMES` is recognized as already having the `fra` prefix and preserved
- **Type _T fallback removal (circular reference bug)**: `AddTypeReferenceAt`, `FindTypeStartChar`, and `BuildReplaces` all had `_T` suffix fallback logic that tried to match type names with/without `_T`. This was fundamentally wrong: VB6 requires exact type names (`As MesMachineState` when the enum is `MesMachineState`, NOT the type `MesMachineState_T`). The parser always stores the exact source text, so the fallback handled a case that cannot exist in valid VB6 code. The fallback caused `MesMachineState` (enum ref) to be attributed to `MesMachineState_T` (type) → `BuildReplaces` generated replace `MesMachineState` → `MesMachineState_T` → circular type reference. Fix: removed ALL `_T` fallback logic — `FindTypeStartChar` returns `-1` if exact match fails; `AddTypeReferenceAt` returns if exact `typeIndex` lookup fails; removed `GetTypeAlternateName` and `effectiveOldName` overrides from `BuildReplaces`; removed the Type exception from the `oldName == newName` skip at line 257
- **Control prefix case-sensitive guard (lbLonPosition bug)**: the `lb` → `lbl` and `fr` → `fra` normalization blocks in `ApplyControlNaming` used `!StartsWith("lbl"/"fra", OrdinalIgnoreCase)` to skip controls that already have the full prefix. This caused false negatives: `lbLonPosition` has `lbL` which matches `lbl` case-insensitively → guard rejected it → control fell through to the generic no-prefix handler → `LblLbLonPosition` instead of `LblLonPosition`. Same issue for `fr` → `fra` (e.g., `frAnchor` has `frA` = `fra` case-insensitive). Fix: changed both guards to `StringComparison.Ordinal` (case-sensitive) so `lbL` ≠ `lbl` and `frA` ≠ `fra`, allowing the normalization to proceed correctly

## Performance Considerations
**Why is VB6 IDE faster?**
1. **Incremental parsing** — VB6 keeps parsed project in memory, doesn't reparse on every load
2. **Lazy resolution** — resolves references only when needed (compile/intellisense), not upfront
3. **Native code** — compiled C++ vs. managed .NET with GC overhead
4. **No semantic analysis** — VB6 doesn't track every reference until compile time
5. **Optimized for single-file edits** — our tool analyzes entire project every run

**Our bottlenecks:**
- Multiple file reads (Parse, Resolve, BuildReplaces) — **FIXED with file cache** (`BuildFileCache` in `Parser.Resolve.cs`)
- Full semantic analysis (all References tracked) — needed for precise refactoring
- Regex matching for every symbol on every line — **REDUCED** (single-pass uses char-level scanning for chains/tokens, regex only for hot-path patterns)
- Multiple passes over same data — **FIXED**: resolution is now single-pass per line

**Potential optimizations:**
1. ~~**Single-pass parsing**~~ — ✅ **DONE**: `ResolveTypesAndCalls` scans each line once
2. **Parallel processing** — analyze modules in parallel (Thread.ParallelForEach)
3. **Compiled Regex** — use `RegexOptions.Compiled` for hot-path patterns (partially done)
4. **Span<char>** — avoid string allocations in parsing
5. **Incremental mode** — cache parsed project, only reparse changed files
6. **Lazy References** — optionally skip reference tracking if not needed
7. **Streaming parser** — process line-by-line without loading entire file in memory

