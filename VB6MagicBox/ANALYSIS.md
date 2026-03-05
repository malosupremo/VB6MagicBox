# VB6MagicBox - Technical Summary to start a new chat with github Copilot

## Purpose
VB6 parser/refactoring tool. Pipeline: parse VB6 project, resolve references (types/calls/fields), build dependencies/usage, apply naming conventions, **build precise replaces**, export outputs, apply refactoring.

## Architecture
**Phase 1 (Analysis)**: `Parser.ParseAndResolve`
1. Parse project files (`Parser.Core` + `Parser.Core.Module`)
2. **Single-pass reference resolution** (`Parser.Resolve.SinglePass` + `.Members` + `.Indexes` + `.Helpers` + `.Types`)
3. Build dependencies & mark used symbols (`Parser.Resolve.Dependencies`)
4. Apply naming conventions & sort (`Parser.Export` → `NamingConvention.Apply`)
5. **Build Replaces** (`Parser.Replaces`) - pre-calculate exact character positions for all renames

**Phase 2 (Refactoring)**: `Refactoring.ApplyRenames`
- ⚡ Ultra-fast: uses pre-calculated `VbModule.Replaces[]` (line + char position)
- No regex matching, no re-parsing
- Applies substitutions from end to start (preserves positions)

**Phase 3 (Type Annotation)**: `TypeAnnotator.AddMissingTypes`
- Runs AFTER refactoring (uses conventional names already applied)
- Adds cleanup: visibility normalization in forms/classes, Call removal, `For ... Step 1` cleanup

## Key Models
- `VbModule`: `Procedures`, `Properties` (separate), `Types`, `Enums`, `Constants`, `GlobalVariables`, `Controls`, `ModuleReferences`, `Used`, **`Replaces`** (NEW - list of `LineReplace`), `StartLine/EndLine` lookups via `ContainsLine`/`GetProcedureAtLine`.
- `VbProcedure`: `StartLine/EndLine`, `Parameters`, `LocalVariables`, `Calls`, `References`, `ReturnType`.
- `VbProperty`: same as procedure (separate list), `ReturnType`, `Parameters`, `References`.
- `VbTypeDef`/`VbField`, `VbEnumDef`/`VbEnumValue` with `References`.
- `VbControl`: `LineNumbers` for control arrays, `References`.
- `VbReference`: now tracks `StartChars` alongside `LineNumbers`/`OccurrenceIndexes` for exact replace positions.
- **`LineReplace`** (NEW): `LineNumber`, `StartChar`, `EndChar`, `OldText`, `NewText`, `Category` - precise substitution with exact character position.

## Parsing (Parser.Core)
- Handles `Function/Sub/Property` and `Declare` (supports `Alias`, optional visibility).
- Collapses line continuations `_` with line mapping; multiline signature parameter line numbers are corrected.
- Regexes for array parens use `\([^)]*\)`; `ReFieldAccess` pattern: `([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)`.
- `Implements` lines recorded in `VbModule.ImplementsInterfaces`.

## Resolution (Single-Pass Architecture)

### Overview
`ResolveTypesAndCalls` replaces the old multi-pass approach with a **single-pass** resolver.
Files: `Parser.Resolve.SinglePass.cs` (main entry + procedure body), `Parser.Resolve.Indexes.cs` (GlobalIndexes),
`Parser.Resolve.Members.cs` (helpers: chain detection, token enumeration, position mapping),
`Parser.Resolve.Helpers.cs` (MarkControlAsUsed, RecordReference, utility), `Parser.Resolve.Types.cs` (post-processing type/class refs).

### GlobalIndexes (`Parser.Resolve.Indexes.cs`)
Built **once** before scanning, reused for every module/procedure:
- `ProcIndex` — Procedure (Sub/Function/Declare, NOT Property) by name → list of (Module, Proc)
- `PropIndex` — Property (Get/Let/Set) by name → list of (Module, Prop)
- `TypeIndex` — UDT (Type) by name
- `EnumDefIndex` — Enum by name → list
- `EnumValueIndex` — Enum value name → list of VbEnumValue
- `EnumValueOwners` — Reverse lookup: VbEnumValue → owning VbEnumDef
- `ClassIndex` — Class modules + forms by name (forms added for `Dim obj As FrmXxx`)
- `ModuleByName` — All modules by VB_Name
- `ConstantIndex` — Global constants by name → list of (Module, Constant)
- `GlobalVarIndex` — Global variables by name → list of (Module, Variable)
- `EnumValueNames` — Fast existence check set

### Per-Module / Per-Procedure Indexes
- `controlIndex` = `mod.Controls.ToDictionary(c => c.Name, OrdinalIgnoreCase)` — one per module
- `paramIndex`, `localVarIndex`, `localConstIndex`, `globalVarModIndex` — one per procedure
- `localNames` — HashSet of param + local var + local const names (shadow guard)
- `env` — variable→type map (global type map, all modules' globals, params, locals, function return, local Set assignments)

### Two-Step Line Processing (`ResolveProcedureBody`)
For each line in the procedure body (original file lines, not collapsed):

**STEP 1 — Dot-chain resolution:**
1. `EnumerateDotChains(maskedEffective)` detects `identifier.identifier` chains
2. `EnumerateParenContents(maskedEffective)` → inner chains inside `(...)`
3. `TryUnwrapFunctionChain` → unwrap `CStr(obj.Field)` patterns
4. Each chain → `ResolveChain` which claims token positions via `chainTokensClaimed`

**STEP 2 — Bare-token resolution:**
`EnumerateTokens(masked)` scans all identifiers; tokens already in `chainTokensClaimed` or after a dot are skipped.

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
8. Procedure (Sub/Function) via `SelectProcTarget`
9. Property (bare cross-module, e.g., `If ExecSts = ...`)
10. Enum value (bare, with context filtering)
11. Module name

### Chain Resolution (`ResolveChain`)
- **Enum.Value**: 2-part chains matching EnumDefIndex are handled first
- **Module-qualified**: `Module.GlobalVar`, `Module.Property`, `Module.Procedure`, `Module.Constant`, `Module.Control`
- **Me keyword**: `Me.Member` resolved as self-reference to current class/form module
- **Chain walk**: uses `env` for type resolution → ClassIndex → Properties → Procedures → Controls → UDT fields; stops at external (non-project) types
- **Position mapping**: `GetDepthZeroTokenPositions` extracts structural tokens (not inside parens); `FindTokenInRawLine` maps effective-line positions → raw-line positions (handles With-expanded lines)

### Key Helpers (`Parser.Resolve.Members.cs`)
- `SkipOptionalParentheses` — returns index unchanged for unbalanced parens (fix for multi-line calls with `_`)
- `EnumerateParenContents` — extracts balanced `()` content for inner chain resolution
- `TryUnwrapFunctionChain` — unwraps function wrappers; skips array access `var(i).Member`
- `MaskStringLiterals` — replaces string content with spaces, preserving positions
- `StripInlineComment` — removes `'` comments (string-aware)
- `PrunePropertyReferenceOverlaps` — removes references overlapping between Property Get/Let/Set variants of the same name

### Post-Processing
- `ResolveTypeReferences` — adds References for `As TypeName` in type fields, global vars, params, locals
- `ResolveClassModuleReferences` — adds References for `As [New] ClassName`
- `MarkUsedTypes` — marks types from declarations
- Module `Used` propagated from any used member

## BuildReplaces (Parser.Replaces) - NEW
- **Pre-calculates ALL substitutions** during analysis phase (after naming conventions applied).
- For each symbol: finds exact character position in source using `References.LineNumbers`.
- Special handling:
  - Properties in other modules: only `.PropertyName` (dot-prefixed)
  - Property/global variable shadowing: if a local parameter/variable matches the renamed symbol (case-insensitive), bare references are qualified as `ModuleName.Symbol`
  - Controls in `Begin` lines: only control name (skip library/type)
  - Attributes: `Attribute VB_Name = "..."` and `Attribute VarName.VB_VarHelpID`
  - Constants: `skipStringLiterals=true` to avoid replacing inside string values
- **Progress logging**: the console line
  `Processando simboli: [i/N] <Category>: <Name>...` shows the current symbol being analyzed
  while building replacements (e.g., `GlobalVariable: inUplStatus`). It iterates all symbols
  that can produce renames or have references so their exact replace positions are recorded.
- **File cache**: reads each file once, reuses for all symbols
- Output: `VbModule.Replaces[]` ordered by line (desc) + char (desc) for safe application

## Refactoring (Refactoring.cs) - REWRITTEN
- ⚡ **Ultra-fast**: uses pre-calculated `Replaces[]` from Phase 1
- No regex, no parsing, no matching logic - just mechanical application
- Groups by line, applies substitutions right-to-left
- Validates each replace (checks OldText still matches)
- Backup only modified files

## Exports
`ExportProjectFiles` writes: 
- `*.symbols.json` - complete analysis
- `*.rename.json` - only non-conventional symbols
- `*.rename.csv` - CSV format
- **`*.linereplace.json`** - **NEW**: precise line-by-line substitutions (LineNumber, StartChar, EndChar, OldText, NewText, Category)
- `*.dependencies.md` - Mermaid graph

## TypeAnnotator
- Adds missing types only when suffix infers type. No default `As Object`.
- Collects missing type cases (no suffix) and exports `*.missingTypes.csv` with `Module,Procedure,Name,ConventionalName,Kind`.
- Tracks missing return types for `Function` and `Property Get` (kind `FunctionReturn`/`PropertyReturn`).
- Skips `As Variant` defaults for constants; non-inferable constants go to missing types.
- **Const expressions with `Or`/`And` or parentheses** are treated as non-inferable (no numeric type assigned) and go to missing types.
- For forms/classes: adds explicit `Public` on procedures/properties without visibility and `Private` on module-level `Dim`/`Const`.
- Cleans VB6 `Call` statements and removes `Step 1` from `For` loops.
- **Runs AFTER refactoring** (uses conventional names already applied).

## Naming
- Modules: `Frm`/`Cls` prefix, PascalCase.
- Procedures: detect standard events, control events, WithEvents, interface-implemented (`Implements`) event handlers keep name (underscore preserved).
- Properties share `ConventionalName` across Get/Let/Set.
- Enum naming preserves short acronyms (e.g., `SQM`).

## Known Pitfalls / Rules
- Always use `StartLine/EndLine` bounds when scanning procedures; use `GetProcedureAtLine(lineNum)` (not `FirstOrDefault(p => lineNum >= p.LineNumber)`).
- Regex for arrays must escape parentheses correctly (`\([^)]*\)`).
- `ReFieldAccess` and nested chain regex must allow array indexing.
- Constants: use `skipStringLiterals=true` when building replaces to avoid altering constant values in quotes.
- `chainTokensClaimed` claims positions BEFORE resolution; if a chain is detected but resolution fails, STEP 2 cannot process those tokens. This is by design to avoid double-counting.
- Local form controls always take priority over cross-module global variables (VB6 scoping).
- `SkipOptionalParentheses` must return the original index for unbalanced parens (multi-line calls with `_`).
- `FindTokenInRawLine` returns -1 for synthetic With-prefix tokens; callers must guard with `pos >= 0`.
- Procedure conflict disambiguation ignores class modules; class members are accessed via object and are not ambiguous.
- `LineReplace.StartChar`/`EndChar` refer to positions in the original source before substitutions; only `NewText` changes later.

## Recent Fixes (2024)
- **External object members**: member-access tokens (e.g., `obj.Prop`) are excluded from parameter/local reference scans so external member names are not renamed.
- **Enum value collisions**: when multiple enum values converge to the same `ConventionalName`, references are qualified as `EnumName.Value` to avoid ambiguity.
- **Property/global variable shadowing**: when a public property or module-level variable is renamed and a local parameter/variable uses the same name (case-insensitive), bare references are qualified as `ModuleName.Symbol` to avoid shadowing.
- **Enum qualification guard**: if the value is already qualified (`EnumName.Value`), only the value is renamed, preventing `Enum.Enum.Value` or wrong enum prefixes.
- **Inline calls after `Then`**: procedure references are recorded even if a call is already present in `Calls`, ensuring inline statements are renamed.
- **Screaming snake to PascalCase**: enum value naming preserves only known acronyms, so `RIC_RUN_CMD_CALLER` -> `RicRunCmdCaller`.
- **String-aware comment stripping**: comment removal ignores apostrophes inside string literals, preventing truncation of lines with text (e.g., `"E' ..."`).
- **StartChar-first replaces**: references now carry `StartChar`, so `BuildReplaces` can apply substitutions without re-scanning the line (fallbacks only when missing).
- **StartChar-targeted replaces**: `BuildReplaces` now uses `StartChar` for all reference categories (module members, properties, constants, enum values) before falling back to occurrence-based matches.
- **Case-only renames**: replaces now apply even when only casing changes (case-sensitive compare), so `filterMode` → `FilterMode` is not skipped.
- **Enum reference filtering**: member-access tokens (e.g., `Enum.Value`) are excluded from bare enum-value matches; qualified tokens record both enum/value positions.
- **Type name normalization**: trailing single-letter `T` is stripped before appending `_T` (e.g., `ParamT` → `Param_T`).
- **Local multi-declaration parsing**: comma-separated `Dim` declarations are split into multiple locals so all occurrences are tracked and renamed.
- **Base variable in chains**: base identifiers used in member chains are marked as references, ensuring parameters are renamed even in `obj.Member...` expressions.
- **String-literal safe replaces**: reference renames skip string literals, with a special allowance for `Attribute ...` lines to keep VB attributes in sync.
- **Function-chain unwrap**: member chains inside function calls (e.g., `CStr(...)`) are unwrapped for resolution; array access (`var(i).Member`) is excluded from unwrap.
- **Chain splitting outside parentheses**: dot-chain parsing ignores dots inside parentheses, enabling nested calls like `Evap_Info(...)` without breaking earlier cases.
- **Control arrays renaming**: controls use all `LineNumbers` for `Begin` lines so every element in a control array is renamed consistently.
- **Global constant shadowing**: global constant references are skipped when a local variable/parameter shadows the constant name.
- **Self module references**: module references are recorded even inside their own module, so form/module names passed as values are renamed consistently.
- **Event handler guard for double underscores**: procedures/properties with an extra underscore in the event part are treated as normal routines (PascalCase), not control/WithEvents events.
- **Function/Property event handler guard**: control event references are only added for `Sub` handlers, avoiding false matches on `Function`/`Property` names that share control prefixes.
- **Property global references**: global variables referenced inside `Property` blocks are now tracked with `StartChar` even when the property has no parameters.
- **Local declaration ordering**: procedure headers keep `Attribute` lines attached; local declarations are grouped as comments → constants → static → Dim, without extra blank lines or alphabetic sorting.
- **Spacing rules tweaks**: no blank after initial file `Attribute` block; pre-procedure comment blocks stay contiguous with a single blank line before them; no blank at the start of a procedure even if the first statement is a comment for a following block.
- **Single-line If spacing**: single-line `If` statements always add a blank line after; they add a blank line before unless preceded by a comment.
- **Property spacing**: paired `Property Get/Let/Set` blocks with the same name stay adjacent (no blank line inserted before/after the pair).
- **Spacing rules**: blank line inserted after local declaration blocks and before `For/Do` loops when preceded by non-block statements; `End With` inserts a blank line unless followed by another block end.
- **Disambiguations CSV**: `*.disambiguations.csv` now only lists lines where a qualified prefix was actually applied; sorted by module and line number.
- **Shadows CSV**: includes `LineNumber`, `LocalType`, and `ShadowedType`, sorted by module and line number.
- **Conflict handling**: procedure/property/constant disambiguation ignores class modules; class members are accessed via object and are not ambiguous.
- **Console output styling**: `[OK]` in green, `[WARN]` in yellow, `[X]` in red, `[i]` in cyan.
- **Control StartChar precision**: control references now carry exact `StartChar` from raw lines (indent preserved) and overwrite coarse positions when more accurate data arrives.
- **Reference cleanup**: coarse references without `StartChar`/`OccurrenceIndex` are skipped for replaces, and precise references update earlier `-1`/incorrect entries.
- **StartChar mismatch report**: `_CHECK_startchars.csv` is emitted when a reference has a valid `StartChar` but no replace is produced (skips `-1`).
- **Reference debug export**: `*.refdebug.csv` now includes `SymbolKind/SymbolName` captured at insertion time (no post-search), plus caller source location for any `-1` entries.
- **Reference debug/check outputs**: `*.refdebug.csv` and `_CHECK_startchars.csv` are always emitted with headers, even when empty.
- **Reference issues export**: `*.refissues.csv` lists every reference still containing `-1` to drive cleanup.
- **Dependencies scan precision**: variables/constants/properties found by `BuildDependenciesAndUsage` now record `OccurrenceIndex` + `StartChar` and owner info.
- **Qualified enum references**: `Enum.Value` matches use a pair regex to avoid false hits like `AdvancedFeatures` and capture correct positions.
- **Frame control prefix**: `Frame` controls starting with `frX` preserve the `fr` prefix (e.g., `frVariable` -> `fraFrVariable`).
- **Global g+Hungarian prefix**: globals like `glngLanguage` strip `g` + type prefix to `Language` (conflicts/reserved handled as usual).
- **Control naming PascalCase**: control conventional names start with uppercase, so event handlers become `LstSel_Click`.
- **BuildReplaces optimization**: references are aggregated into a per-line list, sorted with `List.Sort` and module index, then applied with cached line parsing; entries are de-duplicated.

## Recent Fixes (2025 — Single-Pass Resolver)
- **Single-pass resolver**: replaced old multi-pass approach (PASS 1, 1.2, 1.5, 1.5b + ResolveFieldAccesses + ResolveControlAccesses + ResolveParameterAndLocalVariableReferences + ResolveEnumValueReferences) with a single `ResolveTypesAndCalls` that scans each line once.
- **GlobalIndexes**: 11 pre-built dictionaries replace ad-hoc per-pass index building.
- **OccurrenceIndexes removal**: dead code cleanup across 8 files.
- **Visibility filter**: cross-module resolution now skips `Private` symbols.
- **Form control parsing**: handles indented `Begin` blocks and nested `controlDepth` for container controls.
- **Module.Control lookup**: module-qualified section of ResolveChain checks `moduleMatch.Controls` for `Module.Control` patterns (e.g., `FrmLoading.progress`).
- **Forms in ClassIndex**: forms are added to ClassIndex (alongside classes) so `Dim obj As FrmXxx` chain walks find form controls/properties.
- **Control naming prefixes**: `Frame` → `Fra` (not `Fr`), `Label` → `Lbl` (not `Lb`), `Panel` → `Pnl` (not `Pn`); numeric-only suffixes retain the stem (e.g., `Label1` → `LblLabel1`).
- **SkipOptionalParentheses fix**: returns original index for unbalanced parentheses in multi-line calls (lines with `_` continuation); prevents `EnumerateDotChains` from consuming the entire line.
- **Me keyword**: `Me.Control`, `Me.Property`, etc. in class/form modules resolved as self-reference.
- **Control priority over cross-module globals**: `controlIndex` check moved before `GlobalVarIndex` in both ResolveChain and STEP 2, matching VB6 scoping semantics. Also fixes the `else if` fall-through bug where `GlobalVarIndex` matching the key but failing the inner filter would block the control check.
- **With block expansion**: `ReWithDotReplacement` expands `.Member` references inside `With` blocks; negative lookbehind `(?<![\w)])` prevents false matches on `identifier.member`.
- **Array dimension constants**: constants used as array dimensions in `Dim arr(CONST)` are resolved.
- **Multi-line Declare params**: parameters in multiline `Declare` statements are resolved with corrected line numbers.
- **Space-before-paren**: VB6 call syntax `FuncName (args)` with space before paren is handled.
- **Position collision fix**: `FindTokenInRawLine` maps effective-line positions to raw-line positions, preventing collisions between With-expanded and raw coordinates.

## Performance Considerations
**Why is VB6 IDE faster?**
1. **Incremental parsing** - VB6 keeps parsed project in memory, doesn't reparse on every load
2. **Lazy resolution** - resolves references only when needed (compile/intellisense), not upfront
3. **Native code** - compiled C++ vs. managed .NET with GC overhead
4. **No semantic analysis** - VB6 doesn't track every reference until compile time
5. **Optimized for single-file edits** - our tool analyzes entire project every run

**Our bottlenecks:**
- Multiple file reads (Parse, Resolve, BuildReplaces) - **FIXED with file cache**
- Full semantic analysis (all References tracked) - needed for precise refactoring
- Regex matching for every symbol on every line - **REDUCED** (single-pass uses char-level scanning for chains/tokens, regex only for hot-path patterns)
- Multiple passes over same data - **FIXED**: resolution is now single-pass per line

**Potential optimizations:**
1. ~~**Single-pass parsing**~~ - ✅ **DONE**: `ResolveTypesAndCalls` scans each line once
2. **Parallel processing** - analyze modules in parallel (Thread.ParallelForEach)
3. **Compiled Regex** - use `RegexOptions.Compiled` for hot-path patterns (partially done)
4. **Span<char>** - avoid string allocations in parsing
5. **Incremental mode** - cache parsed project, only reparse changed files
6. **Lazy References** - optionally skip reference tracking if not needed
7. **Streaming parser** - process line-by-line without loading entire file in memory

