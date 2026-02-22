# VB6MagicBox - Technical Summary (Copilot)

## Purpose
VB6 parser/refactoring tool. Pipeline: parse VB6 project, resolve references (types/calls/fields), build dependencies/usage, apply naming conventions, **build precise replaces**, export outputs, apply refactoring.

## Architecture (NEW - Optimized)
**Phase 1 (Analysis)**: `Parser.ParseAndResolve`
1. Parse project files (`Parser.Core`)
2. Resolve types/calls/fields (`Parser.Resolve` + `.Members`)
3. Build dependencies & mark used symbols (`Parser.Resolve.Dependencies`)
4. Apply naming conventions & sort (`Parser.Export` → `NamingConvention.Apply`)
5. **Build Replaces** (`Parser.Replaces`) - **NEW**: pre-calculate exact character positions for all renames

**Phase 2 (Refactoring)**: `Refactoring.ApplyRenames`
- ⚡ Ultra-fast: uses pre-calculated `VbModule.Replaces[]` (line + char position)
- No regex matching, no re-parsing
- Applies substitutions from end to start (preserves positions)

**Phase 3 (Type Annotation)**: `TypeAnnotator.AddMissingTypes`
- Runs AFTER refactoring (uses conventional names already applied)

## Key Models
- `VbModule`: `Procedures`, `Properties` (separate), `Types`, `Enums`, `Constants`, `GlobalVariables`, `Controls`, `ModuleReferences`, `Used`, **`Replaces`** (NEW - list of `LineReplace`), `StartLine/EndLine` lookups via `ContainsLine`/`GetProcedureAtLine`.
- `VbProcedure`: `StartLine/EndLine`, `Parameters`, `LocalVariables`, `Calls`, `References`, `ReturnType`.
- `VbProperty`: same as procedure (separate list), `ReturnType`, `Parameters`, `References`.
- `VbTypeDef`/`VbField`, `VbEnumDef`/`VbEnumValue` with `References`.
- `VbControl`: `LineNumbers` for control arrays, `References`.
- **`LineReplace`** (NEW): `LineNumber`, `StartChar`, `EndChar`, `OldText`, `NewText`, `Category` - precise substitution with exact character position.

## Parsing (Parser.Core)
- Handles `Function/Sub/Property` and `Declare` (supports `Alias`, optional visibility).
- Collapses line continuations `_` with line mapping; multiline signature parameter line numbers are corrected.
- Regexes for array parens use `\([^)]*\)`; `ReFieldAccess` pattern: `([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)`.
- `Implements` lines recorded in `VbModule.ImplementsInterfaces`.

## Resolution (Parser.Resolve + .Members)
- `ResolveTypesAndCalls` builds `procIndex` (procedures only) + `propIndex`.
- **Bare property resolution**: Public properties (PropertyGet) usable bare across modules (e.g., `If ExecSts = ...`) resolved in PASS 1.2 alongside bare procedure calls.
- **External type exclusion**: Field access chains stop if base type not in `typeIndex` or `classIndex` (e.g., `gobjFM489.ActualState.Program.Frequency_Long` where `gobjFM489` is external COM object).
- Field access resolution handles:
  - nested dot chains with arrays at any segment,
  - `With` blocks (prefix `With` expression and replace inline `.Member`),
  - module-prefixed chains (e.g., `Cnf.Config...`) by resolving module globals/properties to a type,
  - class-property chaining (if type is class, use property `ReturnType` to continue chain).
- References are accumulated per occurrence (no skipping for properties).
- Property blocks are resolved separately (field/parameter/return references).
- Function return variable is typed in `env` (function name -> return type) and return assignments are referenced for renames.
- Enum values: explicit pass `ResolveEnumValueReferences` adds references for bare enum values and qualified `EnumName.Value`, respecting shadowing.
- Class usage via `ResolveClassModuleReferences`.
- Type references via `ResolveTypeReferences` for every `As TypeName` occurrence.
- `MarkUsedTypes` also scans Type fields.
- Module `Used` propagated from any used member.
- Public properties are scanned across modules (like globals) for bare uses.

## BuildReplaces (Parser.Replaces) - NEW
- **Pre-calculates ALL substitutions** during analysis phase (after naming conventions applied).
- For each symbol: finds exact character position in source using `References.LineNumbers`.
- Special handling:
  - Properties in other modules: only `.PropertyName` (dot-prefixed)
  - Controls in `Begin` lines: only control name (skip library/type)
  - Attributes: `Attribute VB_Name = "..."` and `Attribute VarName.VB_VarHelpID`
  - Constants: `skipStringLiterals=true` to avoid replacing inside string values
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
- Collects missing type cases (no suffix) and exports `*.missingTypes.csv` with `Module,Procedure,Name,Kind`.
- **Runs AFTER refactoring** (uses conventional names already applied).

## Naming
- Modules: `Frm`/`Cls` prefix, PascalCase.
- Procedures: detect standard events, control events, WithEvents, interface-implemented (`Implements`) event handlers keep name (underscore preserved).
- Properties share `ConventionalName` across Get/Let/Set.
- Enum naming preserves short acronyms (e.g., `SQM`).

## Known Pitfalls / Rules
- Always use `StartLine/EndLine` bounds when scanning.
- Regex for arrays must escape parentheses correctly (`\([^)]*\)`).
- `ReFieldAccess` and nested chain regex must allow array indexing.
- Constants: use `skipStringLiterals=true` when building replaces to avoid altering constant values in quotes.

## Recent Fixes (2024)
- **External object members**: member-access tokens (e.g., `obj.Prop`) are excluded from parameter/local reference scans so external member names are not renamed.
- **Enum value collisions**: when multiple enum values converge to the same `ConventionalName`, references are qualified as `EnumName.Value` to avoid ambiguity.
- **Enum qualification guard**: if the value is already qualified (`EnumName.Value`), only the value is renamed, preventing `Enum.Enum.Value` or wrong enum prefixes.
- **Inline calls after `Then`**: procedure references are recorded even if a call is already present in `Calls`, ensuring inline statements are renamed.
- **Screaming snake to PascalCase**: enum value naming preserves only known acronyms, so `RIC_RUN_CMD_CALLER` -> `RicRunCmdCaller`.
- **String-aware comment stripping**: comment removal ignores apostrophes inside string literals, preventing truncation of lines with text (e.g., `"E' ..."`).

## Performance Considerations
**Why is VB6 IDE faster?**
1. **Incremental parsing** - VB6 keeps parsed project in memory, doesn't reparse on every load
2. **Lazy resolution** - resolves references only when needed (compile/intellisense), not upfront
3. **Native code** - compiled C++ vs. managed .NET with GC overhead
4. **No semantic analysis** - VB6 doesn't track every reference until compile time
5. **Optimized for single-file edits** - our tool analyzes entire project every run

**Our bottlenecks:**
- Multiple file reads (Parse, Resolve, BuildReplaces) - **FIXED with file cache in BuildReplaces**
- Full semantic analysis (all References tracked) - needed for precise refactoring
- Regex matching for every symbol on every line - expensive
- Multiple passes over same data (Parse → Resolve → References → Replaces)

**Potential optimizations:**
1. **Single-pass parsing** - combine parsing + resolution in one pass
2. **Parallel processing** - analyze modules in parallel (Thread.ParallelForEach)
3. **Compiled Regex** - use `RegexOptions.Compiled` for hot-path patterns
4. **Span<char>** - avoid string allocations in parsing
5. **Incremental mode** - cache parsed project, only reparse changed files
6. **Lazy References** - optionally skip reference tracking if not needed
7. **Streaming parser** - process line-by-line without loading entire file in memory
