# VB6MagicBox - Technical Summary (Copilot)

## Purpose
VB6 parser/refactoring tool. Pipeline: parse VB6 project, resolve references (types/calls/fields), build dependencies/usage, apply naming conventions, export outputs, optionally refactor sources.

## Key Models
- `VbModule`: `Procedures`, `Properties` (separate), `Types`, `Enums`, `Constants`, `GlobalVariables`, `Controls`, `ModuleReferences`, `Used`, `StartLine/EndLine` lookups via `ContainsLine`/`GetProcedureAtLine`.
- `VbProcedure`: `StartLine/EndLine`, `Parameters`, `LocalVariables`, `Calls`, `References`, `ReturnType`.
- `VbProperty`: same as procedure (separate list), `ReturnType`, `Parameters`, `References`.
- `VbTypeDef`/`VbField`, `VbEnumDef`/`VbEnumValue` with `References`.
- `VbControl`: `LineNumbers` for control arrays, `References`.

## Parsing (Parser.Core)
- Handles `Function/Sub/Property` and `Declare` (supports `Alias`, optional visibility).
- Collapses line continuations `_` with line mapping; multiline signature parameter line numbers are corrected.
- Regexes for array parens use `\([^)]*\)`; `ReFieldAccess` pattern: `([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)`.
- `Implements` lines recorded in `VbModule.ImplementsInterfaces`.

## Resolution (Parser.Resolve + .Members)
- `ResolveTypesAndCalls` builds `procIndex` (procedures only) + `propIndex`.
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

## Refactoring (Refactoring.cs)
- Applies renames using `References` line numbers only (no global scan).
- Special cases:
  - `Begin <lib> <control>` renames only control name.
  - Properties outside defining module use dot-prefixed rename.
  - Attribute `VB_Name` updated by module rename.
- For constants only: replacements avoid string literals (do not alter constant values). Other categories allow string replacements (needed for `Attribute VB_Name = "..."`).

## Exports
`ExportProjectFiles` writes: `*.symbols.json`, `*.rename.json`, `*.rename.csv`, `*.dependencies.md`.

## TypeAnnotator
- Adds missing types only when suffix infers type. No default `As Object`.
- Collects missing type cases (no suffix) and exports `*.missingTypes.csv` with `Module,Procedure,Name,Kind`.

## Naming
- Modules: `Frm`/`Cls` prefix, PascalCase.
- Procedures: detect standard events, control events, WithEvents, interface-implemented (`Implements`) event handlers keep name (underscore preserved).
- Properties share `ConventionalName` across Get/Let/Set.
- Enum naming preserves short acronyms (e.g., `SQM`).

## Known Pitfalls / Rules
- Always use `StartLine/EndLine` bounds when scanning.
- Regex for arrays must escape parentheses correctly (`\([^)]*\)`).
- `ReFieldAccess` and nested chain regex must allow array indexing.
- Avoid renaming inside string literals except for constants (to keep constant values intact).
