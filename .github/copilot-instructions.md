# Copilot Instructions

## General Guidelines
- First general instruction
- Second general instruction
- When fixing issues, add handling unless there's an obvious error, to avoid regressions.
- Ensure zero references with -1 and include symbol info in refdebug without post-search.
- Throttle progress logging only for inner counters; module header should always print and counters must remain accurate.
- Reference resolution must avoid homonym false positives: when a token is part of a dot-notation member access chain, do not record it as a local/parameter reference; always identify references with precise context.
- When a reported reference fix still reproduces, add defensive pruning to remove local/parameter references whose StartChar points to member-access tokens in dot-notation, even after initial filtering.
- For reference generation, avoid duplicate ownership of the same token position: each line/startChar should resolve to a single symbol, and local/parameter references colliding with member-access symbols must be pruned.

## Code Style
- Use specific formatting rules
- Follow naming conventions, ensuring conventional names are in PascalCase (initial uppercase) so event procedures use uppercase initial.
- For Frame, Panel, and Label controls, if the original name starts with 'Frame', 'Panel', or 'Label' followed by uppercase (e.g., FrameX, PanelY, LabelZ), produce FraX, PnlX, or LblX respectively, only if the remaining suffix is not numeric-only (e.g., keep LblLabel1 as LblLabel1). If the suffix is numeric-only (e.g., Label1), retain the stem (LblLabel1, similarly for Frame/Panel).
- When using regex for arrays, ensure proper escaping in parentheses (e.g., use `"\([^)]*\)"` instead of `"\[^)]*\)"`).
- For `ReFieldAccess`, the correct pattern is `([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)`.
- Follow spacing rules: S0 fallback; S14 wins for initial declaration groups; S7 comments before block use blank + comments + block; S8 comments inside block use block start + comments + blank; single-line If statements must always have a blank line after them unless adjacent to If/End If block boundaries.
- Prefer property Get/Let/Set blocks to be adjacent without blank lines between them.

## Project-Specific Rules
- When parsing VB6 procedures, always use StartLine/EndLine bounds instead of scanning the entire file from LineNumber to prevent duplicate references.
- Key fix: ResolveFieldAccesses and similar functions should scan `for (int i = proc.StartLine - 1; i < proc.EndLine; i++)` not `i < fileLines.Length`.
- Always add safety checks for array bounds with Math.Max/Math.Min.
- For procedure identification in the VB6MagicBox project, use `Parser.Core.cs` to set StartLine/EndLine before reference resolution in `Parser.Resolve.cs`.
- Always use `GetProcedureAtLine(lineNum)` instead of `FirstOrDefault(p => lineNum >= p.LineNumber)`; this ensures the use of the `ContainsLine()` method for accurate procedure identification.
- When reordering variables, keep any `Attribute` line immediately following a procedure signature attached (do not move it).
- When moving local constants/variables to the top of a procedure, left-trim and indent them with two spaces for alignment.
- In this project, `LineReplace.StartChar`/`EndChar` refer to positions in the original source before substitutions; only `NewText` changes later.
- Prefer procedure conflict disambiguation to ignore class modules; class members are accessed via object and not ambiguous.
- For diagnostics exports about references without generated line replaces, include only symbols that actually require a rename (old name differs from new name), excluding already-conventional symbols.