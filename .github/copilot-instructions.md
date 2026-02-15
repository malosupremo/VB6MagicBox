# Copilot Instructions

## General Guidelines
- First general instruction
- Second general instruction

## Code Style
- Use specific formatting rules
- Follow naming conventions

## Project-Specific Rules
- When parsing VB6 procedures, always use StartLine/EndLine bounds instead of scanning the entire file from LineNumber to prevent duplicate references.
- Key fix: ResolveFieldAccesses and similar functions should scan `for (int i = proc.StartLine - 1; i < proc.EndLine; i++)` not `i < fileLines.Length`.
- Always add safety checks for array bounds with Math.Max/Math.Min.
- For procedure identification in the VB6MagicBox project, use `Parser.Core.cs` to set StartLine/EndLine before reference resolution in `Parser.Resolve.cs`.
- Always use `GetProcedureAtLine(lineNum)` instead of `FirstOrDefault(p => lineNum >= p.LineNumber)`; this ensures the use of the `ContainsLine()` method for accurate procedure identification.