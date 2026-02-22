# Spacing Rules (VB6MagicBox)

## Priority (high -> low)
1. **SR10** Header form section is immutable (skip until first `Attribute` or `Option`).
2. **SR11** Add exactly one blank line before `Option Explicit` / `Option Base` (after header).
3. **SR4/SR5** Add one blank line after `End Sub/Function/Property/Type/Enum`, except between `Property Get/Let/Set` of the same property (no blank).
4. **SR7/SR8** Comment attachment rules (before/inside blocks).
5. **SR14** Initial procedure groups separated by one blank: comments → const → static → dim.
6. **SR15/SR16/SR17** Select Case spacing, labels, single-line If spacing.
7. **SR1/SR2/SR3/SR12/SR13** Blank line normalization (collapse, trim, no blank between consecutive declarations, compact enum values).
8. **SR0** Fallback: no blank lines between ordinary statements.

## Rules
- **SR0** Default: no blank lines between consecutive code lines.
- **SR1** Collapse consecutive blank lines to a single blank line.
- **SR2** Lines containing only whitespace become empty lines.
- **SR3** No blank line between consecutive `Dim`/`Static` declarations.
- **SR4** Ensure one blank line after `End Sub/Function/Property/Type/Enum`.
- **SR5** Exception: `Property Get/Let/Set` with the same name remain attached (no blank between).
- **SR6** No blank line immediately inside a block (first/last line inside a block).
- **SR7** Comments before a block: blank line, then comments, then the block.
- **SR8** Comments at block start: block start, comments, then a blank line.
- **SR9** Nested blocks remain attached (no blank between `Then` and `If`, etc.).
- **SR10** Form header (`VERSION/Begin...`) is untouched until the first `Attribute` or `Option`.
- **SR11** Insert one blank line before `Option Explicit` / `Option Base` (after header).
- **SR12** No blank line between consecutive `Const` declarations.
- **SR13** No blank lines between enum values.
- **SR14** Procedure top groups are separated by one blank line: header comments → const → static → dim.
- **SR15** `Select Case`: blank line before each `Case` except the first.
- **SR16** Labels (`Name:`) must have a blank line before and no blank immediately after.
- **SR17** Single-line `If` gets a blank line before/after, unless adjacent to `If`/`End If` boundaries.
