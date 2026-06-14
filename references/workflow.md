# MathType Inline OLE Reference

## Fast Path

1. Produce a LaTeX-in-body DOCX from the source manuscript:
   - Use `$Con_{itn,a,c,h}$` style inline LaTeX.
   - Keep display equations and equation-number paragraphs unchanged.
   - For PDF source, verify inline variable order and naming against PDF text or page images.
2. Convert inline LaTeX text to placeholders:
   - `@@MT0001@@`, `@@MT0002@@`, ...
   - JSON map stores `placeholder -> latex`.
3. Replace placeholders one at a time through Word selection:
   - Open Word hidden.
   - Select last remaining placeholder with `Selection.Find`.
   - Delete the selected placeholder.
   - Call `Selection.InlineShapes.AddOLEObject(ClassType="Equation.DSMT4")`.
   - Paste normalized MathML into MathType editor.
   - Update host and close MathType.
   - Save and close Word.
4. Repeat until no placeholders remain.
5. Verify placement by paragraph OLE counts.

## Why Placeholder Mode Is Required

Do not search for raw `$...$` after insertion. MathType OLE native/cache layers may expose formula text to Word search and cause repeated hits on the same already-replaced formula.

Unique placeholders avoid this because `@@MT0001@@` only exists in the body text and disappears after replacement.

## Why One Formula Per Word Session Is Recommended

After MathType returns control to Word, adjacent Range objects in the same paragraph can become non-editable or shift unexpectedly. Closing and reopening Word after each formula is slower but stable.

## Placement Debugging

Bad output symptom:

- Paragraph 0 contains many OLE objects.
- Target body paragraphs contain blank spaces where formulas should be.
- Formulas appear visually piled near the page top.

Cause:

- The script used `AddOLEObject(..., Range=...)` instead of `Selection.InlineShapes.AddOLEObject(...)`.

Correct output:

- Paragraph 0 has only the original display equation OLE for that paragraph.
- Target body paragraphs contain OLE counts matching formula counts.
- Body text around formulas remains intact.

## Useful Checks

Expected XML checks:

- `//o:OLEObject[@ProgID="Equation.DSMT4"]`
- `//w:p[w:pPr/w:pStyle/@w:val="MTDisplayEquation"]`
- no `@@MT` in `//w:t`
- no `$` delimiters in `//w:t`

Use an ASCII temp file path for checks on Windows if Python or PowerShell mangles Chinese paths.
