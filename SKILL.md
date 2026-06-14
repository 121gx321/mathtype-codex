---
name: mathtype
description: Convert DOCX body inline LaTeX placeholders into real editable MathType Equation.DSMT4 OLE objects in their original paragraph positions while preserving existing MathType display equations. Use for Word manuscripts where inline formulas, variables, symbols, or PDF-transcribed formulas must become editable MathType OLE rather than OMML, images, or plain LaTeX.
---

# MathType Inline OLE

Use this skill for fragile Word/MathType manuscript work where body inline formulas must be real editable MathType `Equation.DSMT4` OLE objects. The key lesson is that MathType OLE insertion is Windows desktop automation, not pure OOXML editing.

## Required Environment

Use a real Windows desktop session with:

- Microsoft Word COM available.
- Activated MathType with `Equation.DSMT4` registered.
- Python with `pywin32`, `lxml`, and `latex2mathml`.
- The `MathType-Word-WPS` CLI available, usually `scripts/mathtype_word_wps.py` from `https://github.com/pcdp577/MathType-Word-WPS`.

Run a quick preflight on new machines or after Office/MathType changes:

```powershell
python path\to\mathtype_word_wps.py check-env --probe-com --probe-mathtype
```

Proceed only if `pywin32=ok`, `latex2mathml=ok`, `ole_class=Equation.DSMT4 registered=True`, `mathtype_runtime_probe=True`, and `backend=word ... available=True`.

## Workflow

1. Preserve the original DOCX. Never overwrite the user's source file.
2. Identify existing display formula paragraphs. Do not edit `MTDisplayEquation` paragraphs or existing MathType OLE display formulas.
3. Put inline formulas into correct body positions as `$...$` LaTeX text. For PDF-based jobs, transcribe/correct variables from the PDF first rather than guessing from flawed Word text.
4. Run `scripts/make_placeholders.py` to replace each `$...$` with a unique `@@MT0001@@` placeholder and write a JSON map.
5. Run `scripts/replace_placeholders_with_mathtype.py` with `--limit 1` repeatedly until all placeholders are replaced.
6. Run `scripts/check_docx_mathtype.py` and verify counts and placement.
7. Run the MathType-Word-WPS `cleanup-leftovers` command.

## Critical Insertion Rule

Always insert inline MathType OLE through the current Word selection:

```text
Selection.Find placeholder -> Selection.Delete -> Selection.InlineShapes.AddOLEObject(...)
```

Do not use:

```text
doc.InlineShapes.AddOLEObject(..., Range=...)
```

For `Equation.DSMT4`, Word may ignore the `Range` argument and put every inline OLE at the start of the document, producing a page-top pile of formulas.

## Commands

Create placeholders:

```powershell
python scripts\make_placeholders.py `
  --input input_latex.docx `
  --output placeholders.docx `
  --map-json placeholders.json
```

Replace placeholders one at a time:

```powershell
Copy-Item placeholders.docx final_mathtype.docx -Force
for ($i = 1; $i -le 64; $i++) {
  python scripts\replace_placeholders_with_mathtype.py `
    --input final_mathtype.docx `
    --output final_mathtype.docx `
    --map-json placeholders.json `
    --tool-script path\to\mathtype_word_wps.py `
    --limit 1 `
    --backend Word.Application `
    --wait 1.0
}
```

Check the result:

```powershell
python scripts\check_docx_mathtype.py --docx final_mathtype.docx --paragraphs 4,10,14,19,25,41,59
```

## Acceptance Criteria

The final DOCX must satisfy:

- All expected OLE objects have `ProgID="Equation.DSMT4"`.
- Existing display formula paragraph count is unchanged.
- `@@MT` placeholder count is `0`.
- `$` LaTeX delimiter count is `0`.
- Inline OLE objects are distributed across the target body paragraphs, not concentrated in paragraph 0.

If formulas appear piled together near the page top, discard the output and rerun with the selection-based insertion script.


