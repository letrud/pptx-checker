---
name: office-file-validate
description: Use after creating or editing any .pptx, .docx or .xlsx file, before delivering it, and whenever a user says Office wants to repair a file or can't read it. Auto-fixes known generator defects, then checks with the Open XML SDK and a PowerPoint-behaviour check.
---

# Office file validation

Schema-valid is not enough. These were all verified in desktop PowerPoint:
- A deck that passes the Open XML SDK validator with 0 errors can still be **unreadable**
  ("Sorry, PowerPoint can't read …").
- LibreOffice and python-pptx open files that PowerPoint rejects.

So run three steps, and deliver only when all of them pass.

## 1. Fix known defects (.pptx, in place)

```bash
python3 <skill-dir>/scripts/fix_known.py deck.pptx
```
| # | Defect (pptxgenjs produces all of these) | What PowerPoint does |
|---|---|---|
| 1 | `<a:gd name="adj">` on an ellipse/rect, from `rectRadius` on an OVAL | drops that shape **and every shape after it** on the slide |
| 2 | `<a:pPr>` after a run in a paragraph | repair prompt |
| 3 | duplicate `<p:cNvPr id>` | repair prompt |
| 4 | `notesMasterIdLst` after `sldIdLst` | schema error (PowerPoint ignores the notes master) |
| 5 | lower-case section GUIDs | schema error |
| 6 | notes master shares the slide master's theme | **unreadable once 4 is fixed**, so 4 and 6 must go together |
| 7 | folder entries / odd order in the zip | repacked like PowerPoint writes it |

Never apply fix 4 without fix 6. `fix_known.py` always does both.

## 2. Check what the schema can't see

```bash
python3 <skill-dir>/scripts/check_semantics.py deck.pptx      # exit 0 = PASS
```
Checks:
- guides on presets that don't define them
- shared notes theme
- duplicate shape IDs
- relationship targets that are missing, and r:ids that are dangling
- slide master and layout IDs
- the section list against the slides

The `phantom-override` warning (Content_Types entries for parts that don't exist) is harmless.

## 3. Open XML SDK validation

```bash
bash <skill-dir>/scripts/validate.sh deck.pptx                # exit 0 = PASS
```
- Works for .pptx, .docx and .xlsx.
- Each error gives the part, XPath, element and description.
- Add `--json` for machine-readable output.
- Exit 3 means there's no binary for this platform (see Setup).

## If anything fails

- **Fix it in the generator** and rebuild. Don't hand-edit the packed file.
- **Find the shape:** shape names (`objectName` in pptxgenjs) appear as `<p:cNvPr name="…">`, which leads back to the code that made it.
- **pptxgenjs: only use `rectRadius` with `ROUNDED_RECTANGLE`.**
- **If a file still fails in real PowerPoint:**
  1. Let PowerPoint repair it and save a copy.
  2. Diff the copy with the original. What PowerPoint dropped or added is the defect.
  3. Add a check here.

| Validator error | Cause | Fix |
|---|---|---|
| `unexpected child element 'pPr'` | rich-text run array | fix_known.py, or give every run options |
| `unexpected child element 'notesMasterIdLst'` | pptxgenjs element order | fix_known.py (with fix 6) |
| `section` `id` Pattern constraint failed | lower-case GUIDs | fix_known.py |
| `attribute 'val' has invalid value` on a colour | `#` or 8-digit hex | 6-digit hex without `#` |

## Setup

`bin/` must contain the binary for this platform, usually `ooxml-validate-linux-x64`. Download it from the GitHub releases of `letrud/pptx-checker`.
