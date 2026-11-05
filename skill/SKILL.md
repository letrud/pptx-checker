---
name: office-file-validate
description: Validate a generated .pptx, .docx or .xlsx against Microsoft's Open XML SDK rules before handing it over, and auto-fix known generator defects that make PowerPoint ask to repair the file. Use after building any Office file, and whenever a user reports "needs repair" or "corrupt".
---

# Office file validation (Open XML SDK)

LibreOffice renders, and the bundled XSD checks pass, on files that real PowerPoint
still calls corrupt. This skill runs Microsoft's own validator (Open XML SDK, the rules
Office applies), shipped as a self-contained binary in `bin/` — no network, no .NET install.

## Workflow

Run these after every build of an Office file, before sending it:

1. **Fix known defects** (pptx only, in place):
   ```bash
   python3 <skill-dir>/scripts/fix_known.py deck.pptx
   ```
   Repairs duplicate shape IDs, `<a:pPr>` placed after a run, `notesMasterIdLst` out of order and lower-case section GUIDs — both produced by
   pptxgenjs and both trigger PowerPoint's repair prompt.

2. **Validate**:
   ```bash
   bash <skill-dir>/scripts/validate.sh deck.pptx
   ```
   Exit 0 = valid. Exit 1 = errors, each printed with the part, XPath, element and
   description. Exit 3 = no binary for this platform (see Setup).
   Add `--json` for machine-readable output, `--format Office2016` to target older Office.

3. **If errors remain**, fix them in the generator, not by hand-editing the packed file.
   Use the part + XPath to find the shape; generator shape names (`objectName`) show in
   the XML as `<p:cNvPr name="…">`, which points back to the code that made it.
   Rebuild, re-run steps 1–2, and only deliver when the result is PASS.

## Reading common errors

| Error description contains | Usual cause | Fix in the generator |
|---|---|---|
| `The element has unexpected child element 'pPr'` | rich-text run array mixing styled runs | `fix_known.py`, or give every run explicit options |
| duplicate `id` / "id attribute" | shape ID clash with slide-number placeholder | `fix_known.py` |
| `unexpected child element 'notesMasterIdLst'` in presentation.xml | pptxgenjs writes it after `sldIdLst` | `fix_known.py` |
| `section` `id` ... Pattern constraint failed | pptxgenjs writes lower-case GUIDs for `sectionTitle` | `fix_known.py` |
| `attribute 'val' has invalid value` on a colour | `#` or 8-digit hex colour | 6-digit hex without `#` |
| errors in `chart*.xml` | pptxgenjs combo/axis options | see the pptx skill's chart notes |

## Setup

The `bin/` folder must contain the binary for the sandbox platform, usually
`ooxml-validate-linux-x64`. Binaries are built by GitHub Actions in
`github.com/letrud/pptx-checker` and attached to each release; download the one you need
into `bin/` when (re)packaging this skill.
