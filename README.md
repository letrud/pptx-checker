# pptx-checker

Validate `.pptx`, `.docx` and `.xlsx` files with Microsoft's Open XML SDK — the rule set
Office applies — from a single self-contained executable. Built for agent sandboxes where
NuGet is unreachable and nothing can be installed, but useful anywhere.

## Use

```bash
ooxml-validate deck.pptx                      # PASS / FAIL with part, XPath, description
ooxml-validate --json deck.pptx report.docx   # machine-readable
ooxml-validate --format Office2016 deck.pptx  # validate against an older Office version
```

Exit codes: `0` valid · `1` errors found · `2` usage error.

### Schema-valid is not enough

A file can pass the SDK validator with 0 errors and still be unreadable in PowerPoint. Two scripts cover what the schema can't see. Each defect they handle was found by opening real files in desktop PowerPoint (Microsoft 365, Windows) and diffing PowerPoint's repaired copy against the original.

```bash
python3 scripts/fix_known.py deck.pptx        # auto-fix known pptxgenjs defects, in place
python3 scripts/check_semantics.py deck.pptx  # PowerPoint-behaviour checks; exit 0 = PASS
```

Two of the findings:

- **Guides a preset doesn't define.** An `adj` guide on an ellipse (pptxgenjs `rectRadius` on an OVAL) makes PowerPoint drop that shape and every shape after it on the slide.
- **Shared theme.** pptxgenjs gives the notes master the slide master's theme. That's harmless only while `notesMasterIdLst` is out of schema order, because PowerPoint then ignores the notes master. Put the list in order without giving the notes master its own theme, and PowerPoint can't read the file. `fix_known.py` always does both.

The full list is in [`skill/SKILL.md`](skill/SKILL.md).

## Get the binary

Download from the [latest release](../../releases/latest):
`ooxml-validate-linux-x64`, `-linux-arm64`, `-osx-arm64`, `-win-x64.exe`, plus `SHA256SUMS`.

Each is self-contained: the .NET runtime and the Open XML SDK are inside. No install needed.

## Build it yourself

```bash
dotnet publish src/OoxmlValidate.csproj -c Release -r linux-x64 -o out
```

## Enable the build

The workflow is kept at `ci/build.yml`. Move it into place once (needs a token with the
`workflow` permission, e.g. your own account):

```bash
git mv ci/build.yml .github/workflows/build.yml && git commit -m "Enable CI" && git push
```

Or in the GitHub web UI: **Add file → Create new file**, name it `.github/workflows/build.yml`,
and paste the contents of `ci/build.yml`.

## Release

Push a tag `v*` (for example `v0.1.0`). The `build` workflow publishes all four binaries and
attaches them, with checksums, to a GitHub release.

## Agent skill

`skill/SKILL.md` plus `scripts/` and a `bin/` folder holding the release binary make up an
agent skill that validates and auto-fixes generated Office files before delivery. To
package it: copy `skill/SKILL.md`, `scripts/`, and the downloaded binary into `bin/` of one
folder, then upload that folder as a skill.
