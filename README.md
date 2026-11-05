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

`scripts/fix_known.py deck.pptx` repairs two pptxgenjs defects in place (duplicate shape
IDs, misplaced `<a:pPr>`) that make PowerPoint ask to repair the file.

## Get the binary

Download from the [latest release](../../releases/latest):
`ooxml-validate-linux-x64`, `-linux-arm64`, `-osx-arm64`, `-win-x64.exe`, plus `SHA256SUMS`.

Each is self-contained: the .NET runtime and the Open XML SDK are inside. No install needed.

## Build it yourself

```bash
dotnet publish src/OoxmlValidate.csproj -c Release -r linux-x64 -o out
```

## Release

Push a tag `v*` (for example `v0.1.0`). The `build` workflow publishes all four binaries and
attaches them, with checksums, to a GitHub release.

## Agent skill

`skill/SKILL.md` plus `scripts/` and a `bin/` folder holding the release binary make up an
agent skill that validates and auto-fixes generated Office files before delivery. To
package it: copy `skill/SKILL.md`, `scripts/`, and the downloaded binary into `bin/` of one
folder, then upload that folder as a skill.
