"""Checks for .pptx defects that pass Open XML schema validation but still break PowerPoint.

Every check here was found by opening real files in desktop PowerPoint:

  guide-on-preset   An <a:gd> guide on a preset shape that does not define it, e.g. an "adj" guide on
                    an ellipse (pptxgenjs writes one when rectRadius is passed with an OVAL).
                    PowerPoint keeps the shape and silently drops EVERY shape after it on the slide
                    during repair; if other defects are present the file will not open at all.
  shared-theme      The notes master uses the same theme part as a slide master. pptxgenjs does this;
                    PowerPoint tolerates it only while it skips the notes master, which it does when
                    notesMasterIdLst is out of schema order. Put the list in order and the file breaks.
  dup-shape-id      Two shapes in one part with the same <p:cNvPr id>.
  missing-target    A relationship pointing at a part that does not exist.
  dangling-rid      An r:id / r:embed in a part with no matching relationship.
  layout-id         Slide master / layout IDs not unique or below 2^31.
  sections          The section list does not cover every slide exactly once.

Warnings (reported, do not fail):
  phantom-override  [Content_Types].xml overrides for parts that do not exist (pptxgenjs writes these).

Usage:  python3 check_semantics.py deck.pptx [--json]     exit 0 = no errors, 1 = errors
"""
import collections
import json
import posixpath
import re
import sys
import zipfile

# Presets that define no adjust values: any <a:gd> in their avLst is invalid.
NO_ADJUST = {
    "rect", "ellipse", "line", "lineInv", "rtTriangle", "diamond", "heart", "lightningBolt", "cloud",
    "flowChartProcess", "flowChartDecision", "flowChartInputOutput", "flowChartPredefinedProcess",
    "flowChartInternalStorage", "flowChartDocument", "flowChartMultidocument", "flowChartTerminator",
    "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation", "flowChartConnector",
    "flowChartPunchedCard", "flowChartPunchedTape", "flowChartSummingJunction", "flowChartOr",
    "flowChartCollate", "flowChartSort", "flowChartExtract", "flowChartMerge", "flowChartOnlineStorage",
    "flowChartDelay", "flowChartMagneticTape", "flowChartMagneticDisk", "flowChartMagneticDrum",
    "flowChartDisplay", "flowChartOffpageConnector", "flowChartAlternateProcess",
}
# Presets with a single adjust value named "adj".
SINGLE_ADJ = {"roundRect", "snip1Rect", "round1Rect", "triangle", "parallelogram", "trapezoid", "octagon",
              "plaque", "can", "cube", "bevel", "donut", "foldedCorner", "frame", "smileyFace", "sun", "moon",
              "bracketPair", "bracePair", "leftBracket", "rightBracket", "chevron", "homePlate", "plus"}

SHAPE_PARTS = re.compile(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster|notesSlides/notesSlide|notesMasters/notesMaster)\d+\.xml$")
GEOM = re.compile(r'<a:prstGeom prst="([^"]+)"[^>]*>\s*<a:avLst>(.*?)</a:avLst>', re.S)


def check(path):
    errors, warnings = [], []
    z = zipfile.ZipFile(path)
    names = {n for n in z.namelist() if not n.endswith("/")}
    read = lambda n: z.read(n).decode("utf-8", "replace")

    for n in sorted(names):
        if not SHAPE_PARTS.match(n):
            continue
        x = read(n)
        for m in GEOM.finditer(x):
            prst, av = m.group(1), m.group(2)
            gds = re.findall(r'<a:gd name="([^"]+)"', av)
            bad = gds if prst in NO_ADJUST else [g for g in gds if g != "adj"] if prst in SINGLE_ADJ else []
            if bad:
                shape = re.findall(r'name="([^"]*)"', x[:m.start()])
                errors.append(("guide-on-preset", n, f'prst="{prst}" has guide(s) {bad} it does not define (shape "{shape[-1] if shape else "?"}")'))
        ids = re.findall(r'<p:cNvPr\b[^>]*?\bid="(\d+)"', x)
        dup = sorted(k for k, v in collections.Counter(ids).items() if v > 1)
        if dup:
            errors.append(("dup-shape-id", n, f"duplicate shape ids {dup}"))

    # relationships
    for r in sorted(n for n in names if n.endswith(".rels")):
        base = posixpath.dirname(posixpath.dirname(r))
        rx = read(r)
        for tgt, mode in re.findall(r'Target="([^"]+)"(?:\s+TargetMode="(\w+)")?', rx):
            if mode == "External":
                continue
            p = tgt[1:] if tgt.startswith("/") else posixpath.normpath(posixpath.join(base, tgt))
            if p not in names:
                errors.append(("missing-target", r, f"target {tgt} does not exist"))
        src = posixpath.join(base, posixpath.basename(r)[:-5])
        if src in names and src.endswith(".xml"):
            have = set(re.findall(r'Id="([^"]+)"', rx))
            used = set(re.findall(r'r:(?:id|embed|link|pict)="([^"]+)"', read(src)))
            if used - have:
                errors.append(("dangling-rid", src, f"r:ids with no relationship {sorted(used - have)}"))

    # notes master must not share a theme with a slide master
    def themes(pattern):
        out = {}
        for n in names:
            if re.match(pattern, n):
                out[n] = set(re.findall(r'Target="\.\./theme/(theme\d+\.xml)"', read(n)))
        return out
    master_t = set().union(*themes(r"ppt/slideMasters/_rels/slideMaster\d+\.xml\.rels$").values() or [set()])
    for n, t in themes(r"ppt/notesMasters/_rels/notesMaster\d+\.xml\.rels$").items():
        if t & master_t:
            errors.append(("shared-theme", n, f"notes master uses {sorted(t & master_t)}, which a slide master also uses"))

    # presentation.xml ids and sections
    if "ppt/presentation.xml" in names:
        P = read("ppt/presentation.xml")
        ids = re.findall(r'<p:sldMasterId id="(\d+)"', P)
        for m in sorted(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)):
            ids += re.findall(r'<p:sldLayoutId id="(\d+)"', read(m))
        if len(ids) != len(set(ids)) or any(int(i) < 2**31 for i in ids):
            errors.append(("layout-id", "ppt/presentation.xml", "slide master/layout ids must be unique and >= 2147483648"))
        slides = re.findall(r'<p:sldId id="(\d+)"', P)
        sec = re.findall(r'<p14:sldId id="(\d+)"', P)
        if sec and sorted(sec) != sorted(slides):
            errors.append(("sections", "ppt/presentation.xml", "section list does not match the slide list"))

    ct = read("[Content_Types].xml")
    phantom = [o for o in re.findall(r'PartName="/([^"]+)"', ct) if o not in names]
    if phantom:
        warnings.append(("phantom-override", "[Content_Types].xml", f"{len(phantom)} override(s) for missing parts, e.g. {phantom[0]}"))
    return errors, warnings


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--json"]
    if len(args) != 1:
        sys.exit("usage: check_semantics.py deck.pptx [--json]")
    errs, warns = check(args[0])
    if "--json" in sys.argv:
        print(json.dumps({"file": args[0], "errors": errs, "warnings": warns}, indent=2))
    else:
        print(("FAIL" if errs else "PASS") + f"  {args[0]}  ({len(errs)} error(s), {len(warns)} warning(s))")
        for kind, part, msg in errs:
            print(f"  [{kind}] {part}: {msg}")
        for kind, part, msg in warns:
            print(f"  warning [{kind}] {part}: {msg}")
    sys.exit(1 if errs else 0)
