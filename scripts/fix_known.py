"""Auto-fix known generator defects in a .pptx that make PowerPoint repair or refuse it.

All verified against desktop PowerPoint (Microsoft 365, Windows). Fixes:

  Slide, layout and master parts
    1. Invalid guides on presets that do not define them (e.g. "adj" on an ellipse, written by
       pptxgenjs when rectRadius is given to an OVAL) — removed. PowerPoint otherwise drops every
       shape after it on the slide.
    2. <a:pPr> placed after a run inside a paragraph — removed (it must be the first child).
    3. Duplicate shape IDs (<p:cNvPr id>) within a part — renumbered sequentially.

  ppt/presentation.xml
    4. <p:notesMasterIdLst> after <p:sldIdLst> — moved before it (schema order).
    5. Lower-case section GUIDs (<p14:section id>) — upper-cased (schema pattern).

  Package
    6. Notes master sharing the slide master's theme part — given its own copy (themeN.xml).
       REQUIRED together with fix 4: once the notes master list is in schema order PowerPoint
       loads the notes master, and a shared theme makes the file unreadable.
    7. Repacked like PowerPoint writes: [Content_Types].xml first, no folder entries, all deflated.

Usage:  python3 fix_known.py deck.pptx   (fixes in place, prints what it changed)
Run check_semantics.py and ooxml-validate afterwards.
"""
import re
import sys
import zipfile

PART = re.compile(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$")
ID = re.compile(r'(<p:cNvPr\b[^>]*?\bid=")(\d+)(")')
STRAY_PPR = re.compile(r"(</a:r>)<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)", re.S)
NO_ADJ_GUIDES = re.compile(r'(<a:prstGeom prst="(?:rect|ellipse|line|lineInv|rtTriangle|diamond|heart|lightningBolt|cloud|flowChart\w+)"[^>]*>\s*<a:avLst>)(.*?)(</a:avLst>)', re.S)
THEME_CT = "application/vnd.openxmlformats-officedocument.theme+xml"


def fix(path: str) -> dict:
    c = dict.fromkeys(["bad_guides", "stray_pPr", "duplicate_ids", "notes_master_order", "section_guids", "notes_theme"], 0)
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}

    for n in list(files):
        if PART.match(n):
            x = files[n].decode("utf-8")
            x, k = NO_ADJ_GUIDES.subn(lambda m: m.group(1) + m.group(3) if "<a:gd" in m.group(2) else m.group(0), x)
            c["bad_guides"] += len(re.findall(r"<a:gd", "".join(m.group(2) for m in NO_ADJ_GUIDES.finditer(files[n].decode("utf-8")))))
            x, k = STRAY_PPR.subn(r"\1", x)
            c["stray_pPr"] += k
            ids = [m.group(2) for m in ID.finditer(x)]
            if len(ids) != len(set(ids)):
                counter = iter(range(1, 1_000_000))
                x = ID.sub(lambda m: f"{m.group(1)}{next(counter)}{m.group(3)}", x)
                c["duplicate_ids"] += 1
            files[n] = x.encode("utf-8")

    if "ppt/presentation.xml" in files:
        x = files["ppt/presentation.xml"].decode("utf-8")
        m = re.search(r"<p:notesMasterIdLst>.*?</p:notesMasterIdLst>", x, re.S)
        if m and x.find("<p:sldIdLst") < m.start():
            x = x[: m.start()] + x[m.end():]
            i = x.find("<p:sldIdLst")
            x = x[:i] + m.group(0) + x[i:]
            c["notes_master_order"] += 1
        c["section_guids"] += len(re.findall(r'<p14:section\b[^>]*?\bid="\{[^"]*[a-f][^"]*\}"', x))
        x = re.sub(r'(<p14:section\b[^>]*?\bid=")(\{[^"]+\})', lambda g: g.group(1) + g.group(2).upper(), x)
        files["ppt/presentation.xml"] = x.encode("utf-8")

    master_themes = set()
    for n in files:
        if re.match(r"ppt/slideMasters/_rels/slideMaster\d+\.xml\.rels$", n):
            master_themes |= set(re.findall(r'Target="\.\./theme/(theme\d+\.xml)"', files[n].decode()))
    for r in [n for n in files if re.match(r"ppt/notesMasters/_rels/notesMaster\d+\.xml\.rels$", n)]:
        x = files[r].decode()
        m = re.search(r'Target="\.\./theme/(theme\d+\.xml)"', x)
        if m and m.group(1) in master_themes:
            k = 1
            while f"ppt/theme/theme{k}.xml" in files:
                k += 1
            files[f"ppt/theme/theme{k}.xml"] = files[f"ppt/theme/{m.group(1)}"]
            files[r] = x.replace(m.group(0), f'Target="../theme/theme{k}.xml"').encode()
            ct = files["[Content_Types].xml"].decode()
            files["[Content_Types].xml"] = ct.replace("</Types>", f'<Override PartName="/ppt/theme/theme{k}.xml" ContentType="{THEME_CT}"/></Types>').encode()
            c["notes_theme"] += 1

    order = sorted(files, key=lambda n: (n != "[Content_Types].xml", n != "_rels/.rels"))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0)), files[n], zipfile.ZIP_DEFLATED)
    return c


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: fix_known.py deck.pptx")
    print("fixed:", ", ".join(f"{k}={v}" for k, v in fix(sys.argv[1]).items()))
