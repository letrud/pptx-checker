"""Auto-fix known generator defects in a .pptx that make PowerPoint offer to repair it.

Fixes, in slide, layout and master parts:
  1. Duplicate shape IDs (<p:cNvPr id>) within a part — renumbered sequentially.
     pptxgenjs can give a shape the same ID as the layout's slide-number placeholder.
  2. <a:pPr> placed after a run inside a paragraph — removed (it must be the first child).
     pptxgenjs emits this for rich-text arrays that mix styled runs.

And in ppt/presentation.xml:
  3. <p:notesMasterIdLst> after <p:sldIdLst> — moved before it (schema order).
  4. Lower-case section GUIDs (<p14:section id>) — upper-cased (schema pattern).

Usage:  python3 fix_known.py deck.pptx   (fixes in place, prints what it changed)
"""
import re
import shutil
import sys
import tempfile
import zipfile

PART = re.compile(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$")
ID = re.compile(r'(<p:cNvPr\b[^>]*?\bid=")(\d+)(")')
STRAY_PPR = re.compile(r"(</a:r>)<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)", re.S)


def fix(path: str) -> dict:
    changes = {"duplicate_ids": 0, "stray_pPr": 0, "notes_master_order": 0, "section_guids": 0}
    tmp = tempfile.mktemp(suffix=".pptx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if PART.match(item.filename):
                xml = data.decode("utf-8")
                xml, n = STRAY_PPR.subn(r"\1", xml)
                changes["stray_pPr"] += n
                ids = [m.group(2) for m in ID.finditer(xml)]
                if len(ids) != len(set(ids)):
                    counter = iter(range(1, 1_000_000))
                    xml = ID.sub(lambda m: f"{m.group(1)}{next(counter)}{m.group(3)}", xml)
                    changes["duplicate_ids"] += 1
                data = xml.encode("utf-8")
            elif item.filename == "ppt/presentation.xml":
                xml = data.decode("utf-8")
                m = re.search(r"<p:notesMasterIdLst>.*?</p:notesMasterIdLst>", xml, re.S)
                if m and xml.find("<p:sldIdLst") < m.start():
                    xml = xml[: m.start()] + xml[m.end():]
                    i = xml.find("<p:sldIdLst")
                    xml = xml[:i] + m.group(0) + xml[i:]
                    changes["notes_master_order"] += 1
                xml, n = re.subn(r'(<p14:section\b[^>]*?\bid=")(\{[^"]+\})',
                                 lambda g: g.group(1) + g.group(2).upper(), xml)
                changes["section_guids"] += sum(1 for _ in re.finditer(r'<p14:section\b[^>]*?\bid="\{[^"]*[a-f]', data.decode("utf-8")))
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return changes


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: fix_known.py deck.pptx")
    c = fix(sys.argv[1])
    print(f"parts with duplicate IDs renumbered: {c['duplicate_ids']}; misplaced <a:pPr> removed: {c['stray_pPr']}; "
          f"notesMasterIdLst moved: {c['notes_master_order']}; section GUIDs upper-cased: {c['section_guids']}")
