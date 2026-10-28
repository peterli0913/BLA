#!/usr/bin/env python3
"""Print the accepted text of a .docx (or the cells of a .xlsx) one searchable line per row.

A "line" is one Word paragraph (body, table cell, header, footer, footnote) with tracked
deletions dropped and insertions kept, i.e. what Word's Find sees after Accept All.
For .xlsx, a line is one non-empty cell.

    python docx_lines.py MOR.docx                 # all lines, numbered
    python docx_lines.py MOR.docx --grep "NMT"    # only lines matching a regex
    python docx_lines.py MOR.docx --out mor.txt   # save for reading or grepping later
"""

import argparse
import re
import sys
import zipfile
from collections import namedtuple
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
SKIP = {W + "del", W + "moveFrom", W + "txbxContent", MC + "Fallback"}

Line = namedtuple("Line", "part index text")


def _para_text(p):
    parts = []

    def walk(node):
        if node.tag in SKIP:
            return
        if node.tag == W + "t" and node.text:
            parts.append(node.text)
        elif node.tag == W + "tab":
            parts.append("\t")
        elif node.tag in (W + "br", W + "cr"):
            parts.append("\n")
        elif node.tag == W + "noBreakHyphen":
            parts.append("\u2011")
        for child in node:
            walk(child)

    walk(p)
    return "".join(parts)


def docx_lines(path):
    z = zipfile.ZipFile(path)
    names = ["word/document.xml"] + sorted(
        n for n in z.namelist() if re.match(r"word/(header|footer)\d*\.xml$|word/(footnotes|endnotes)\.xml$", n)
    )
    out = []
    for name in names:
        root = etree.fromstring(z.read(name))
        part = Path(name).stem
        for i, p in enumerate(root.iter(W + "p")):
            text = _para_text(p)
            if text.strip():
                out.append(Line(part, i, text))
    return out


def xlsx_lines(path):
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and str(cell.value).strip():
                    out.append(Line(ws.title, cell.coordinate, str(cell.value)))
    return out


def lines(path):
    path = Path(path)
    return xlsx_lines(path) if path.suffix.lower() in (".xlsx", ".xlsm") else docx_lines(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--grep", help="regex; print only matching lines")
    ap.add_argument("--out", help="write to this file instead of stdout")
    args = ap.parse_args()
    pattern = re.compile(args.grep) if args.grep else None
    rows = [ln for ln in lines(args.file) if not pattern or pattern.search(ln.text)]
    sink = open(args.out, "w", encoding="utf-8") if args.out else sys.stdout
    for ln in rows:
        sink.write(f"{ln.part}:{ln.index}\t{ln.text!r}\n")
    if args.out:
        sink.close()
        print(f"wrote {len(rows)} lines to {args.out}")


if __name__ == "__main__":
    main()
