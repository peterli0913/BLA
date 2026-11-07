#!/usr/bin/env python3
"""Merge cursor and Qianwen review comments onto one PPQ protocol copy.

Body text and existing reviewer comments stay as they are. AI小鹰 comments from
the earlier review are renamed to cursor小鹰. The 40 Qianwen comments are
copied onto the same anchors and renamed to 千问小鹰.
"""

import hashlib
import shutil
import zipfile
from pathlib import Path

from lxml import etree

BASE = Path("/workspace/2. L2 Step1 PVP-1400200301.01_-10-24-2026-AI小鹰批注.docx")
QWEN = Path("/workspace/PVP-L2-Step1_AI小鹰批注(千问)_20261027.docx")
DST = Path("/workspace/2. L2 Step1 PVP-1400200301.01_cursor小鹰+千问小鹰批注_20261027.docx")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
XML = "http://www.w3.org/XML/1998/namespace"


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def in_del(el):
    parent = el
    while parent is not None:
        if parent.tag == qn(W, "del"):
            return True
        parent = parent.getparent()
    return False


def accepted_text(root):
    parts = []
    for el in root.iter(qn(W, "t")):
        if el.text and not in_del(el):
            parts.append(el.text)
    return "".join(parts)


def text_nodes(root):
    nodes = []
    offset = 0
    for el in root.iter(qn(W, "t")):
        if not el.text or in_del(el):
            continue
        nodes.append([el, offset, offset + len(el.text)])
        offset += len(el.text)
    return nodes, offset


def marker_offsets(root, author_filter):
    comments = etree.fromstring(zipfile.ZipFile(QWEN).read("word/comments.xml"))
    wanted = {
        c.get(qn(W, "id"))
        for c in comments.findall(qn(W, "comment"))
        if c.get(qn(W, "author")) == author_filter
    }
    positions = {}
    offset = 0
    for el in root.iter():
        if el.tag == qn(W, "commentRangeStart") and el.get(qn(W, "id")) in wanted:
            positions.setdefault(el.get(qn(W, "id")), {})["start"] = offset
        elif el.tag == qn(W, "commentRangeEnd") and el.get(qn(W, "id")) in wanted:
            positions.setdefault(el.get(qn(W, "id")), {})["end"] = offset
        elif el.tag == qn(W, "t") and el.text and not in_del(el):
            offset += len(el.text)
    missing = [cid for cid, pos in positions.items() if "start" not in pos or "end" not in pos]
    if missing or len(positions) != len(wanted):
        raise SystemExit(f"Incomplete Qianwen anchors: {missing} count {len(positions)}/{len(wanted)}")
    return positions, offset


def needs_preserve(text, source):
    if source is not None and source.get(qn(XML, "space")) == "preserve":
        return True
    return text.startswith(" ") or text.endswith(" ")


def split_text_node(t_el, index):
    text = t_el.text
    left, right = text[:index], text[index:]
    if not left or not right:
        raise SystemExit("Refusing to split at an edge")
    run = t_el.getparent()
    if run.tag != qn(W, "r"):
        raise SystemExit(f"Text parent is {run.tag}, not a run")
    t_el.text = left
    if needs_preserve(left, t_el):
        t_el.set(qn(XML, "space"), "preserve")
    new_run = etree.Element(qn(W, "r"))
    rpr = run.find(qn(W, "rPr"))
    if rpr is not None:
        new_run.append(etree.fromstring(etree.tostring(rpr)))
    new_t = etree.SubElement(new_run, qn(W, "t"))
    new_t.text = right
    if needs_preserve(right, t_el):
        new_t.set(qn(XML, "space"), "preserve")
    following = list(t_el.itersiblings())
    for sib in following:
        new_run.append(sib)
    run.addnext(new_run)
    return new_t


def split_all(root, points):
    nodes, total = text_nodes(root)
    for el, start, end in reversed(nodes):
        cuts = sorted(p for p in points if start < p < end)
        current = el
        current_start = start
        for cut in reversed(cuts):
            current = split_text_node(current, cut - current_start)
            current_start = cut
    nodes, total2 = text_nodes(root)
    if total2 != total:
        raise SystemExit("Split changed accepted text length")
    return nodes


def insert_ranges(root, ranges):
    nodes, _ = text_nodes(root)
    by_start = {start: (el, end) for el, start, end in nodes}
    by_end = {end: el for el, start, end in nodes}
    starts = {}
    ends = {}
    for cid, start, end in ranges:
        starts.setdefault(start, []).append(cid)
        ends.setdefault(end, []).append(cid)
    for offset, cids in ends.items():
        if offset not in by_end:
            raise SystemExit(f"No text node ends at {offset}")
        run = by_end[offset].getparent()
        for cid in reversed(cids):
            end = etree.Element(qn(W, "commentRangeEnd"))
            end.set(qn(W, "id"), cid)
            ref_run = etree.Element(qn(W, "r"))
            rpr = etree.SubElement(ref_run, qn(W, "rPr"))
            style = etree.SubElement(rpr, qn(W, "rStyle"))
            style.set(qn(W, "val"), "CommentReference")
            ref = etree.SubElement(ref_run, qn(W, "commentReference"))
            ref.set(qn(W, "id"), cid)
            run.addnext(ref_run)
            run.addnext(end)
    for offset, cids in starts.items():
        if offset not in by_start:
            raise SystemExit(f"No text node starts at {offset}")
        run = by_start[offset][0].getparent()
        for cid in reversed(cids):
            start = etree.Element(qn(W, "commentRangeStart"))
            start.set(qn(W, "id"), cid)
            run.addprevious(start)


def existing_para_ids(*roots):
    found = set()
    for root in roots:
        for el in root.iter():
            for attr, val in el.attrib.items():
                if attr.endswith("paraId"):
                    found.add(val)
    return found


def append_qwen_comments(comments, extended, qwen_comments, id_map, used_para_ids):
    for old in qwen_comments:
        old_id = old.get(qn(W, "id"))
        new_id = id_map[old_id]
        text = "".join(t.text or "" for t in old.iter(qn(W, "t")))
        para_id = f"{0xB2000000 + int(new_id):08X}"
        if para_id in used_para_ids:
            raise SystemExit(f"paraId collision {para_id}")
        used_para_ids.add(para_id)
        comment = etree.Element(qn(W, "comment"))
        comment.set(qn(W, "id"), new_id)
        comment.set(qn(W, "author"), "千问小鹰")
        comment.set(qn(W, "date"), old.get(qn(W, "date")))
        comment.set(qn(W, "initials"), "QW")
        paragraph = etree.SubElement(comment, qn(W, "p"))
        paragraph.set(qn(W14, "paraId"), para_id)
        paragraph.set(qn(W14, "textId"), "77777777")
        ppr = etree.SubElement(paragraph, qn(W, "pPr"))
        style = etree.SubElement(ppr, qn(W, "pStyle"))
        style.set(qn(W, "val"), "CommentText2")
        ref_run = etree.SubElement(paragraph, qn(W, "r"))
        ref_rpr = etree.SubElement(ref_run, qn(W, "rPr"))
        ref_style = etree.SubElement(ref_rpr, qn(W, "rStyle"))
        ref_style.set(qn(W, "val"), "CommentReference")
        etree.SubElement(ref_run, qn(W, "annotationRef"))
        run = etree.SubElement(paragraph, qn(W, "r"))
        rpr = etree.SubElement(run, qn(W, "rPr"))
        fonts = etree.SubElement(rpr, qn(W, "rFonts"))
        fonts.set(qn(W, "ascii"), "Calibri")
        fonts.set(qn(W, "hAnsi"), "Calibri")
        fonts.set(qn(W, "eastAsia"), "微软雅黑")
        color = etree.SubElement(rpr, qn(W, "color"))
        color.set(qn(W, "val"), "000000")
        sz = etree.SubElement(rpr, qn(W, "sz"))
        sz.set(qn(W, "val"), "21")
        szcs = etree.SubElement(rpr, qn(W, "szCs"))
        szcs.set(qn(W, "val"), "21")
        lang = etree.SubElement(rpr, qn(W, "lang"))
        lang.set(qn(W, "eastAsia"), "zh-CN")
        t = etree.SubElement(run, qn(W, "t"))
        t.set(qn(XML, "space"), "preserve")
        t.text = text
        comments.append(comment)
        ex = etree.Element(qn(W15, "commentEx"))
        ex.set(qn(W15, "paraId"), para_id)
        ex.set(qn(W15, "done"), "0")
        extended.append(ex)


def rename_author(comments, people):
    for comment in comments.findall(qn(W, "comment")):
        if comment.get(qn(W, "author")) == "AI小鹰":
            comment.set(qn(W, "author"), "cursor小鹰")
            comment.set(qn(W, "initials"), "CS")
    for person in people.findall(qn(W15, "person")):
        if person.get(qn(W15, "author")) == "AI小鹰":
            person.set(qn(W15, "author"), "cursor小鹰")
    if not any(p.get(qn(W15, "author")) == "千问小鹰" for p in people.findall(qn(W15, "person"))):
        person = etree.Element(qn(W15, "person"))
        person.set(qn(W15, "author"), "千问小鹰")
        info = etree.SubElement(person, qn(W15, "presenceInfo"))
        info.set(qn(W15, "providerId"), "None")
        info.set(qn(W15, "userId"), "千问小鹰")
        people.append(person)


def raw_hash(root):
    chunks = []
    for node in root.iter():
        if node.tag in (qn(W, "t"), qn(W, "delText")) and node.text:
            chunks.append(node.text)
    return hashlib.sha256("".join(chunks).encode()).hexdigest()


def main():
    qwen_doc = etree.fromstring(zipfile.ZipFile(QWEN).read("word/document.xml"))
    qwen_comments_root = etree.fromstring(zipfile.ZipFile(QWEN).read("word/comments.xml"))
    qwen_ai = [
        c
        for c in qwen_comments_root.findall(qn(W, "comment"))
        if c.get(qn(W, "author")) == "AI小鹰"
    ]
    positions, qwen_len = marker_offsets(qwen_doc, "AI小鹰")

    if DST.exists():
        DST.unlink()
    shutil.copy(BASE, DST)
    with zipfile.ZipFile(DST, "r") as zf:
        files = {name: zf.read(name) for name in zf.namelist()}
    doc = etree.fromstring(files["word/document.xml"])
    comments = etree.fromstring(files["word/comments.xml"])
    extended = etree.fromstring(files["word/commentsExtended.xml"])
    people = etree.fromstring(files["word/people.xml"])
    before = raw_hash(doc)
    before_accepted = accepted_text(doc)
    if before_accepted != accepted_text(qwen_doc) or len(before_accepted) != qwen_len:
        raise SystemExit("Accepted text differs; cannot map comment anchors")
    old_comments = {
        c.get(qn(W, "id")): etree.tostring(c)
        for c in comments.findall(qn(W, "comment"))
        if c.get(qn(W, "author")) != "AI小鹰"
    }
    cursor_texts = [
        "".join(t.text or "" for t in c.iter(qn(W, "t")))
        for c in comments.findall(qn(W, "comment"))
        if c.get(qn(W, "author")) == "AI小鹰"
    ]

    points = set()
    ranges = []
    id_map = {}
    for index, comment in enumerate(qwen_ai):
        old_id = comment.get(qn(W, "id"))
        new_id = str(40001 + index)
        id_map[old_id] = new_id
        start = positions[old_id]["start"]
        end = positions[old_id]["end"]
        if end <= start:
            raise SystemExit(f"Empty range for {old_id}")
        points.add(start)
        points.add(end)
        ranges.append((new_id, start, end))

    split_all(doc, points)
    insert_ranges(doc, ranges)
    rename_author(comments, people)
    append_qwen_comments(comments, extended, qwen_ai, id_map, existing_para_ids(doc, comments, extended))

    if raw_hash(doc) != before or accepted_text(doc) != before_accepted:
        raise SystemExit("Body text changed")
    for cid, blob in old_comments.items():
        found = [c for c in comments.findall(qn(W, "comment")) if c.get(qn(W, "id")) == cid]
        if len(found) != 1 or etree.tostring(found[0]) != blob:
            raise SystemExit(f"Existing comment {cid} changed")
    cursor = [c for c in comments.findall(qn(W, "comment")) if c.get(qn(W, "author")) == "cursor小鹰"]
    qwen = [c for c in comments.findall(qn(W, "comment")) if c.get(qn(W, "author")) == "千问小鹰"]
    if ["".join(t.text or "" for t in c.iter(qn(W, "t"))) for c in cursor] != cursor_texts:
        raise SystemExit("cursor小鹰 comment text changed")
    qwen_texts = ["".join(t.text or "" for t in c.iter(qn(W, "t"))) for c in qwen_ai]
    if ["".join(t.text or "" for t in c.iter(qn(W, "t"))) for c in qwen] != qwen_texts:
        raise SystemExit("千问小鹰 comment text changed")
    ids = {c.get(qn(W, "id")) for c in cursor + qwen}
    starts = {el.get(qn(W, "id")) for el in doc.iter(qn(W, "commentRangeStart"))}
    ends = {el.get(qn(W, "id")) for el in doc.iter(qn(W, "commentRangeEnd"))}
    refs = {el.get(qn(W, "id")) for el in doc.iter(qn(W, "commentReference"))}
    if not ids <= starts or not ids <= ends or not ids <= refs:
        raise SystemExit("Missing comment markers")
    if any(c.get(qn(W, "author")) == "AI小鹰" for c in comments.findall(qn(W, "comment"))):
        raise SystemExit("AI小鹰 author still present")

    files["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/comments.xml"] = etree.tostring(comments, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/commentsExtended.xml"] = etree.tostring(
        extended, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    files["word/people.xml"] = etree.tostring(people, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = DST.with_suffix(".tmp.docx")
    with zipfile.ZipFile(BASE, "r") as src, zipfile.ZipFile(tmp, "w") as out:
        for info in src.infolist():
            out.writestr(info, files[info.filename])
    tmp.replace(DST)
    print(f"wrote {DST.name} cursor={len(cursor)} qwen={len(qwen)}")


if __name__ == "__main__":
    main()
