#!/usr/bin/env python3
"""Add English to the new (blue) Chinese text in the user-revised bilingual Step 1 cleaning SOP.

Same rules as build_cleaning_sop_bilingual.py: "English/" goes in front of the Chinese as its own run
(Times New Roman; 8 pt inside tables), 8 pt table paragraphs get exact 10 pt line spacing, source typos are
fixed in both languages, activation tank / Tri-clamp terminology. Only paragraphs that still have Chinese
without English are touched; everything else in word/document.xml and all other parts stay as they are.

Usage: python3 scripts/build_cleaning_sop_bilingual_blue.py
"""
import re
import zipfile
from pathlib import Path

from lxml import etree

import build_cleaning_sop_bilingual as base
from build_cleaning_sop_bilingual import ADJ_WRENCH, HAN, SS_WRENCH, W, reassemble, wipe, wipe_ends

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Cleaning and disinfection SOP for step 1-0930新_中英对照 (2).docx"
OUT = ROOT / "deliverables" / "Cleaning and disinfection SOP for step 1-0930新_中英对照 (2).docx"

BLUE = {"0070C0", "548DD4"}

# Source issues in the new text, fixed in both languages (in addition to base.CORRECTIONS).
CORRECTIONS = base.CORRECTIONS + [
    ("对拆除的加料管路内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端",
     "对拆除的加料管路内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从加料管路两端"),
    ("此釜用于DMF溶剂预冷使用，回流管路139788-step01-P27", "此釜用于DMF溶剂预冷使用，回流管路139788-step01-P28"),
]

# Second half of a header split across paragraphs ("Equipment name/设" + "备名称").
SKIP = base.SKIP | {"备名称"}

NOT_CLEANED = "cleaning not applicable"


def in_situ_p16(table):
    return (f"This synthesizer valve is cleaned in situ together with drain line 139788-step01-P16; see the online "
            f"cleaning section for 139788-step01-P16 in Table {table} for the cleaning method")


def dmf_precool(rest):
    return f"This reactor is used for DMF solvent pre-cooling; {rest} not in contact with material, {NOT_CLEANED}"


T = {
    "设备清洁记录": "Equipment Cleaning Record",
    "—适用于产品139788 step1固相合成工序的批间清洁": [
        ("适用于", "Applicable to between-batch cleaning of the solid-phase synthesis process of product 139788 step1"),
    ],
    "适用于产品139788 step1固相合成工序生产结束后清洁":
        "Applicable to end-of-production cleaning of the solid-phase synthesis process of product 139788 step1",
    "项目号：                                                                      清洗方法：": [
        ("项目号：", "Project No."), ("清洗方法：", "Cleaning method"),
    ],
    "下口鼓氮气": "Bottom nitrogen bubbling",
    "此合成仪阀门同排液管线139788-step01-P16 进行原位清洗，清洗方式详见表4中139788-step01-P16在线清洗部分": in_situ_p16(4),
    "此合成仪阀门同排液管线139788-step01-P16 进行原位清洗，清洗方式详见表18中139788-step01-P16在线清洗部分": in_situ_p16(18),
    "衬PTFE": "PTFE-lined",
    "气动阀": "Pneumatic valve",
    "下口排液气动阀": "Bottom drain pneumatic valve",
    "下口尾气真空阀门": "Bottom vent vacuum valve",
    "排废管路": "Waste line",
    "316不锈钢": "316 stainless steel",
    "管路为新鲜溶剂充满状态，不涉及清洗": f"The line is kept full of fresh solvent; {NOT_CLEANED}",
    "新鲜溶剂管路": "Fresh solvent line",
    "脱保护管路": "Deprotection line",
    "溶剂管路氮气": "Solvent line nitrogen",
    "316L不锈钢软管": "316L stainless steel hose",
    "316L不锈钢弯头": "316L stainless steel elbow",
    "不锈钢软管": "Stainless steel hose",
    "软管": "Hose",
    "临时软管": "Temporary hose",
    "使用不锈钢扳手将软管两端的快开卡盘松动拆卸，取下管件;":
        f"{SS_WRENCH} loosen and remove the clamps at both ends of the hose, and remove the pipe fitting;",
    "使用洁净擦拭布蘸取甲醇对管径部件进行擦拭清洗；": "Wipe the pipe fitting with a clean wipe dipped in methanol;",
    "使用不锈钢扳手将视镜两端快开卡盘松开，取下玻璃视镜":
        f"{SS_WRENCH} release the clamps at both ends of the sight glass, and remove the glass",
    "使用洁净擦拭布蘸取纯化水后对拆卸后视镜内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the inner surface of the removed sight glass"),
    "将拆解的视镜依次完成复位组装，更换视镜两端连接的垫片": reassemble("sight glass", "gaskets", "sight glass"),
    "将拆解的视镜依次完成复位组装，更换视镜两端连接的快开垫片": reassemble("sight glass", "clamp gaskets", "sight glass"),
    "激活釜视镜": "Activation tank sight glass",
    "使用不锈钢扳手将视镜四个固定螺杆逆时针方向拧松，取下四个螺杆及下方螺母，将视镜解体并将玻璃视镜取出；":
        f"{SS_WRENCH} loosen the four fixing bolts of the sight glass counterclockwise, remove the four bolts and "
        "the nuts underneath, disassemble the sight glass and take out the glass;",
    "使用洁净擦拭布蘸取纯化水后对解体后视镜内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the inner surface of the disassembled sight glass"),
    "激活釜排废管线阀门，不涉及清洗": f"Valve on the activation tank waste line; {NOT_CLEANED}",
    "新鲜溶剂氮气吹扫管路阀门，不涉及清洗": f"Valve on the fresh solvent nitrogen purge line; {NOT_CLEANED}",
    "使用活口扳手将与加料管道连接端的卡盘松开拆卸": f"{ADJ_WRENCH} release and remove the clamp at the feed pipe connection end",
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的加料管路内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从加料管路两端进行吹干，QA验收；":
        wipe_ends("the inner surface of the removed feed line", "feed line"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的弯头，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从弯头两端进行吹干，QA验收":
        wipe_ends("the removed elbow and seal surfaces", "elbow", end=""),
    "异径管": "Reducer",
    "液体加料口": "Liquid charging port",
    "使用活口扳手将与阀门的连接端的法兰上螺杆逆时针拧松取下;":
        f"{ADJ_WRENCH} loosen counterclockwise and remove the bolts on the flange at the valve connection end;",
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后管道进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the removed pipe", "reducer"),
    "馏分管路": "Distillate line",
    "此釜用于DMF溶剂预冷使用，此釜回流管路及冷凝器到接收罐和TJ4S-1213-R23釜管路未接触物料不涉及清洗":
        dmf_precool("its reflux line and the lines from the condenser to the receiver tank and reactor "
                    "TJ4S-1213-R23 are"),
    "此釜用于DMF溶剂预冷使用，蒸馏管路139788-step01-P27未接触物料不涉及清洗":
        dmf_precool("distillation line 139788-step01-P27 is"),
    "此釜用于DMF溶剂预冷使用，回流管路139788-step01-P28未接触物料不涉及清洗":
        dmf_precool("reflux line 139788-step01-P28 is"),
    "接收罐磁板液位计连接阀门": "Receiver tank magnetic level gauge connection valve",
    "接收罐磁板液位计排尽阀门": "Receiver tank magnetic level gauge drain valve",
    "釜上视镜": "Reactor sight glass",
    "釜上口氮气": "Reactor top nitrogen",
    "接收罐压力表": "Receiver tank pressure gauge",
    "接收罐下口": "Receiver tank bottom outlet",
}

# "316L" and "不锈钢软管" typed as two paragraphs in one cell: the English goes in front of "316L".
SPLIT_PREFIX = {("316L", "不锈钢软管"): "316L stainless steel hose"}


def translated(s):
    m = HAN.search(s)
    head = s[:m.start()] if m else s
    return "/" in head and re.search("[A-Za-z]", head) is not None


def continued(s, prev):
    """True when s is the Chinese half of an English/Chinese pair split across two paragraphs."""
    if prev is None or not re.search("[A-Za-z]", prev):
        return False
    if prev.strip().endswith("/") or s.strip().startswith("/"):
        return True
    m = re.search(r"[A-Za-z][^\u4e00-\u9fff]*/([^/]*)$", prev)
    return bool(m) and m.group(1).strip() != "" and not HAN.search(m.group(1))


def is_blue(p):
    for r in p.iter(W + "r"):
        c = r.find(f"{W}rPr/{W}color")
        if c is not None and c.get(W + "val") in BLUE and "".join(t.text or "" for t in r.findall(W + "t")).strip():
            return True
    return False


def norm(s):
    return re.sub(r"\s+", " ", s.strip())


def lookup_table():
    table = {norm(k): v for k, v in base.build_lookup().items()}
    for k, v in T.items():
        v = base.en_terms(v) if isinstance(v, str) else [(a, base.en_terms(e)) for a, e in v]
        assert table.get(norm(k), v) == v, k
        table[norm(k)] = v
    return table


def set_latin_font(p):
    for r in p.iter(W + "r"):
        if not re.search(r"[A-Za-z0-9]", "".join(t.text or "" for t in r.findall(W + "t"))):
            continue
        rpr = r.find(W + "rPr")
        if rpr is None:
            rpr = etree.Element(W + "rPr")
            r.insert(0, rpr)
        fonts = base.child(rpr, "rFonts", base.RPR_ORDER)
        if fonts.get(W + "ascii") in base.SYMBOL_FONTS:
            continue
        for a in ("ascii", "hAnsi", "cs"):
            fonts.set(W + a, base.EN_FONT)


def set_table_spacing(p):
    sizes = set()
    for r in p.iter(W + "r"):
        if "".join(t.text or "" for t in r.findall(W + "t")):
            sz = r.find(W + "rPr/" + W + "sz")
            sizes.add(None if sz is None else sz.get(W + "val"))
    if sizes != {base.TABLE_SZ}:
        return False
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        p.insert(0, ppr)
    sp = base.child(ppr, "spacing", base.PPR_ORDER)
    sp.set(W + "line", base.TABLE_LINE)
    sp.set(W + "lineRule", "exact")
    return True


def main():
    lookup = lookup_table()
    with zipfile.ZipFile(SRC) as zin:
        tree = etree.fromstring(zin.read("word/document.xml"))
        targets, not_blue, missing = [], [], set()
        last = {}
        for p in tree.iter(W + "p"):
            s = base.ptext(p)
            if not s.strip():
                continue
            c = base.container(p)
            prev_p = last.get(c)
            prev = base.ptext(prev_p) if prev_p is not None else None
            last[c] = p
            key = s.strip()
            if not HAN.search(s) or key in SKIP or translated(s) or base.has_inline_english(s) or continued(s, prev):
                continue
            targets.append((p, prev_p))
            if not is_blue(p):
                not_blue.append(key)

        fixed = translated_n = spaced = 0
        for p, prev_p in targets:
            for old, new in CORRECTIONS:
                fixed += base.replace_in_paragraph(p, old, new)
            s = base.ptext(p)
            key = norm(s)
            touched = [p]
            prefix = SPLIT_PREFIX.get((base.ptext(prev_p).strip() if prev_p is not None else None, key))
            if prefix:
                base.translate_paragraph(prev_p, base.ptext(prev_p), prefix)
                touched.append(prev_p)
            elif key in lookup:
                base.translate_paragraph(p, s, lookup[key])
            else:
                missing.add(key)
                continue
            translated_n += 1
            for q in touched:
                set_latin_font(q)
                if base.in_table(q):
                    spaced += set_table_spacing(q)
        assert not missing, sorted(missing)

        xml = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        OUT.parent.mkdir(exist_ok=True)
        with zipfile.ZipFile(OUT, "w") as zout:
            for item in zin.infolist():
                data = xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    print(f"{translated_n} paragraphs translated, {fixed} source corrections, {spaced} set to exact 10 pt")
    print(f"non-blue paragraphs translated: {not_blue}")
    print(f"-> {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
