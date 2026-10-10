#!/usr/bin/env python3
"""Bilingual (English above Chinese) version of Nitrogen Humidifier Operation Manual-1009.docx.

Follows the format of section 1 "Applicable Objects/适用对象": headings and table header labels become
"English/中文" on one line; every other Chinese paragraph gets an English paragraph inserted above it.
Chinese text is untouched. English runs are Times New Roman, 10.5 pt (五号) in body text and 9 pt (小五)
in tables and figure labels. Only word/document.xml is rewritten.

Usage: python3 scripts/build_humidifier_manual_bilingual.py
"""
import copy
import re
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Nitrogen Humidifier Operation Manual-1009.docx"
OUT = ROOT / "deliverables" / "Nitrogen Humidifier Operation Manual-1009_中英对照.docx"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
V = "{urn:schemas-microsoft-com:vml}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
FONT = "Times New Roman"
BODY_SZ, TABLE_SZ = "21", "18"
HEADING_STYLES = {"2", "3", "50"}
CJK = re.compile(r"[\u3400-\u9fff]")

WFI = "water for injection (WFI)"
DOLLY = "Stainless steel transport dolly"
FINISHED = "Before transferring WFI"

# headings and table header labels: rendered as "English/中文"
INLINE = {
    "设备功能介绍": "Equipment Function",
    "设备结构介绍": "Equipment Structure",
    "设备操作流程": "Operating Procedure",
    "使用储液袋和转移桶取用注射用水的操作方法": "Collecting WFI with a Storage Bag and Transfer Drum",
    "转移桶结构": "Transfer Drum Structure",
    "储液袋结构": "Storage Bag Structure",
    "操作前准备": "Preparation",
    "操作前检查": "Pre-operation Checks",
    "取用注射用水": "Collecting WFI",
    "将储液袋连接至氮气加湿器": "Connecting the Storage Bag to the Nitrogen Humidifier",
    "无菌连接接头使用方法": "Using the Aseptic Connector",
    "使用不锈钢水桶取用注射用水的操作方法": "Collecting WFI with a Stainless Steel Water Drum",
    "不锈钢水桶结构": "Stainless Steel Water Drum Structure",
    "序号": "No.",
    "名称": "Name",
    "检查内容": "Check Item",
    "操作描述": "Operation",
}

# everything else: English paragraph above the Chinese one
ABOVE = {
    # ---- equipment description
    "本设备以注射用水作为加湿介质，用于对氮气进行加湿处理。氮气通入设备内与注射用水充分接触，实现增湿，输出湿度稳定的湿氮气。":
        f"This equipment uses {WFI} as the humidification medium to humidify nitrogen. Nitrogen fed into the "
        "equipment is brought into full contact with WFI and is humidified, delivering humidified nitrogen at a "
        "stable humidity.",
    "氮气加湿器主要由鼓泡罐、输液泵、过滤器、呼吸器、调节阀、压力传感器、温度传感器，转移管路、阀门等部分组成。":
        "The nitrogen humidifier mainly consists of a bubbler tank, transfer pump, filters, vent filter, control "
        "valve, pressure sensor, temperature sensor, transfer lines and valves.",
    # ---- transfer drum
    "转移桶由桶盖、桶身、不锈钢转运底座组成。桶身标有刻度，用于提示人员储液袋内注射用水的体积。桶下配有侧底部出液口，用于放置储液袋的放液管路。示意图见图4.1.1，表4.1.1。":
        "The transfer drum consists of a lid, a drum body and a stainless steel transport dolly. The drum body "
        "is graduated to show operators the volume of WFI in the storage bag. An outlet at the lower side of the "
        "drum is used to route the drain line of the storage bag. See Figure 4.1.1 and Table 4.1.1.",
    "图4.1.1 转移桶结构图": "Figure 4.1.1 Transfer Drum Structure",
    "表4.1.1 转移桶部件清单": "Table 4.1.1 Transfer Drum Parts List",
    "桶盖": "Lid",
    "桶身": "Drum body",
    "不锈钢转运底座": DOLLY,
    "侧底部出液口": "Lower side outlet",
    # ---- storage bag
    "储液袋主要由袋身、进液管路、放液管路、取样管路组成。示意图见图4.1.2，表4.1.2。":
        "The storage bag mainly consists of the bag body, fill line, drain line and sampling line. "
        "See Figure 4.1.2 and Table 4.1.2.",
    "图4.1.2 储液袋结构图": "Figure 4.1.2 Storage Bag Structure",
    "表4.1.2 储液袋部件清单": "Table 4.1.2 Storage Bag Parts List",
    "袋身": "Bag body",
    "进液管路": "Fill line",
    "放液管路": "Drain line",
    "取样管路": "Sampling line",
    "止液夹": "Pinch clamp",
    "每条管路上都配有一个止液夹，止液夹用于储液袋软管的通断与流量调节。向后拨卡扣松开软管，管路导通；滑动夹体可微调流速；推至前端压紧软管，即可关断管路，实现流体隔离。示意图见图4.1.3。":
        "Each line is fitted with a pinch clamp, which opens and closes the storage bag tubing and regulates "
        "the flow. Push the latch back to release the tubing and open the line; slide the clamp body to "
        "fine-tune the flow rate; push it to the front end to compress the tubing and close the line, "
        "isolating the fluid. See Figure 4.1.3.",
    "图4.1.3 止液夹结构图": "Figure 4.1.3 Pinch Clamp Structure",
    "进液管路配有快开密封，主要由堵头、密封O型圈、快开卡子组成，使用前和接水结束后需要确保该接口密封完好。示意图及结构介绍见图4.1.4，表4.1.3。":
        "The fill line has a sanitary quick-release closure, mainly consisting of an end cap, a sealing O-ring "
        "and a Tri-clamp. Make sure this connection is properly sealed before use and after WFI collection. "
        "See Figure 4.1.4 and Table 4.1.3 for the diagram and components.",
    "图4.1.4 进液管路接头结构图": "Figure 4.1.4 Fill Line Connector Structure",
    "表4.1.3 进液管路接头部件清单": "Table 4.1.3 Fill Line Connector Parts List",
    "密封堵头": "Sealing end cap",
    "密封O型圈": "Sealing O-ring",
    "快开卡子": "Tri-clamp",
    "放液管路配有无菌连接接口，用于一次性储液袋管路之间的密闭无菌对接，实现注射用水等液体的无菌转移，该接头为一次性使用，连接后不可重复拆装。":
        "The drain line is fitted with an aseptic connector for closed, aseptic connection between single-use "
        "storage bag lines, enabling aseptic transfer of WFI and other liquids. The connector is single-use; "
        "once connected, it cannot be disconnected and reconnected.",
    "该接头主要由保护拉片、无菌隔离膜、外壳卡扣锁止机构、软管插接端、密封O型圈组成。示意图见图4.1.5，表4.1.4。":
        "The connector mainly consists of protective pull tabs, a sterile barrier membrane, a snap-lock "
        "housing, a hose barb and a sealing O-ring. See Figure 4.1.5 and Table 4.1.4.",
    "图4.1.5 无菌连接接头结构图": "Figure 4.1.5 Aseptic Connector Structure",
    "表4.1.4 无菌连接接头部件清单": "Table 4.1.4 Aseptic Connector Parts List",
    "保护拉片": "Protective pull tab",
    "无菌隔离膜": "Sterile barrier membrane",
    "外壳卡扣锁止机构": "Snap-lock housing",
    "软管插接端": "Hose barb",
    "密封O型圈（内部）": "Sealing O-ring (internal)",
    # ---- preparation and checks (storage bag)
    "使用储液袋和转移桶进行注射用水的转移前，需确认如下器具已准备完全：转移桶、储液袋、注射用水专用软管、快开卡子、垫片等。":
        f"{FINISHED} with the storage bag and transfer drum, confirm that the following items are ready: "
        "transfer drum, storage bag, dedicated WFI hose, Tri-clamps, gaskets, etc.",
    "检查注射用水使用点状态标识处于“合格准用”状态，注射用水温度传感器在校验有效期内":
        "Check that the status label at the WFI point of use shows \u201cQualified \u2013 Released for Use\u201d "
        "and that the WFI temperature sensor is within its calibration period",
    "检查注射用水专用软管在有效期内，软管完好洁净，无破损、无外来异物等异常情况":
        "Check that the dedicated WFI hose is within its expiry date and is intact and clean, with no damage, "
        "foreign matter or other abnormalities",
    "检查储液袋袋身、管路等所有配件完好洁净，无破损、无外来异物等异常情况":
        "Check that the storage bag body, lines and all other components are intact and clean, with no damage, "
        "foreign matter or other abnormalities",
    "检查储液袋各管路上的止液夹均处于压紧隔断状态":
        "Check that the pinch clamps on all storage bag lines are fully closed, compressing the tubing",
    "检查储液袋放液管路无菌连接器保护盖完好，无破损":
        "Check that the protective cover of the aseptic connector on the storage bag drain line is intact and "
        "undamaged",
    # ---- collecting WFI
    "将注射用水专用软管一端连接到注射用水使用点上":
        "Connect one end of the dedicated WFI hose to the WFI point of use",
    "按下注射用水面板上的“启动”按钮，等待注射用水温度降至常温后自动出水。":
        "Press the \u201cStart\u201d button on the WFI panel and wait for the WFI to cool to ambient temperature; "
        "water is then dispensed automatically.",
    "注射用水出水后，持续放水至少1分钟，充分冲洗注射用水专用软管内壁（如需变径连接，则需要同变径一起冲洗）":
        "Once WFI starts flowing, let it run for at least 1 minute to thoroughly flush the inner wall of the "
        "dedicated WFI hose (if a reducer is needed for the connection, flush it together with the hose)",
    "将储液袋平整放入转移桶内，将储液袋放液管路从侧底部出液口拿出（管路和接头不可接触地面）":
        "Lay the storage bag flat inside the transfer drum and route its drain line out through the lower side "
        "outlet (the line and connector must not touch the floor)",
    "冲洗结束后，将注射用水专用软管另一端连接至储液袋进液管接口":
        "After flushing, connect the other end of the dedicated WFI hose to the fill line connection of the "
        "storage bag",
    "打开储液袋进液管路止液阀，打开注射用水使用点阀门，从注射用水使用点取用注射用水":
        "Open the pinch clamp on the storage bag fill line, then open the WFI point-of-use valve to collect WFI",
    "达到所需体积后依次关闭注射用水使用点阀门、储液袋进液管路止液阀，断开注射用水专用软管与储液袋进液管，将储液袋进液管路使用堵头密封。断开注射用水专用软管与注射用水使用点连接，排净干燥后悬挂存放":
        "When the required volume is reached, close the WFI point-of-use valve and then the pinch clamp on the "
        "storage bag fill line. Disconnect the dedicated WFI hose from the fill line and seal the fill line with "
        "the end cap. Disconnect the dedicated WFI hose from the point of use, drain and dry it, then hang it for "
        "storage",
    "将进液管路平整放入桶内，盖上转移桶盖，等待使用":
        "Lay the fill line flat inside the drum, close the transfer drum lid and hold for use",
    # ---- aseptic connector
    "掰开两侧连接器半体上的蓝色保护拉片盖，向下翻折":
        "Open the blue protective pull-tab covers on both connector halves and fold them down",
    "对齐连接器两半体，保持拉片盖朝下。按照箭头指示，将两个半体对推合拢":
        "Align the two connector halves with the pull-tab covers facing down. Push the halves together in the "
        "direction of the arrows",
    "分别挤压连接器两侧，直至听到卡扣“咔哒”声，确认形成牢固连接。目视检查两侧端面齐平、相互平行":
        "Squeeze both sides of the connector until the latch clicks, confirming a secure connection. Visually "
        "check that both end faces are flush and parallel",
    "按压蓝色拉片，将拉片盖合上": "Press the blue pull tabs to close the pull-tab covers",
    "拉动蓝色拉片，从连接器上撕下无菌隔离膜":
        "Pull the blue pull tab to peel the sterile barrier membrane out of the connector",
    "连接完毕，可以进行液体输送": "The connection is complete and liquid transfer can begin",
    # ---- stainless steel water drum
    "不锈钢水桶由桶身、不锈钢转运底座组成。桶上设计有一个进水接口、一个气体接口、一个视镜口，桶下设计一个放水接口，桶上阀门均为隔膜阀。示意图见图4.2.1，表4.2.1。":
        "The stainless steel water drum consists of a drum body and a stainless steel transport dolly. The top "
        "of the drum has a water inlet, a gas port and a sight glass port, and the bottom has a water outlet. "
        "All valves on the drum are diaphragm valves. See Figure 4.2.1 and Table 4.2.1.",
    "图4.2.1 不锈钢水桶结构图": "Figure 4.2.1 Stainless Steel Water Drum Structure",
    "表4.2.1 不锈钢水桶部件清单": "Table 4.2.1 Stainless Steel Water Drum Parts List",
    "进水口": "Water inlet",
    "气体接口": "Gas port",
    "视镜接口": "Sight glass port",
    "放水口": "Water outlet",
    "隔膜阀": "Diaphragm valve",
    "使用不锈钢水桶进行注射用水的转移前，需确认如下器具已准备完全：注射用水专用不锈钢水桶、注射用水专用软管、快开卡子、垫片、压力表、过滤器等。":
        f"{FINISHED} with the stainless steel water drum, confirm that the following items are ready: dedicated "
        "stainless steel WFI drum, dedicated WFI hose, Tri-clamps, gaskets, pressure gauge, filter, etc.",
    "将过滤器、压力表安装在不锈钢水桶上，安装示意图见图4.2.2，表4.2.2。":
        "Install the filter and pressure gauge on the stainless steel water drum. See Figure 4.2.2 and "
        "Table 4.2.2 for the installation diagram.",
    "图4.2.2 安装示意图": "Figure 4.2.2 Installation Diagram",
    "表4.2.2 安装后部件清单": "Table 4.2.2 Parts List after Installation",
    "压力表": "Pressure gauge",
    "过滤器": "Filter",
}

# Chinese labels inside figure text boxes: English line added above, box grown upward by one line
FIGURE_LABELS = {
    "注射用水使用点状态标识": "WFI point-of-use status label",
    "温度传感器校验标识": "Temp. sensor calibration label",
}
LINE_EMU = 114300  # 9 pt exact line spacing used in these text boxes


def in_textbox(el):
    return any(a.tag == W + "txbxContent" for a in el.iterancestors())


def own_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t") if not in_textbox(t))


# w:rPr children that must follow w:sz / w:szCs (schema order)
AFTER_SZ = {W + t for t in ("highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em",
                            "lang", "eastAsianLayout", "specVanish", "oMath")}


def en_rpr(base, size):
    rpr = copy.deepcopy(base) if base is not None else etree.Element(W + "rPr")
    for tag in ("rFonts", "sz", "szCs", "lang"):
        for el in rpr.findall(W + tag):
            rpr.remove(el)
    fonts = etree.Element(W + "rFonts")
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(W + k, FONT)
    rstyle = rpr.find(W + "rStyle")
    if rstyle is None:
        rpr.insert(0, fonts)
    else:
        rstyle.addnext(fonts)
    if size:
        anchor = next((c for c in rpr if c.tag in AFTER_SZ or not c.tag.startswith(W)), None)
        for tag in ("sz", "szCs"):
            el = etree.Element(W + tag)
            el.set(W + "val", size)
            if anchor is None:
                rpr.append(el)
            else:
                anchor.addprevious(el)
    return rpr


def make_run(text, base_rpr, size):
    r = etree.Element(W + "r")
    r.append(en_rpr(base_rpr, size))
    t = etree.SubElement(r, W + "t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r


def first_text_rpr(p):
    for r in p.findall(W + "r"):
        if r.find(W + "t") is not None:
            return r.find(W + "rPr")
    return None


def insert_inline(p, english, size):
    first = next(r for r in p.findall(W + "r") if r.find(W + "t") is not None)
    first.addprevious(make_run(english + "/", first.find(W + "rPr"), size))


def insert_above(p, english, size, keep_next=False):
    new = etree.Element(W + "p")
    ppr = p.find(W + "pPr")
    new_ppr = copy.deepcopy(ppr) if ppr is not None else etree.Element(W + "pPr")
    mark = new_ppr.find(W + "rPr")
    if mark is not None:
        new_ppr.replace(mark, en_rpr(mark, size))
    if keep_next and new_ppr.find(W + "keepNext") is None:
        style = new_ppr.find(W + "pStyle")
        keep = etree.Element(W + "keepNext")
        if style is None:
            new_ppr.insert(0, keep)
        else:
            style.addnext(keep)
    if len(new_ppr):
        new.append(new_ppr)
    new.append(make_run(english, first_text_rpr(p), size))
    p.addprevious(new)


def grow_label_box(txbx):
    """Move the text box up by one line and make it one line taller, so the Chinese label stays in place."""
    for a in txbx.iterancestors():
        if a.tag == WPS + "wsp":
            xfrm = a.find(f"{WPS}spPr/{A}xfrm")
            off, ext = xfrm.find(A + "off"), xfrm.find(A + "ext")
            off.set("y", str(int(off.get("y")) - LINE_EMU))
            ext.set("cy", str(int(ext.get("cy")) + LINE_EMU))
            a.find(WPS + "bodyPr").set("rIns", "0")
            return
        if a.tag == V + "shape":
            style = dict(kv.split(":", 1) for kv in a.get("style").strip(";").split(";"))
            style["top"] = str(int(style["top"]) - LINE_EMU)
            style["height"] = str(int(style["height"]) + LINE_EMU)
            a.set("style", ";".join(f"{k}:{v}" for k, v in style.items()) + ";")
            a.find(V + "textbox").set("inset", "7.2pt,3.6pt,0pt,3.6pt")
            return
    raise ValueError("text box without shape")


def main():
    zin = zipfile.ZipFile(SRC)
    root = etree.fromstring(zin.read("word/document.xml"))
    body = root.find(W + "body")
    first_table = body.find(W + "tbl")
    missing, done = [], 0

    for txbx in list(root.iter(W + "txbxContent")):
        text = "".join(t.text or "" for t in txbx.iter(W + "t"))
        if text in FIGURE_LABELS:
            p = txbx.find(W + "p")
            insert_above(p, FIGURE_LABELS[text], TABLE_SZ)
            grow_label_box(txbx)
            done += 1

    for p in list(body.iter(W + "p")):
        if in_textbox(p):
            continue
        text = own_text(p).strip()
        if not CJK.search(text):
            continue
        if "/" in text or any(a is first_table for a in p.iterancestors()):
            continue  # section 1 is already bilingual
        in_table = any(a.tag == W + "tbl" for a in p.iterancestors())
        size = TABLE_SZ if in_table else BODY_SZ
        ppr = p.find(W + "pPr")
        style = ppr.find(W + "pStyle").get(W + "val") if ppr is not None and ppr.find(W + "pStyle") is not None \
            else None
        if text in INLINE and (style in HEADING_STYLES or in_table):
            insert_inline(p, INLINE[text], size if in_table else None)
        elif text in ABOVE:
            insert_above(p, ABOVE[text], size, keep_next=not in_table)
        else:
            missing.append(text)
            continue
        done += 1
    assert not missing, "\n".join(missing)

    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, xml if item.filename == "word/document.xml" else zin.read(item.filename))
    print(OUT.relative_to(ROOT), done, "paragraphs")


if __name__ == "__main__":
    main()
