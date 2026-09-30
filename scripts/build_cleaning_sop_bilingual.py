#!/usr/bin/env python3
"""Bilingual (English/Chinese) version of the Step 1 cleaning and disinfection SOP.

Every Chinese paragraph that has no English yet gets "English/" inserted in front of the
Chinese text, inside the existing run so fonts and sizes follow the original. Only
word/document.xml is rewritten; headers, footers, watermark and all other parts are copied as-is.

Usage: python3 scripts/build_cleaning_sop_bilingual.py
"""
import re
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Cleaning and disinfection SOP for step 1-0930新.docx"
OUT = ROOT / "deliverables" / "Cleaning and disinfection SOP for step 1-0930新_中英对照.docx"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
HAN = re.compile(r"[\u4e00-\u9fff]")
SEP = "/"

# Paragraphs left untouched: English already sits in the previous paragraph, or the source text is garbled.
SKIP = {
    "质量控制部，生产部，质量保证部",
    "冷凝器到接收罐和TJ4S-1213-R23釜管路）",
    "Connection loc原位清洁tion/",
}


def wipe(obj, blow="blow dry with compressed air", qa=True, moist=False, end=""):
    cloth = "a clean wipe moistened with purified water" if moist else "a clean wipe dipped in purified water"
    s = f"Wipe {obj} with {cloth} until no material residue is visible, {blow}"
    return s + (", QA acceptance" if qa else "") + end


def wipe_ends(obj, ends, end=";"):
    return wipe(obj, blow=f"blow dry with compressed air alternately from both ends of the {ends}", moist=True, end=end)


def reinstall(part, what="gaskets at both ends"):
    return f"Reinstall the {part} and replace the {what}"


def reassemble(part, gasket, where):
    return f"Reassemble the disassembled {part} in sequence and replace the {gasket} at both ends of the {where}"


def see_online(kind, line, table):
    return (f"{kind} is in line 139788-step01-{line}; see the online cleaning section of main equipment "
            f"cleaning in Table {table} for the cleaning method")


def this_line(desc, line, table):
    return (f"For this {desc}139788-step01-{line}, see the online cleaning section of main equipment "
            f"cleaning in Table {table} for the cleaning method")


AFTER_DONE = "after completion, attach it to the back of this record and sign across the seam"
SS_WRENCH = "Use a stainless steel wrench to"
ADJ_WRENCH = "Use an adjustable wrench to"
FOUR_NUTS = f"{SS_WRENCH} loosen the nuts of the four fixing bolts counterclockwise, remove the four bolts,"
P01_LINE = "material transfer line 139788-step01-P01 from mixing tank TJ4S-1213-AT08 to solid-phase synthesizer TJ4S-1213-Sy12"
P17_LINE = ("material transfer line 139788-step01-P17 from activation reactor TJ4S-1213-Sy12-R01 "
            "to solid-phase synthesizer TJ4S-1213-Sy12")
MAIN_BODY_18 = "Cleaning of solid-phase synthesizer and activation reactor main body"
VALVES_SYN = "Cleaning of valves and components installed on solid-phase synthesizer"
LINES_SYN = "Cleaning of synthesizer's own lines"
VALVES_ACT = "Cleaning of valves and components installed on activation reactor"
LINES_ACT = "Cleaning of activation reactor's own lines"
TANK_BODY = "Cleaning of mixing tank main body"
TANK_ACC = "Cleaning of mixing tank accessories"
TANK_LINES = "Cleaning of mixing tank lines"
DISASSEMBLE_TAIL = (
    "see the table below for disassembly cleaning requirements. Disassemble the {eq}'s nitrogen/vent lines and "
    "valves, rupture disc, solid charging valve, other material charging valves and sight glass down to the "
    "smallest parts, wipe with purified water until no material residue remains, and blow dry with compressed air."
)
REINSTALL_TANK = (
    "After acceptance of the equipment vessel and accessories, install the sight glass, other material charging "
    "valves, solid charging valve, rupture disc, and nitrogen/vent lines and valves in reverse order; after "
    "installation, blow dry the mixing tank and transfer lines with nitrogen"
)

# Chinese paragraph text (stripped) -> English. A plain string is inserted before the paragraph text;
# a list of (anchor, english) inserts each English part right before its Chinese anchor.
T = {
    "1．自动清洗": [("自动清洗", "Automatic Cleaning")],
    "2．手动清洗": [("手动清洗", "Manual Cleaning")],
    "304不锈钢": "304 stainless steel",
    "316L不锈钢": "316L stainless steel",
    "CPo139788-   -01-W-    ，完成后粘贴于记录背面骑缝签字": [("完成后", AFTER_DONE)],
    "DMF称重参见称重记录表，编号为：               ，": "For DMF weighing, refer to the weighing record No.",
    "DMF称重参见称重记录表，编号为：               ，完成后粘贴于记录背面骑缝签字": [
        ("DMF称重", "For DMF weighing, refer to the weighing record No."), ("完成后", AFTER_DONE)],
    "DMF进料管路": "DMF feed line",
    "IPAC进料管路": "IPAC feed line",
    "MTBE进料管路": "MTBE feed line",
    "QA放行人及日期：": "QA released by&Date",
    "QC人员按照标准取样方式执行擦拭取样": "QC personnel shall perform swab sampling following the standard sampling procedure",
    "Step 1完整的单体设备P&ID图(用于显示此步骤所有设备上连接的每一个管路，每一个阀门，每一个配件等细节)，见如下表2：":
        "The complete P&ID of each individual Step 1 equipment (showing details of every pipeline, valve and "
        "fitting connected to all equipment in this step) is listed in Table 2 below:",
    "Step 1所有相关的设备的清洁策略整体计划如下表3：":
        "The overall cleaning strategy plan for all Step 1 related equipment is shown in Table 3 below:",
    "Step 1每个设备的名称，编号，用途，材质见如下表1。":
        "The name, number, purpose and material of each Step 1 equipment are shown in Table 1 below.",
    "TJ4S-1213-AT08到TJ4S-1213-Sy12转移管路": "Transfer line from TJ4S-1213-AT08 to TJ4S-1213-Sy12",
    "☐是      ☐否": [("是", "Yes"), ("否", "No")],
    "三通": "Tee",
    "上展阀": "Upward-opening valve",
    "上音叉液位计": "Upper tuning-fork level switch",
    "下口氮气管路": "Bottom outlet nitrogen line",
    "下口第一个阀门": "First valve at bottom outlet",
    "下口转移物料管路": "Bottom outlet material transfer line",
    "下口转移管路": "Bottom outlet transfer line",
    "不接触物料，不涉及清洁": "No product contact; cleaning not applicable",
    "乙腈称重参见称重记录表，编号为：": "For acetonitrile weighing, refer to the weighing record No.",
    "乙腈称重参见称重记录表，编号为：CPo139788-   -01-W-    ，完成后粘贴于记录背面骑缝签字": [
        ("乙腈称重", "For acetonitrile weighing, refer to the weighing record No."), ("完成后", AFTER_DONE)],
    "乳胶": "Latex",
    "仅涉及到溶剂，不涉及化学残留取样": "Solvent only; chemical residue sampling not applicable",
    "从TJ4S-1213-R23-V75到TJ4S-1213-R23氮气吹至无溶剂残留":
        "From TJ4S-1213-R23-V75 to TJ4S-1213-R23, purge with nitrogen until no solvent remains",
    "从TJ4S-1213-R23-V75到TJ4S-1213-Sy07氮气吹至无溶剂残留":
        "From TJ4S-1213-R23-V75 to TJ4S-1213-Sy07, purge with nitrogen until no solvent remains",
    "从TJ4S-1213-R23-V78到TJ4S-1213-R23氮气吹至无溶剂残留氮气吹至无溶剂残留":
        "From TJ4S-1213-R23-V78 to TJ4S-1213-R23, purge with nitrogen until no solvent remains",
    "使用DMF溶剂按照取样点评估示意图QM04053依次擦拭搅拌桨、内表面、侧出料口、取样手套箱":
        "Using DMF solvent, swab the agitator, inner surface, side discharge port and sampling glove box in "
        "sequence according to sampling point assessment diagram QM04053",
    "使用不锈钢扳手将尾气过滤器与管路上下两端的快开卡盘松动拆卸，取下管径部件;":
        f"{SS_WRENCH} loosen and remove the clamps connecting the vent filter to the pipeline at the upper and "
        "lower ends, and remove the pipe fitting;",
    "使用不锈钢扳手将氮气过滤器与弯头上下两端的快开卡盘松动拆卸，取下管径部件;":
        f"{SS_WRENCH} loosen and remove the clamps connecting the nitrogen filter to the elbow at the upper and "
        "lower ends, and remove the pipe fitting;",
    "使用不锈钢扳手将氮气过滤器与管路上下两端的快开卡盘松动拆卸，取下卡箍、滤芯、过滤器外壳及滤芯底座;":
        f"{SS_WRENCH} loosen and remove the clamps connecting the nitrogen filter to the pipeline at the upper "
        "and lower ends, and remove the clamp, filter cartridge, filter housing and cartridge base;",
    "使用不锈钢扳手将氮气过滤器与管路上下两端的快开卡盘松动拆卸，取下管径部件;":
        f"{SS_WRENCH} loosen and remove the clamps connecting the nitrogen filter to the pipeline at the upper "
        "and lower ends, and remove the pipe fitting;",
    "使用不锈钢扳手将氮气过滤器套筒快开卡箍固定螺栓逆时针松动取下，打开过滤器套筒，将滤芯取下废弃":
        f"{SS_WRENCH} loosen counterclockwise and remove the fixing bolt of the nitrogen filter housing clamp, "
        "open the filter housing, and remove and discard the filter cartridge",
    "使用不锈钢扳手将法兰盘四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将雷达液位计取下；":
        f"{SS_WRENCH} loosen the nuts of the four fixing bolts on the flange counterclockwise, remove the four "
        "bolts, and remove the radar level gauge;",
    "使用不锈钢扳手将淋洗球四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将淋洗球取出；":
        f"{SS_WRENCH} loosen the nuts of the four fixing bolts of the spray ball counterclockwise, remove the "
        "four bolts, and take out the spray ball;",
    "使用不锈钢扳手将视镜四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将试镜解体并将玻璃垫片取出取出；":
        f"{SS_WRENCH} loosen the nuts of the four fixing bolts of the sight glass counterclockwise, remove the "
        "four bolts, disassemble the sight glass and take out the glass and gasket;",
    "使用不锈钢扳手将阀门两端快开卡子拆除，将气动阀门拆下；":
        f"{SS_WRENCH} remove the clamps at both ends of the valve, and remove the pneumatic valve;",
    "使用不锈钢扳手将阀门中间连接的快开卡子紧固螺栓逆时针拧开取下，将单向阀解体":
        f"{SS_WRENCH} unscrew counterclockwise and remove the fastening bolt of the clamp in the middle of the "
        "valve, and disassemble the check valve",
    "使用不锈钢扳手将阀门从管路上拆除；": f"{SS_WRENCH} remove the valve from the pipeline;",
    "使用不锈钢扳手将阀门八个固定螺杆的螺母逆时针方向拧松，取下八个螺杆，将试镜与釜口分离并将垫片取下；":
        f"{SS_WRENCH} loosen the nuts of the eight fixing bolts counterclockwise, remove the eight bolts, "
        "separate the sight glass from the reactor nozzle and remove the gasket;",
    "使用不锈钢扳手将阀门四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将试镜解体并将玻璃取出；":
        f"{FOUR_NUTS} disassemble the sight glass and take out the glass;",
    "使用不锈钢扳手将阀门四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将试镜解体并将玻璃垫片取出取出；":
        f"{FOUR_NUTS} disassemble the sight glass and take out the glass and gasket;",
    "使用不锈钢扳手将阀门四个固定螺杆的螺母逆时针方向拧松，取下四个螺杆，将阀门解体并将阀球取出；":
        f"{FOUR_NUTS} disassemble the valve and take out the ball;",
    "使用不锈钢扳手将阀门四个固定螺杆逆时针方向拧松，取下四个螺杆及下方螺母，将球阀解体并将阀球取出；":
        f"{SS_WRENCH} loosen the four fixing bolts of the valve counterclockwise, remove the four bolts and the "
        "nuts underneath, disassemble the ball valve and take out the ball;",
    "使用洁净擦拭布蘸取纯化水后对取下后玻璃以及法兰，密封垫片表面，进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the removed glass, flange and gasket surfaces"),
    "使用洁净擦拭布蘸取纯化水后对拆卸后淋洗球，表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the surface of the removed spray ball"),
    "使用洁净擦拭布蘸取纯化水后对拆卸后雷达液位计与罐体接触面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the contact surface between the removed radar level gauge and the tank"),
    "使用洁净擦拭布蘸取纯化水后对拆除后阀门内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the inner surface of the removed valve"),
    "使用洁净擦拭布蘸取纯化水后对解体后玻璃以及密封垫片表面，两端不锈钢头进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the disassembled glass, gasket surfaces and stainless steel ends on both sides"),
    "使用洁净擦拭布蘸取纯化水后对解体后玻璃以及密封垫片表面，两端法兰进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the disassembled glass, gasket surfaces and flanges on both ends"),
    "使用洁净擦拭布蘸取纯化水后对解体后阀门及球体，表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the surfaces of the disassembled valve and ball"),
    "使用洁净擦拭布蘸取纯化水后对阀门及球体，表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the surfaces of the valve and ball"),
    "使用洁净擦拭布蘸取纯化水对卡箍、滤芯、过滤器外壳及滤芯底座进行擦拭清洗，直至目视无物料残留，使用压缩气体交替吹干，QA验收；":
        wipe("the clamp, filter cartridge, filter housing and cartridge base",
             blow="blow dry alternately with compressed air", end=";"),
    "使用洁净擦拭布蘸取纯化水对套筒内外表面以及密封连接处，进行擦拭清洗，直至目视无物料残留，使用压缩气体交替吹干，QA验收；":
        wipe("the inner and outer surfaces of the housing and the sealing joints",
             blow="blow dry alternately with compressed air", end=";"),
    "使用洁净擦拭布蘸取纯化水对弯头部件进行擦拭清洗，直至目视无物料残留，使用压缩气体交替吹干，QA验收；":
        wipe("the elbow", blow="blow dry alternately with compressed air", end=";"),
    "使用洁净擦拭布蘸取纯化水对滤芯表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干":
        wipe("the filter cartridge surface", qa=False),
    "使用洁净擦拭布蘸取纯化水对管径部件进行擦拭清洗，直至目视无物料残留，使用压缩气体交替吹干，QA验收；":
        wipe("the pipe fitting", blow="blow dry alternately with compressed air", end=";"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后Y型管进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从Y型管两端进行吹干，QA验收；":
        wipe_ends("the removed Y-pipe", "Y-pipe"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后三通进行擦洗，直至目视无物料残留，使用压缩气体交替从三通两端进行吹干，QA验收；":
        wipe_ends("the removed tee", "tee"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后变径=进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the removed reducer", "reducer"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后变径进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the removed reducer", "reducer"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后变径，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the removed reducer and seal surfaces", "reducer"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后变径，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹扫，直至无任何物料痕迹及水渍后，检查变径无划痕/腐蚀/老化/破损情况，有则更换，联系QA进行验收放行；":
        "Wipe the removed reducer and seal surfaces with a clean wipe moistened with purified water until no "
        "material residue is visible, purge with compressed air alternately from both ends of the reducer until "
        "no trace of material or water stains remains, then check the reducer for scratches/corrosion/aging/damage "
        "and replace it if any, and contact QA for acceptance and release;",
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后弯头进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从弯头两端进行吹干，QA验收；":
        wipe_ends("the removed elbow", "elbow"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后的三通进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从三通两端进行吹干， QA验收；":
        wipe_ends("the removed tee", "tee"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后直管进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从直管两端进行吹干，QA验收；":
        wipe_ends("the removed straight pipe", "straight pipe"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后直管进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从直管两端进行吹干，QA验收；将直管进行复位安装，更换部件两端连接的垫片":
        wipe_ends("the removed straight pipe", "straight pipe") + " " + reinstall("straight pipe"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后管件内表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the inner surface of the removed pipe fitting", "reducer"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后管路进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从管路两端进行吹干，QA验收；":
        wipe_ends("the removed pipe", "pipe"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除后阀门及内部阀芯，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体吹干，QA验收":
        wipe("the removed valve, internal valve core and seal surfaces", moist=True),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的三通，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从三通两端进行吹干，QA验收；":
        wipe_ends("the removed tee and seal surfaces", "tee"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的加料管路，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从加料管路两端进行吹干，QA验收；":
        wipe_ends("the removed feed line and seal surfaces", "feed line"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的变径，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the removed reducer and seal surfaces", "reducer"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的延伸管，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从延伸管两端进行吹干，QA验收；":
        wipe_ends("the removed extension pipe and seal surfaces", "extension pipe"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的弯头，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the removed elbow and seal surfaces", "reducer"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的弯头，密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从弯头两端进行吹干，QA验收；":
        wipe_ends("the removed elbow and seal surfaces", "elbow"),
    "使用洁净擦拭布蘸取纯化水浸湿后对拆除的过滤器底座与壳体密封件表面进行擦拭清洗，直至目视无物料残留，使用压缩气体交替从变径两端进行吹干，QA验收；":
        wipe_ends("the seal surfaces of the removed filter base and housing", "reducer"),
    "使用活口扳手将三通三端连接的螺丝逆时针旋转松动取下":
        f"{ADJ_WRENCH} loosen counterclockwise and remove the screws at the three ends of the tee",
    "使用活口扳手将与三通的连接端的快开卡盘松开;": f"{ADJ_WRENCH} release the clamp at the tee connection end;",
    "使用活口扳手将与三通连接端的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts at the tee connection end;",
    "使用活口扳手将与加料管路连接端的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts at the feed line connection end;",
    "使用活口扳手将与变径连接端的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts at the reducer connection end;",
    "使用活口扳手将与弯头连接端的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts at the elbow connection end;",
    "使用活口扳手将与管件的连接端的快开卡盘打开": f"{ADJ_WRENCH} open the clamp at the pipe fitting connection end",
    "使用活口扳手将与过滤器连接端的卡盘松开，拆卸过滤器底座与壳体":
        f"{ADJ_WRENCH} release the clamp at the filter connection end, and remove the filter base and housing",
    "使用活口扳手将与阀门的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts connected to the valve;",
    "使用活口扳手将与阀门的连接端的快开卡盘上两个固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the two fixing bolts on the clamp at the valve connection end;",
    "使用活口扳手将与阀门的连接端的快开卡盘上松开;": f"{ADJ_WRENCH} release the clamp at the valve connection end;",
    "使用活口扳手将与阀门的连接端的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts at the valve connection end;",
    "使用活口扳手将延伸管连接端的法兰固定螺杆逆时针方向旋转松动;":
        f"{ADJ_WRENCH} loosen counterclockwise the flange fixing bolts at the extension pipe connection end;",
    "使用活口扳手或对应尺寸的梅花扳手将固体加料口变径与搅拌罐连接的八颗螺丝逆时针旋转松动，再将与阀门的连接端的快开卡盘上两个固定螺杆逆时针方向旋转松动取下螺杆，快开卡箍分解为两瓣，取下卡箍、螺杆及四氟快开垫片放置在低密度聚乙烯袋中，并放在配件车上;":
        "Use an adjustable wrench or a box wrench of the appropriate size to loosen counterclockwise the eight "
        "screws connecting the solid charging port reducer to the mixing tank, then loosen counterclockwise and "
        "remove the two fixing bolts on the clamp at the valve connection end; separate the clamp into two halves, "
        "and place the clamp, bolts and PTFE clamp gasket in a low-density polyethylene bag on the parts cart;",
    "全面检查蝶阀整体状态，查看阀座密封面、阀杆及碟片 Qa验收":
        "Fully inspect the overall condition of the butterfly valve, including the seat sealing surface, stem "
        "and disc, QA acceptance",
    "其他溶剂加料泵": "Other solvent feed pump",
    "其他溶剂加料管路": "Other solvent feed line",
    "其他溶剂加料管路备用进液口": "Spare liquid inlet of other solvent feed line",
    "其他溶剂进料管路": "Other solvent feed line",
    "冲洗完成后按照P&ID图所示，依次将反应釜上的所有管路和阀门进行拆卸，拆卸清洁要求参见如下表所示，反应釜的氮气/尾气管路及阀门、爆破片、固体加料阀、其他物料加料阀、视镜拆卸至最小节，使用纯化水擦拭至无物料残留，压缩气体吹干。":
        "After rinsing, disassemble all pipelines and valves on the reactor in sequence as shown in the P&ID; "
        + DISASSEMBLE_TAIL.format(eq="reactor"),
    "冲洗完成后按照P&ID图所示，依次将固相合成仪上的需要拆卸清洁的管路和阀门进行拆卸，所有拆卸，清洁，验收及组装要求参见如下表5":
        "After rinsing, disassemble the pipelines and valves on the solid-phase synthesizer that require "
        "disassembly cleaning in sequence as shown in the P&ID; see Table 5 below for all disassembly, cleaning, "
        "acceptance and assembly requirements",
    "冲洗完成后按照P&ID图所示，依次将搅拌罐上的所有管路和阀门进行拆卸，拆卸清洁要求参见如下表所示，搅拌罐的氮气/尾气管路及阀门、爆破片、固体加料阀、其他物料加料阀、视镜拆卸至最小节，使用纯化水擦拭至无物料残留，压缩气体吹干。":
        "After rinsing, disassemble all pipelines and valves on the mixing tank in sequence as shown in the P&ID; "
        + DISASSEMBLE_TAIL.format(eq="mixing tank"),
    "出料手套箱出料口": "Discharge port of discharge glove box",
    "加DMF溶剂的专用溶剂管路": "Dedicated solvent line for DMF addition",
    "加DMF管路": "DMF addition line",
    "加DMF管路139788-step01-P19": "DMF addition line 139788-step01-P19",
    "加IPAC管路": "IPAC addition line",
    "加MTBE管路": "MTBE addition line",
    "加专用DMF溶剂最低点阀门": "Low-point valve of dedicated DMF solvent addition line",
    "加固体料": "Solid charging",
    "加料管线进搅拌罐阀门TJ4S-1213-AT08-V16": "Feed line inlet valve to mixing tank TJ4S-1213-AT08-V16",
    "加新鲜溶剂管线进釜阀门TJ4S-1213-Sy12-V21, TJ4S-1213-Sy12-V75":
        "Fresh solvent feed line inlet valves to reactor TJ4S-1213-Sy12-V21, TJ4S-1213-Sy12-V75",
    "加桶装溶剂管路进釜阀门": "Drummed solvent feed line inlet valve to reactor",
    "加溶剂管路上阀门": "Valve on solvent addition line",
    "单向阀": "Check valve",
    "压力表下阀门": "Valve below pressure gauge",
    "原位清洁": "In-situ cleaning",
    "原位清洁（在线清洗）": "In-situ cleaning (online cleaning)",
    "原位清洁（氮气吹扫）": "In-situ cleaning (nitrogen purge)",
    "参照工艺辅助材料使用计划": "Refer to the process auxiliary material usage plan",
    "取样人及日期:": "Sampled by&Date",
    "取样口": "Sampling port",
    "取样方式": "Sampling method",
    "取样方式和位置*": "Sampling method and location*",
    "取样记录": "Sampling record",
    "变径": "Reducer",
    "合成": "Synthesis",
    "合成仪TJ4S-1213-Sy12自带手套箱": "Built-in glove box of synthesizer TJ4S-1213-Sy12",
    "合成釜压力变送器阀门": "Synthesis reactor pressure transmitter valve",
    "合成釜手套箱尾气管路": "Synthesis reactor glove box vent line",
    "合成釜手套箱氮气管路": "Synthesis reactor glove box nitrogen line",
    "合成釜氮气管路": "Synthesis reactor nitrogen line",
    "合成釜视镜": "Synthesis reactor sight glass",
    "合成釜音叉及压力变送器连接口": "Synthesis reactor tuning fork and pressure transmitter connection port",
    "合成釜音叉及压力检测口": "Synthesis reactor tuning fork and pressure detection port",
    "向激活釜加入新鲜溶剂管路上": "On the fresh solvent line to the activation reactor",
    "哈氏合金": "Hastelloy",
    "回流管路阀门": "Reflux line valve",
    "回流管路（": "Reflux line (condenser to receiver tank and TJ4S-1213-R23 reactor line)",
    "回流阀": "Reflux valve",
    "固体加料口": "Solid charging port",
    "固体加料口阀门": "Solid charging port valve",
    "固体加料管": "Solid charging pipe",
    "固体投料口": "Solid charging port",
    "固相合成仪下口排液管路阀门TJ4S-1213-Sy12-XV29": "Solid-phase synthesizer bottom drain line valve TJ4S-1213-Sy12-XV29",
    "在激活釜内通过淋洗球循环清洗激活液转移管路139788-step01-P17，固体加料口TJ4S-1213-Sy12-V22和液体加料口TJ4S-1213-Sy12-V14，每条线路各清洗3min。":
        "In the activation reactor, circulate through the spray ball to clean activation solution transfer line "
        "139788-step01-P17, solid charging port TJ4S-1213-Sy12-V22 and liquid charging port TJ4S-1213-Sy12-V14, "
        "3 min for each line.",
    "在阀门TJ4S-1213-Sy12-V32处连接氮气管路，开启氮气，并打开TJ4S-1213-Sy12-V32和进液阀门TJ4S-1213-Sy12-V51，TJ4S-1213-Sy12-V52、TJ4S-1213-Sy12-V50使用氮气吹扫管路":
        "Connect a nitrogen line at valve TJ4S-1213-Sy12-V32, turn on nitrogen, open TJ4S-1213-Sy12-V32 and "
        "liquid inlet valves TJ4S-1213-Sy12-V51, TJ4S-1213-Sy12-V52 and TJ4S-1213-Sy12-V50, and purge the line "
        "with nitrogen",
    "在阀门TJ4S-1213-Sy12-V32处连接氮气管路，开启氮气，并打开TJ4S-1213-Sy12-V32和进液阀门TJ4S-1213-Sy12-V51，TJ4S-1213-Sy12-V52、TJ4S-1213-Sy12-V50使用氮气吹扫管路至无液滴":
        "Connect a nitrogen line at valve TJ4S-1213-Sy12-V32, turn on nitrogen, open TJ4S-1213-Sy12-V32 and "
        "liquid inlet valves TJ4S-1213-Sy12-V51, TJ4S-1213-Sy12-V52 and TJ4S-1213-Sy12-V50, and purge the line "
        "with nitrogen until no droplets remain",
    "填写人及日期：": "Filled by&Date",
    "备用液体进料管": "Spare liquid feed pipe",
    "多肽固相合成仪": "Peptide solid-phase synthesizer",
    "多肽固相合成仪&激活釜": "Peptide solid-phase synthesizer & activation reactor",
    "大清洁清洁方法": "Major cleaning method",
    "安全阀": "Safety valve",
    "安装新的折叠滤芯，将套筒依次完成复位组装，并更换过滤器两端连接的垫片":
        "Install a new pleated filter cartridge, reassemble the housing in sequence, and replace the gaskets at "
        "both ends of the filter",
    "完成后粘贴于记录背面骑缝签字": "After completion, attach it to the back of this record and sign across the seam",
    "对固相合成仪上安装的手套进行更换": "Replace the gloves installed on the solid-phase synthesizer",
    "对尾气滤芯进行清洁": "Clean the vent filter cartridge",
    "对氮气滤芯进行清洁": "Clean the nitrogen filter cartridge",
    "对转移管路139788-step01-P21的最低点阀门139788- step01-P21-009进行排净。":
        "Drain transfer line 139788-step01-P21 completely through its low-point valve 139788- step01-P21-009.",
    "将Y型管进行复位安装，更换部件两端连接的垫片": reinstall("Y-pipe"),
    "将三通进行回装对应管路，装上垫片并在插固定螺杆顺时针拧紧":
        "Reinstall the tee onto the corresponding pipeline, fit the gaskets, insert the fixing bolts and "
        "tighten clockwise",
    "将三通进行复位安装，更换部件三端连接的快开垫片": reinstall("tee", "clamp gaskets at all three ends"),
    "将三通进行复位安装，更换部件两端连接的垫片": reinstall("tee"),
    "将加料管路进行复位安装，更换部件两端连接的垫片": reinstall("feed line"),
    "将变径进行复位安装，更换部件两端连接的垫片": reinstall("reducer"),
    "将合成仪内的 N,N-二甲基甲酰胺通过废液管139788-step01-P16排至废液罐。":
        "Drain the N,N-dimethylformamide in the synthesizer to the waste tank through waste line 139788-step01-P16.",
    "将合成仪内的纯化水通过排放管路139788-step01-P16排至废液罐":
        "Drain the purified water in the synthesizer to the waste tank through drain line 139788-step01-P16",
    "将延伸管进行复位安装，更换部件两端连接的垫片": reinstall("extension pipe"),
    "将弯头进行复位安装，更换部件两端连接的垫片": reinstall("elbow"),
    "将拆卸后雷达液位计依次完成复位组装，更换阀门两端连接的快开垫片":
        "Reassemble the removed radar level gauge in sequence and replace the clamp gaskets at both ends of the valve",
    "将拆卸清洗验收完毕的阀门安装至对应位置":
        "Install the disassembled, cleaned and accepted valves back to their corresponding positions",
    "将拆解的单向阀依次完成复位组装，更换阀门两端连接的快开垫片": reassemble("check valve", "clamp gaskets", "valve"),
    "将拆解的淋洗球依次完成复位组装，更换阀门两端连接的快开垫片": reassemble("spray ball", "clamp gaskets", "valve"),
    "将拆解的管径依次完成复位组装，更换管件两端连接的快开垫片": reassemble("pipe fitting", "clamp gaskets", "fitting"),
    "将拆解的管径弯头依次完成复位组装，更换弯头两端连接的快开垫片": reassemble("pipe elbow", "clamp gaskets", "elbow"),
    "将拆解的试镜依次完成复位组装，更换阀门两端连接的快开垫片": reassemble("sight glass", "clamp gaskets", "valve"),
    "将拆解的试镜依次完成复位组装，更换阀门两端连接的法兰垫片": reassemble("sight glass", "flange gaskets", "valve"),
    "将拆解的试镜法兰进行回装": "Reinstall the disassembled sight glass flange",
    "将拆解的过滤器依次完成复位组装，更换过滤器两端连接的快开垫片": reassemble("filter", "clamp gaskets", "filter"),
    "将拆解的阀门依次完成复位组装，更换阀门两端连接的四氟法兰垫片": reassemble("valve", "PTFE flange gaskets", "valve"),
    "将拆解的阀门依次完成复位组装，更换阀门两端连接的密封垫片": reassemble("valve", "sealing gaskets", "valve"),
    "将拆解的阀门依次完成复位组装，更换阀门两端连接的快开垫片": reassemble("valve", "clamp gaskets", "valve"),
    "将激活釜内的 N,N-二甲基甲酰胺通过转移管路139788-step01-P17及转移管路139788-step01-P11经淋洗球淋洗固相合成仪，以及氮气吹扫转移管路139788-step01-P17及转移管路139788-step01-P11 1min":
        "Transfer the N,N-dimethylformamide in the activation reactor through transfer lines 139788-step01-P17 "
        "and 139788-step01-P11 to rinse the solid-phase synthesizer via the spray ball, and purge transfer lines "
        "139788-step01-P17 and 139788-step01-P11 with nitrogen for 1 min",
    "将直管进行复位安装，更换部件两端连接的垫片": reinstall("straight pipe"),
    "将管件进行复位安装，更换部件两端连接的垫片": reinstall("pipe fitting"),
    "将管路进行复位安装，更换部件两端连接的垫片": reinstall("pipe"),
    "将过滤器底座及壳体进行复位安装，更换部件两端连接的垫片": reinstall("filter base and housing"),
    "尾气滤芯": "Vent filter cartridge",
    "尾气管路": "Vent line",
    "序号": "No.",
    "开启合成仪搅拌和下口阀门TJ4S-1213-Sy12-V57鼓氮气程序搅洗15min. 搅拌转速 25±5rpm, 氮气流量 10±3m³/h":
        "Start the synthesizer agitation and the nitrogen bubbling program through bottom valve "
        "TJ4S-1213-Sy12-V57, and wash with agitation for 15 min. Agitation speed 25±5 rpm, nitrogen flow 10±3 m³/h",
    "开启合成仪搅拌和下口鼓氮气程序搅洗5min. 搅拌转速 25±5rpm, 氮气流量 10±3m³/h":
        "Start the synthesizer agitation and the bottom nitrogen bubbling program, and wash with agitation for "
        "5 min. Agitation speed 25±5 rpm, nitrogen flow 10±3 m³/h",
    "弯头": "Elbow",
    "弯管": "Bent pipe",
    "循环清洗进液管路": "Circulation cleaning of liquid inlet line",
    "快开球阀": "Clamp-type ball valve",
    "快开阀门": "Clamp-type valve",
    "手动清洁(清洗筛板及相关阀门)：": "Manual cleaning (cleaning of sieve plate and related valves):",
    "手动隔膜阀": "Manual diaphragm valve",
    "手套": "Gloves",
    "手套箱尾气管路": "Glove box vent line",
    "手套箱氮气管路": "Glove box nitrogen line",
    "打开新鲜溶剂管路处氮气阀门TJ4S-1213-Sy12-R1XV1和进液阀门TJ4S-1213-Sy12-V21、TJ4S-1213-Sy12-V75使用氮气吹扫管路":
        "Open nitrogen valve TJ4S-1213-Sy12-R1XV1 at the fresh solvent line and liquid inlet valves "
        "TJ4S-1213-Sy12-V21 and TJ4S-1213-Sy12-V75, and purge the line with nitrogen",
    "批次间和生产结束后拆卸大清洁": "Major disassembly cleaning between batches and after end of production",
    "批次间更换": "Replace between batches",
    "批次间清洁": "Between-batch cleaning",
    "投料口": "Charging port",
    "拆卸清洁": "Disassembly cleaning",
    "拆卸清洗": "Disassembly cleaning",
    "按照标准清洁方式进行设备淋洗取样": "Perform equipment rinse sampling following the standard cleaning procedure",
    "按照标准清洁方式，调用配方CMaTJ413978801A01进行在线清洁":
        "Following the standard cleaning procedure, run recipe CMaTJ413978801A01 for online cleaning",
    "按照标准清洁方式，调用配方CMaTJ413978801B01进行在线清洁":
        "Following the standard cleaning procedure, run recipe CMaTJ413978801B01 for online cleaning",
    "按照标准清洁方式，进行设备清洁": "Clean the equipment following the standard cleaning procedure",
    "排尽新鲜溶剂气动隔膜泵及管路，氮气吹干新鲜溶剂泵TJ4S-1213-Sy12-P05及其加料管路139788-step01-P25及139788-step01-P11，吹扫时间20±5min":
        "Drain the fresh solvent pneumatic diaphragm pump and lines completely, and blow dry fresh solvent pump "
        "TJ4S-1213-Sy12-P05 and its feed lines 139788-step01-P25 and 139788-step01-P11 with nitrogen for 20±5 min",
    "接收罐尾气": "Receiver tank vent",
    "接收罐氮气": "Receiver tank nitrogen",
    "搅拌桨、内表面、侧出料口、取样手套箱": "Agitator, inner surface, side discharge port, sampling glove box",
    "搅拌罐": "Mixing tank",
    "搪玻璃": "Glass-lined",
    "擦拭取样；": "Swab sampling;",
    "放行后将三通进行回装对应管路，装上垫片并在插固定螺杆顺时针拧紧":
        "After release, reinstall the tee onto the corresponding pipeline, fit the gaskets, insert the fixing "
        "bolts and tighten clockwise",
    "放行后将变径进行回装合成仪上对应口，装上垫片并插入八个固定螺杆顺时针拧紧，安装阀门，两侧放入垫片，接阀门端安装快开卡箍，每侧插入两个螺杆，并顺时针禁":
        "After release, reinstall the reducer onto the corresponding port of the synthesizer, fit the gasket, "
        "insert the eight fixing bolts and tighten clockwise; install the valve with gaskets on both sides, fit "
        "the clamp at the valve end, insert two bolts on each side and tighten clockwise",
    "料液管线，搅拌罐阀门TJ4S-1213-AT08-V12": "Feed liquid line, mixing tank valve TJ4S-1213-AT08-V12",
    "新鲜溶剂加料管线进釜阀门": "Fresh solvent feed line inlet valve to reactor",
    "新鲜溶剂氮气吹扫管路": "Fresh solvent nitrogen purge line",
    "更换后的手套批号：": "Lot No. of replacement gloves",
    "更换蝶阀两端垫片": "Replace the gaskets at both ends of the butterfly valve",
    "样品编号：": "Sample No.",
    "桶装溶剂ACN，通过阀门TJ4S-1213-AT08-V07及淋洗球向搅拌罐TJ4S-1213-AT08中加入55±5 kg溶剂，通过下口连接管路139788-step01-P01放出，并在末端阀门TJ4S-1213-Sy12-V16位置进行取样":
        "Add 55±5 kg of drummed solvent ACN into mixing tank TJ4S-1213-AT08 through valve TJ4S-1213-AT08-V07 "
        "and the spray ball, discharge through bottom connecting line 139788-step01-P01, and take samples at end "
        "valve TJ4S-1213-Sy12-V16",
    "检测结果是否合格": "Is the test result acceptable?",
    "检测结果：": "Test result",
    "止回阀": "Non-return valve",
    "此合成仪排液管线139788-step01-P16 ，清洗方式详见设备主体清洁方式表20中在线清洗部分":
        this_line("synthesizer drain line ", "P16", 20),
    "此合成仪排液管线139788-step01-P16 ，清洗方式详见设备主体清洁方式表4中在线清洗部分":
        this_line("synthesizer drain line ", "P16", 4),
    "此激活釜罐底转移管线139788-step01-P17 ，清洗方式详见设备主体清洁方式表20中在线清洗部分":
        this_line("activation reactor bottom transfer line ", "P17", 20),
    "此激活釜罐底转移管线139788-step01-P17 ，清洗方式详见设备主体清洁方式表4中在线清洗部分":
        this_line("activation reactor bottom transfer line ", "P17", 4),
    "此管线139788-step01-P01清洁方式已包含在表25清洗部分":
        "Cleaning of this line 139788-step01-P01 is included in the cleaning section of Table 25",
    "此管线139788-step01-P11，清洗方式详见设备主体清洁方式表20中在线清洗部分": this_line("line ", "P11", 20),
    "此管线139788-step01-P11，清洗方式详见设备主体清洁方式表4中在线清洗部分": this_line("line ", "P11", 4),
    "此管线139788-step01-P19清洁方式已包含在表20清洗部分":
        "Cleaning of this line 139788-step01-P19 is included in the cleaning section of Table 20",
    "此管线139788-step01-P19清洁方式已包含在表4在线清洗部分":
        "Cleaning of this line 139788-step01-P19 is included in the online cleaning section of Table 4",
    "此转移管线与主设备搅拌罐XX同步通过139788-step01-P01进行清洁 ，清洗方式详见设备主体清洁方式表8中清洗部分":
        "This transfer line is cleaned through 139788-step01-P01 together with main equipment mixing tank XX; "
        "see the cleaning section of main equipment cleaning in Table 8 for the cleaning method",
    "此部件在139788-step01-P11管线中 ，清洗方式详见设备主体清洁方式表20中在线清洗部分": see_online("This component", "P11", 20),
    "此部件在139788-step01-P11管线中 ，清洗方式详见设备主体清洁方式表4中在线清洗部分": see_online("This component", "P11", 4),
    "此部件在139788-step01-P17管线中 ，清洗方式详见设备主体清洁方式表20中在线清洗部分": see_online("This component", "P17", 20),
    "此部件在139788-step01-P17管线中 ，清洗方式详见设备主体清洁方式表4中在线清洗部分": see_online("This component", "P17", 4),
    "此部件在139788-step01-P19管线中 ，清洗方式详见设备主体清洁方式表20中在线清洗部分": see_online("This component", "P19", 20),
    "此部件在139788-step01-P19管线中 ，清洗方式详见设备主体清洁方式表4中在线清洗部分": see_online("This component", "P19", 4),
    "此部件在139788-step01-P25管线中 ，清洗方式详见管件清洗部分28":
        "This component is in line 139788-step01-P25; see pipe fitting cleaning section 28 for the cleaning method",
    "此部件在139788-step01-P26管线中 ，清洗方式详见管件清洗部分28":
        "This component is in line 139788-step01-P26; see pipe fitting cleaning section 28 for the cleaning method",
    "此部分阀门包含在表22中139788-step01-P13管线中，同该管线一并进行清洁":
        "These valves are included in line 139788-step01-P13 in Table 22 and are cleaned together with that line",
    "此部分阀门包含在表8中139788-step01-P19管线中，同该管线一并进行清洁":
        "These valves are included in line 139788-step01-P19 in Table 8 and are cleaned together with that line",
    "此部分阀门安装在表6中的139788-step01-P13管线中，同该管线一并进行清洁":
        "These valves are installed in line 139788-step01-P13 in Table 6 and are cleaned together with that line",
    "此阀门包含在表8中139788-step01-P19管线中，同该管线一并进行清洁":
        "This valve is included in line 139788-step01-P19 in Table 8 and is cleaned together with that line",
    "此阀门在139788-step01-P01管线中 ，清洗方式详见设备主体清洁方式表20中在线清洗部分": see_online("This valve", "P01", 20),
    "此阀门在139788-step01-P01管线中 ，清洗方式详见设备主体清洁方式表8中在线清洗部分": see_online("This valve", "P01", 8),
    "此阀门在139788-step01-P17管线中 ，清洗方式详见设备主体清洁方式表20中在线清洗部分": see_online("This valve", "P17", 20),
    "此阀门在139788-step01-P17管线中 ，清洗方式详见设备主体清洁方式表4中在线清洗部分": see_online("This valve", "P17", 4),
    "此阀门在139788-step01-P17管线中 ，清洗方式详见设备主体清洁方式表8中在线清洗部分": see_online("This valve", "P17", 8),
    "气动球阀": "Pneumatic ball valve",
    "气动隔膜泵": "Pneumatic diaphragm pump",
    "气动隔膜阀": "Pneumatic diaphragm valve",
    "氮气吹干激活釜和合成仪以及转移管路139788-step01-P16，各吹扫时间 30±10min.":
        "Blow dry the activation reactor, the synthesizer and transfer line 139788-step01-P16 with nitrogen, "
        "30±10 min each.",
    "氮气滤芯": "Nitrogen filter cartridge",
    "氮气管路": "Nitrogen line",
    "氮气连接口三通": "Nitrogen connection port tee",
    "法兰三通": "Flanged tee",
    "法兰变径": "Flanged reducer",
    "法兰球阀": "Flanged ball valve",
    "泵TJ4S-1213-R23-P01到TJ4S-1213-Sy12和泵TJ4S-1213-R23-P01到TJ4S-1213-R23的自循环管路":
        "Self-circulation lines from pump TJ4S-1213-R23-P01 to TJ4S-1213-Sy12 and from pump TJ4S-1213-R23-P01 "
        "to TJ4S-1213-R23",
    "活化": "Activation",
    "流体过滤器进口": "Fluid filter inlet",
    "液体投料口": "Liquid charging port",
    "淋洗取样": "Rinse sampling",
    "淋洗口": "Rinse port",
    "淋洗球": "Spray ball",
    "淋洗球上口": "Spray ball top port",
    "淋洗球进料口": "Spray ball inlet",
    "淋洗球阀门": "Spray ball valve",
    "清洁策略": "Cleaning strategy",
    "清洁策略的出处": "Source of cleaning strategy",
    "清洁记录": "Cleaning record",
    "溶剂冲洗，氮气吹干": "Solvent flush, nitrogen blow-dry",
    "溶剂加料口": "Solvent charging port",
    "滴加DIC管线进釜阀门TJ4S-1213-Sy12-R1XV5": "DIC dosing line inlet valve to reactor TJ4S-1213-Sy12-R1XV5",
    "激活液转移管路": "Activation solution transfer line",
    "激活釜": "Activation reactor",
    "激活釜上口": "Activation reactor top port",
    "激活釜下口转移至固相合成仪及循环管路":
        "Transfer and circulation lines from activation reactor bottom outlet to solid-phase synthesizer",
    "激活釜下口阀门": "Activation reactor bottom valve",
    "激活釜内的纯化水通过转移管路139788-step01-P17及转移管路139788-step01-P11和合成仪淋洗球淋洗合成仪，氮气吹扫转移管1min。":
        "Rinse the synthesizer with the purified water in the activation reactor through transfer lines "
        "139788-step01-P17 and 139788-step01-P11 and the synthesizer spray ball, and purge the transfer lines "
        "with nitrogen for 1 min.",
    "激活釜备用滴加管线阀门": "Activation reactor spare dosing line valve",
    "激活釜滴加DIC管线阀门": "Activation reactor DIC dosing line valve",
    "激活釜通过纯化水专用软管通过TJ4S-1213-Sy12-V14阀门及淋洗球加入210±10kg纯化水，循环清洗激活液转移管路139788-step01-P17及转移管路139788-step01-P11，固体加料口TJ4S-1213-Sy12-V22和液体加料口TJ4S-1213-Sy18，每条线路各清洗3min":
        "Add 210±10 kg of purified water to the activation reactor through the dedicated purified water hose, "
        "valve TJ4S-1213-Sy12-V14 and the spray ball; circulate to clean activation solution transfer line "
        "139788-step01-P17, transfer line 139788-step01-P11, solid charging port TJ4S-1213-Sy12-V22 and liquid "
        "charging port TJ4S-1213-Sy18, 3 min for each line",
    "版本号": "Version No.",
    "物料转移管路": "Material transfer line",
    "玻璃": "Glass",
    "生产结束后清洁": "End-of-production cleaning",
    "生效日期": "Effective date",
    "甲苯进料管路": "Toluene feed line",
    "直管": "Straight pipe",
    "短接": "Spool piece",
    "管线编号": "Line No.",
    "管路编号": "Line No.",
    "纯化水称重参见称重记录表，编号为：": "For purified water weighing, refer to the weighing record No.",
    "纯化水称重参见称重记录表，编号为：CPo139788-   -01-W-    ，完成后粘贴于记录背面骑缝签字": [
        ("纯化水称重", "For purified water weighing, refer to the weighing record No."), ("完成后", AFTER_DONE)],
    "编号": "Number",
    "罐区溶剂DMF，通过139788-step01-P19管路及淋洗球向激活釜TJ4S-1213-Sy12-R01中加入65±5 kg溶剂，并通过转移管路139788-step01-P17及139788-step01-P11，转移至固相合成仪TJ4S-1213-Sy12，通过固相合成仪下口连接管路139788-step01-P16放出，并在下口阀TJ4S-1213-Sy12-XV29位置进行取样":
        "Add 65±5 kg of tank-farm solvent DMF into activation reactor TJ4S-1213-Sy12-R01 through line "
        "139788-step01-P19 and the spray ball, transfer it to solid-phase synthesizer TJ4S-1213-Sy12 through "
        "transfer lines 139788-step01-P17 and 139788-step01-P11, discharge through the synthesizer bottom "
        "connecting line 139788-step01-P16, and take samples at bottom valve TJ4S-1213-Sy12-XV29",
    "罐区溶剂通过 N,N-二甲基甲酰胺加料管路139788-step01-P19向激活釜TJ4S-1213-Sy12-R01内加入 210±10kg N,N-二甲基甲酰胺":
        "Add 210±10 kg of N,N-dimethylformamide from the tank farm into activation reactor TJ4S-1213-Sy12-R01 "
        "through N,N-dimethylformamide feed line 139788-step01-P19",
    "罐区溶剂通过N,N-二甲基甲酰胺加料管路139788-step01-P19向激活釜TJ4S-1213-Sy12-R01内加入65±5kg N,N-二甲基酰胺淋洗激活液，循环 3min，后通过转移管路139788-step01-P17转移至固相合成仪内，开启搅拌和下口鼓氮气，持续 15min，搅拌转速 25±5rpm，氮气流量 10±3m³/h，再通过废液管路139788-step01-P16排空":
        "Add 65±5 kg of N,N-dimethylformamide from the tank farm into activation reactor TJ4S-1213-Sy12-R01 "
        "through N,N-dimethylformamide feed line 139788-step01-P19 to rinse off the activation solution, "
        "circulate for 3 min, then transfer into the solid-phase synthesizer through transfer line "
        "139788-step01-P17; start agitation and bottom nitrogen bubbling for 15 min, agitation speed 25±5 rpm, "
        "nitrogen flow 10±3 m³/h, then drain through waste line 139788-step01-P16",
    "罐区溶剂通过甲基叔丁基醚加料管路139788-step01-P19 向激活釜TJ4S-1213-Sy12-R01加入137.5±12.5kg 甲基叔丁基醚淋洗激活液，循环 3min，后通过转移管路139788-step01-P17转移至合成仪内，开启搅拌和下口鼓氮气，持续 15min，再通过废液管路139788-step01-P16排空。搅拌转速 25±5rpm，氮气流量 10±3m³/h。":
        "Add 137.5±12.5 kg of methyl tert-butyl ether from the tank farm into activation reactor "
        "TJ4S-1213-Sy12-R01 through methyl tert-butyl ether feed line 139788-step01-P19 to rinse off the "
        "activation solution, circulate for 3 min, then transfer into the synthesizer through transfer line "
        "139788-step01-P17; start agitation and bottom nitrogen bubbling for 15 min, then drain through waste "
        "line 139788-step01-P16. Agitation speed 25±5 rpm, nitrogen flow 10±3 m³/h.",
    "脱保护管路氮气吹扫管路": "Nitrogen purge line of deprotection line",
    "脱保护进料管路": "Deprotection feed line",
    "自动方法清洗": "Automatic cleaning",
    "自循环管路": "Self-circulation line",
    "蒸馏管路（TJ4S-1213-R23到冷凝器的管路）": "Distillation line (line from TJ4S-1213-R23 to condenser)",
    "蝶阀": "Butterfly valve",
    "表10，搅拌罐自身配件清洗": f"Table 10, {TANK_ACC}",
    "表11，搅拌罐管线清洗": f"Table 11, {TANK_LINES}",
    "表12，从搅拌罐TJ4S-1213-AT08到固相合成仪TJ4S-1213-Sy12物料转移管路的管路139788-step01-P01清洁如下":
        f"Table 12, Cleaning of {P01_LINE}",
    "表13，从激活釜TJ4S-1213-Sy12-R01到固相合成仪TJ4S-1213-Sy12物料转移管路的管路139788-step01-P17清洁如下":
        f"Table 13, Cleaning of {P17_LINE}",
    "表14，从搅拌罐TJ4S-1213-AT08到固相合成仪TJ4S-1213-Sy12物料转移管路的管路139788-step01-P01上连接的阀门清洁如下":
        f"Table 14, Cleaning of valves connected on {P01_LINE}",
    "表15，从激活釜TJ4S-1213-Sy12-R01到固相合成仪TJ4S-1213-Sy12物料转移管路139788-step01-P17上连接的阀门清洁如下":
        f"Table 15, Cleaning of valves connected on {P17_LINE}",
    "表16，step1工序相关耗材": "Table 16, Consumables related to Step 1",
    "表17，设备清洁后取样信息": "Table 17, Sampling information after equipment cleaning",
    "表18，固相合成仪及激活釜主体清洗": f"Table 18, {MAIN_BODY_18}",
    "表19，固相合成仪上安装的阀门及部件清洗": f"Table 19, {VALVES_SYN}",
    "表1，设备概况": "Table 1, Equipment overview",
    "表20，合成仪自身配备的管路清洗": f"Table 20, {LINES_SYN}",
    "表21，激活釜上安装的阀门及部件清洗": f"Table 21, {VALVES_ACT}",
    "表22，激活釜自身配备的管路清洗": f"Table 22, {LINES_ACT}",
    "表23，搅拌罐主体清洗": f"Table 23, {TANK_BODY}",
    "表24，搅拌罐自身配件清洗": f"Table 24, {TANK_ACC}",
    "表25，搅拌罐管线清洗": f"Table 25, {TANK_LINES}",
    "表26，预冷釜主体清洗": "Table 26, Cleaning of pre-cooling reactor main body",
    "表27，预冷釜阀门部件清洗": "Table 27, Cleaning of pre-cooling reactor valves and components",
    "表28，预冷釜管件清洗": "Table 28, Cleaning of pre-cooling reactor pipe fittings",
    "表29，从搅拌罐TJ4S-1213-AT08到固相合成仪TJ4S-1213-Sy12物料转移管路的管路139788-step01-P01清洁如下":
        f"Table 29, Cleaning of {P01_LINE}",
    "表2，单体设备P&ID图信息": "Table 2, P&ID information of individual equipment",
    "表30，从激活釜TJ4S-1213-Sy12-R01到固相合成仪TJ4S-1213-Sy12物料转移管路的管路139788-step01-P17清洁如下":
        f"Table 30, Cleaning of {P17_LINE}",
    "表31，从搅拌罐TJ4S-1213-AT08到固相合成仪TJ4S-1213-Sy12物料转移管路的阀门清洁如下":
        "Table 31, Cleaning of valves on the material transfer line from mixing tank TJ4S-1213-AT08 to "
        "solid-phase synthesizer TJ4S-1213-Sy12",
    "表32，从激活釜TJ4S-1213-Sy12-R01到固相合成仪TJ4S-1213-Sy12物料转移管路139788-step01-P17上连接的阀门清洁如下":
        f"Table 32, Cleaning of valves connected on {P17_LINE}",
    "表33，设备相关耗材清洁方式如下": "Table 33, Cleaning methods of equipment-related consumables",
    "表34，设备清洁后取样位置及方式如下": "Table 34, Sampling locations and methods after equipment cleaning",
    "表3设备清洁策略": "Table 3 Equipment cleaning strategy",
    "表4，固相合成仪及激活釜主体清洗": f"Table 4, {MAIN_BODY_18}",
    "表5，固相合成仪上安装的阀门及部件清洗": f"Table 5, {VALVES_SYN}",
    "表6，合成仪自身配备的管路清洗": f"Table 6, {LINES_SYN}",
    "表7，激活釜上安装的阀门及部件清洗": f"Table 7, {VALVES_ACT}",
    "表8，激活釜自身配备的管路清洗": f"Table 8, {LINES_ACT}",
    "表9，搅拌罐主体清洗": f"Table 9, {TANK_BODY}",
    "衬halar": "Halar-lined",
    "衬四氟": "PTFE-lined",
    "衬塑": "Plastic-lined",
    "观察视镜": "Observation sight glass",
    "视镜": "Sight glass",
    "设备名称": "Equipment name",
    "设备本身清洁方式见如下表4：": "The cleaning method of the equipment itself is shown in Table 4 below:",
    "设备材质": "Equipment material",
    "设备氮气连接口变径": "Equipment nitrogen connection port reducer",
    "设备用途": "Equipment purpose",
    "设备编号": "Equipment No.",
    "设备罐体及配件验收完毕后，反向依次将视镜、其他物料加料阀、固体加料阀、爆破片、氮气/尾气管路及阀门进行安装，安装完毕后，使用氮气将搅拌罐和转移管路吹干":
        REINSTALL_TANK,
    "设备罐体及配件验收完毕后，反向依次将视镜、其他物料加料阀、固体加料阀、爆破片、氮气/尾气管路及阀门进行安装，安装完毕后，使用氮气将搅拌罐和转移管路吹干。":
        REINSTALL_TANK + ".",
    "设备链不涉及消毒，具体评估方式参考CVP-1397880301.01文件":
        "Disinfection is not applicable to the equipment train; refer to document CVP-1397880301.01 for the assessment",
    "试镜": "Sight glass",
    "该设备的清洁从清洁时间节点上分为批次间清洁和生产结束后清洁。":
        "By timing, cleaning of this equipment is divided into between-batch cleaning and end-of-production cleaning.",
    "调用自动清洗配方执行设备主体的在线清洗(清洗设备主体及转移管路)，配方编号: CMaTJ413978801A01":
        "Run the automatic cleaning recipe to perform online cleaning of the main equipment (cleaning the main "
        "equipment and transfer lines), recipe No.: CMaTJ413978801A01",
    "调用自动清洗配方进行工艺溶剂淋洗及干燥，配方编号CMaTJ413978801B01":
        "Run the automatic cleaning recipe for process solvent rinsing and drying, recipe No. CMaTJ413978801B01",
    "过滤器出口": "Filter outlet",
    "运行配方流程如下：": "The recipe runs as follows:",
    "进氮气阀门": "Nitrogen inlet valve",
    "连接位置": "Connection location",
    "通过反应釜TJ4S-1213-R23上的固定淋洗口TJ4S-1213-R23-V62通过纯化水专用管路及专用质量流量计使用 220±10kg 纯化水冲洗反应釜和转移管路139788-step01-P21，放出至废液罐。":
        "Through fixed rinse port TJ4S-1213-R23-V62 on reactor TJ4S-1213-R23, rinse the reactor and transfer "
        "line 139788-step01-P21 with 220±10 kg of purified water via the dedicated purified water line and "
        "dedicated mass flowmeter, and discharge to the waste tank.",
    "通过固相合成釜加料口TJ4S-1213-Sy12-V11，通过纯化水专用管路及专用质量流量计使用 90±10kg 纯化水，冲洗合成仪内表面、筛板表面，冲洗完成后，冲洗水通过下口阀TJ4S-1213-Sy12-V57及管路139788-step01-P16排至废液罐":
        "Through solid-phase synthesis reactor charging port TJ4S-1213-Sy12-V11, rinse the synthesizer inner "
        "surface and sieve plate surface with 90±10 kg of purified water via the dedicated purified water line and "
        "dedicated mass flowmeter; after rinsing, drain the rinse water to the waste tank through bottom valve "
        "TJ4S-1213-Sy12-V57 and line 139788-step01-P16",
    "通过搅拌罐上淋洗球阀门TJ4S-1213-Sy12-V07，使用新鲜溶剂专用泵将77.5±2.5kg 桶装乙腈加入至储罐进行淋洗，通过气动隔膜泵TJ4S-1200-PDP029和转移管路139788-step01-P01放出至废液罐。":
        "Through spray ball valve TJ4S-1213-Sy12-V07 on the mixing tank, add 77.5±2.5 kg of drummed acetonitrile "
        "into the tank with the dedicated fresh solvent pump for rinsing, and discharge to the waste tank through "
        "pneumatic diaphragm pump TJ4S-1200-PDP029 and transfer line 139788-step01-P01.",
    "通过搅拌罐上淋洗球阀门TJ4S-1213-Sy12-V07，使用纯化水专用软管及专用质量流量计，向储罐内加入105±5kg 纯化水淋洗储罐内表面，通过气动隔膜泵TJ4S-1200-PDP029和转移管路139788-step01-P01放出至废液罐。":
        "Through spray ball valve TJ4S-1213-Sy12-V07 on the mixing tank, add 105±5 kg of purified water into the "
        "tank via the dedicated purified water hose and dedicated mass flowmeter to rinse the inner surface of the "
        "tank, and discharge to the waste tank through pneumatic diaphragm pump TJ4S-1200-PDP029 and transfer line "
        "139788-step01-P01.",
    "通过阀门TJ4S-1213-Sy12-V140，使用氮气吹扫管线139788-step01-P20，至管路中无液体":
        "Purge line 139788-step01-P20 with nitrogen through valve TJ4S-1213-Sy12-V140 until no liquid remains in the line",
    "配件名称": "Accessory name",
    "配方运行加料数据详见：": "For recipe charging data, see:",
    "配方运行流程如下：": "The recipe runs as follows:",
    "配邻羟基苯甲腈乙酸溶液": "Preparation of 2-hydroxybenzonitrile acetic acid solution",
    "采用洁净擦拭布充分蘸取纯化水，对蝶阀外部阀体、阀杆、连接口处、手柄部位进行全方位擦拭清洗；直至目视无物料残留，":
        "Thoroughly wipe the butterfly valve body exterior, stem, connections and handle with a clean wipe well "
        "soaked in purified water until no material residue is visible,",
    "重复 1–5 操作1次，总共2次": "Repeat steps 1–5 once, 2 times in total",
    "重复 h–k操作2次，总共3次": "Repeat steps h–k twice, 3 times in total",
    "釜下口TJ4S-1213-R23-V47到泵1213-R23-P01之间管路":
        "Line between reactor bottom outlet TJ4S-1213-R23-V47 and pump 1213-R23-P01",
    "釜头氮气": "Reactor head nitrogen",
    "针形阀": "Needle valve",
    "雷达液位计": "Radar level gauge",
    "非专用溶剂管路": "Non-dedicated solvent line",
    "音叉液位计连接处三通": "Tee at tuning-fork level switch connection",
    "项目结束后拆卸大清洁": "Major disassembly cleaning after end of project",
    "预冷DMF": "DMF pre-cooling",
    "预冷DMF进料口": "Pre-cooled DMF inlet",
    "预冷釜": "Pre-cooling reactor",
    "预留加溶剂口": "Reserved solvent addition port",
    "预留液体加料口": "Reserved liquid charging port",
}


def ptext(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def container(p):
    for a in p.iterancestors():
        if a.tag in (W + "tc", W + "body", W + "sdtContent"):
            return a


def has_inline_english(s):
    return bool(re.search(r"[A-Za-z][A-Za-z .&()\-,;:]*\s*/\s*[\u4e00-\u9fff]", s))


def insert_at(p, offset, text):
    pos = 0
    for t in p.iter(W + "t"):
        n = len(t.text or "")
        if pos <= offset < pos + n:
            k = offset - pos
            t.text = (t.text or "")[:k] + text + (t.text or "")[k:]
            t.set(XML_SPACE, "preserve")
            return
        pos += n
    raise ValueError(f"offset {offset} not found in {ptext(p)!r}")


def translate_paragraph(p, s, en):
    if isinstance(en, str):
        en = [(s.strip(), en)]
    # Insert from the end so earlier offsets stay valid.
    spots = []
    start = 0
    for anchor, text in en:
        i = s.index(anchor, start)
        spots.append((i, text + SEP))
        start = i + len(anchor)
    for i, text in reversed(spots):
        insert_at(p, i, text)


def main():
    with zipfile.ZipFile(SRC) as zin:
        tree = etree.fromstring(zin.read("word/document.xml"))
        done, missing = 0, set()
        last_by_cont = {}
        for p in tree.iter(W + "p"):
            s = ptext(p)
            if not s.strip():
                continue
            c = container(p)
            prev = last_by_cont.get(id(c))
            last_by_cont[id(c)] = s
            key = s.strip()
            if not HAN.search(s) or key in SKIP or has_inline_english(s):
                continue
            if prev is not None and prev.strip().endswith("/") and re.search("[A-Za-z]", prev):
                continue
            if key not in T:
                missing.add(key)
                continue
            translate_paragraph(p, s, T[key])
            done += 1
        assert not missing, sorted(missing)
        xml = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        OUT.parent.mkdir(exist_ok=True)
        with zipfile.ZipFile(OUT, "w") as zout:
            for item in zin.infolist():
                data = xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    print(f"{done} paragraphs translated -> {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
