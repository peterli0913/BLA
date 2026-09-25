#!/usr/bin/env python3
"""Lilly 45-min call (09/25/2026 evening): run sheet + bilingual talking script as Word."""

import re

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT = "/workspace/deliverables/2026-09-25_Lilly客户会_45分钟讲稿与流程_中英对照.docx"

NAVY = RGBColor(0x1B, 0x4F, 0x72)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
INK = RGBColor(0x1C, 0x28, 0x33)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
TBD_RED = RGBColor(0xC0, 0x39, 0x2B)
ROW_ALT = "F4F8FB"
HEAD_BG = "1B4F72"
NOTE_BG = "FEF5E7"
SECTION_BG = "EAF2F8"

TBD_PATTERN = re.compile(r"(【待定[^】]*】|\[TBD[^\]]*\]|（需核实）|\(to be confirmed\)|需核实)")


def set_run_font(run, name="Arial", size=10.5, bold=False, color=INK, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), "微软雅黑")


def add_rich(p, text, *, size=10.5, bold=False, color=INK):
    """Add text, rendering 【待定】/[TBD]/需核实 markers in bold red."""
    for part in TBD_PATTERN.split(text):
        if not part:
            continue
        if TBD_PATTERN.fullmatch(part):
            set_run_font(p.add_run(part), size=size, bold=True, color=TBD_RED)
        else:
            set_run_font(p.add_run(part), size=size, bold=bold, color=color)


def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell, color="D5D8DC"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def cell_text(cell, text, *, bold=False, size=10, color=INK, center=False, fill=None):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    if fill:
        shade(cell, fill)
    set_cell_border(cell)
    lines = text.split("\n")
    p = cell.paragraphs[0]
    p.clear()
    for i, line in enumerate(lines):
        if i:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.2
        add_rich(p, line, size=size, bold=bold, color=color)


def set_col_widths(table, widths_cm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for idx, w in enumerate(widths_cm):
        table.columns[idx].width = Cm(w)
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            row.cells[idx].width = Cm(w)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def add_table(doc, header, rows, widths, *, size=10, first_col_bold=False):
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        cell_text(t.rows[0].cells[i], h, bold=True, size=size, color=WHITE, fill=HEAD_BG)
    repeat_header(t.rows[0])
    for r_idx, row in enumerate(rows, start=1):
        fill = ROW_ALT if r_idx % 2 == 0 else None
        no_split(t.rows[r_idx])
        for c_idx, val in enumerate(row):
            cell_text(t.rows[r_idx].cells[c_idx], val, size=size, fill=fill,
                      bold=first_col_bold and c_idx == 0)
    set_col_widths(t, widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:color"), "C9A227")
    bottom.set(qn("w:space"), "2")
    bdr.append(bottom)
    pPr.append(bdr)
    set_run_font(p.add_run(text), size=14, bold=True, color=NAVY)


def add_h2(doc, text, meta=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    set_run_font(p.add_run(text), size=12, bold=True, color=NAVY)
    if meta:
        m = doc.add_paragraph()
        m.paragraph_format.space_after = Pt(4)
        m.paragraph_format.keep_with_next = True
        add_rich(m, meta, size=9.5, color=MUTED)


def add_body(doc, text, *, size=10.5, bold=False, color=INK, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    add_rich(p, text, size=size, bold=bold, color=color)
    return p


def add_bullet(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.2
    add_rich(p, text, size=size)


def add_note(doc, text):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_text(t.rows[0].cells[0], text, size=9.5, fill=NOTE_BG)
    set_col_widths(t, [17.6])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_script(doc, pairs):
    """Side-by-side CN | EN script; one row per spoken beat."""
    t = doc.add_table(rows=1 + len(pairs), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_text(t.rows[0].cells[0], "中文讲稿", bold=True, color=WHITE, fill=HEAD_BG)
    cell_text(t.rows[0].cells[1], "English script", bold=True, color=WHITE, fill=HEAD_BG)
    repeat_header(t.rows[0])
    for i, (cn, en) in enumerate(pairs, start=1):
        no_split(t.rows[i])
        if cn.startswith("§"):
            cell_text(t.rows[i].cells[0], cn[1:], bold=True, size=10, color=NAVY, fill=SECTION_BG)
            cell_text(t.rows[i].cells[1], en[1:], bold=True, size=10, color=NAVY, fill=SECTION_BG)
            continue
        cell_text(t.rows[i].cells[0], cn, size=10.5)
        cell_text(t.rows[i].cells[1], en, size=10.5)
    set_col_widths(t, [8.4, 9.2])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_page_number(paragraph):
    run = paragraph.add_run()
    for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tag)
            run._r.append(fc)
        else:
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = text
            run._r.append(instr)
    set_run_font(run, size=8.5, color=MUTED)


# ---------------------------------------------------------------- content

RUN_SHEET = [
    ["1", "0:00–0:04", "4", "开场、议程、Tracker 总体状态\nOpening, agenda, overall status",
     "Tracker（EN）首屏 / 筛选状态列", "全部"],
    ["2", "0:04–0:19", "15", "Tracker 重点项（4 个主题，按 P 列分组、组内按行号）\nKey tracker items (4 themes)",
     "Tracker（EN），按 P 列筛选", "A 清洁: 2, 12, 13, 29, 34, 35, 38, 42, 43, 44\nB 设备改造: 5, 6, 7, 36, 37, 54, 55\nC 物流通道 + CNC 湿度: 10, 22, 23, 30\nD 钝化 / CCS / 过滤: 20, 21 / 24, 32 / 27, 28, 49"],
    ["3", "0:19–0:26", "7", "湿氮气装置使用策略 + 2D/3D 确认\nHumidified N2 strategy + 2D/3D",
     "Humidified_Nitrogen_..._EN.pptx 第 1–3 页", "37, 54"],
    ["4", "0:26–0:34", "8", "纯水使用方案 + PPQ 后改造提议\nPurified water plan + post-PPQ proposal",
     "Purified_Water_Use_Plan_EN.pptx 第 1–6 页", "纯水 PW: 11, 25, 26, 39（第 5 页）/ 33, 40, 41, 52（第 6 页）/ 47"],
    ["5", "0:34–0:40", "6", "其余项快速过（按 P 列：DIPEA、其他）\nRemaining items by category",
     "Tracker（EN），按 P 列筛选", "DIPEA: 15, 16, 17, 18, 19\n其他 Other: 3, 4, 8, 9, 14, 31, 45, 46, 48, 50, 51, 53, 56"],
    ["6", "0:40–0:45", "5", "待礼来确认的问题 + 下一步\nDecisions needed + next steps",
     "本讲稿第四部分（可直接共享）", "—"],
]

CHECKLIST = [
    ["1", "把讲稿中所有【待定】日期填好（汇总见第六部分）；尤其是 2D/3D 两个方案的完成时间（Tracker 第 54 行 Step 3A 现写 11/18，内部认为难以实现）、CNC 物流通道时间", "□"],
    ["2", "纯水 8 项：Tracker 与纯水 PPT 日期先统一（参考《纯水日期对齐_Tracker与PPT对照_09-25-2026.xlsx》），会上只讲一套日期", "□"],
    ["3", "决定是否向礼来说明“访问后新增罐改造 → 整体计划重排”（讲稿第 2 段有一句可选话术，标了［可选］）", "□"],
    ["4", "打开并按顺序排好：Tracker（EN）→ 湿氮气 PPT（EN）→ 纯水 PPT（EN）→ 本讲稿第四部分", "□"],
    ["5", "Tracker 按 P 列（Asymchem Internal Category）筛选，讲到哪个主题切到哪个分类，组内按行号往下走；隐藏列 A（Number）如需报 Item 号可临时显示", "□"],
    ["6", "请工程（坡度、2D/3D、隔膜阀）、纯水、QA（清洗验证）同事会上待命，追问时可请其补充", "□"],
]

SEG1 = [
    ("感谢大家今晚的时间。今天大约 45 分钟。",
     "Thanks everyone for joining tonight. We have about 45 minutes."),
    ("今天以回应上次现场访问的 Action Tracker 为主，重点讲每一项的措施和时间点。",
     "The main goal today is to walk you through our responses to the site-visit action tracker — the actions and the timelines."),
    ("然后用两份简短的 PPT 补充两个专题：湿氮气装置的使用策略，和纯水的使用方案。",
     "Then we'll use two short decks for two topics: the humidified nitrogen system, and the purified water use plan."),
    ("最后留 5 分钟，列出几个需要礼来确认的问题。",
     "We'll keep the last five minutes for a few decisions we need from Lilly."),
    ("先说总体状态：Tracker 一共 55 项，5 项已完成，1 项进行中，其余 49 项都已明确负责人，正在推进。",
     "First, the overall status. There are 55 items in total. Five are completed, one is in progress, and the other 49 all have owners and are being worked on."),
    ("时间方面，我们不是给一个统一的完成日期，而是按工艺步骤给：Step 1、Step 2、Step 3A、Step 3B。这样每一步开工前该完成的内容更清楚。",
     "On timing — instead of one single completion date, we're giving dates by process step: Step 1, Step 2, Step 3A and Step 3B. That makes it clear what has to be done before each step starts."),
    ("有少数几项的时间还在根据最新的改造范围确认，今天会标出来，确认后在 Tracker 里更新。",
     "A few dates are still being confirmed against the latest modification scope. I'll flag those today, and we'll update the tracker once they're confirmed."),
]

SEG2 = [
    ("§主题 A：清洁 Cleaning（第 2、12、13、29、34、35、38、42、43、44 行，约 5 分钟）",
     "§Theme A: Cleaning (Items 1, 11, 12, 28, 33, 34, 37, 41, 42, 43 — about 5 min)"),
    ("先讲清洁，一共 10 项，按顺序过。",
     "Let's start with cleaning. There are ten items, and I'll go in order."),
    ("第 2 行｜清洗验证与确认策略：按步骤提交——Step 1 在 10/03，Step 2 在 10/20，Step 3A 在 11/09，Step 3B 在 11/28。我们理解礼来需要在生产可销售批之前，对我们的清洁能力有信心。",
     "Item 1 — cleaning validation and verification strategy. We'll submit it step by step: Step 1 by October 3rd, Step 2 by October 20th, Step 3A by November 9th, and Step 3B by November 28th. We understand you need confidence in our cleaning before we make saleable batches."),
    ("第 12 行｜专用产品接触软管和溶剂软管的生命周期管理：修订 SOP，明确相应要求，包括不论批次长短都定期目视检查。11/30。",
     "Item 11 — lifecycle of dedicated product-contact and solvent hoses. We'll revise the SOP to define the requirements, including regular visual inspection regardless of campaign length. By November 30th."),
    ("第 13 行｜溶剂和工艺液转移软管的使用前清洁：现在新软管首次安装后用甲醇、水冲洗，再氮气吹干。按礼来建议，Step 3 Demo 前改用工艺溶剂冲洗一次，并记录在相应记录里。时间【待定】。",
     "Item 12 — pre-use cleaning of transfer hoses. Today, new hoses are rinsed with methanol, then water, and dried with nitrogen after first installation. As you recommended, before the Step 3 demo we'll rinse once with process solvent and record it. Date [TBD]."),
    ("第 29 行｜手套箱清洁：水擦之后增加酒精擦拭，避免残留水分。10/04。",
     "Item 28 — glovebox cleaning. After the water wipe, we add an alcohol wipe so no water is left behind. By October 4th."),
    ("第 34 行｜第 4 类系统 PD 泵进出口低点排净，清洁前后从这些点排净并用氮气吹扫：这个能力已经具备。",
     "Item 33 — low-point drains at the inlet and outlet of PD pumps in Category 4 systems, with draining and nitrogen purge before and after cleaning. That's already in place."),
    ("第 35 行｜PD 泵清洁程序：目前用淋洗取样做残留检测。SOP 增加擦拭取样，考虑大面积软部件的回收率和有代表性的擦拭位置。10/04。",
     "Item 34 — PD pump cleaning program. Today we use rinse sampling for residues. We'll add swab sampling to the SOP, covering recovery for the large soft parts and a representative swab location. By October 4th."),
    ("第 38 行｜生产前来不及整改的设备和管路差距：在清洁文件里逐项建立对应关系，包括人工清洁、重点检查，以及死角的清洁或排净要求。时间同第 2 行。",
     "Item 37 — equipment and piping gaps not fixed before production. We'll build direct traceability in the cleaning documents: manual cleaning, targeted inspection, and cleaning or drain instructions for dead legs. Same timeline as Item 1."),
    ("第 42 行｜重复使用设备的切换策略：现在设备拆开清洗，回流溶剂由 ARL 做二级检测；活化罐到合成仪的管线用溶剂冲洗后重复使用。我们会准备一份 PPT 给礼来说明，同时评估是否需要擦拭取样。时间【待定】。",
     "Item 41 — changeover strategy for reused equipment. Today, equipment is taken apart for cleaning, and ARL does Level 2 testing on the reflux solvent. The line from the activation tank to the synthesizer is reused after a solvent flush. We'll prepare a presentation for you and assess whether swabbing is needed. Date [TBD]."),
    ("第 43 行｜新设备放行控制（去除加工残留、碎屑和颗粒，微生物负荷可接受）：修订 SOP。时间【待定】。",
     "Item 42 — release controls for new equipment: removing fabrication residues, debris and particles, with acceptable bioburden. We'll revise the SOP. Date [TBD]."),
    ("第 44 行｜喷淋覆盖测试：设备已做喷淋覆盖测试。下一步把确认过的压力、时间、报警限、流量时长，以及难清洁部位的人工清洁或擦拭，写进操作指导。时间【待定】。",
     "Item 43 — spray coverage testing. It's been done on the equipment. Next, we'll put the qualified pressure, time, alarm limits and flow durations, plus manual cleaning or swabbing of hard-to-clean areas, into the operating instructions. Date [TBD]."),

    ("§主题 B：设备改造 Equipment modification（第 5、6、7、36、37、54、55 行，约 4 分钟）",
     "§Theme B: Equipment modification (Items 4, 5, 6, 35, 36, 53, 54 — about 4 min)"),
    ("第 5 行｜管路坡度：C3/C4 的管路已改造安装，坡度正在确认。C1/C2 之前评估为没有坡度要求，只设了排净点，要求是可排净，这一点需要跟礼来重新确认。",
     "Item 4 — line slope. The C3 and C4 lines are modified and installed, and we're verifying the slope now. For C1 and C2, we'd assessed earlier there's no slope requirement — only drainability, with drain points provided. We'd like to re-confirm that with you."),
    ("我们会按工艺步骤一条线一条线排查，每个系统完成后把照片和视频发给礼来。C1/C2 的坡度怎么确认，想听听你们的意见。",
     "We'll check each step and each line. When each system is done, we'll send you photos and videos. For C1 and C2, we'd like your input on how to confirm the slope."),
    ("第 6 行｜视镜灯缺失：全线排查，缺的补采购，11/30（PPQ 后项）。正在建检查清单；过渡期按礼来建议用手电筒通过视镜目视检查。",
     "Item 5 — missing sight-glass lights. We're reviewing the whole train and buying what's missing, by November 30th — this is a post-PPQ item. We're building an inspection checklist, and in the meantime we use flashlights through the sight glasses, as you suggested."),
    ("第 7 行｜盐酸加料插底管：调整插底管长度，提供材质证明，确认具体合金牌号。10/01。",
     "Item 6 — HCl charging dip tube. We'll adjust the length and provide the material certificate, including the exact alloy. By October 1st."),
    ("第 36 行｜PD 泵长期替换：卫生级离心泵正在采购，PPQ 后安装，01/30/2027。过渡期每次用完两端排空、吹干，会提供详细操作程序。",
     "Item 35 — long-term PD pump replacement. Hygienic centrifugal pumps are being procured and will be installed after PPQ, by January 30th, 2027. Until then, both ends are drained and blown dry after each use, and we'll provide the detailed procedure."),
    ("第 37 行｜设备/管路标准差距（球阀、长接管、挂钩插底管、低点缺排净、仪表和流量计、法兰、螺栓连接桨叶、死角等）：重新设计正在进行。设计确认后，关键项先跟礼来沟通再执行。",
     "Item 36 — gaps against equipment and piping standards: ball valves, long nozzles, hook-supported dip tubes, missing low-point drains, instruments and flow meters, flanges, bolted agitator blades, dead legs and so on. The redesign is in progress. Once it's confirmed, we'll review the key items with you before we execute."),
    ("有两点需要礼来：一是物料管线排净设计按 2D 还是 3D，二是部分罐的改造方案已确认，请礼来帮忙确认。各步骤完成时间【待定】。",
     "Two things we need from you: first, 2D or 3D for drainability on the material lines; second, some tank modification plans are ready, and we'd like your team to confirm them. Dates by step: [TBD]."),
    ("［可选］访问后我们把几台罐的改造也加进了范围，所以整体计划在重排，更新后的计划【待定】给到。",
     "[Optional] After your visit we added some vessel modifications to the scope, so we're re-baselining the overall schedule. We'll share the updated plan by [TBD]."),
    ("第 54 行｜隔膜阀安装角度：逐条排查含隔膜阀的管线，不合规的整改，跟管路改造一起做（现场进度需核实）。时间取决于 2D/3D，【待定】。这一点等会儿在湿氮气 PPT 里一起说。",
     "Item 53 — diaphragm valve angles. We're walking down every line with diaphragm valves and correcting any that are wrong, together with the piping work (to be confirmed). Timing depends on 2D or 3D — [TBD]. I'll come back to this in the nitrogen deck."),
    ("第 55 行｜有代表性的取样能力：在 100 L 配液罐上加装取样阀。时间【待定】。",
     "Item 54 — representative sampling. We'll install a sampling valve on the 100 L prep vessel. Date [TBD]."),

    ("§主题 C：物流通道 Material logistics corridor + CNC 湿度 CNC humidity（第 10、22、23、30 行，约 2.5 分钟）",
     "§Theme C: Material logistics corridor + CNC humidity (Items 9, 21, 22, 29 — about 2.5 min)"),
    ("第 10 行｜12 车间 CNC 原料暂存区缺少缓冲区和压差联锁：短期方案是提交变更，改造厂区物流通道，原料进 CNC 前增加一个气闸间。完成时间【待定】。",
     "Item 9 — the CNC raw material staging area in Workshop 12 has no buffer zone or pressure interlock. The short-term fix: a change control to modify the logistics corridor and add an airlock before materials go into the CNC area. Completion date [TBD]."),
    ("Tracker 写的是 Pre-PPQ 前完成，但 Step 1 在 10/05 就开始，时间上来不及。我们建议以 Step 3 Pre-PPQ 作为完成节点，请礼来确认。",
     "The tracker says before Pre-PPQ, but Step 1 starts on October 5th, so we can't make that. We propose Step 3 Pre-PPQ as the milestone. We'd like Lilly to confirm."),
    ("第 22 行｜原料接收和转入 CNC 的管理措施：现在规定物料到车间后由专人转入 CNC。SOP 补充物料全流程管理，避免转运中出现空档。10/20。",
     "Item 21 — admin controls for receiving materials and moving them into CNC. Today, a dedicated operator moves the material into the CNC once it arrives. We'll add full material-flow management to the SOP so there are no gaps. By October 20th."),
    ("第 23 行｜长期的气闸或物理隔离（PPQ 后项）：提交变更，增加缓冲间。变更 10/20 提交。",
     "Item 22 — long-term airlock or physical segregation, a post-PPQ item. We'll raise a change control to add a buffer room, by October 20th."),
    ("第 30 行｜CNC 区夏季湿度低于 60%（PPQ 后项）：空调机组加冷却盘管，在生产间隙完成施工和确认，受影响房间需要停产，05/30/2027。过渡期按礼来建议用移动除湿机，相应程序【待定】（需核实）。",
     "Item 29 — keeping CNC humidity below 60% in summer, a post-PPQ item. We'll add cooling coils to the AHUs and qualify during a production gap. The affected rooms have to stop, so the date is May 30th, 2027. In the short term we'll use portable dehumidifiers, as you suggested. The procedure is [TBD] (to be confirmed)."),

    ("§主题 D：钝化 Passivation + CCS + 过滤 Filtration strategy（第 20、21 / 24、32 / 27、28、49 行，约 3.5 分钟）",
     "§Theme D: Passivation + CCS + Filtration strategy (Items 19, 20 / 23, 31 / 26, 27, 48 — about 3.5 min)"),
    ("第 20 行｜GG917 和 LAARA 工艺管线钝化：L1、L2、L3 的 Step 1 都已完成，Step 2 在 09/29；L3 的 Step 3 在 10/07。GG917 的 Step 1 和 Step 3 已完成，Step 2 在 10/02。",
     "Item 19 — passivation of the GG917 and LAARA process lines. Step 1 is done for L1, L2 and L3, and Step 2 is on September 29th. L3 Step 3 is October 7th. For GG917, Steps 1 and 3 are done, and Step 2 is October 2nd."),
    ("第 21 行｜溶剂储罐和总管钝化计划：一期罐区暂不钝化，用金属离子检测来控制。二期罐区交付后，管线和储罐钝化，之后所有项目都用二期罐区的溶剂。02/17/2027。",
     "Item 20 — passivation plan for bulk solvent tanks and headers. The Phase I tank farm won't be passivated for now; we control it with metal-ion testing. Once the Phase II tank farm is handed over, its lines and tanks will be passivated, and all later projects will use Phase II solvents. By February 17th, 2027."),
    ("第 24 行｜全厂多产品交叉污染风险评估：CCS 文件缺少详细描述。在各步骤的 CCS 里补充风险和控制措施的汇总。10/20。",
     "Item 23 — site multiproduct cross-contamination risk assessment. Our CCS documents don't have enough detail, so we'll add a summary of risks and mitigations to each step's CCS. By October 20th."),
    ("第 32 行｜工艺连接和公用管线的交叉污染风险，包括 GG917 和 LAARA 的共用公用系统和设备隔离：在 CCS 里补充详细评估。10/04。",
     "Item 31 — cross-contamination risk from process connections and utility lines, including shared utilities and equipment isolation for GG917 and LAARA. We'll add a detailed assessment to the CCS. By October 4th."),
    ("第 27 行｜各步骤过滤策略（位置、类型、更换标准、依据、图表），与礼来公司过滤指南一致：更新 SOP 和 CCS。10/04。",
     "Item 26 — the filtration strategy for each step — location, type, replacement criteria, justification, diagrams — in line with your corporate guidance. We'll update the SOP and CCS by October 4th."),
    ("第 28 行｜所有步骤的液体原料过滤，以及 SPPS 原料桶加料用 75 微米过滤的依据：在 SOP 和 CCS 里补充。10/04。",
     "Item 27 — filtration of liquid raw materials for all steps, and the justification for 75-micron filtration when charging SPPS drums. We'll add it to the SOP and CCS by October 4th."),
    ("第 49 行｜过滤器使用前后完整性测试策略：我们有内部程序，10/20 前给礼来说明。",
     "Item 48 — pre-use and post-use filter integrity testing. We have an internal procedure, and we'll present it to you by October 20th."),
]

SEG3 = [
    ("§第 1 页：总体使用策略", "§Slide 1: Overall operating strategy"),
    ("接下来看湿氮气装置。第一页是总体策略，分三部分。",
     "Next, the humidified nitrogen system. Slide 1 is the overall strategy, in three parts."),
    ("第一，首次使用：先用碱液循环，再用纯水循环，然后氮气吹扫、低点排净。接着用无菌袋和无菌接头，用注射用水循环并喷淋冲洗，低点排净。最后再用无菌袋加注射用水，投入使用。",
     "One, first use. We recirculate alkaline solution, then purified water, then purge with nitrogen and drain at the low points. Next, using a sterile bag and an aseptic connector, we recirculate and spray-rinse with WFI, and drain again. Finally, we charge fresh WFI through a sterile bag, and it's ready to use."),
    ("第二，定期更换注射用水：低点排净，加新鲜注射用水循环冲洗，再排净，重新加水使用。更换频次【待定】（需核实）。",
     "Two, periodic WFI replacement. Drain at the low points, add fresh WFI, recirculate to rinse, drain, then recharge. The interval is [TBD] (to be confirmed)."),
    ("第三，两台加湿器交替使用：在换水周期内，两台按第二部分的方法定期循环清洗、交替使用。",
     "Three, two humidifiers used alternately. Within the replacement interval, both units are cleaned by recirculation, as in part two, and used in turn."),
    ("§第 2 页：低点排净和高点吹扫位置（前段）", "§Slide 2: Low-point drain and high-point purge locations (upstream)"),
    ("第二页标了前段的排净和吹扫点：进洁净区的物料管线设了清洗/吹扫口，可以双向清洗和吹扫；过滤器离线清洗；有注射用水预冲洗的排放管；设备低点有排净管。",
     "Slide 2 shows the upstream drain and purge points. The material line into the clean area has a cleaning and purge port, so we can clean and purge in both directions. Filters are cleaned off-line. There's a drain line for the WFI pre-flush, and a low-point drain on the unit."),
    ("§第 3 页：去三合一的管线 + 2D/3D", "§Slide 3: Line to the AFD + 2D vs 3D"),
    ("第三页是去三合一的湿氮气管线：同样设清洗/吹扫口，可双向清洗吹扫，管线低点设排净。",
     "Slide 3 is the humidified nitrogen line to the AFD. It also has a cleaning and purge port for both directions, and a low-point drain on the line."),
    ("这里有个问题要请礼来确认：排净点和吹扫点的阀门，按 2D 还是 3D 死角要求安装？",
     "Here's one decision we need from you: should the valves at the drain and purge points meet a 2D or a 3D dead-leg requirement?"),
    ("2D：阀门采购周期更长（内部口径约 6 周），完成时间【待定】。",
     "2D: the valves take longer to procure — about six weeks. Completion: [TBD]."),
    ("3D：阀门采购周期短，国内设备供应商确认 3D 设计可以有效清洁，完成时间【待定】。",
     "3D: shorter lead time, and our domestic equipment suppliers have confirmed a 3D design can be cleaned effectively. Completion: [TBD]."),
    ("这个选择也会影响刚才说的管路改造、隔膜阀调整的时间，纯水改造方案里的阀门也是按 3D 设计的。",
     "This choice also drives the piping and diaphragm valve timing I mentioned. And the purified water design, which we'll see next, also assumes 3D valves."),
]

SEG4 = [
    ("§开场", "§Intro"),
    ("下面是纯水。这部分对应 Tracker 里的纯水项：第 11、25、26、33、39、40、41、47、52 行。",
     "Now purified water. This covers the PW items in the tracker: Items 10, 24, 25, 32, 38, 39, 40, 46 and 51."),
    ("前四页是纯水总管改造后的用水方案，后两页是每一项的措施和时间。",
     "The first four slides show how water will be used after the header modification. The last two list the actions and dates for each item."),
    ("§第 1 页：排废与供水流程", "§Slide 1: Drain and water-fill flow"),
    ("用水点通过软管接到固定管线上。先走路径①排废冲洗，冲洗完再走路径②给设备供水。",
     "The PW outlet connects to a fixed line with a hose. We flush to drain first — path one — and then fill the equipment — path two."),
    ("§第 2–3 页：取样与吹扫", "§Slides 2–3: Sampling and purge"),
    ("取样也是先排废，再取样。用完后氮气先吹①，再吹②。",
     "Sampling is the same idea: drain first, then sample. After use, we purge path one first, then path two, with nitrogen."),
    ("§第 4 页：操作程序与过滤器管理", "§Slide 4: Operating procedure and filter management"),
    ("用水：接软管、开阀、冲洗排废，再给用水点供水；然后开氮气阀，先向设备侧吹扫，再反向吹回用水口，最后关闭所有阀门。",
     "For water use: connect the hose, open the valve, flush to waste, then supply the point of use. Then open the nitrogen valve, purge toward the equipment, then back toward the PW outlet, and close all valves."),
    ("取样：同样先冲洗排废，再人工取样。过滤器每次使用前安装，使用时间不超过 24 小时。",
     "For sampling: flush to waste first, then take a manual sample. Filters are installed before each use, and each filter is used for no more than 24 hours."),
    ("§第 5 页：行动项（第 11、25、26、39 行；日期按会前统一后的版本讲）",
     "§Slide 5: Items 10, 24, 25, 38 (use the aligned dates)"),
    ("第 11 行｜专用纯水软管一个月寿命的依据：实际做法是每周清洁、每六个月更换，会做使用/消毒周期研究来支持。时间【待定】。",
     "Item 10 — basis for the one-month life of dedicated PW hoses. Our practice is weekly cleaning and replacement every six months, and we'll support it with a use and sanitization study. Date [TBD]."),
    ("第 25 行｜硅胶水管换 PTFE：国产交期约 3 周，进口约 8 周。时间【待定】。",
     "Item 24 — replacing silicone water hoses with PTFE. Domestic lead time is about three weeks, imported about eight. Date [TBD]."),
    ("第 26 行｜用后立即断开水管并受控存放：更新程序文件。时间【待定】。",
     "Item 25 — disconnecting water hoses right after use and storing them properly. We'll update the procedure. Date [TBD]."),
    ("第 39 行｜纯水总管改造（重点）：提交变更，按第 1 页的方案改造，阀门按 3D 设计。周期包括采购、施工、钝化、验证，完成时间【待定】。",
     "Item 38 — the PW header modification, the key one. We'll raise a change control and modify it per Slide 1, with valves designed to 3D. The timeline covers procurement, installation, passivation and validation. Completion: [TBD]."),
    ("在 Pre-PPQ 前完成比较紧。我们建议总管改造放在 PPQ 之后，PPQ 前按现行方式用水，同时执行先冲洗排废、24 小时过滤器这些措施（过渡方案细节需核实）。",
     "Finishing this before Pre-PPQ is tight. Our proposal is to do the header work after PPQ. Before that, we keep the current practice, plus the interim controls — flush to waste before use, and the 24-hour filters (interim details to be confirmed)."),
    ("这样不影响 Pre-PPQ，同时控制风险。请礼来确认这个方向。",
     "That way it doesn't block Pre-PPQ, and the risk stays controlled. We'd like Lilly to confirm this approach."),
    ("§第 6 页：行动项（第 33、40、41、52 行）+ 第 47 行", "§Slide 6: Items 32, 39, 40, 51 + Item 46"),
    ("第 33 行｜软管晾干存放时端口用透气膜覆盖：透气膜已采购。时间【待定】。",
     "Item 32 — covering hose ends with permeable film while they dry. The film is already purchased. Date [TBD]."),
    ("第 40 行｜用水点过滤器：立即纠正，每次使用前安装，使用时间不超过 24 小时，要求写进 MOR。时间【待定】。",
     "Item 39 — point-of-use filters. An immediate corrective action: installed before each use, no more than 24 hours in use, and written into the MOR. Date [TBD]."),
    ("第 41 行｜新用水点排放口取样的代表性，以及是否需要正式再验证：评估进行中。时间【待定】。",
     "Item 40 — representative sampling from the new point-of-use drains, and whether formal revalidation is needed. Assessment in progress. Date [TBD]."),
    ("第 52 行｜软管端离地至少 6 英寸，并有明显标识：短期培训现场人员，长期装不锈钢软管架。时间【待定】。",
     "Item 51 — hose ends at least six inches off the floor, with a clear visual indicator. Short term, we train the operators; long term, we install stainless-steel hose racks. Date [TBD]."),
    ("第 47 行（PPT 里没有）｜移动容器放置时间研究，确认可完全排净，并写进操作文件：时间【待定】。礼来提到即使研究支持更长，也按 48 小时刷新，我们会在方案里考虑。",
     "Item 46, not in the deck — hold-time studies for portable vessels, confirming full drainability, and adding it to the operating documents. Date [TBD]. We note your point on a 48-hour refresh limit, and we'll take it into account."),
]

SEG5 = [
    ("§DIPEA（第 15–19 行，约 2 分钟）：4 项已完成，1 项进行中", "§DIPEA (Items 14–18, about 2 min): four completed, one in progress"),
    ("第 15 行｜乙醛低于 50 ppm 的 DIPEA 供应商：已采购并收到供应商样品，初步检测低于 50 ppm。等分析方法跟礼来对齐后正式放行，09/30。进行中。",
     "Item 14 — a reliable supplier of DIPEA with less than 50 ppm acetaldehyde. We've bought and received supplier samples, and preliminary results are below 50 ppm. Final release waits for us to align the analytical method with you. Target September 30th. In progress."),
    ("第 16 行｜内部处理 DIPEA 降乙醛：小试有一定效果，但决定直接向供应商采购低乙醛 DIPEA，不在凯莱英内部精制。已完成。",
     "Item 15 — in-house DIPEA treatment to reduce acetaldehyde. Lab trials showed some effect, but we'll buy low-acetaldehyde DIPEA directly from the supplier, with no further purification here. Completed."),
    ("第 17 行｜缩短洗涤中 DIPEA 暴露时间：实验证明洗涤时间长不会导致 DIPEA 分解；与 RS 批相比，暴露时间差异主要来自小包装频繁换桶。改用 200 L 大桶。已完成。",
     "Item 16 — shorter DIPEA exposure during washes. Our data show long washes don't decompose DIPEA. Versus the RS batch, the difference came from small packages and frequent drum changes, so we're moving to 200-liter drums. Completed."),
    ("第 18 行｜SC 偶联提高氮气鼓泡流量降 HCN：22 kg 批量下 17 m³/h 已证明有效，流量可调到 30 m³/h，相关文件同步更新。已完成。",
     "Item 17 — higher nitrogen sparge for the SC coupling to reduce HCN. 17 cubic meters per hour is proven effective at the 22-kilogram scale, and it can go up to 30. We'll update the documents. Completed."),
    ("第 19 行｜SC 偶联氮气吹扫降 HCN：双方已达成一致，不建议吹扫/抽真空。已完成。",
     "Item 18 — nitrogen sweep during the SC coupling. We agreed that sweep or vacuum isn't recommended. Completed."),
    ("§其他 Other（第 3、4、8、9、14、31、45、46、48、50、51、53、56 行，约 4 分钟）",
     "§Other (Items 2, 3, 7, 8, 13, 30, 44, 45, 47, 49, 50, 52, 55 — about 4 min)"),
    ("第 3 行｜爆破片铭牌：已完成。",
     "Item 2 — rupture disc nameplate. Completed."),
    ("第 4 行｜进料管径满足 200 L/h：IEPE 理论计算，加现场测试。09/30。",
     "Item 3 — inlet line sizing for 200 liters per hour. IEPE calculation plus an on-site test. By September 30th."),
    ("第 8 行｜蠕动泵排空软管后读移动秤：MOR/SOP 里目前没有细化要求，更新 MOR。10/03。",
     "Item 7 — reading the mobile scale after the peristaltic pump empties the hose. The MOR and SOP don't have this detail today, so we'll update the MOR. By October 3rd."),
    ("第 9 行｜工艺废液总管无止回阀：评估回流风险，需要的加装止回阀，跟整体管路施工一起做。11/03。",
     "Item 8 — no check valves on the process waste header. We'll assess backflow risk and add check valves where needed, together with the overall piping work. By November 3rd."),
    ("第 14 行｜波纹软管换光滑内壁（参考 Pharmaline N）：现在用铂金硫化硅胶管，计划更换。国产约 3 周，进口约 2 个月。请礼来确认国产是否满足要求。",
     "Item 13 — replacing corrugated hoses with smooth-bore ones, such as Pharmaline N. We use platinum-cured silicone today, and replacement is planned. Domestic takes about three weeks, imported about two months. Please confirm whether domestic hoses are acceptable."),
    ("第 31 行｜尾气总管图纸和设计给礼来审阅：本周内上传。",
     "Item 30 — vent header drawings and designs for your review. Uploaded before the end of this week."),
    ("第 45 行｜访客和支持人员在洁净区穿专用内层（PPQ 后项）：修订洁净区更衣程序。11/10。",
     "Item 44 — a dedicated base layer for visitors and support staff in classified areas, a post-PPQ item. We'll revise the gowning procedure by November 10th."),
    ("第 46 行｜人员进入权限管理（PPQ 后项）：我们有内部程序，11/10 前给礼来说明。",
     "Item 45 — access control, a post-PPQ item. We have an internal procedure and will present it to you by November 10th."),
    ("第 48 行｜确保只有合格流动相进入工艺：配液混合器下游在线监测电导率，合格的进流动相罐，不合格的转废液；质量流量计计量，混合器和接收罐都有 pH 监测。现有措施已覆盖。",
     "Item 47 — making sure only in-spec mobile phase reaches the process. Conductivity is monitored after the mixer: in-spec goes to the tank, out-of-spec goes to waste. Dosing is by mass flow meters, and pH is monitored at the mixer and the receiving tank. Existing controls cover this."),
    ("第 50 行｜三合一搅拌桨行程，置换洗涤时桨叶不被浸没：时间【待定】。",
     "Item 49 — AFD stroke height, so the blades aren't submerged during displacement washes. Date [TBD]."),
    ("第 51 行｜粉碎机操作策略（成品取样、惰化、目标桶重不超装、筛网温度监测、人机工程）：09/27。",
     "Item 50 — mill operating strategy: final sampling, inerting, target drum weight without overfill, screen temperature, and ergonomics. By September 27th."),
    ("第 53 行｜所有固体加料用密闭装置，所有步骤（含 SPPS 和裂解）用分体阀加料：时间【待定】。",
     "Item 52 — contained devices for all solids charging, and split-valve charging for every step, including SPPS and cleavage. Date [TBD]."),
    ("第 56 行｜非伸缩式 pH 探头校准时罐保持封盖，评估换伸缩式护套：更新校准指导（时间【待定】）；伸缩式护套采购周期约 6 周。",
     "Item 55 — keeping tanks capped while non-retractable pH probes are calibrated, and looking at retractable housings. We'll update the calibration instruction, date [TBD]. Retractable housings take about six weeks to procure."),
]

SEG6 = [
    ("最后，有 6 个问题需要礼来确认，我共享一下清单。",
     "To wrap up, there are six decisions we need from Lilly. Let me share the list."),
    ("（逐条读第四部分的表格）",
     "(Read through the table in Part 4.)"),
    ("下一步：我们会在【待定】前把更新后的 Tracker 发给大家，之后每周同步一次进度。管路坡度的照片和视频按系统完成情况陆续发送。",
     "Next steps: we'll send the updated tracker by [TBD], and then give you a weekly progress update. Photos and videos of the line slopes will follow as each system is done."),
    ("谢谢大家，还有什么问题吗？",
     "Thank you. Any other questions?"),
]

DECISIONS = [
    ["1", "湿氮气及管路改造的死角按 2D 还是 3D 做？", "2D or 3D for the dead legs (N2 system and piping)?", "37, 54；湿氮气 PPT 第 3 页"],
    ["2", "C1/C2 管路坡度，礼来希望用什么方式确认？（照片/视频/现场）", "How would Lilly like to confirm slope for C1/C2 — photos, video or on site?", "5"],
    ["3", "CNC 物流通道改造以 Step 3 Pre-PPQ 作为完成节点，是否可以？", "Can Step 3 Pre-PPQ be the milestone for the CNC logistics corridor?", "10"],
    ["4", "光滑内壁软管用国产（约 3 周）是否可以？", "Are domestic smooth-bore hoses (about 3 weeks) acceptable?", "14"],
    ["5", "纯水总管完整改造放在 PPQ 后，PPQ 前用过渡措施，是否可以？", "Can the full PW header modification be done after PPQ, with interim controls before?", "39, 41"],
    ["6", "需要礼来确认的罐改造方案，何时方便评审？", "When can Lilly review the tank modification plans that need your sign-off?", "37"],
]

QA = [
    ("3D 真的能清洁干净吗？依据是什么？",
     "国内供应商确认 3D 可清洁；我们会在清洗验证里确认。具体数据可会后提供（需核实）。",
     "Can a 3D dead leg really be cleaned? What's the basis?",
     "Our domestic suppliers confirm it's cleanable, and we'll confirm it through cleaning validation. We can share the details after the call (to be confirmed)."),
    ("Tracker 第 54 行 Step 3A 写的 11/18 还算数吗？",
     "这个时间取决于 2D/3D 的选择和访问后增加的改造范围，正在重排，【待定】前确认。",
     "Is November 18th for Step 3A (Item 53) still valid?",
     "It depends on the 2D or 3D decision and on the scope we added after your visit. We're re-baselining and will confirm by [TBD]."),
    ("C1/C2 为什么不做坡度？",
     "之前评估 C1/C2 的要求是可排净，没有具体坡度要求，已设排净点；C3/C4 已改造，坡度在确认。坡度怎么确认想听礼来意见，我们会提供照片/视频。",
     "Why no slope for C1 and C2?",
     "We'd assessed that C1 and C2 need drainability, not a specific slope, and drain points are provided. C3 and C4 are modified and the slope is being verified. We'd like your view on how to confirm it, and we'll share photos and videos."),
    ("为什么一期罐区不钝化？",
     "一期罐区用金属离子检测控制；二期罐区交付后钝化，之后项目都用二期溶剂，02/17/2027。",
     "Why not passivate the Phase I tank farm?",
     "We control Phase I with metal-ion testing. Phase II will be passivated after handover, and all later projects will use Phase II solvents — by February 17th, 2027."),
    ("CNC 湿度在 2027 年 5 月前怎么控制？",
     "空调改造需要受影响房间停产；过渡期用移动除湿机，程序【待定】（需核实），会在 Tracker 里补充。",
     "How do you control CNC humidity before May 2027?",
     "The AHU work needs the affected rooms to stop. In the short term we use portable dehumidifiers; the procedure is [TBD] (to be confirmed), and we'll add it to the tracker."),
    ("新用水点排放口取样有代表性吗？需要重新验证吗？",
     "取样前先冲洗排废；代表性以及是否需要正式再验证正在评估（第 41 行），结论【待定】。",
     "Is sampling from the new point-of-use drains representative? Do you need revalidation?",
     "We always flush to waste before sampling. Whether it's representative, and whether formal revalidation is needed, is being assessed — Item 40. Result [TBD]."),
    ("纯水总管为什么要到 2027 年？",
     "包括阀门采购、施工、钝化、验证几个阶段；Pre-PPQ 前完成很紧，所以建议 PPQ 后做，PPQ 前执行冲洗排废和 24 小时过滤器等措施。",
     "Why does the PW header take until 2027?",
     "It includes valve procurement, installation, passivation and validation. Doing it before Pre-PPQ is tight, so we propose after PPQ, with flush-to-waste and 24-hour filters in the meantime."),
    ("移动容器的放置时间能控制在 48 小时内吗？",
     "我们会在放置时间研究方案里考虑礼来的 48 小时要求，研究时间【待定】（需核实）。",
     "Can you limit portable vessel refresh time to 48 hours?",
     "We'll take your 48-hour limit into account in the hold-time study plan. Study timing is [TBD] (to be confirmed)."),
    ("喷淋覆盖参数会固定吗？",
     "会。确认后把压力、流量、时间写进程序并固定（第 44 行）。",
     "Will the spray coverage parameters be fixed?",
     "Yes. After verification, pressure, flow and time will be written into the procedure and fixed — Item 43."),
    ("PD 泵换泵前怎么保证清洁？",
     "每次用完两端排空、吹干，会提供详细操作程序；擦拭取样 SOP 10/04 更新；低点排净已有。",
     "How do you keep the PD pumps clean before replacement?",
     "After each use both ends are drained and blown dry, and we'll provide the detailed procedure. The swab SOP is updated by October 4th, and the low-point drains are already in place."),
    ("pH 护套 6 周能更快吗？",
     "在跟供应商沟通，有更新会同步（需核实）。",
     "Can the six-week pH housing lead time be shortened?",
     "We're working with the supplier and will update you (to be confirmed)."),
    ("湿氮气多久换一次水？",
     "频次【待定】（需核实），会写进 SOP；两台交替使用，换水不影响生产。",
     "How often do you swap the WFI in the N2 humidifier?",
     "The interval is [TBD] (to be confirmed), and it'll be in the SOP. With two units used alternately, water changes don't affect production."),
]

TBD_LIST = [
    ["清洗：喷淋覆盖参数固定", "44", "QA / 工程"],
    ["清洗：Step 3 软管使用前冲洗", "13", "刘林冲"],
    ["清洗：切换策略 / 新设备放行说明", "42, 43", "QA"],
    ["设备差距各步骤时间；访问后新增罐改造、整体计划重排后的交付时间", "37", "工程"],
    ["管路坡度各步骤时间", "5", "工程"],
    ["隔膜阀排查进度；2D、3D 两个方案的完成时间（现写 Step 3A 11/18）；2D 阀门交期约 6 周是否准确", "54；湿氮气 PPT 第 3 页", "工程"],
    ["CNC 物流通道完成时间（Tracker 10/20，内部讨论为 10 月底）", "10", "工程 / 生产"],
    ["CNC 过渡期移动除湿机程序", "30", "生产设备部"],
    ["湿氮气换水频次", "湿氮气 PPT 第 1 页", "生产 / QA"],
    ["纯水 8 项日期（Tracker 与 PPT 统一后再填）", "11, 25, 26, 33, 39, 40, 41, 52", "纯水负责人"],
    ["移动容器放置时间研究（礼来提 48 小时刷新上限）", "47", "研发 / 生产 / QA"],
    ["纯水 PPQ 前过渡措施细节（内部会议提过每个用水点加三通、先排废）", "39, 41", "纯水负责人"],
    ["三合一搅拌桨行程 / 100 L 配液罐取样阀", "50, 55", "郭泽鹏"],
    ["密闭固体加料 / 分体阀加料", "53", "刘林冲"],
    ["pH 探头校准指导更新时间", "56", "生产"],
    ["更新版 Tracker 发给礼来的日期", "—", "项目"],
]


def check_coverage():
    """Every tracker row 2–56 must be named (as 第 N 行) in segments 2, 4 or 5."""
    mentioned = set()
    for cn, _ in SEG2 + SEG4 + SEG5:
        if not cn.startswith("§"):
            mentioned.update(int(n) for n in re.findall(r"第 (\d+) 行", cn))
    missing = sorted(set(range(2, 57)) - mentioned)
    assert not missing, f"tracker rows not covered: {missing}"


def build():
    check_coverage()
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(1.7)
    sec.right_margin = Cm(1.7)
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)

    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run_font(hp.add_run("Project 139788 BLA Retrofit｜内部讲稿 Internal – Do not share"), size=8.5, color=MUTED)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(fp)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    set_run_font(p.add_run("礼来客户会 45 分钟讲稿与流程"), size=20, bold=True, color=NAVY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    set_run_font(p.add_run("Lilly Call – Run Sheet & Bilingual Talking Script"), size=13, bold=True, color=GOLD)

    add_table(doc, ["项目", "内容"], [
        ["日期 / 时长", "2026-09-25 晚（具体时间需核实）｜45 分钟"],
        ["主讲", "本人（单人主讲），工程 / QA / 纯水同事待命补充"],
        ["主线", "回应 Site Visit Action Tracker 的措施与时间点（约 30 分钟）"],
        ["辅助材料", "湿氮气装置使用策略 PPT（EN）、纯水使用方案 PPT（EN）"],
        ["标记说明", "红色【待定】/ [TBD] = 会前需填写；红色“需核实” = 事实未确认，会上谨慎表述"],
    ], [3.6, 14.0], size=10, first_col_bold=True)

    add_note(doc, "讲法原则：①不承诺做不到的日期，按 Step 1 / 2 / 3A / 3B 分步给；②讲“排净”，不把不是 CIP 的内容说成 CIP；"
                  "③每个需要礼来决定的点，讲的时候先点到，最后第 6 段统一确认；④Item 号 = Tracker 行号 − 1（英文稿按 Item 号说）。")

    add_h1(doc, "一、会前准备清单")
    add_table(doc, ["#", "事项", "完成"], CHECKLIST, [0.8, 15.4, 1.4], size=10)

    add_h1(doc, "二、流程与时间分配（45 分钟）")
    add_table(doc, ["#", "时间", "分钟", "内容", "屏幕", "Tracker 行"], RUN_SHEET,
              [0.7, 2.0, 1.1, 5.2, 4.2, 4.4], size=9.5)
    add_note(doc, "覆盖范围：Tracker 第 2–56 行共 55 项全部写进讲稿——第 2 段 28 项、第 4 段 9 项、第 5 段 18 项；"
                  "第 3 段湿氮气 PPT 再次引用第 37、54 行。\n"
                  "控时提示：第 2 段超时时，主题 D 每项只报措施和日期；第 5 段的“其他”可只读行号、事项和日期，"
                  "把时间留给第 6 段的决策确认。")

    add_h1(doc, "三、分段讲稿（中英对照）")
    add_h2(doc, "第 1 段｜开场与总体状态（0:00–0:04，4 分钟）", "屏幕：Tracker（EN）首屏")
    add_script(doc, SEG1)
    add_h2(doc, "第 2 段｜Tracker 重点项（0:04–0:19，15 分钟）",
           "屏幕：Tracker（EN），按 P 列分类筛选，组内按行号往下讲。中文稿标 Tracker 行号，英文稿用 Item 号（= 行号 − 1）。")
    add_script(doc, SEG2)
    add_h2(doc, "第 3 段｜湿氮气装置使用策略（0:19–0:26，7 分钟）",
           "屏幕：Humidified_Nitrogen_System_Operating_Strategy_2026-09-24_EN.pptx")
    add_script(doc, SEG3)
    add_h2(doc, "第 4 段｜纯水使用方案（0:26–0:34，8 分钟）",
           "屏幕：Purified_Water_Use_Plan_EN.pptx。会前先统一 Tracker 与 PPT 的日期，会上只讲一套。")
    add_script(doc, SEG4)
    add_h2(doc, "第 5 段｜其余项快速过（0:34–0:40，6 分钟）",
           "屏幕：Tracker（EN），按 P 列筛选：先 DIPEA，再“其他”，组内按行号。")
    add_script(doc, SEG5)
    add_h2(doc, "第 6 段｜决策确认与下一步（0:40–0:45，5 分钟）", "屏幕：本讲稿第四部分")
    add_script(doc, SEG6)

    add_h1(doc, "四、待礼来确认的问题（可直接共享）")
    add_table(doc, ["#", "问题", "Question for Lilly", "Tracker 行"], DECISIONS,
              [0.7, 6.9, 7.4, 2.6], size=10)

    add_h1(doc, "五、可能的追问与建议回答")
    add_table(doc, ["中文：问题 / 回答", "English: Question / Answer"],
              [[f"问：{q}\n答：{a}", f"Q: {eq}\nA: {ea}"] for q, a, eq, ea in QA],
              [8.4, 9.2], size=10)

    add_h1(doc, "六、【待定】项汇总（会前填写）")
    add_table(doc, ["待定内容", "Tracker 行 / 来源", "找谁确认"], TBD_LIST, [9.6, 4.6, 3.4], size=10)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
