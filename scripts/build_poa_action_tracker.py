#!/usr/bin/env python3
"""Build the POA retrofit action tracker from the 09/21 deployment meeting."""

from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/meetings/POA生物制品项目改造-行动项跟踪.xlsx"

NAVY = "1B4F72"
NAVY2 = "154360"
GOLD = "C9A227"
WHITE = "FFFFFF"
INK = "1C2833"
MUTED = "5D6D7E"
LINE = "D5D8DC"
ALT = "F4F8FB"
OPEN_BG = "FBF8F1"
PROG_BG = "D6EAF8"
BLOCK_BG = "FADBD8"
DONE_BG = "D5F5E3"
HIGH_BG = "FDEBD0"
CARD_OPEN = "1B4F72"
CARD_BLOCK = "922B21"
CARD_DUE = "B9770E"
CARD_OVER = "7B241C"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
thick_bottom = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="medium", color=NAVY),
)

font = Font(name="Arial", size=10, color=INK)
font_b = Font(name="Arial", size=10, bold=True, color=INK)
font_white = Font(name="Arial", size=10, bold=True, color=WHITE)
font_title = Font(name="Arial", size=18, bold=True, color=WHITE)
font_sub = Font(name="Arial", size=10, color="D4E6F1")
font_card_n = Font(name="Arial", size=22, bold=True, color=WHITE)
font_card_l = Font(name="Arial", size=9, bold=True, color="FDEBD0")
font_small = Font(name="Arial", size=8, italic=True, color=MUTED)
font_hdr = Font(name="Arial", size=9, bold=True, color=WHITE)
wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_c = Alignment(wrap_text=True, vertical="center", horizontal="center")
left = Alignment(vertical="center", horizontal="left", wrap_text=True)
center = Alignment(vertical="center", horizontal="center", wrap_text=True)

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_navy2 = PatternFill("solid", fgColor=NAVY2)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_alt = PatternFill("solid", fgColor=ALT)
fill_head = PatternFill("solid", fgColor=NAVY)
fill_hint = PatternFill("solid", fgColor="EBF5FB")
fill_input = PatternFill("solid", fgColor="FFF3CD")

STATUSES = ["未开始", "进行中", "阻塞", "已完成", "取消"]
PRIORS = ["高", "中", "低"]
STREAMS = [
    "制备配液",
    "超滤",
    "沉淀",
    "采购到货",
    "设计方案",
    "QA质量",
    "客户沟通",
    "综合协调",
]

# source: 2026-09-21 POA deployment meeting transcript
ACTIONS = [
    dict(id="A-20260921-01", raised=date(2026, 9, 21), stream="设计方案", origin="王总",
         action="牵头东富龙现场，按设备本体 / 管道 / 仪表 / 阀门 / 泵查漏补缺，更新制备配液与超滤改造须知。",
         done="缺项已写入表，可提交合并会。", owner="志恒", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置。上午对接东富龙，作为 IP 设计总牵头，配合两名管理人员细化。",
         src="2026-09-21 部署会", note="厂名转写或作东福龙 / 多福龙，按东富龙，待核实。"),
    dict(id="A-20260921-02", raised=date(2026, 9, 21), stream="制备配液", origin="王总",
         action="对接东富龙现场人员，完成上午车间勘察。",
         done="厂家到场并走完相关管路。", owner="李主任", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置，上午完成。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-03", raised=date(2026, 9, 21), stream="超滤", origin="王总",
         action="跟点超滤片断现场勘察并记录缺项。",
         done="超滤缺项已记入改造须知表。", owner="张东", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置。张东超滤设备接触少，现场以跟点记录为主。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-04", raised=date(2026, 9, 21), stream="设计方案", origin="王总",
         action="厂家技术配合结束后，梳理超滤 / 制备配液改造事项并提交下午会。",
         done="有可过会的梳理稿。", owner="志鹏", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置。电子 / 智能技术员与张东配合其现场工作。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-05", raised=date(2026, 9, 21), stream="客户沟通", origin="王总",
         action="按沉淀 / 超滤 / 制备三片断，评估客户访问问题是否举一反三全覆盖，写入表单。",
         done="表单能看出跨片断覆盖与缺口。", owner="松涛", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置。避免只改被点名的片断。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-06", raised=date(2026, 9, 21), stream="制备配液", origin="王（IAP）",
         action="通知东富龙曹工 / 赵工上午到场勘察利旧，下午交出利旧与推倒两套方案对比。",
         done="两套方案的事项与时间可对比。", owner="王（IAP，全名待补）", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：会上要求马上打电话；群里已催。下午与内部清单一起看。",
         src="2026-09-21 部署会", note="责任人全名待补。"),
    dict(id="A-20260921-07", raised=date(2026, 9, 21), stream="设计方案", origin="王总",
         action="13:00 与亚光（菊工）合并昨天下午勘察事项与内部清单。",
         done="合并清单一版。", owner="王总", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置。昨日已把陈店 / 沉淀相关区域设备管道勘察一遍。",
         src="2026-09-21 部署会", note="亚光转写或作雅光，待核实。"),
    dict(id="A-20260921-08", raised=date(2026, 9, 21), stream="设计方案", origin="王总",
         action="汉邦到场后梳理事项并出方案，次日早上评审。",
         done="方案可上会。", owner="汉邦对接人（待补）", due=date(2026, 9, 22),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置，下午到场。",
         src="2026-09-21 部署会", note="对接人待补。"),
    dict(id="A-20260921-09", raised=date(2026, 9, 21), stream="设计方案", origin="王总",
         action="三片断厂家方案与原方案查漏补缺，合并为一套改造方案。",
         done="合并方案一版。", owner="志恒", due=date(2026, 9, 22),
         status="未开始", pri="高",
         progress="2026-09-21：部署会明确：三片断方案合并后才算整套改造方案形成。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-10", raised=date(2026, 9, 21), stream="制备配液", origin="王总",
         action="列出配液「利旧改造 vs 推倒重建」对比，单独报领导。",
         done="对比表 + 汇报口径。未选出之前不得写成已定。", owner="王总", due=None,
         status="未开始", pri="高",
         progress="2026-09-21：会上明确未决，方案列出后单独汇报。",
         src="2026-09-21 部署会", note="截止日期待方案列出后指定。"),
    dict(id="A-20260921-11", raised=date(2026, 9, 21), stream="采购到货", origin="方哥/王总",
         action="输出采购三类台账：已采买、询价中、未采购；此后每日更新到货进展。",
         done="物料可点名归入三类之一；每日有更新。", owner="宋金柱", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会布置。田子才、冯毅配合；三人无论谁在场，每天向主持更新到货。",
         src="2026-09-21 部署会", note="单一责任人按宋金柱；配合人见进展。"),
    dict(id="A-20260921-12", raised=date(2026, 9, 21), stream="采购到货", origin="方哥",
         action="核实东富龙 600 搅拌罐采购订货状态。",
         done="订货状态写入台账。", owner="宋金柱", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：部署会指定为金柱今日重点之一。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-13", raised=date(2026, 9, 21), stream="综合协调", origin="方哥",
         action="湿氮气擦净、消毒、推至车间门口，再拆卸转移至 AFD（三合一）背后；中午前完成。",
         done="2026-09-21 13:00 客户进车间前就位并复确认擦洗。", owner="宋金柱", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：泽鹏进场后第一时间擦净推到物流通道，消完毒推到车间门口；其后金柱安排拆卸，转移到 AFD 背后。中午前完成。",
         src="2026-09-21 部署会", note="门口搬运：泽鹏。拆卸转移：宋金柱。术语按项目指定：三合一 = AFD。"),
    dict(id="A-20260921-14", raised=date(2026, 9, 21), stream="综合协调", origin="王总",
         action="到岗后与于总确认工作分工及下一步执行计划。",
         done="分工与计划已对接主持侧。", owner="刘维江", due=date(2026, 9, 23),
         status="未开始", pri="中",
         progress="2026-09-21：王总已电话沟通；会中明确后天到岗。QA 美佳后续与其交底。",
         src="2026-09-21 部署会", note=""),
    dict(id="A-20260921-15", raised=date(2026, 9, 21), stream="制备配液", origin="美佳",
         action="向腾信压缩 D 区现有搅拌罐改口工期；达不成则把完不成节点提到中午沟通会。",
         done="新工期，或正式报出的冲突节点（洪杰 / 美佳 / 张豹可上会）。", owner="美佳", due=date(2026, 9, 21),
         status="阻塞", pri="高",
         progress="2026-09-21：美佳反馈约 5 个罐、工期 20 天（15 日起改到下月 5 日），与 11/18 交付冲突；并判断 20 天也不一定做完。王总/方哥要求继续压，资源不够找郭宏杰。",
         src="2026-09-21 部署会", note="会上先说 C 区、后更正为 D 区现有罐改口，不是新建。位号待补。"),
    dict(id="A-20260921-16", raised=date(2026, 9, 21), stream="采购到货", origin="方哥",
         action="先出既有采购清单，再把客户新加项与厂家新识别改造项补入。",
         done="新旧项可区分，不回头混在一起。", owner="美佳", due=date(2026, 9, 21),
         status="未开始", pri="高",
         progress="2026-09-21：客户刚新加多项；厂家今日仍在梳理，梳理完往清单里补。自控阀门货期可能拖交付。",
         src="2026-09-21 部署会", note="与 A-20260921-11 联动：11 出架子，16 补新项。"),
    dict(id="A-20260921-17", raised=date(2026, 9, 21), stream="综合协调", origin="王总",
         action="组织每日 19:00 改造沟通会（含线上）。",
         done="会议召开，行动项有当日进展。", owner="德鹏", due=date(2026, 9, 21),
         status="进行中", pri="中",
         progress="2026-09-21：部署会确定每晚 19:00；可线上。另议周一晨会半小时，是否固定待确认。",
         src="2026-09-21 部署会", note="截止日期按首场；此后每日重复。"),
    dict(id="A-20260921-18", raised=date(2026, 9, 21), stream="综合协调", origin="李方向",
         action="为张东申请电脑。",
         done="电脑到位。", owner="李方向", due=None,
         status="未开始", pri="低",
         progress="2026-09-21：会上确认张东无电脑，李方向答应申请。",
         src="2026-09-21 部署会", note="截止日期会上未指定。"),
    dict(id="A-20260921-19", raised=date(2026, 9, 21), stream="设计方案", origin="李方向",
         action="在 demo 批次启动前完成设计，并报客户沟通敲定，避免边改边设计。",
         done="客户确认方案；不耽误 10 月中下旬开工。", owner="待指定（生产部 + IPE）", due=date(2026, 10, 20),
         status="未开始", pri="高",
         progress="2026-09-21：李方向强调设计须在 10 月中下旬前搞定。IPE 厂家与生产部整体负责，确认后报客户。",
         src="2026-09-21 部署会", note="截止日期 10/20 为纪要推断，待核实。责任人待落到人名。"),
]


def style_range(cell, fill=None, font_=None, align=None, border=None, num=None):
    if fill:
        cell.fill = fill
    cell.font = font_ or font
    cell.alignment = align or wrap
    cell.border = border if border is not None else thin
    if num:
        cell.number_format = num


def apply_cf_status(ws, col, first, last):
    rng = f"{col}{first}:{col}{last}"
    mapping = {
        "未开始": "D5D8DC",
        "进行中": "AED6F1",
        "阻塞": "F5B7B1",
        "已完成": "ABEBC6",
        "取消": "E5E8E8",
    }
    for val, color in mapping.items():
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="equal",
                formula=[f'"{val}"'],
                fill=PatternFill("solid", fgColor=color),
                font=Font(name="Arial", size=10, bold=True, color=INK),
            ),
        )


def apply_cf_pri(ws, col, first, last):
    rng = f"{col}{first}:{col}{last}"
    for val, color in (("高", "F5B041"), ("中", "F9E79F"), ("低", "D5DBDB")):
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="equal",
                formula=[f'"{val}"'],
                fill=PatternFill("solid", fgColor=color),
                font=Font(name="Arial", size=10, bold=True, color=INK),
            ),
        )


def apply_cf_traffic(ws, col, first, last):
    rng = f"{col}{first}:{col}{last}"
    for val, color in (("逾期", "F5B7B1"), ("今日到期", "FAD7A0"), ("按期", "D5F5E3"), ("—", "F4F6F7")):
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="equal",
                formula=[f'"{val}"'],
                fill=PatternFill("solid", fgColor=color),
                font=Font(name="Arial", size=10, bold=True, color=INK),
            ),
        )


def build():
    wb = Workbook()

    # ----- 总览 -----
    dash = wb.active
    dash.title = "总览"
    dash.sheet_view.showGridLines = False
    dash.page_setup.orientation = "landscape"
    dash.page_setup.fitToPage = True
    dash.page_setup.fitToWidth = 1
    dash.page_setup.fitToHeight = 1
    dash.page_setup.paperSize = dash.PAPERSIZE_A4
    dash.page_margins = PageMargins(0.5, 0.5, 0.6, 0.5)
    dash.oddHeader.left.text = "POA 生物制品改造 · 内部"
    dash.oddFooter.right.text = "Source: 2026-09-21 deployment meeting"

    dash.merge_cells("B2:I3")
    dash["B2"] = "POA 生物制品项目改造 · 行动项跟踪"
    style_range(dash["B2"], fill=fill_navy, font_=font_title, align=Alignment(vertical="center", horizontal="left"))
    for col in range(3, 10):
        dash.cell(2, col).fill = fill_navy
        dash.cell(3, col).fill = fill_navy
        dash.cell(2, col).border = Border()
        dash.cell(3, col).border = Border()
    dash.merge_cells("B4:I4")
    dash["B4"] = "持续总账，不再按会议日期新开主表。模板列习惯保留：提出方 / 内容 / 进展 / 完成 / 负责人。黄色格子为里程碑输入，改这里，下面卡片用公式重算。"
    style_range(dash["B4"], fill=fill_navy2, font_=font_sub, align=Alignment(vertical="center", wrap_text=True))
    for col in range(3, 10):
        dash.cell(4, col).fill = fill_navy2
        dash.cell(4, col).border = Border()

    # milestone inputs
    dash["B6"] = "里程碑（方案一，会上：领导要求）"
    style_range(dash["B6"], fill=fill_gold, font_=Font(name="Arial", size=10, bold=True, color=NAVY), border=Border())
    dash.merge_cells("B6:C6")
    labels = [
        (7, "Demo 期", date(2026, 10, 5), "C7", "会上：按 10/5 demo 倒推 10/15 开工"),
        (8, "计划开工", date(2026, 10, 15), "C8", "方案一"),
        (9, "建设完成", date(2026, 11, 3), "C9", "会上约数，制备配液最紧"),
        (10, "验证完成", date(2026, 11, 16), "C10", "会上约数"),
        (11, "交付", date(2026, 11, 18), "C11", "硬门禁：所有工作不得晚于该日"),
    ]
    dash["B7"] = "Demo 期"
    dash["B8"] = "计划开工"
    dash["B9"] = "建设完成"
    dash["B10"] = "验证完成"
    dash["B11"] = "交付"
    dash["C7"] = date(2026, 10, 5)
    dash["C8"] = date(2026, 10, 15)
    dash["C9"] = date(2026, 11, 3)
    dash["C10"] = date(2026, 11, 16)
    dash["C11"] = date(2026, 11, 18)
    notes = {
        7: "Source: 2026-09-21 部署会。demo 推迟则开工推迟；提前则可能提前 1–2 天拆设备。",
        8: "Source: 2026-09-21 部署会，方案一。",
        9: "Source: 2026-09-21 部署会约数。",
        10: "Source: 2026-09-21 部署会约数。",
        11: "Source: 2026-09-21 部署会硬门禁。",
    }
    for r in range(7, 12):
        style_range(dash.cell(r, 2), fill=fill_hint, font_=font_b, align=center)
        style_range(dash.cell(r, 3), fill=fill_input, font_=Font(name="Arial", size=11, bold=True, color="1A5276"),
                    align=center, num="YYYY-MM-DD")
        dash.cell(r, 3).comment = None
        dash[f"D{r}"] = notes[r]
        style_range(dash[f"D{r}"], font_=font_small, align=left, border=Border())
        dash.merge_cells(f"D{r}:I{r}")

    dash["C11"].comment = Comment("Hard gate from 21 Sep 2026 meeting. Do not treat other dates as equally firm.", "纪要")

    # KPI cards — formulas against 行动项总账
    # Data starts row 5 on 行动项总账; status col I, due col H, traffic col N
    cards = [
        (6, 6, "未关闭", "=COUNTIFS('行动项总账'!I:I,\"未开始\")+COUNTIFS('行动项总账'!I:I,\"进行中\")+COUNTIFS('行动项总账'!I:I,\"阻塞\")", CARD_OPEN),
        (6, 7, "阻塞", "=COUNTIF('行动项总账'!I:I,\"阻塞\")", CARD_BLOCK),
        (6, 8, "今日到期", "=COUNTIFS('行动项总账'!H:H,TODAY(),'行动项总账'!I:I,\"<>已完成\",'行动项总账'!I:I,\"<>取消\")", CARD_DUE),
        (6, 9, "逾期未完成", "=COUNTIFS('行动项总账'!N:N,\"逾期\")", CARD_OVER),
    ]
    # place cards on row 13-15
    dash.merge_cells("B13:C13")
    dash["B13"] = "未关闭"
    dash.merge_cells("D13:E13")
    dash["D13"] = "阻塞"
    dash.merge_cells("F13:G13")
    dash["F13"] = "今日到期"
    dash.merge_cells("H13:I13")
    dash["H13"] = "逾期未完成"
    dash.merge_cells("B14:C15")
    dash["B14"] = '=COUNTIFS(\'行动项总账\'!I:I,"未开始")+COUNTIFS(\'行动项总账\'!I:I,"进行中")+COUNTIFS(\'行动项总账\'!I:I,"阻塞")'
    dash.merge_cells("D14:E15")
    dash["D14"] = '=COUNTIF(\'行动项总账\'!I:I,"阻塞")'
    dash.merge_cells("F14:G15")
    dash["F14"] = '=COUNTIFS(\'行动项总账\'!H:H,TODAY(),\'行动项总账\'!I:I,"<>已完成",\'行动项总账\'!I:I,"<>取消")'
    dash.merge_cells("H14:I15")
    dash["H14"] = '=COUNTIF(\'行动项总账\'!N:N,"逾期")'

    card_fills = [
        (2, 3, CARD_OPEN),
        (4, 5, CARD_BLOCK),
        (6, 7, CARD_DUE),
        (8, 9, CARD_OVER),
    ]
    for r in (13, 14, 15):
        for c1, c2, color in card_fills:
            for c in range(c1, c2 + 1):
                cell = dash.cell(r, c)
                cell.fill = PatternFill("solid", fgColor=color)
                cell.border = Border()
                cell.font = font_card_l if r == 13 else font_card_n
                cell.alignment = Alignment(horizontal="center", vertical="center")
    dash["B13"].alignment = Alignment(horizontal="center", vertical="center")
    dash["D13"].alignment = Alignment(horizontal="center", vertical="center")
    dash["F13"].alignment = Alignment(horizontal="center", vertical="center")
    dash["H13"].alignment = Alignment(horizontal="center", vertical="center")

    dash["B17"] = "按片断未关闭（公式）"
    style_range(dash["B17"], fill=fill_gold, font_=Font(name="Arial", size=10, bold=True, color=NAVY), border=Border())
    dash.merge_cells("B17:C17")
    dash["B18"] = "片断"
    dash["C18"] = "未关闭"
    dash["D18"] = "阻塞"
    for col in (2, 3, 4):
        style_range(dash.cell(18, col), fill=fill_head, font_=font_hdr, align=center)
    for i, stream in enumerate(STREAMS, start=19):
        dash.cell(i, 2).value = stream
        dash.cell(i, 3).value = f'=COUNTIFS(\'行动项总账\'!$D:$D,B{i},\'行动项总账\'!$I:$I,"<>已完成",\'行动项总账\'!$I:$I,"<>取消",\'行动项总账\'!$I:$I,"<>")'
        dash.cell(i, 4).value = f'=COUNTIFS(\'行动项总账\'!$D:$D,B{i},\'行动项总账\'!$I:$I,"阻塞")'
        for col in (2, 3, 4):
            style_range(dash.cell(i, col), fill=fill_alt if i % 2 else fill_white, font_=font, align=center if col > 2 else left)

    dash.merge_cells("F17:I17")
    dash["F17"] = "用法（1 分钟）"
    style_range(dash["F17"], fill=fill_gold, font_=Font(name="Arial", size=10, bold=True, color=NAVY), border=Border())
    dash.merge_cells("F18:I26")
    dash["F18"] = (
        "1. 只改「行动项总账」。不要再为每次会新建主表。\n"
        "2. 进展日志只追加，格式：YYYY-MM-DD：一句话。不要覆盖旧进展。\n"
        "3. 完成后：状态=已完成，在进展里写证据（文件名 / 版本 / 会议日期）。\n"
        "4. 用表头筛选看「未开始 / 进行中 / 阻塞」。\n"
        "5. 截止日期空着时，交通灯为「—」，请尽快补日期。\n"
        "6. 黄色里程碑可改；卡片和片断统计全部是公式。\n"
        "7. 三合一对外写 AFD（agitated filter dryer）。\n"
        "8. 本文件为内部跟踪，不直接发给客户。"
    )
    style_range(dash["F18"], fill=fill_hint, font_=font, align=Alignment(wrap_text=True, vertical="top"))
    for r in range(18, 27):
        for c in range(6, 10):
            dash.cell(r, c).fill = fill_hint
            dash.cell(r, c).border = Border()

    dash.merge_cells("B28:I28")
    dash["B28"] = "最近会议：2026-09-21 部署会 09:03–09:29。纪要：meetings/2026/2026-09-21-poa-deployment.md。项目代码书面 POA、口述 KOA，待核实。"
    style_range(dash["B28"], font_=font_small, align=left, border=Border())

    dash.column_dimensions["A"].width = 3
    dash.column_dimensions["B"].width = 16
    dash.column_dimensions["C"].width = 16
    dash.column_dimensions["D"].width = 14
    dash.column_dimensions["E"].width = 14
    dash.column_dimensions["F"].width = 14
    dash.column_dimensions["G"].width = 14
    dash.column_dimensions["H"].width = 14
    dash.column_dimensions["I"].width = 18
    dash.row_dimensions[2].height = 24
    dash.row_dimensions[3].height = 18
    dash.row_dimensions[4].height = 32
    dash.row_dimensions[13].height = 18
    dash.row_dimensions[14].height = 22
    dash.row_dimensions[15].height = 22
    dash.row_dimensions[18].height = 18
    dash.freeze_panes = "B6"

    # ----- 行动项总账 -----
    ws = wb.create_sheet("行动项总账")
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.5)
    ws.print_title_rows = "1:4"
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "&B POA 改造行动项总账"
    ws.oddFooter.left.text = "内部 · 只追加进展，不删历史行"
    ws.oddFooter.right.text = "&P / &N"

    ws.merge_cells("A1:N2")
    ws["A1"] = "POA 生物制品项目改造 · 行动项总账（持续）"
    style_range(ws["A1"], fill=fill_navy, font_=font_title, align=Alignment(vertical="center", horizontal="left"))
    for c in range(1, 15):
        ws.cell(1, c).fill = fill_navy
        ws.cell(2, c).fill = fill_navy
        ws.cell(1, c).border = Border()
        ws.cell(2, c).border = Border()
    ws.merge_cells("A3:N3")
    ws["A3"] = (
        "筛选「状态 / 片断 / 负责人」。进展只追加。交通灯=公式（对比截止日期与 TODAY）。"
        "ID 不换号。黄色状态/日期可改。"
    )
    style_range(ws["A3"], fill=fill_navy2, font_=font_sub, align=Alignment(vertical="center"))
    for c in range(1, 15):
        ws.cell(3, c).fill = fill_navy2
        ws.cell(3, c).border = Border()

    headers = [
        "ID",
        "提出日\nRaised",
        "提出方\nOriginator",
        "片断 / 专业\nWorkstream",
        "行动项\nAction",
        "验收标准\nDone means",
        "负责人\nOwner",
        "截止日期\nDue",
        "状态\nStatus",
        "优先级\nPriority",
        "进展日志（只追加）\nProgress log",
        "更新日\nUpdated",
        "来源会议\nSource",
        "交通灯\nTraffic",
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(4, i, h)
        style_range(cell, fill=fill_head, font_=font_hdr, align=wrap_c, border=thick_bottom)

    first_data = 5
    for i, a in enumerate(ACTIONS):
        r = first_data + i
        due_cell = f"H{r}"
        status_cell = f"I{r}"
        values = [
            a["id"],
            a["raised"],
            a["origin"],
            a["stream"],
            a["action"],
            a["done"],
            a["owner"],
            a["due"],
            a["status"],
            a["pri"],
            a["progress"],
            date(2026, 9, 21),
            a["src"],
            f'=IF(OR({status_cell}="已完成",{status_cell}="取消"),"—",IF({due_cell}="",IF({status_cell}="阻塞","逾期","—"),IF({due_cell}<TODAY(),"逾期",IF({due_cell}=TODAY(),"今日到期","按期"))))',
        ]
        fills_status = {
            "未开始": OPEN_BG,
            "进行中": PROG_BG,
            "阻塞": BLOCK_BG,
            "已完成": DONE_BG,
        }
        row_fill = PatternFill("solid", fgColor=fills_status.get(a["status"], ALT if i % 2 else WHITE))
        for c, v in enumerate(values, 1):
            cell = ws.cell(r, c, v)
            num = None
            if c in (2, 8, 12) and isinstance(v, date):
                num = "YYYY-MM-DD"
            elif c in (2, 8, 12) and v is None:
                num = "YYYY-MM-DD"
            al = wrap_c if c in (1, 2, 4, 7, 8, 9, 10, 12, 14) else wrap
            fnt = font_b if c in (1, 7, 9) else font
            style_range(cell, fill=row_fill, font_=fnt, align=al, num=num)
        if a["note"]:
            ws.cell(r, 5).comment = Comment(a["note"], "纪要")
        ws.row_dimensions[r].height = 56

    last = first_data + len(ACTIONS) - 1
    # extra blank rows for new items
    for r in range(last + 1, last + 16):
        for c in range(1, 15):
            cell = ws.cell(r, c, None)
            style_range(cell, fill=fill_white, font_=font, align=wrap)
            if c in (2, 8, 12):
                cell.number_format = "YYYY-MM-DD"
        ws.cell(r, 14).value = (
            f'=IF(OR(I{r}="",I{r}="已完成",I{r}="取消"),IF(I{r}="","","—"),'
            f'IF(H{r}="",IF(I{r}="阻塞","逾期","—"),'
            f'IF(H{r}<TODAY(),"逾期",IF(H{r}=TODAY(),"今日到期","按期"))))'
        )
        ws.row_dimensions[r].height = 22
    blank_last = last + 15

    dv_status = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    dv_status.error = "请选：未开始 / 进行中 / 阻塞 / 已完成 / 取消"
    dv_status.errorTitle = "状态"
    dv_status.prompt = "选择状态"
    dv_status.promptTitle = "状态"
    dv_pri = DataValidation(type="list", formula1='"' + ",".join(PRIORS) + '"', allow_blank=True)
    dv_stream = DataValidation(type="list", formula1='"' + ",".join(STREAMS) + '"', allow_blank=True)
    ws.add_data_validation(dv_status)
    ws.add_data_validation(dv_pri)
    ws.add_data_validation(dv_stream)
    dv_status.add(f"I{first_data}:I{blank_last}")
    dv_pri.add(f"J{first_data}:J{blank_last}")
    dv_stream.add(f"D{first_data}:D{blank_last}")

    apply_cf_status(ws, "I", first_data, blank_last)
    apply_cf_pri(ws, "J", first_data, blank_last)
    apply_cf_traffic(ws, "N", first_data, blank_last)

    ws.auto_filter.ref = f"A4:N{blank_last}"
    ws.freeze_panes = "A5"
    ws.auto_filter.add_sort_condition("H5")

    widths = {
        "A": 16, "B": 12, "C": 16, "D": 14, "E": 42, "F": 28, "G": 20,
        "H": 13, "I": 12, "J": 10, "K": 46, "L": 12, "M": 18, "N": 12,
    }
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 32
    ws.sheet_view.zoomScale = 110

    # ----- 0921 会议快照（贴近原模板） -----
    snap = wb.create_sheet("20260921")
    snap.sheet_view.showGridLines = False
    snap.page_setup.orientation = "landscape"
    snap.page_setup.fitToPage = True
    snap.page_setup.fitToWidth = 1
    snap.page_setup.paperSize = snap.PAPERSIZE_A3
    snap.merge_cells("A1:G2")
    snap["A1"] = "2026-09-21 部署会快照（模板列）· 主跟踪请回「行动项总账」"
    style_range(snap["A1"], fill=fill_navy, font_=font_title, align=Alignment(vertical="center"))
    for c in range(1, 8):
        snap.cell(1, c).fill = fill_navy
        snap.cell(2, c).fill = fill_navy
        snap.cell(1, c).border = Border()
        snap.cell(2, c).border = Border()
    old_headers = [
        "日期\nDate",
        "议题提出方\nIssue Originator",
        "会议主题内容\nMeeting Theme and Content",
        "进展\nProgress",
        "是否完成\nCompleted or not",
        "负责人\nPerson in charge",
        "对应总账 ID\nRegister ID",
    ]
    for i, h in enumerate(old_headers, 1):
        style_range(snap.cell(3, i, h), fill=fill_head, font_=font_hdr, align=wrap_c, border=thick_bottom)
    for i, a in enumerate(ACTIONS):
        r = 4 + i
        completed = "YES" if a["status"] == "已完成" else ("BLOCKED" if a["status"] == "阻塞" else "NO")
        vals = [a["raised"], a["origin"], a["action"], a["progress"], completed, a["owner"], a["id"]]
        bg = PatternFill("solid", fgColor=BLOCK_BG if completed == "BLOCKED" else (ALT if i % 2 else WHITE))
        for c, v in enumerate(vals, 1):
            cell = snap.cell(r, c, v)
            num = "YYYY-MM-DD" if c == 1 else None
            style_range(cell, fill=bg, font_=font_b if c in (5, 6, 7) else font,
                        align=wrap_c if c in (1, 5, 7) else wrap, num=num)
        snap.row_dimensions[r].height = 48
        # formula pull status from master by ID
        snap.cell(r, 5).value = (
            f'=IFERROR(IF(INDEX(\'行动项总账\'!$I:$I,MATCH(G{r},\'行动项总账\'!$A:$A,0))="已完成","YES",'
            f'IF(INDEX(\'行动项总账\'!$I:$I,MATCH(G{r},\'行动项总账\'!$A:$A,0))="阻塞","BLOCKED","NO")),"NO")'
        )
        snap.cell(r, 6).value = f'=IFERROR(INDEX(\'行动项总账\'!$G:$G,MATCH(G{r},\'行动项总账\'!$A:$A,0)),"{a["owner"]}")'
        snap.cell(r, 4).value = f'=IFERROR(INDEX(\'行动项总账\'!$K:$K,MATCH(G{r},\'行动项总账\'!$A:$A,0)),"")'
    snap.auto_filter.ref = f"A3:G{3 + len(ACTIONS)}"
    snap.freeze_panes = "C4"
    for col, w in zip("ABCDEFG", (12, 16, 48, 48, 14, 22, 18)):
        snap.column_dimensions[col].width = w
    snap.row_dimensions[1].height = 22
    snap.row_dimensions[2].height = 16
    snap.row_dimensions[3].height = 32
    snap.auto_filter.add_filter_column(4, ["NO", "BLOCKED"], blank=False)

    # ----- 使用说明 -----
    guide = wb.create_sheet("使用说明")
    guide.sheet_view.showGridLines = False
    guide.merge_cells("B2:F2")
    guide["B2"] = "与另一项目「行动项跟踪-0921-0928.xlsx」的对照"
    style_range(guide["B2"], fill=fill_navy, font_=font_title, align=Alignment(vertical="center"))
    for c in range(3, 7):
        guide.cell(2, c).fill = fill_navy
        guide.cell(2, c).border = Border()
    rows = [
        ("原模板", "本文件"),
        ("每个会议日期一张主表（0915、0907…）", "一张「行动项总账」跨会跟踪，日期表只作快照"),
        ("是否完成 = YES/空", "状态五档 + 交通灯公式，能看出逾期"),
        ("进展中英混写、越写越长", "中文为主；只追加带日期的一行"),
        ("无 ID，难引用", "A-YYYYMMDD-xx，纪要与 Excel 同一编号"),
        ("无片断 / 截止日期 / 优先级", "补齐片断、到期日、优先级、验收标准"),
        ("无总览", "总览卡片和片断统计全部用公式"),
        ("适合客户双语周会", "本项目先做内部执行账；对外稿另出"),
    ]
    for i, (a, b) in enumerate(rows, start=4):
        guide.cell(i, 2).value = a
        guide.cell(i, 3).value = b
        guide.merge_cells(f"C{i}:F{i}")
        fill = fill_head if i == 4 else (fill_alt if i % 2 == 0 else fill_white)
        fnt = font_hdr if i == 4 else font
        style_range(guide.cell(i, 2), fill=fill, font_=fnt, align=wrap)
        for c in range(3, 7):
            style_range(guide.cell(i, c), fill=fill, font_=fnt, align=wrap)
        guide.row_dimensions[i].height = 28
    guide.merge_cells("B13:F16")
    guide["B13"] = (
        "每周更新建议：会后 2 小时内改总账（状态、截止日期、进展）。"
        "不要复制整行到新 sheet。关闭项留在表里，用筛选看未关闭。"
        "术语：三合一 = AFD（agitated filter dryer）。"
    )
    style_range(guide["B13"], fill=fill_hint, font_=font, align=Alignment(wrap_text=True, vertical="top"))
    for r in range(13, 17):
        for c in range(2, 7):
            guide.cell(r, c).fill = fill_hint
            guide.cell(r, c).border = Border()
    guide.column_dimensions["A"].width = 3
    guide.column_dimensions["B"].width = 42
    guide.column_dimensions["C"].width = 28
    guide.column_dimensions["D"].width = 18
    guide.column_dimensions["E"].width = 18
    guide.column_dimensions["F"].width = 18
    guide.row_dimensions[2].height = 28

    # named ranges for milestones
    wb.defined_names.add(DefinedName(name="MilestoneHandover", attr_text="'总览'!$C$11"))
    wb.defined_names.add(DefinedName(name="MilestoneStart", attr_text="'总览'!$C$8"))

    ws.sheet_properties.tabColor = NAVY
    dash.sheet_properties.tabColor = GOLD
    snap.sheet_properties.tabColor = "5D6D7E"

    wb.save(OUT)
    print("saved", OUT, "rows", len(ACTIONS))


if __name__ == "__main__":
    build()
