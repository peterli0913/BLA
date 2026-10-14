#!/usr/bin/env python3
"""Internal minutes of the 10/14/2026 Lilly on-site wrap-up (useful content through 00:14:21)."""

import subprocess
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from build_lilly_call_script import (
    GOLD, MUTED, NAVY, add_bullet, add_h1, add_h2, add_page_number, add_table, set_run_font,
)

OUT = Path("/workspace/meetings/2026/2026-10-14-礼来客户会会议纪要.docx")

SUMMARY = [
    "今天和现场三位礼来把在审文件过了一遍。只是描述性修改的，由礼来回去催 feedback、关闭 comment；有实质问题的，今天当场讨论。今晚仍要看 Ewan、Sonia 是否还有新 comment，有的话明天继续在现场谈。",
    "GAP 的 comment 已全部回复，上午与客户过完，没有留下待讨论项，待 Sonia、Ewan 确认后即可关闭。PVP 氧化后洗涤树脂的温度已与 Ewan 定为 −5～0℃、中心点 −3℃，并按他的建议增加预冷；礼来要求其他相关文件同步改，不能只改 PVP。",
    "CCS 的风险写法已定：严重程度为 5 时，结论最低按中风险来写，不在结论里降级；结论之后的 action 说明当前能做的控制都已做，风险处于可控状态。TSP 切割小试，礼来坚持在 QC 检测；我方已说明 QC 不接受实验室研发样品，Jerry、Sara 回去对齐。若仍要求 QC 做，只能用纸质记录，并要确认是仅 PPQ 三批，还是后续商业化批次同样执行。",
    "ORA 和 GAP 在签批时仍 open 的 action，用 QMS 跟踪，完成期限为生产前，证据上传 QMS。CVP 与现场两位没有新的 open 项，等今晚 Sonia 的反馈。PQ readiness 已与 Sara 一致：全部 action 完成后，由凯莱英 QA 和礼来 QA 终批；模板已接受，其他项目也可按该 readiness 和 checklist 做。",
    "变更及附件的 comment 下午已讨论完。三片段 SPPS-1 启动前的 DMF 管路清洁验证，礼来想了解除 level 1 data check 之外有没有分析方法介入，明天准备说明。报警管理和配方 TCU 手动调温已解释，Jerry 没有进一步意见。明天现场看 DIPA 投料；激活罐须在天窗关闭后才能启动 SPPS；DIPA 从普通区到受控区的转移要给出短期措施和现状。球阀材质与 DMS、DMF 不兼容，今晚看能否解决，否则明天现场说明。调试和验证期间使用乙二醇一事已解释并关闭。QC 昨日遗留，明天由韩一杰解释。",
]

TOPICS = [
    {
        "title": "1. 文件审阅的处理方式",
        "rows": [
            ("礼来关注 / 要求",
             "• 描述性修改：回去催相关人员尽快 feedback，关掉 comment。\n"
             "• 有实质问题的文件，今天作为重点讨论。\n"
             "• 若今晚 Ewan、Sonia 还有新 comment，明天继续在现场谈。离场前已与 Sara 说好。"),
            ("会上共识",
             "没有实质问题的 comment 由礼来侧关闭；有问题的今天已讨论。今晚核对有无新 comment，有则明天现场继续。"),
        ],
    },
    {
        "title": "2. GAP",
        "rows": [
            ("礼来关注 / 要求",
             "• Sonia 和 Ewan 确认我方回复是否满足要求。\n"
             "• 不排除今晚还有新的 comment，明天需要再核对一遍。"),
            ("我方会上回应",
             "• GAP 由王杰负责，comment 已全部回复，上午已与客户过完。\n"
             "• 上午没有留下需要再讨论的议题。"),
            ("会上共识",
             "若今晚没有新 comment，这两个 GAP 文件可以关闭，并更新 GAP 状态。"),
        ],
    },
    {
        "title": "3. PVP：氧化后洗涤温度与其他文件同步",
        "rows": [
            ("礼来关注 / 要求",
             "• 氧化后洗涤树脂的中心点温度，今天与 Ewan 对齐。\n"
             "• Ewan 建议增加预冷。\n"
             "• 不能只改 PVP。温度和预冷写进 PVP 之后，其他相关文件要一起改。"),
            ("我方会上回应",
             "• 温度范围改为 −5～0℃，中心点 −3℃。这一点已经解决，PVP 新增批复也已完成。\n"
             "• 车间确认预冷可以执行。下午已按这个要求改相关文件。\n"
             "• 客户离场前，还有部分文件需要核实是否都改完。"),
            ("会上共识",
             "温度范围和预冷已定，PVP 新增批复已完成。其他相关文件同步修改，并核实是否改完。"),
        ],
    },
    {
        "title": "4. CCS",
        "rows": [
            ("礼来关注 / 要求",
             "• 文件里要分清原始风险，以及 action 之后的风险。这一处是 CCS 当前的主要问题。"),
            ("我方会上回应",
             "• 严重程度为 5 时，风险最低按中风险评估，结论就写中风险，不在结论里把等级降下来。\n"
             "• 结论之后的 action 写明：当前能做的控制都已做，该风险处于可控状态。"),
            ("会上共识",
             "按上述写法处理。CCS 今天主要对齐的就是这一点。"),
        ],
    },
    {
        "title": "5. TSP：切割小试的检测地点",
        "rows": [
            ("礼来关注 / 要求",
             "• Step 1 固相合成之后，Step 2 要做切割小试。\n"
             "• 切割小试的结果，礼来坚持放在 QC 检测。"),
            ("我方会上回应",
             "• 下午 QA、QC 和分析同事讨论后，已告知现场 Jerry 和 Sara：QC 不接受实验室研发样品的检测。\n"
             "• 若礼来仍要求在 QC 检测，只能用纸质记录把流程补齐。\n"
             "• 还要确认适用范围：仅 PPQ 三批，还是后续商业化批次也同样执行。"),
            ("会上共识",
             "Jerry、Sara 回去内部对齐后再定后续做法。TSP 其余事项没有遗留。"),
        ],
    },
    {
        "title": "6. ORA 与 GAP 的 action 跟踪",
        "rows": [
            ("礼来关注 / 要求",
             "• ORA 和 GAP 里的 action 都要有跟踪。签批时，部分 action 可能仍是 open。"),
            ("我方会上回应",
             "• QMS 里已有 action 跟踪模块。签批时仍 open 的 action，放入 QMS 跟踪。\n"
             "• 完成期限为生产前。相关证据上传到 QMS。"),
            ("会上共识",
             "ORA 和 GAP 的 action 都按上述方式 follow up，即可接受。"),
        ],
    },
    {
        "title": "7. CVP",
        "rows": [
            ("礼来关注 / 要求",
             "• 现场两位没有留下新的 open 项。\n"
             "• 他们看到的描述仍较概括，回去催 Sonia 关闭 Sonia 名下的 comment。"),
            ("会上共识",
             "CVP 今天没有新增 open 项。最终以今晚 Sonia 的反馈为准。"),
        ],
    },
    {
        "title": "8. PQ readiness",
        "rows": [
            ("礼来关注 / 要求",
             "• 与 Sara 讨论 PQ readiness。模板本身没有被指出问题。"),
            ("我方会上回应",
             "• 全部 action 完成之后，由凯莱英 QA 和礼来 QA 做最终审批。"),
            ("会上共识",
             "模板接受。全部 action 完成后走双方 QA 终批。其他项目也可按该 PPQ readiness 和 checklist 执行。"),
        ],
    },
    {
        "title": "9. 球阀材质与 DMS、DMF 的相容性",
        "rows": [
            ("礼来关注 / 要求",
             "• Ewan 提出材质问题。信息今天已转给 Jerry 和 Sara。\n"
             "• 他们可能明天到现场看：替换后的球阀与 DMS、DMF 不兼容。"),
            ("我方会上回应",
             "• 今晚先看能否解决。解决不了，明天在现场说明。"),
            ("会上共识",
             "今晚处理；处理不完则明天现场看或说明。"),
        ],
    },
    {
        "title": "10. 变更，以及 SPPS-1 启动前的 DMF 管路清洁验证",
        "rows": [
            ("礼来关注 / 要求",
             "• 变更和变更附件，下午已充分讨论，没有留下未关闭的 comment，仍待礼来审阅这些回复是否都解决问题。\n"
             "• 三片段 SPPS-1 启动前的 DMF 管路清洁验证：此前收到的反馈是 level 1 data check。今天他们想知道有没有分析方法可以介入。这一点没有说成硬性要求，是要这个信息。"),
            ("我方会上回应",
             "• 此前向 Ewan 反馈的是 level 1 data check。\n"
             "• 明天准备材料，向他们说明。"),
            ("会上共识",
             "变更 comment 待礼来审阅关闭。分析方法是否介入，明天说明。"),
        ],
    },
    {
        "title": "11. 报警管理与配方 TCU 温度",
        "rows": [
            ("礼来关注 / 要求",
             "• 报警管理，以及配方里 TCU 温度如何调整。"),
            ("我方会上回应",
             "• 当前不能完全自动控温，调 TCU 温度需要人员手动介入。已说明现在的管理方式。"),
            ("会上共识",
             "Jerry 对当前说明没有进一步意见。"),
        ],
    },
    {
        "title": "12. 明日现场：DIPA 投料",
        "rows": [
            ("礼来关注 / 要求",
             "• 明天指定看 DIPA 投料细节。物料流程已经确认，接下来要看现场怎么把料投进去。\n"
             "• 关心投料过程如何密封加入。只有两个点加入 DIPA。"),
            ("我方会上回应",
             "• 用实物 DIPA，还是用溶液做模拟，下来再定，也可以直接问客户。客户若要看实物，再领料实投。\n"
             "• 相关设备仍在钝化和清洁检测中。倾向到现场做模拟讲解，由熟悉流程的人说明怎么投、怎么控制流速和温度，以及与工艺风险相关的控制点。\n"
             "• 现场听众一位是工程背景、一位是 QA 背景，讲解人要能把工艺讲清楚。\n"
             "• 需要讲清的控制：DIPA 见光会生成乙醛，是 9788 杂质的一个来源，与光照和暴露时间有关。开盖、关盖时做氮气保护，尽量避光。"),
            ("会上共识",
             "明天到现场讲 DIPA 投料。实物还是模拟，投料前与客户确认。"),
        ],
    },
    {
        "title": "13. 激活罐异物偏差与天窗",
        "rows": [
            ("礼来关注 / 要求",
             "• 激活罐异物偏差今天已讨论。\n"
             "• 客户明确要求：天窗关闭之后，才能启动 SPPS。"),
            ("我方会上回应",
             "• 按现场讨论结果，偏差文件已经更新，会后上传。"),
            ("会上共识",
             "偏差本身没有遗留问题。文件更新后上传。天窗关闭是启动 SPPS 的前提。"),
        ],
    },
    {
        "title": "14. DIPA 从普通区到受控区的转移",
        "rows": [
            ("礼来关注 / 要求",
             "• 本次访问中看过 DIPA 的转移路线。从普通区到受控区这一段缺少控制。\n"
             "• 要了解短期措施，以及现在具体做到什么程度。"),
            ("我方会上回应",
             "• 相关信息已整理进群，并 @ 到责任人。"),
            ("会上共识",
             "需要向礼来反馈短期措施和当前实施情况。"),
        ],
    },
    {
        "title": "15. 设备调试、验证期间使用乙二醇",
        "rows": [
            ("礼来关注 / 要求",
             "• 设备调试和验证期间使用过乙二醇。"),
            ("我方会上回应",
             "• 已经向客户解释。说明已发到微信群。"),
            ("会上共识",
             "该问题已经解决。"),
        ],
    },
    {
        "title": "16. QC 昨日遗留",
        "rows": [
            ("礼来关注 / 要求",
             "• 昨天 QC 有遗留问题，本次总结没有展开具体内容。"),
            ("我方会上回应",
             "• 明天由韩一杰下来解释。"),
            ("会上共识",
             "明天解释。"),
        ],
    },
]

ACTIONS = [
    ["1", "今晚核对 Ewan、Sonia 是否还有新 comment。没有的话，请三萌关闭两个 GAP 文件并更新状态；有的话，明天在现场继续讨论。",
     "GAP", "今晚核对；明天关闭或继续讨论"],
    ["2", "把氧化后洗涤树脂温度（−5～0℃，中心点 −3℃）和预冷措施同步写入其他相关文件，并核实是否都已改完。",
     "PVP", "今天未核实完的继续核实"],
    ["3", "按今天定的口径修改 CCS：严重程度为 5 时结论保持中风险，在后续 action 中写明现有控制已做、风险处于可控状态。",
     "CCS", "未定"],
    ["4", "等 Jerry、Sara 对齐切割小试的检测地点。若仍要求 QC 检测，用纸质记录完善流程，并确认适用于 PPQ 三批，还是后续商业化批次同样执行。",
     "TSP", "未定"],
    ["5", "签批时仍为 open 的 ORA、GAP action 放入 QMS 跟踪，生产前完成，证据上传 QMS。",
     "ORA / GAP", "生产前"],
    ["6", "处理球阀材质与 DMS、DMF 不兼容的问题。今晚解决不了，明天在现场说明。",
     "ORA", "今晚；否则明天现场"],
    ["7", "准备三片段 SPPS-1 启动前 DMF 管路清洁验证的说明，回答除 level 1 data check 之外有没有分析方法介入。",
     "变更", "明天"],
    ["8", "明天现场讲解 DIPA 投料。先与客户确认用实物还是模拟；说明密封加入、流速和温度控制，以及避光、开盖时氮气保护。",
     "DIPA 投料", "明天现场"],
    ["9", "上传已按现场讨论更新的激活罐异物偏差。",
     "激活罐偏差", "会后"],
    ["10", "SPPS 启动前关闭激活罐天窗。",
     "激活罐 / SPPS", "SPPS 启动前"],
    ["11", "向礼来反馈 DIPA 从普通区转移到受控区的短期措施和当前实施情况。",
     "DIPA 转移", "未定"],
    ["12", "由韩一杰解释 QC 昨天的遗留问题。",
     "QC", "明天"],
]


def add_topic(doc, topic):
    add_h2(doc, topic["title"])
    add_table(doc, ["项目", "内容"], topic["rows"], [3.2, 14.4], size=10, first_col_bold=True)


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.7)
    sec.top_margin = sec.bottom_margin = Cm(1.5)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(fp)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    set_run_font(p.add_run("礼来客户会会议纪要"), size=20, bold=True, color=NAVY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    set_run_font(
        p.add_run("Project Documents, Open Actions and Site Topics – 10/14/2026"),
        size=12, bold=True, color=GOLD,
    )

    add_h1(doc, "一、会议要点")
    for s in SUMMARY:
        add_bullet(doc, s)

    add_h1(doc, "二、分议题纪要")
    for topic in TOPICS:
        add_topic(doc, topic)

    add_h1(doc, "三、后续 Action Items")
    add_table(doc, ["#", "行动", "议题", "时间"], ACTIONS, [1.0, 10.6, 3.0, 3.0], size=9.5)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
