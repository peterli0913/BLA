#!/usr/bin/env python3
"""Make column J (Lilly Response) of the 09-29 action tracker bilingual: English on top, Chinese below.

Only xl/sharedStrings.xml and xl/worksheets/sheet1.xml are rewritten: the J strings get the Chinese
appended, column J is widened, and rows are made taller (never shorter) where J needs it.
All other parts are copied byte-for-byte.

Usage: python3 scripts/build_tracker_lilly_response_bilingual.py
"""
import math
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Asymchem_Visit_Action_Tracker_09-29-2026 - Lilly Response.xlsx"
OUT = ROOT / "deliverables" / "Asymchem_Visit_Action_Tracker_09-29-2026 - Lilly Response_中英对照.xlsx"

SHEET = "xl/worksheets/sheet1.xml"
SST = "xl/sharedStrings.xml"
J_WIDTH = 60
MAX_HT = 409.5
CJK = re.compile(r"[\u3000-\u303f\u3400-\u9fff\uff00-\uffef]")

NO_FOLLOW_UP = "无需进一步跟进。"
REVIEW_REVERT = "AJM：我会审阅所提供的信息后再回复。"

# Row number -> Chinese translation of the Lilly response in J.
ZH = {
    2: "需要就各步骤的清洁间隔长度达成一致，这将决定“大清洁”的频次。采用清洁间隔时，同一间隔内批与批之间的保持时间采用什么策略？\n\n"
       "SPPS 何时需要清洁主批记录？第一批之后吗？这可能决定清洁策略和文件对齐的时间表及紧迫程度。在该记录完成之前，"
       "Asymchem 将依据什么指令在 Step 3 Demo 与 PPQ 开始之间完成适当的清洁，以确保设备足够洁净？"
       "Step 1/Step 2 Demo 之后执行的是哪份清洁记录？我们能否审阅？\n\n"
       "请确保清洁 SOP/记录涵盖所有需要清洁的产品接触类零散部件（过滤器外壳、泵等），这些部件可能无法通过 CIP 程序有效清洁。\n\n"
       "另外，请针对每个步骤告知淋洗/擦拭限度的计算在哪里，以便我们审阅。此外，请说明每个步骤将分别执行多少次清洁确认（verification）"
       "和清洁验证（validation），以及清洁验证何时开始（相对于 PPQ 的时间节点）。",
    4: "AJM：可以接受，此项视为关闭。",
    5: "需要进一步讨论，以澄清 Asymchem 提出的方案。我们不太理解 F 列中各 Option 的含义。",
    6: "过渡期内，请确保在 MOR 中写明指令：必要时使用手电筒完成罐内目视检查。",
    7: "AJM：好的。关于插底管：如在液面上方加入，请确保液流导向罐底存液，而不是冲向任何表面；如采用液面下加入，"
       "请确保能够避免管线内的不均匀问题。材质方面，如使用 PTFE，请考虑渗透的影响，可考虑以 PFA 作为替代。",
    8: "AJM：如已在 MOR 中确认，则无需进一步跟进。我们只是希望记录中有足够的细节，确保能够准确核算实际进入罐内的物料。",
    9: "请评估 Step 1 和 Step 2 的设备装置，确认是否存在类似差距。",
    10: "在长期厂房改造完成之前，需要提出并实施短期管理控制措施来弥补这一差距。",
    12: "无需进一步跟进。但请确保关键软管的备用软管已下单或已到货，以备需要更换时使用。",
    13: "请更新程序，确保今后新安装的软管在使用前用溶剂冲洗（在用于工艺之前，最后流经软管的不应是水/甲醇）。",
    14: "预期是在各步骤 PPQ 开始前更换为光滑内壁软管。如无法做到，请说明短期内哪些位置仍将使用波纹软管。"
        "请说明接触工艺液/溶剂的软管材质（MOC）。是硅胶吗？我们原先的理解是这些软管为 PTFE 材质，只有水管是硅胶的。",
    20: "对该方案无需进一步跟进。但请说明 F 列中两个“Step 3”的区别。",
    21: "基于近期偏差中的观察，需要进一步讨论如何应对未钝化的溶剂高位槽、管道以及焊接质量方面的问题。"
        "钝化溶剂总管和/或通过目视检查或内窥镜抽查部分焊缝是否生锈，需要多长时间？残留物调查中涉及的焊缝是否为手工焊？",
    22: "需要明确的是，我们的期望是由生产部门专门记录物料在室外卸货区的接收时间及转入厂房的时间。"
        "这样可以追溯物料暴露于室外环境的时长。物料在室外的时间应尽量缩短。",
    23: "无需进一步跟进。",
    24: "JOB：Lilly 将审阅 CCS（污染控制策略）并提供反馈。",
    25: "无需进一步跟进。",
    26: "AJM：另外，请确认在连接水管时已考虑使用后排空软管，以防断开连接时水洒到设备/地面上。",
    27: "AJM：请尽早提供一份示例，说明最终的形式，以便 Lilly 团队提供反馈。示例中请涵盖多种过滤器类型"
        "（仅用于颗粒控制、用于微生物负荷控制、原料/公用工程等），以便我们给出有意义的反馈。",
    28: "AJM：请确认这适用于整个工艺中的所有原料投料（Step 1、2、3A/3B）。10 微米更为合理，但仍与你们分享的其他颗粒控制过滤器的指南不一致"
        "（按你们另一份指南为 0.45 微米）。现场我们曾讨论使用类似 Parker ZCTP1-045C-N-PP Fluroplus 的过滤器，你们决定不采用该方案是出于什么原因？",
    29: "AJM：没有其他顾虑。\n\nJOB：我们是否有溶解度数据，支持用水和 IPA 擦拭能够去除潜在的残留物料？还是说目视洁净是唯一的控制手段？",
    30: "无需进一步跟进，此项视为关闭。",
    32: "是否有通用的程序/流程，描述 Asymchem 在换产和隔离方面的做法？\n\nJOB：Lilly 将审阅 CCS（污染控制策略）并提供反馈。",
    33: "请确认透气膜与所有工艺液/溶剂均相容。如不相容，且工艺液/溶剂不促进微生物生长，请改用洁净的实心盖。",
    34: "AJM：请确认 PD 泵（容积泵）两侧均设有低点排放口。",
    35: NO_FOLLOW_UP,
    36: NO_FOLLOW_UP,
    37: "AJM：请更详细地解释 F 列中各 Option 的含义。如有可能，请分享进展视频，以便 Lilly 直观确认是否可接受。"
        "另外，如有需要，请在现场安装前发送设计图纸供我们确认。我最关心的是你们将如何解决液位变送器、温度元件这类不理想的仪表，"
        "因为对我来说，这些不像其他一些事项那样一目了然。",
    38: "JOB：更新清洁风险评估。",
    39: "请说明 Step 1/Step 2 的工艺用水或清洁用水供应是否存在类似差距。\n\n"
        "另外，请通过视频向我们更新改造进展。还请确认在等待零死角阀期间能否先采用临时方案（使用软管/便携式流量计，"
        "或仅使用最短距离三通/普通阀门等）。我们需要了解在长期改造到位之前将采取哪些风险缓解措施。",
    40: "AJM：再确认一下：只有在向促进微生物生长的溶液中加水时，我们才要求在水管连接处下游安装这些过滤器。",
    41: "AJM：见上方意见。我们需要了解临时方案的具体形式，以及短期内能在更贴近使用端的位置产生哪些数据（简化验证/专项取样方案等）。",
    42: "AJM：请告知何时可提供给我们审阅并提出意见。我们希望确保本工艺中所有重复使用（非新购）的设备都已充分完成换产/清洁/更换。"
        "能否按步骤提供一份清单，说明哪些设备曾用于其他工艺？",
    43: "AJM：请告知何时可提供给我们审阅。",
    44: "AJM：请确保将喷淋覆盖测试结果（时间、压力、擦拭取样位置等）直接纳入清洁记录/SOP。",
    45: "无需进一步跟进。",
    46: "AJM：请告知何时提供该程序。",
    48: REVIEW_REVERT,
    49: "AJM：请告知何时提供该策略。",
    50: REVIEW_REVERT,
    52: "AJM：请确认将在现场存放位置设置 6 英寸高度的目视标识（如涂刷标线等）。",
    53: "按所提供的时间表，这是否意味着我们计划在这些装置尚未到位的情况下开始 PPQ？\n\n"
        "JOB：增加 SBV 需要相应更新清洁文件、更新 MOR 操作指令，",
    54: "AJM：没看明白这里的回复。这只是纠正现有隔膜阀的一个简单动作。",
    55: "AJM：没看明白这个回复。我希望看到一份该工艺的取样图（列出每一个需要取的样品），"
        "然后确认具备相应的取样能力——即每个样品将从哪里取（阀门编号、位置）。",
    56: "AJM：既然该安装只能在 PPQ 开始后进行，PPQ 期间的临时方案是什么？",
}


def bilingual(en: str, zh: str) -> str:
    en = en.rstrip()
    return f"{en}\n\n{zh}" if "\n" in en else f"{en}\n{zh}"


def text_lines(text: str, width: float) -> int:
    n = 0
    for para in text.split("\n"):
        units = sum(2.0 if CJK.match(ch) else 1.1 for ch in para)
        n += max(1, math.ceil(units / max(width - 1.5, 4)))
    return n


def unescape(s: str) -> str:
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def main() -> None:
    with zipfile.ZipFile(SRC) as z:
        parts = {i.filename: z.read(i.filename) for i in z.infolist()}
        infos = z.infolist()
    sheet = parts[SHEET].decode("utf8")
    sst = parts[SST].decode("utf8")

    sis = re.findall(r"<si>.*?</si>", sst, re.S)
    uses: dict[int, list[str]] = {}
    for name, data in parts.items():
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            for ref, idx in re.findall(r'<c r="([A-Z]+\d+)"[^>]*t="s"[^>]*><v>(\d+)</v>', data.decode("utf8")):
                uses.setdefault(int(idx), []).append(f"{name}!{ref}")

    j_cells = {int(r): int(i) for r, i in re.findall(r'<c r="J(\d+)"[^>]*t="s"[^>]*><v>(\d+)</v>', sheet) if r != "1"}
    assert set(j_cells) == set(ZH), sorted(set(j_cells) ^ set(ZH))

    new_text: dict[int, str] = {}
    for row, idx in j_cells.items():
        body = sis[idx]
        assert "<r>" not in body, f"J{row} is rich text"
        en = unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S)))
        assert all(u.startswith(SHEET + "!J") for u in uses[idx]), f"J{row} string shared with {uses[idx]}"
        text = bilingual(en, ZH[row])
        assert new_text.get(idx, text) == text, f"J{row}: same source string, different translation"
        new_text[idx] = text

    for idx, text in new_text.items():
        sis[idx] = f'<si><t xml:space="preserve">{escape(text)}</t></si>'
    head = sst[: sst.index("<si>")]
    tail = sst[sst.rindex("</si>") + len("</si>"):]
    parts[SST] = (head + "".join(sis) + tail).encode("utf8")

    sheet, n = re.subn(r'(<col min="10" max="10" width=")[\d.]+(")', rf"\g<1>{J_WIDTH}\g<2>", sheet)
    assert n == 1
    raised = []
    for row, idx in j_cells.items():
        need = min(text_lines(new_text[idx], J_WIDTH) * 16.5 + 10, MAX_HT)
        m = re.search(rf'<row r="{row}"[^>]*>', sheet)
        tag = m.group(0)
        cur = float(re.search(r' ht="([\d.]+)"', tag).group(1))
        if need > cur:
            new_tag = re.sub(r' ht="[\d.]+"', f' ht="{need:g}"', tag)
            if "customHeight" not in new_tag:
                new_tag = new_tag.replace(">", ' customHeight="1">', 1)
            sheet = sheet[: m.start()] + new_tag + sheet[m.end():]
            raised.append((row, cur, need))
    parts[SHEET] = sheet.encode("utf8")

    OUT.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, parts[info.filename])
    print(f"{len(j_cells)} J cells made bilingual, J width {J_WIDTH}, {len(raised)} rows raised:")
    for row, cur, need in raised:
        print(f"  row {row}: {cur:g} -> {need:g}")


if __name__ == "__main__":
    main()
