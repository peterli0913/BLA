#!/usr/bin/env python3
"""Add review comments to the L2 Step 1 PPQ protocol. Body text and existing comments are not edited."""

import hashlib
import shutil
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from lxml import etree

SRC = Path("/workspace/2. L2 Step1 PVP-1400200301.01_-10-24-2026. (5).docx")
DST = Path("/workspace/2. L2 Step1 PVP-1400200301.01_-10-24-2026-AI小鹰批注.docx")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
XML = "http://www.w3.org/XML/1998/namespace"

NSMAP_W = {"w": W}

# (anchor, occurrence, comment). Occurrence is among paragraphs whose accepted text contains the anchor.
COMMENTS = [
    (
        "peptide fragment s product",
        0,
        "首段把终产品 ASYM-140020 / LSN4182924（Eloralintide Fragment L2）的化学名和树脂上中间体 ASYM-141366 接在同一句里，句末 “peptide fragment s product” 是残句，读不成 “fragment” 或 “fragment product”。"
        "请按 L3 PVP 的写法拆开：先写 Fragment L2 是哪一个化合物，再写 ASYM-141366 是它的 peptide-on-resin 中间体。化学名本身不要改成树脂上产物的名字。",
    ),
    (
        "2- Chlorotritly chloride",
        0,
        "2-CTC 的英文写错了。这里是 “2- Chlorotritly chloride”，Table 4 又写成 “2-Chlorotritlyl chloride”。正确拼写是 2-chlorotrityl chloride。后文 Regulatory Starting Material 一节的 “2-Chlorotrityl Chloride” 是对的，三处请统一。",
    ),
    (
        "will consistently produce.",
        0,
        "目标句在 produce 后面被句号截断：“will consistently produce. ASYM-141366 meets its IPC acceptance criteria”。"
        "读者读不到“稳定生产出满足 IPC、并可进入下游的 ASYM-141366”。建议改成 will consistently produce ASYM-141366 that meets its IPC acceptance criteria and is suitable for forward processing。"
        "L3 模板在 produce 和产品代码之间也没有把句子接顺，这里又多了一个句号，不能直接沿用。",
    ),
    (
        "Asychem TJ4",
        0,
        "本段有三处需要改。"
        "（1）Asychem TJ4 少了一个 m，应为 Asymchem TJ4。"
        "（2）范围写本方案评价 CPPs and CQAs，但第 9 节明确 ASYM-141366 不制定 CQA。L3 对应句写的是 CPPs and IPCs，与“树脂上中间体无放行标准”一致。这里应改为 CPPs and IPCs。"
        "（3）括号里的 Sections 16–18 对不上目录：稳定性是第 15 节，均匀性研究是第 16 节，清洁验证是第 17 节，第 18 节是变更和偏差。这句话点名的是清洁验证、稳定性和支持性研究，不能写成 16–18。",
    ),
    (
        "release specifications",
        0,
        "本句写 PPQ 要证明物料 meets its predefined quality attributes, in-process requirements and release specifications。"
        "同一方案第 1 节和第 9 节已经写明 ASYM-141366 没有批准的放行标准，成功标准是 IPC 和 CPP，不是 release specification。"
        "L3 这一段已经改成 meets its IPC acceptance criteria and is suitable for forward processing。L2 这里仍留着放行标准，和自身结论矛盾。",
    ),
    (
        "The personals who are executing",
        0,
        "personals 不是人员的英文，应为 personnel。后半句 shall be attached report 也不通，培训签到应附在报告里，建议写成 The training signature record shall be attached to the report。"
        "执行本方案的人员与 Appendix 1 现行版 SOP 的关系要保留。",
    ),
    (
        "105.63 kg",
        0,
        "批产量三套数字对不上，而且按本表的注算不回输出范围。"
        "Table 4 输出是 105.63–278.62 kg，收率 60%–130%，投料 47.90–48.10 kg。注中说输出下限 = 最低投料 × 最低收率、上限 = 最高投料 × 最高收率。若收率是相对树脂投料的质量收率，下限约为 28.7 kg，不是 105.63 kg；若收率是相对同一个理论产量，105.63/0.60 = 176.05 kg，278.62/1.30 = 214.32 kg，两个理论产量也不相同。"
        "页眉 header2 和 header6 仍保留 Batch Size: 145.12–271.65 kg (output)，与 Table 4 又是第三套数字。既有批注曾要求去掉页眉批产量，这两处页眉还在。"
        "批号从 CPo140020-03-01-02 起到 -04，没有 -01。请确认 -01 是否另有用途，并在三个批次上使用同一套、可复核的投料、产量和收率。"
        "三批的投料、产量和收率是纵向合并单元格，改数字时要同时改合并区域，避免只改到第一批。",
    ),
    (
        "-14~57",
        0,
        "Sy15 的设备能力写了两个互相矛盾的温度控制范围：-14~57 °C 和 -4~53 °C。ESA 对同一台设备只写 -4~53 °C，没有 -14~57 °C。请删掉多出来的一组，并改成一次 “control range”。"
        "ESA 还写了搅拌桨在最高转速下的最小体积 51.84 L，本表只保留了静止时的 51.54 L。若最高转速下的值仍有效，应补上。"
        "确认文件号也不一致：本表 Sy15 和 R01 都引用 SVP-TJ4-334.01-089，ESA 引用的是 VR-TJ4-334.01-065.02。请明确哪一份是现行 IQ/OQ/PQ 报告，不能两份同时作为依据而不说明替代关系。",
    ),
    (
        "0-119.84",
        0,
        "TJ4S-1200-AT02 的搅拌能力写成 0–119.84 rpm。ESA 对这台釜写的是确认最大搅拌转速 125 rpm，没有 119.84 这个数。119.84 看起来像计算残留，不像设备确认范围。请改回确认报告中的范围。",
    ),
    (
        "Glass-line Reactor",
        0,
        "Glass-line Reactor 应写成 Glass-lined reactor。ESA 和本方案后文对 TJ4S-1213-R22 用的也是 glass-lined。设备名称会进入确认和批记录，不要留 line。",
    ),
    (
        "Mass flow mete",
        0,
        "Mass flow mete 少了 r，应为 mass flow meter。活化釜接触材料一栏里同一错误还出现一次，请一并改。",
    ),
    (
        "bottom valve of tank is opened",
        0,
        "分装和溶剂确认这一步，和 ESA 的工艺主体相比少了两类操作，不只是参数调整。"
        "ESA 写氨基酸、Oxyma、PyBOP 和 2-CTC 树脂在手套箱 TJ4S-1200-GB01 至 GB05 中分装；本方案设备表没有这些手套箱，正文也没写分装方式。若分装仍在 PPQ 范围内，设备表和本步都要补上；若已改到别处，应写明。"
        "ESA 有单独的 20% piperidine/DMF 配制（向 TJ4S-1213-SPS01-AT01/AT02 加 DMF 和哌啶后搅拌待用）。这里只写成溶液已配制、底阀已打开。设备表对这两台釜的功能仍是 Prepared piperidine/DMF solution。配制步骤和设备功能不一致。"
        "列表里是 PyBop，Phe23 活化步是 PyBOP。请统一为 PyBOP。",
    ),
    (
        "proceed to next operation if Water content",
        0,
        "含水判定两句粘在一起：“proceed to next operation if Water content >500 ppm, repeat”。合格和不合格之间缺句号，不合格时也没写重复哪一步。ESA 写的是重复 Operation 3，直到 ≤500 ppm。请拆成两句，并写明重复的是 DMF 冲洗和取样，而不是整段预处理里含糊的 repeat。",
    ),
    (
        "Fmoc-Gly-OH coupling",
        0,
        "No. 列几乎是空的，但本行是 5，后面两行 Re-swelling / IPAc 洗涤是 1。这三个数字对不上氨基酸循环，也不像步骤号，像修订后留下的残值。执行时可能被当成步骤号。请整列重排，或删掉没有编号规则的数字。",
    ),
    (
        "110 -150 min",
        0,
        "标准 De-Fmoc 写成 110–150 min（目标 2 h）。目标 120 min 不在这个区间的中点。现行参数范围表这一项的 NOR 是 110–130 min，PAR 是 68–180 min，目标 120 min。150 min 落在 PAR 里，但不在 NOR 里。"
        "同一句在后面每个标准 De-Fmoc 都重复了。请按 110–130 min（目标 120 min）统一，除非另有已批准的范围替换参数范围表。Pre-Ser20 的 11–36 min 和 20–51 min 与参数范围表一致，不要改成这一条。",
    ),
    (
        "Start timing when temperature reaches 20 °C",
        0,
        "Phe23 的 NOR 已改为 19–28 °C（与 Table 7 一致），但计时仍写“温度达到 20 °C 才开始”。若批次控制在 19 °C，计时条件不会出现。ESA 旧范围是 25±5 °C，20 °C 正好是下限，所以旧句说得通。范围改完后这句要一起改，例如从进入 19 °C 开始计时，或把计时点和 NOR 下限写成同一个数。",
    ),
    (
        "Incomplete coupling, 6.5 h",
        0,
        "NMeAsn22 的偶联时间被写进了茚三酮判定的括号里：“Blue to Purple Positive = Incomplete coupling, 6.5 h ≤ Coupling time ≤ 13.5 h”。时间窗口看起来像颜色标准的一部分。Phe23 那一行是把时间单独成句的，请改成同样的句子结构。6.5–13.5 h 本身与参数范围表一致，不必改数值。",
    ),
    (
        "13.5 h ≤ total time ≤ 27.5 h",
        0,
        "Asn21 先写总时间到 8 h ± 15 min 后补加 Oxyma 和 DIC，再写 5 h ± 15 min，最后又写总时间 13.5–27.5 h。两段上限相加大约是 13.5 h，解释不了上限 27.5 h。本步没有茚三酮或溴酚蓝终点，IPC 表也不包含 Asn21，27.5 h 不是“延长到阴性”的窗口。"
        "参数范围表同时有 Coupling time 1 = 8 h、Coupling time 2 = 5 h，以及另一条 NLT 780 min、PAR 780–1680 min（13–28 h）。13.5–27.5 h 和 13–28 h 也不一致。请统一总时间定义：是两段固定时间之和，还是允许延长的总时钟，并写清到点后做什么。",
    ),
    (
        "Add DMF (9.0 V, 8.53 kg/kg)  into the synthesis reactor",
        0,
        "这是 Asn21 醋酐封端之后的 DMF 洗涤，写成 9.0 V / 8.53 kg/kg。Table 7 的 Post Asn21 capping DMF wash NOR 是 7.13–8.03 kg/kg，PAR 为 NLT 6.75，目标 NLT 7.58。8.53 高于 NOR 上限。"
        "参数范围表这一步也是 NLT 7.58 kg/kg，NOR 约 7.20–7.96 kg/kg。9.0 V / 8.53 kg/kg 是前面甲醇封端后洗涤（以及参数范围表 DMF post-coupling wash）的数值，不像这一步。ESA 旧工艺这里是 8.0 V / 7.584 kg/kg，与目标 NLT 7.58 相符。请改回这一步自己的体积，不要沿用 9.0 V。",
    ),
    (
        "9.68 kg)",
        0,
        "Pre-Ser20 第二釜写成 “10.39 V, 9.68 kg”，单位少了 /kg。按 48 kg 树脂理解，9.68 kg 总量只有约 0.2 V，和 10.39 V 矛盾。参数范围表这一项是 10.39 V、9.68 kg/kg，NOR 约 9.20–10.16 kg/kg。体积和当量是相对 ESA 两釜都为 8.0 V 的有意调整，请只补单位，不要改回 8.0 V。11–36 min 和 20–51 min 与参数范围表一致。",
    ),
    (
        "colorless, react completely",
        0,
        "第一个 Ser（Ser20）的茚三酮合格标准写成 “colorless, react completely”。同方案 IPC 表、Appendix 3 以及下一个 Ser 偶联写的都是 Yellow to light pink = negative，Blue to purple = positive。colorless 和 react completely 都对不上方法，react completely 也不是检验结论的写法。请改成与 AM-140020-029 和其他偶联步相同的颜色判定。",
    ),
    (
        "target: 20.0 °C, activate",
        0,
        "活化温度的括号没有关上，后面的 activate 被装进了 target 括号：“at 20.0 ± 5.0 °C (target: 20.0 °C, activate for 20±5 min”。同一残句在 Ser、Arg、Val、Leu、Glu、Ala、Leu 和 Orn 的活化步重复。请改成 at 20.0 ± 5.0 °C (target: 20.0 °C). Activate for 20 ± 5 min (target: 20 min) at 20.0 ± 5.0 °C (target: 20.0 °C)。数值不用改。",
    ),
    (
        "9.60 kg/kg))",
        0,
        "这是 Ser20 脱保护后的常规 IPAc 洗涤，写成 11.0 V / 9.60 kg/kg，末尾多了一个右括号。"
        "参数范围表里，常规脱保护后 IPAc 是 10.0 V / 8.73 kg/kg（NOR 约 8.38–9.08）。11.0 V / 9.60 kg/kg 只用于 Orn11 偶联之后的 5 次 IPAc。本行前后其他脱保护后洗涤也都是 10.0 V / 8.73 kg/kg。请改回 10.0 V / 8.73 kg/kg，并去掉多余括号。",
    ),
    (
        "Starting with the 10th wash",
        0,
        "Leu12 脱保护后的 10 次 IPAc 写成 11.0 V / 9.60 kg/kg。参数范围表这 10 次仍是 10.0 V / 8.73 kg/kg；11.0 V 是 Orn 偶联后的洗涤，不是这一步。"
        "句子停在 “lighter or consistent with the reference solution and then”，没有写合格后进入下一步，也不完整。IPC 表写的是颜色深于对照时再洗两次 IPAc，直到浅于或一致于对照。请把体积改回 10.0 V，并把 and then 后面的动作写完。",
    ),
    (
        "Stirring and turn off the nitrogen",
        0,
        "“Bubble with nitrogen, then start the agitator. Stirring and turn off the nitrogen.” 缺少搅拌时间，Stirring 也不是谓语。再溶胀的其他行写的是 after stirring。同样的断句在 Leu12 前的再溶胀又出现一次。请改成和其他再溶胀行一样的完整句子；若搅拌时间仍要控制，把时间写上。",
    ),
    (
        "Heat the mixture to 25.0 ± 5.0 °C. The mixture is allowed to react at 25.0 ± 5.0 °C",
        0,
        "Orn11 偶联温度写成 25.0 ± 5.0 °C，即 20–30 °C。标准氨基酸的 Table 7 NOR 是 20.0–27.0 °C，PAR 是 18.6–28.4 °C。30 °C 超出 NOR，也超出 PAR。Ser、Arg、Val、Leu、Glu、Ala 和第二个 Leu 都已经写成 20.0–27.0 °C。这里像是 ESA 旧的 25±5 °C 没有改完。请与标准氨基酸 NOR 对齐。",
    ),
    (
        "Fmoc-L- Orn(Boc)-OH (2.00 eq.) coupling",
        0,
        "物料名称中间多了一个空格：Fmoc-L- Orn(Boc)-OH。投料句末尾的 coupling 是多出来的，这一行是活化和转移，不是偶联。建议改为 Add Fmoc-L-Orn(Boc)-OH (2.00 eq.)。2.00 eq 与参数范围表一致，不要改当量。",
    ),
    (
        "repeat Step 93",
        0,
        "“repeat Step 93” 对不上任何步骤号。本表 No. 列没有 93，93 只是表的行序。IPC 表写的复测单位是 1 次甲苯加 2 次 MTBE，直到第二次 MTBE 洗涤液的氯冉酸颜色浅于或一致于 500 ppm 对照。"
        "“during the second stirring” 也没写清是第二次 MTBE，还是这一行里的第二次搅拌。合格时的句子停在 “If the color is lighter or consistent with the reference solution,”，没有写进入干燥。请按 IPC 表把不合格时的重复单元和合格后的去向写完整，不要引用 Step 93。",
    ),
    (
        "activation tank TJ4S-1213- R22",
        0,
        "TJ4S-1213-R22 在这里被叫做 activation tank。这台是搪玻璃釜，用来配制醋酐/吡啶/DMF，不是活化釜。活化釜是 TJ4S-1213-Sy15-R01。ESA 这一句写的是 Glass-lined reactor (TJ4S-1213-R22)。设备位号中间还有一个空格 “1213- R22”。请改名称并去掉空格。",
    ),
    (
        "Sweep with nitrogen, and vacuum underneath",
        0,
        "干燥相对 ESA 少了两项操作，不是单纯改参数：上口氮气流量 11–15 m³/h（目标 14 m³/h），以及干燥 6 h 后每 4±1 h 以 7±2 rpm 搅拌 2–5 min。参数范围表这一步也只保留夹套 T ≤ 30 °C。若 MOR 仍要求氮气流量和间歇搅拌，这里漏了；若已经取消，应写明，避免和 ESA / MOR-TJ414002001.04 冲突。"
        "LOD ≤9% (w/w) 与 IPC 表一致，相对 ESA 的 ≤10% 是收紧后的标准，不要改回 10%。复测间隔若仍是每 6±4 h，也请写上。",
    ),
    (
        "IPAc wash quantity",
        0,
        "Table 7 和 Appendix 4 应是同一份 CPP 清单，但下面四行的 NOR 不一致：Capping DMF wash 8.06–9.00 对 8.02–9.04 kg/kg；IPAc 8.21–9.25 对 9.02–10.18 kg/kg；Toluene 9.16–10.02 对 9.01–10.17 kg/kg；MTBE 7.78–8.52 对 7.66–8.64 kg/kg。其余行一致。"
        "IPAc 这一行还有范围逻辑错误。Table 7 的 PAR 是 NLT 8.73，NOR 却从 8.21 开始，NOR 没有落在 PAR 里面。低端按 NOR 操作会不合格。Appendix 4 的 IPAc NOR 从 9.02 开始，又包不住正文里的常规目标 8.73 kg/kg。"
        "实际上有两套 IPAc：脱保护后 10.0 V / 8.73 kg/kg，Orn 偶联后 11.0 V / 9.60 kg/kg。合成一行 CPP 哪一套都盖不全。Pre-Ser20 两次哌啶时间也一样：正文和参数范围表是 11–36 min、20–51 min，两张 CPP 表的 NOR 却是 10–60 min 和 15–75 min。"
        "请按参数范围表把 Table 7 和 Appendix 4 收成同一套，并把两种 IPAc 分成两行，保证每一行都满足 NOR 在 PAR 之内、目标在 NOR 之内。",
    ),
    (
        "8.02 – 9.04",
        0,
        "Appendix 4 的 NOR 与 Table 7 不一致，见 Table 7 中 IPAc wash quantity 一行的批注。汇总表如果先被拿去填 PPQ 数据，会用错窗口。请与 Table 7 用同一套 PAR 和 NOR，不要只改其中一张。",
    ),
    (
        "RPT-1190340",
        0,
        "PAR 报告号前后不一致。这里是 RPT-1190340，Appendix 2 第 3 项是 RPT-1194035。两个号只是数字换了位置，必有一个是错的。句式 “<Refer to RPT- (Lilly document number: RPT-1190340.) -and” 也不完整。请核对 Lilly 文件号后只保留一个，并补上 Asymchem 报告号。",
    ),
    (
        "approved Master Operation Record (MOR-TJ414002001.04)",
        0,
        "接受标准要求三批按 MOR-TJ414002001.04 生产。ESA 写明它评估的就是这份 MOR 里的改前工艺。本方案 Table 6 已经使用改后的参数（例如 Gly 偶联 3.5 h、标准 De-Fmoc 时间、Pre-Ser20 第二釜 10.39 V、Orn 后 IPAc 11.0 V、LOD ≤9%）。"
        "若执行时仍用 .04，批记录会和 Table 6、Table 7 不一致；若 MOR 已升版，接受标准和 Appendix 2 都应改成新版本。PPQ 不能同时把旧版 MOR 和改后工艺描述都定为成功标准。",
    ),
    (
        "YL003110",
        0,
        "Fmoc-L-Val-OH 的质量标准号写成 SP-138518-02。138518 是 Fmoc-L-Arg(Pbf)-OH 的物料代码，Arg 那一行已经使用 SP-138518-02。Table 1 中 Val 的 Asymchem 代码是 ASYM-138202，按本表其他氨基酸的对应关系，Val 应为 SP-138202-02。请核对受控标准号，不要沿用 Arg 的号码。",
    ),
    (
        "ethod Number",
        0,
        "表头 “ethod Number/” 少了 M，QualifiedLab、EffectiveDate 也粘在一起。"
        "更重要的是：接受修订后，数据行是空的。被删掉的内容仍是 L3 的方法，例如 AM-140021-011、AM-140021-031，以及 Fmoc-L-Cys(Trt)-OH、Boc-L-Glu-OtBu、ASYM-140850。这些不是 L2。"
        "正文已经引用 L2 的 IPC 方法 AM-YL000336-004、AM-140020-029、AM-140020-031、AM-140020-030 和 AM-140020-016。请用 L2 的 RSM、2-CTC 和 IPC 方法重填此表，不要保留 L3 行，也不要留空表。",
    ),
    (
        "simple technique that require",
        0,
        "“it is simple technique that require no validation or verification” 语法不通，也和前文矛盾。第 3 节要求 qualified analytical methods，第 11 节要求全部 IPC 方法在 PV 开始前完成验证，并点名了 AM-140020-016-VP-01。若方法确认结论确实是 LOD 只需核对使用，请引用那个结论，不要在这里写成“因为简单所以不需要验证”。",
    ),
    (
        "Orn11 post coupling wash",
        0,
        "括号没有关上：“Chloranil test (Orn11 post coupling wash piperidine residue.” Appendix 3 同一条目也缺右括号。不合格措施写“1 次甲苯 + 2 次 MTBE”，与工艺描述里的 “repeat Step 93” 不是同一套指令。请补括号，并和 Table 6 最后两次洗涤使用同一套复洗动作。",
    ),
    (
        "PSP-140020-2026-002.01",
        0,
        "稳定性方案号不一致。本节是 PSP-140020-2026-002.01，Table 10 和 Appendix 2 第 9 项都是 PSP-140020-2026-003.01。保留样品和稳定性监测必须指向同一份现行方案。句末还有一个多余句号。",
    ),
    (
        "after release",
        0,
        "“packaged … after release” 和本物料“没有放行标准、只允许放行至下一步”的结论不一致。建议写成 after QA releases the batch for further processing。"
        "本句在 release 和 Packaging 之间也缺句号。"
        "干燥剂规则夹了中文标点，NLT 和 ≥ 重复：“Net weight ≤ 1kg，NLT desiccant ≥ 10% of Net weight”。请改成英文全句，并与 Table 11 的 100 g 硅胶规格一致：≤1 kg 时干燥剂不少于净重的 10%，>1 kg 时为净重的 1% 加 100 g。",
    ),
    (
        "HPDE drum",
        0,
        "HPDE 应为 HDPE（high-density polyethylene）。外包装正文和 Table 11 各有一处。ESA 收集步骤用的也是 HDPE drums。",
    ),
    (
        "CVP-1400200301,",
        0,
        "正文写 CVP-1400200301，Appendix 2 写 CVP-1400200301.01。清洁验证被接受标准引用，版本号要写全。也请确认 .01 是否就是现行版。",
    ),
    (
        "Quality argeement",
        0,
        "Quality argeement 拼写错误，应为 Quality Agreement。文件号 QAA03550.V01 在后文是对的，请只改这个词。",
    ),
    (
        "registration starting materials",
        0,
        "registration starting materials 应改为 regulatory starting materials。第 1 节和第 2 节用的已经是 regulatory starting materials。注册语境里的起始物料是 regulatory starting material，不是 registration starting material。",
    ),
    (
        "Fu Mongolia county",
        0,
        "Fu Mongolia county 会被读成蒙古国，不是阜蒙县。辽宁凯莱英地址中的阜蒙县应写成 Fuxin Mongol Autonomous County (Fumeng County), Fuxin, Liaoning。Table 3 和后面各供应商地址重复了同一句，请全部替换。",
    ),
    (
        "ASYM-1400020",
        0,
        "下一步产物写成 ASYM-1400020，多了一个 0。Fragment L2 的代码是 ASYM-140020 / LSN4182924。"
        "同一句说 Table 6 代替放行检验。Table 6 是工艺描述，IPC 和接收标准在 Table 9。下一段已经正确引用 Table 9。这里应改为 Table 9。"
        "包装和贮存写成 Section 16，目录里包装是第 14 节，第 16 节是均匀性。稳定性放在 Section 15 是对的。清洁验证的 Section 17 也是对的。",
    ),
    (
        "Section 16",
        0,
        "“refer to Section 20” 指错了。第 20 节是总结报告本身。放行至下一步的条件在第 21 节 Batch Release and Disposition。请改成 Section 21。",
    ),
]


def accepted_text(p):
    parts = []

    def walk(n):
        if n.tag == f"{{{W}}}del":
            return
        if n.tag == f"{{{W}}}t" and n.text:
            parts.append(n.text)
        for c in n:
            walk(c)

    walk(p)
    return "".join(parts)


def text_hash(root):
    chunks = []
    for node in root.iter():
        if node.tag in (f"{{{W}}}t", f"{{{W}}}delText") and node.text:
            chunks.append(node.text)
    return hashlib.sha256("".join(chunks).encode()).hexdigest()


def existing_para_ids(roots):
    found = set()
    for root in roots:
        for el in root.iter():
            for attr, val in el.attrib.items():
                if attr.endswith("paraId"):
                    found.add(val)
    return found


def main():
    if DST.exists():
        DST.unlink()
    shutil.copy(SRC, DST)

    with zipfile.ZipFile(DST, "r") as z:
        files = {name: z.read(name) for name in z.namelist()}

    doc = etree.fromstring(files["word/document.xml"])
    comments = etree.fromstring(files["word/comments.xml"])
    extended = etree.fromstring(files["word/commentsExtended.xml"])
    people = etree.fromstring(files["word/people.xml"])
    before = text_hash(doc)
    old_comment_ids = [c.get(f"{{{W}}}id") for c in comments.findall(f"{{{W}}}comment")]
    old_comment_text = etree.tostring(comments)

    paras = list(doc.iter(f"{{{W}}}p"))
    used = existing_para_ids([doc, comments, extended])
    start = datetime(2026, 10, 27, 6, 8, 0)
    next_id = 30001

    for i, (anchor, occurrence, text) in enumerate(COMMENTS):
        hits = [p for p in paras if anchor in accepted_text(p)]
        if len(hits) <= occurrence:
            raise SystemExit(f"Anchor not found enough times ({len(hits)}): {anchor}")
        p = hits[occurrence]
        para_id = f"{0xA1000000 + i:08X}"
        if para_id in used:
            raise SystemExit(f"paraId collision {para_id}")
        used.add(para_id)
        cid = str(next_id)
        next_id += 1
        when = (start + timedelta(minutes=6 * i)).strftime("%Y-%m-%dT%H:%M:%SZ")
        insert_range(p, cid)
        append_comment(comments, cid, when, para_id, text)
        append_extended(extended, para_id)

    add_person(people)

    after = text_hash(doc)
    if before != after:
        raise SystemExit("Document text changed; aborting")
    new_ids = [c.get(f"{{{W}}}id") for c in comments.findall(f"{{{W}}}comment")]
    if new_ids[: len(old_comment_ids)] != old_comment_ids:
        raise SystemExit("Existing comment order changed")
    old_root = etree.fromstring(old_comment_text)
    for old in old_root.findall(f"{{{W}}}comment"):
        oid = old.get(f"{{{W}}}id")
        found = [c for c in comments.findall(f"{{{W}}}comment") if c.get(f"{{{W}}}id") == oid]
        if len(found) != 1 or etree.tostring(found[0]) != etree.tostring(old):
            raise SystemExit(f"Existing comment {oid} was modified")

    files["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/comments.xml"] = etree.tostring(comments, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/commentsExtended.xml"] = etree.tostring(extended, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/people.xml"] = etree.tostring(people, xml_declaration=True, encoding="UTF-8", standalone=True)

    tmp = DST.with_suffix(".tmp.docx")
    with zipfile.ZipFile(SRC, "r") as src, zipfile.ZipFile(tmp, "w") as out:
        for info in src.infolist():
            data = files[info.filename]
            out.writestr(info, data)
    tmp.replace(DST)
    print(f"wrote {DST} comments {len(COMMENTS)} hash_ok")


def insert_range(p, cid):
    start = etree.Element(f"{{{W}}}commentRangeStart")
    start.set(f"{{{W}}}id", cid)
    pPr = p.find(f"{{{W}}}pPr")
    if pPr is not None:
        pPr.addnext(start)
    else:
        p.insert(0, start)
    end = etree.Element(f"{{{W}}}commentRangeEnd")
    end.set(f"{{{W}}}id", cid)
    p.append(end)
    r = etree.Element(f"{{{W}}}r")
    rPr = etree.SubElement(r, f"{{{W}}}rPr")
    style = etree.SubElement(rPr, f"{{{W}}}rStyle")
    style.set(f"{{{W}}}val", "CommentReference")
    ref = etree.SubElement(r, f"{{{W}}}commentReference")
    ref.set(f"{{{W}}}id", cid)
    p.append(r)


def append_comment(comments, cid, when, para_id, text):
    c = etree.Element(f"{{{W}}}comment")
    c.set(f"{{{W}}}id", cid)
    c.set(f"{{{W}}}author", "AI小鹰")
    c.set(f"{{{W}}}date", when)
    c.set(f"{{{W}}}initials", "AI")
    p = etree.SubElement(c, f"{{{W}}}p")
    p.set(f"{{{W14}}}paraId", para_id)
    p.set(f"{{{W14}}}textId", "77777777")
    pPr = etree.SubElement(p, f"{{{W}}}pPr")
    st = etree.SubElement(pPr, f"{{{W}}}pStyle")
    st.set(f"{{{W}}}val", "CommentText2")
    rPr_p = etree.SubElement(pPr, f"{{{W}}}rPr")
    lang_p = etree.SubElement(rPr_p, f"{{{W}}}lang")
    lang_p.set(f"{{{W}}}eastAsia", "zh-CN")
    r = etree.SubElement(p, f"{{{W}}}r")
    rPr = etree.SubElement(r, f"{{{W}}}rPr")
    fonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
    fonts.set(f"{{{W}}}ascii", "Calibri")
    fonts.set(f"{{{W}}}hAnsi", "Calibri")
    fonts.set(f"{{{W}}}eastAsia", "微软雅黑")
    sz = etree.SubElement(rPr, f"{{{W}}}sz")
    sz.set(f"{{{W}}}val", "21")
    szCs = etree.SubElement(rPr, f"{{{W}}}szCs")
    szCs.set(f"{{{W}}}val", "21")
    lang = etree.SubElement(rPr, f"{{{W}}}lang")
    lang.set(f"{{{W}}}eastAsia", "zh-CN")
    t = etree.SubElement(r, f"{{{W}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML}}}space", "preserve")
    t.text = text
    comments.append(c)


def append_extended(extended, para_id):
    ex = etree.Element(f"{{{W15}}}commentEx")
    ex.set(f"{{{W15}}}paraId", para_id)
    ex.set(f"{{{W15}}}done", "0")
    extended.append(ex)


def add_person(people):
    author = "AI小鹰"
    for person in people.findall(f"{{{W15}}}person"):
        if person.get(f"{{{W15}}}author") == author:
            return
    person = etree.Element(f"{{{W15}}}person")
    person.set(f"{{{W15}}}author", author)
    info = etree.SubElement(person, f"{{{W15}}}presenceInfo")
    info.set(f"{{{W15}}}providerId", "None")
    info.set(f"{{{W15}}}userId", "AI小鹰")
    people.append(person)


if __name__ == "__main__":
    main()
