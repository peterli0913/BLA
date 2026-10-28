#!/usr/bin/env python3
"""List hard errors in the L2 Step 1 MOR (MOR-TJ414002001.04) and write them to Excel."""

import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))
from build_l2_step1_param_crosscheck import style_sheet  # noqa: E402

import openpyxl  # noqa: E402

ROOT = Path("/workspace")
MOR = ROOT / "MOR_10282026.docx"
OUT = ROOT / "L2-Step1-MOR-硬伤清单_10-28-2026.xlsx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# (类别, 严重程度, 位置, 内容（Word 中搜索）, 问题与改法)
ROWS = [
    # 与 PVP / 参数表不一致
    ("与 PVP/参数表不一致", "高", "3.5.5、3.8.5、3.11.5、3.16.8、3.19.5、3.22.5、3.25.5、3.28.5、3.31.5、3.34.5、3.37.5、3.40.5、3.43.5 脱保护后 IPAc 洗涤",
     "Add 445.92 kg",
     "PVP 和参数表是 10.0 V、8.73 kg/kg（419.04 kg），NOR 8.38–9.08 kg/kg（402.3–435.8 kg）。MOR 写 445.92 kg（9.29 kg/kg、10.64 V），★NOR 432.54–459.30 kg，两套范围几乎不重叠。按现在的 MOR 生产，批记录会超出 PVP 表 7 的 NOR。先定数，再改 13 处步骤和 ★NOR 句。"),
    ("与 PVP/参数表不一致", "高", "3.3、3.7、3.10、3.13、3.18、3.21、3.24、3.27、3.30、3.33、3.36、3.39、3.42、3.47 偶联模块",
     "Normal Operating Range (NOR) :20~30℃",
     "所有偶联的 NOR 都写 20–30 °C，正文 25 ± 5 °C。PVP 和参数表已按氨基酸区分：标准 AA 20–27 °C，Phe23 和 Asn21 19–28 °C，NMeAsn22 和 αMePhe15 19–27 °C（Gly24 才是 20–30 °C）。超温预警高限 28 °C 也高于标准 AA 的 NOR 上限 27 °C。逐个改 NOR、正文温度和预警值。"),
    ("与 PVP/参数表不一致", "高", "3.6.3、3.9.3、3.12.3、3.17.3、3.20.3、3.23.3、3.26.3、3.29.3、3.32.3、3.35.3、3.38.3、3.41.3、3.46.3 溶解",
     "Maintain stirring for NLT 15 min",
     "PVP 和参数表把“溶解夹套温度 15–25 °C”定为 CPP，MOR 溶解时不控夹套温度，也没有记录点和 ★；溶解后才调到 20 ± 5 °C（Phe23 则直接降到 5 ± 5 °C）。在加氨基酸前加“夹套调至 20 ± 5 °C”，溶解期间记录最高/最低值并标 ★。"),
    ("与 PVP/参数表不一致", "高", "3.3.5 Gly24 偶联",
     "react fo",
     "MOR 是 3 h ± 15 min，PVP 和参数表是 3.5 h ± 15 min。同句 “fo” 应为 “for”。"),
    ("与 PVP/参数表不一致", "中", "12 处标准 De-Fmoc（3.5.3、3.8.3 … 3.43.3）",
     "2 h±15 min (target: 2 h)",
     "MOR 是 105–135 min，PVP 和参数表 NOR 是 110–130 min。改为 110–130 min（目标 120 min）。"),
    ("与 PVP/参数表不一致", "高", "3.16.9、3.16.11（Ser20 后）；3.25.6、3.25.8（Val17 前）；3.40.6、3.40.8（Leu12 前）再溶胀，及原料信息表同行",
     "3.16.9 Add DMF",
     "这六步仍是 12 V、546.05 kg。PVP 和参数表：Ser20 后 13 V（12.32 kg/kg，591.36 kg）；Val17 前、Leu12 前 10 V（9.48 kg/kg，455.04 kg）。Johan（礼来）也确认 PAR 为 2 × NLT 9.48 kg/kg。"),
    ("与 PVP/参数表不一致", "高", "3.47.6–3.47.11、3.48.1–3.48.3（Orn11 偶联后），及原料信息表同行",
     "3.47.6 Synthesis reactor wash 1",
     "IPAc、甲苯、MTBE 仍是 10 V（419.04 / 418.56 / 355.68 kg），换算 8.73 / 8.72 / 7.41 kg/kg，全部低于 PVP 表 7 NOR 下限。改为 11 V：460.80 / 460.32 / 391.20 kg。"),
    ("与 PVP/参数表不一致", "中", "3.47.7–3.47.11 洗涤顺序",
     "3.47.9 Toluene addition",
     "MOR 常规路径为甲苯、MTBE、甲苯、MTBE、MTBE（3 次 MTBE），PVP 和参数表为“甲苯、MTBE、MTBE”×2（4 次 MTBE）。确认后补一次 MTBE，或改 PVP。"),
    ("与 PVP/参数表不一致", "高", "3.49.1 干燥",
     "until LOD ≤10.0%",
     "正文 LOD ≤10.0%，本 MOR 的 IPC 表和 PVP 都是 ≤9%。中文版优先，执行时会按 10%。改为 ≤9% (w/w)。"),
    ("与 PVP/参数表不一致", "高", "分析方法 / 过程中间体取样检测列表；3.13 Asn21 偶联",
     "Sampling and test list of in-process intermediate",
     "缺 Asn21 补加 Oxyma/DIC 后的氯冉酸 IPC。Johan（礼来）两次要求列为 IPC（补加前一次可作 FIO）。在 3.13 加取样步骤，IPC 表加一行，方法 AM-140020-030。"),
    ("与 PVP/参数表不一致", "中", "3.2 DIPEA ★NOR；3.4 封端 DMF ★NOR；3.16.3 / 3.16.6 Pre-Ser20 时间",
     "【5.82~6.18 eq.】",
     "MOR 用 ±3% 称量范围或 ±2 min 作 NOR：DIPEA 5.82–6.18 eq（PVP 5.70–6.30），封端 DMF 397.25–421.83 kg（PVP 8.10–8.96 kg/kg = 388.8–430.1 kg），Pre-Ser20 15 ± 2 / 30 ± 2 min（PVP 11–36 / 20–51 min）。PVP 表头写“NOR (MOR)”，两边应一致。"),
    ("与 PVP/参数表不一致", "中", "3.2.3 Gly24 溶解",
     "Maintain stirring for 1~45 min",
     "最短 1 min，PVP 要求至少 15 min，其他氨基酸 MOR 也是 NLT 15 min。统一。"),
    ("与 PVP/参数表不一致", "中", "3.7.6、3.33.5、3.36.5、3.47.5",
     "Start timing when temperature reaches 20 ℃",
     "Phe23、αMePhe15 的 NOR 下限是 19 °C，“到 20 °C 才计时”与 NOR 冲突。改为“温度进入 NOR 即开始计时”。"),
    ("与 PVP/参数表不一致", "中", "3.47 Orn11 偶联",
     "The total quantity of DMF is a CPP",
     "MOR 把“DMF 总用量”标为 CPP，PVP 表 7 没有这一项；MOR 另把“DIC 加料温度”“20% 哌啶溶液温度”标为 ★CPP，PVP 也没有。三份文件的 CPP 清单要一致。"),
    ("与 PVP/参数表不一致", "低", "3.0 预处理",
     "Add 200.00 - 300.00 kg DMF",
     "参数表为 5.21 kg/kg（250.08 kg）± 5%，PVP 写 5.50 V。MOR 范围更宽。统一写法，或在参数表注明此步按 200–300 kg。"),

    # MOR 内部：数值和逻辑
    ("内部数值/逻辑", "高", "3.46.7 Orn11 DIC 加料",
     "Add DIC(2.20 eq.)",
     "英文 2.20 eq，中文“(2.00 eq.)”。MOR 规定“本记录以中文为准”，执行时会按 2.00 eq，低于 PVP 表 7 DIC NOR 2.09–2.31 eq。中文改为 2.20 eq。（同一搜索句在 3.29.7 也有一处，那处中英一致。）"),
    ("内部数值/逻辑", "高", "3.17.7 Ser20 DIC 加料下方 ★限值",
     "XoC",
     "★上下限仍是占位符 “X oC”，应按 15–25 °C。紧接着的超温预警是 22–28 °C，本步控温却是 20±5 °C：目标 20 °C 已经低于预警低限，正常加料会一直报警，高限 28 °C 又高于 NOR 上限 25 °C。3.20.7、3.23.7、3.26.7 等同类 DIC 步的预警是 17–23 °C。这里改为 17–23 °C。"),
    ("内部数值/逻辑", "高", "3.13 Asn21 补加 DIC 下方 ★限值",
     "xx",
     "★上下限仍是占位符 “xx oC”。补实际 NOR。"),
    ("内部数值/逻辑", "高", "3.41.7 Leu12 DIC 加料下方 ★限值",
     "32 oC",
     "★限值写成 8–32 °C，其他氨基酸 DIC 加料都是 15–25 °C，且 8–32 °C 超出 PAR 12.6–26.5 °C。改为 15–25 °C。"),
    ("内部数值/逻辑", "高", "3.7、3.9、3.13、3.17、3.46、3.47 偶联温度 ★限值",
     "13 oC",
     "偶联温度的 ★上下限写成 13–37 °C，这是 Ac₂O 封端的 PAR（12.3–36.8 °C），不是偶联。偶联 PAR 为 17–30 °C 左右，NOR 见上方偶联温度一条。按氨基酸改。（3.14 配封端液的 13–37 °C 是对的。）"),
    ("内部数值/逻辑", "高", "3.9.7 NMeAsn22 DIC 加料下方超温预警",
     "Fmoc-L-N-Me-Asn(Trt)-OH activation",
     "预警写成 低限 22 °C、高限 23 °C，低限比目标 20 °C 还高，加料过程会一直报警。3.20.7 等同类步骤为 17–23 °C（3.17.7 是另一套 22–28 °C，见该行）。改为 17–23 °C。"),
    ("内部数值/逻辑", "高", "3.13.7 Asn21 偶联",
     "13 h≤total time≤28 h",
     "同一节又写 “coupling time：NMT 840 min”（14 h），两者矛盾；PVP 为 13.5–27.5 h。同句 “the mixture be prepared in Ins.3.13.4” 引用错，Oxyma 溶液在 3.13.6 配制。"),
    ("内部数值/逻辑", "高", "标准氨基酸、Glu14 偶联",
     "4 h≤Coupling Time≤10 h",
     "同一节既写 4–10 h（Glu14 写 8–18 h），又写 “coupling time：NMT 840 min”。标准 AA 的 840 min 超出 PAR 600 min；Glu14 的 840 min 又低于本节的 18 h。这些窗口是 PAR，PVP 用的是 NOR（4.5–9.5 h、6.5–13.5 h、8.5–17.5 h），两份文件不同。按氨基酸写一套上限，并说明窗口是 NOR 还是 PAR。"),
    ("内部数值/逻辑", "高", "3.3 Gly24 偶联",
     "NLT 600min",
     "英文 NLT（不少于），中文“不超过”，意思相反；且 600 min 超出 PAR 481 min。改为 NMT 481 min。"),
    ("内部数值/逻辑", "高", "3.5.5 等脱保护后 IPAc 的第二句",
     "440.34~467.58",
     "同一步又写 “Add 419.04 kg IPAc (10 V)”，配的称量范围 440.34~467.58 kg 连 419.04 都不包含；另有 4 处写 406.47~431.61 kg。3.11.5 还写成 “432.54 ~ 4459.30 kg”。一步只能有一个量和一个范围。"),
    ("内部数值/逻辑", "高", "原料信息表 3.16.4 行",
     "3.16.4",
     "原料信息表仍是 8 V、7.448 kg/kg、357.50 kg；正文 3.16.4 已改为 10.40 V、9.68 kg/kg、464.64 kg。"),
    ("内部数值/逻辑", "高", "第 7 页“过程中间体信息”、页脚、版本说明",
     "145.12",
     "三处产量不一致：页脚 105.63–278.62 kg；版本说明 145.12–271.65 kg；按 MOR 自己的公式（取代度标准 1.0–1.2 mmol/g、收率 60–130%、投料 47.90–48.10 kg）应为 108.84–271.66 kg。105.63–278.62 kg 对应的取代度约 0.96–1.24 mmol/g，超出标准。统一后同步 PVP 表 4。"),
    ("内部数值/逻辑", "高", "页眉：3.17、3.20、3.23、3.26、3.29、3.32、3.35、3.38、3.41、3.46 节（第 335、366、395、428、457、487、519、551、583、622 页起）",
     "MOR-TJ414001901.04",
     "这 10 节的页眉 MOR 号是 L1 项目的 MOR-TJ414001901.04，应为 MOR-TJ414002001.04。"),
    ("内部数值/逻辑", "高", "设备使用清单 激活罐",
     "Size /设备大小:1500 L",
     "TJ4S-1213-Sy15-R01 的材质写 Stainless Steel / 不锈钢，PVP、ESA 均为 Hastelloy C-276。"),
    ("内部数值/逻辑", "中", "设备使用清单 手套箱 4 行",
     "Glove box/手套箱",
     "只列 4 台手套箱，材质都写不锈钢。ESA 为 GB01–GB05 五台，GB04 内表面为 Halar。补齐编号和材质。"),
    ("内部数值/逻辑", "中", "设备使用清单 操作号",
     "3.13.4",
     "TJ4S-1200-AT02（配 Oxyma/DMF）和 TJ4S-1213-R22（配封端液）的操作号都写 3.13.4，应分别为 3.13.5–3.13.6 和 3.14.1–3.15.1；DIPEA 计量泵写 3.2.1，应为 3.2.4。"),
    ("内部数值/逻辑", "高", "分析方法 / 过程中间体取样检测列表",
     "the resin is yellow to orange",
     "茚三酮（AM-140020-029）的判据写成溴酚蓝的“树脂黄色至橙色”；茚三酮应为“溶液黄色至浅粉色”。氯冉酸（AM-140020-030）一行只有“比色”，没有合格标准；偶联各行的指令号为空。"),
    ("内部数值/逻辑", "高", "3.36 Glu14 取样",
     "perform testing in accordance with the ninhydrin test record",
     "Glu14 用溴酚蓝（AM-140020-031），取样说明却写按茚三酮检测记录检测（本句全文 8 处，Glu14 这一处错）。"),
    ("内部数值/逻辑", "中", "3.7.6 Phe23；3.10.5、3.18.5 的不合格处置",
     "NMT 841 min",
     "Phe23 的偶联上限中英文都写 841 min，其余氨基酸写 840 min。3.10.5（NMeAsn22）、3.18.5（Ser20）本步上限是 840 min，不合格后继续取样的处置句仍写 841 min，处置比本步上限多 1 min。Phe23 改为 840 min；两处处置句改成与本步上限相同。"),
    ("内部数值/逻辑", "中", "3.22.8 与 3.22.9",
     "3.22.9 Resin reswelling wash 2",
     "3.22.8 已经是“溶胀后洗涤2”（加 12 V DMF）。3.22.9 不再加液，只鼓氮气、搅拌、排液，标题仍写洗涤2。批记录里同一次洗涤会出现两个“洗涤2”。3.22.9 改成该步的排液记录，或并入 3.22.8。"),
    ("内部数值/逻辑", "中", "3.44 树脂洗涤",
     "3.44.1 Sample",
     "3.44.1 用了两次：前面是脱保护后 IPAc 洗涤，这里是第 2 次洗涤的氯冉酸取样。两个操作不能共用一个指令号。取样改为 3.44.2。（3.43.5 同样把洗涤和 IPC 编成同一个号。）"),
    ("内部数值/逻辑", "中", "3.28（Leu16 前脱保护）第二次再溶胀",
     "3.25.8 DMF addition",
     "3.28 节里这一步编号写成 3.25.8，应为 3.28.8（全文两处，3.28 节那处错）。"),
    ("内部数值/逻辑", "中", "3.13 Asn21；3.43；3.2.5",
     "documented in Ins3.43.2",
     "交叉引用中英不一致：3.13 英文 “Ins3.13.5”、中文“3.13.7”；3.43 一处英文 3.43.2、中文 3.42.2；3.2.5 一处英文 3.1.3、中文 3.2.1。核对后统一。"),
    ("内部数值/逻辑", "中", "版本说明 第 3 条",
     "20±10 min (target: 10 min)",
     "版本说明写封端 DIPEA 参考加料时间目标 10 min，3.4.3 正文写 20 ± 10 min（目标 20 min）。"),
    ("内部数值/逻辑", "中", "3.17.8、3.26.8 活化",
     "Maintain for NLT 15 min",
     "英文 NLT 15 min（目标 20 min），中文 20 ± 5 min；参数表 NOR 15–25 min。中文优先，英文改为 20 ± 5 min。"),
    ("内部数值/逻辑", "中", "3.0 预处理",
     "搅拌5 min±5 min",
     "英文 5 ± 2 min，中文 5 ± 5 min。"),
    ("内部数值/逻辑", "中", "3.16、3.19、3.25 等脱保护后洗涤",
     "搅拌5±2 min（目标时间：10 min）",
     "中文 ±2 min 配目标 10 min；英文同时写 “Stir for 5±2 min … After stirring for 5±2 min”，搅拌写了两遍。3.4 封端后洗涤也有英文 5 ± 2 min、中文 10 ± 2 min 的不一致。"),
    ("内部数值/逻辑", "中", "3.5 等哌啶溶液确认",
     "17.50%~22.50%",
     "浓度 NOR 17.50–22.50% 和温度 NOR 只写在中文，英文句没有数字。参数表和 PVP 也没有浓度 NOR。"),
    ("内部数值/逻辑", "低", "3.17.1、3.20.1 等",
     "95±5 rpm",
     "其余活化罐搅拌都是 95 ± 10 rpm。"),
    ("内部数值/逻辑", "低", "原料信息表 3.37.1、3.37.5、3.38.1",
     "3.37.1",
     "体积比（8、10、5）错填到“摩尔比”列。"),
    ("内部数值/逻辑", "低", "原料信息表 3.2.4 与 3.4.3",
     "129.24",
     "DIPEA 分子量写 129.25 和 129.24 两个值。"),

    # 中文为准的翻译 / 措辞
    ("中英翻译", "高", "全文（约 300 处）",
     "用于工艺性能执行",
     "“确认”被成批替换成“执行”，且 MOR 规定中文为准：如“工艺性能执行”（应为工艺性能确认，即 PPQ）、“执行物料在复检日期”、“操作人员须执行反应体系温度…”、“拉曼扫描执行”、“各类执行”、“设备执行和配方执行”、“执行合计净重在投料范围内”。全文把这类“执行”改回“确认”，“执行人及日期”除外。"),
    ("中英翻译", "中", "3.5.3、3.8.3、3.11.3",
     "滤除至基本",
     "中文“滤除至基本通过隔膜泵从合成釜中排出溶液”断句错乱；英文 “Filter until nearly no liquid outflow” 与后一句 Drain 重复。删去该半句。"),
    ("中英翻译", "中", "收料 Ins.4",
     "Particulate Drain",
     "颗粒物滤棉译成 Particulate Drain，应为 particulate filter（分装节已写对）。"),
    ("中英翻译", "中", "收料 Ins.4 转移",
     "use a cabinet when transferring",
     "中文“保温箱”，英文 cabinet。2–8 °C 转移应为 insulated box / cool box。"),
    ("中英翻译", "中", "3.5.9、3.11.9、3.31.9",
     "3.5.9 Resin reswelling wash 1",
     "英文写 wash 1，中文写“洗涤2”（3.11.9 少一个“后”）。对照同结构的 3.8.9，英文是 wash 2、中文是洗涤2，前一步是加 DMF。MOR 以中文为准，这三处英文改为 wash 2；3.11.9 中文改为“溶胀后洗涤2”。"),
    ("中英翻译", "低", "英文缩写表",
     "Qty",
     "Qty 的中文写“质量”，应为“数量”；缩写表缺 CPP、NOR、PAR、DMF、DIC、DIPEA、IPAc、MTBE、LDPE、HDPE、OEB、FIO 等正文实际用到的缩写，却列了没用到的 MP、BP。"),
    ("中英翻译", "低", "3.18、3.21、3.24、3.27 取样",
     "Kaiser test record",
     "其余偶联写 ninhydrin test record，这四处写 Kaiser test record，记录名统一。"),

    # 笔误 / 格式
    ("笔误/格式", "低", "多处",
     "Shirnk",
     "Shirnk（2 处）、soulution（3）、diaphragmgm（2）、Ddeprotection、Adust、activation tan、Pyriine（代码表）、hastelly、cleacing、Ckecked、Checeked、Checkmed（3）、Ttransfer、additon、Sloid、LDPEpowder、butunpack（2）、solutionand、drainingdiaphragm、“conten t”、“Solid-peptide peptide synthesizer”、“哈式合金”、3.46 中文标题前多一个“芴”、TJ4S-1213-Sy15-RT01。"),
    ("笔误/格式", "低", "分装各节",
     "Page 32~62 of 624",
     "文件共 674 页（30 处）。改页码或去掉总页数。"),
    ("笔误/格式", "低", "3.40.3；3.45；目录 3.8、3.25",
     "Deprotection of Fmoc-L-Leu-OH activation",
     "标题与内容不符：3.40.3 写“Leu 活化的脱保护”；3.45 只有再溶胀却叫 “(14) Deprotection”；目录 3.8、3.25 把脱保护标成下一个氨基酸名。3.43.5 编号用了两次。"),
    ("笔误/格式", "低", "3.10.5",
     "(target: 4 h) (target：4 h)",
     "目标重复写了两次。"),
    ("笔误/格式", "低", "3.6.5",
     "(target: 5 ℃.",
     "括号没关。"),
    ("笔误/格式", "低", "3.15.6",
     "Charge 364.03 kg DMF addition",
     "英文语法不通，且缺称量范围（中文有）。"),
    ("笔误/格式", "低", "过程中间体信息",
     "÷ 1000 kg/mol]",
     "公式多一个 “]”。"),
]


def body_lines():
    z = zipfile.ZipFile(MOR)
    doc = etree.fromstring(z.read("word/document.xml"))
    out = []
    for p in doc.iter(W + "p"):
        parts = []

        def walk(node):
            if node.tag == W + "del":
                return
            if node.tag == W + "t" and node.text:
                parts.append(node.text)
            for child in node:
                walk(child)

        walk(p)
        out.append("".join(parts))
    for name in z.namelist():
        if re.match(r"word/(header|footer)\d*\.xml", name):
            out.append("".join(t.text or "" for t in etree.fromstring(z.read(name)).iter(W + "t")))
    return out


def main():
    lines = body_lines()
    missing = [q for _, _, _, q, _ in ROWS if not any(q in line for line in lines)]
    if missing:
        raise SystemExit(f"Search strings not found: {missing}")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MOR 硬伤"
    style_sheet(
        ws,
        ["序号", "类别", "严重程度", "位置", "内容（Word 中搜索）", "问题与改法"],
        [6, 16, 9, 34, 36, 92],
        [(i, *row) for i, row in enumerate(ROWS, 1)],
    )
    wb.save(OUT)
    counts = {}
    for row in ROWS:
        counts[row[1]] = counts.get(row[1], 0) + 1
    print(f"wrote {OUT.name}: {len(ROWS)} rows {counts}")


if __name__ == "__main__":
    main()
