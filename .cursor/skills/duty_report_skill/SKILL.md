---
name: duty-report-skill
description: Use when handing off or continuing Asymchem BLA / L2 Step 1 CMC work, when the user says 值班, 交接, 沉淀, duty_report_skill, or asks to apply the recent review and bilingual methods. Use before a new 批注, 硬伤核对, or 中英对照工艺表.
---

# 值班方法

这份技能只记已经定下来的做法。数字和文件号以 [annotating-cmc-drafts/references/l2-step1-project-facts.md](../annotating-cmc-drafts/references/l2-step1-project-facts.md) 为准。检查清单以 [check-logic.md](../annotating-cmc-drafts/references/check-logic.md) 和 [GROK-PROMPT.md](../annotating-cmc-drafts/GROK-PROMPT.md) 为准。两处要写同一句话。

## 接到一件事

| 用户要的 | 先读 | 交付 |
|---|---|---|
| 批注、硬伤、三方核对、客户意见有没有改完 | `annotating-cmc-drafts` | 只出 Excel，不改源 Word，不在 Word 里插批注 |
| 工艺或记录的中英对照，并说内容不变 | 本技能「中英对照」一节，术语先对 `cn-en-regulatory-translation/references/glossary.md` | 新的对照稿。英文原件不覆盖、不删除 |
| 对外邮件、纪要、PPT | `client-comms`、`meeting-minutes-and-actions`、`presentation-deck` | 按那些技能 |

回复用中文。先说结论：共几条、几条高，再列会执行错的几条，最后是需要拍板的问题。

## 审文件

1. 用 `annotating-cmc-drafts/scripts/docx_lines.py` 取接受修订后的文本。先读版本说明，再和上一版做差，再读批注。
2. 依据顺序：客户 PAR / PFD → 参数表（定 NOR）→ PVP → MOR。ESA、GA、ORA、CVP 跟锁定的 PVP 对齐。下游比上游宽报高；比上游窄报中，并写进「需要你确认」。上游自己两套数，不要替用户选。
3. 客户只按邮箱认：`@lilly.com`、`@network.lilly.com`。Yiwen Zhang（yiwen.zhang@lilly.com）是礼来客户。`done=0` 不是已经关闭。回复写了 Revised 时，回到锚点句子核对。
4. 合上清单再按操作顺序读一遍。每一步只问：按中文做会得到什么数，和上游是不是同一个数，指令号是不是唯一。新发现写进同一张 Excel。
5. 用 `scripts/write_review_xlsx.py` 写 Excel。搜索句必须能在 Word 里原样找到，10–40 字，不跨段，不含 `\xa0`。校验不过就不要交。
6. 新的检查点用同一句话写回 `check-logic.md` 和 `GROK-PROMPT.md`。用户确认过的口径写进项目事实，并从「待确认」里删掉。

### 用户后来定死的报与不报

- 小数位两种都搜。4 V 与 4.0 V、25°C 与 25.0°C、2.0 eq 与 2.00 eq，只改了一处、别处还留着，要报。
- 符号用英文半角。中文或全角 ，。：；（）【】 以及 ℃（应为 °C）要报。只差空格的不报，包括 `15minutes`、`Fmoc-L- Orn`、设备号里多的空格。
- 短字符串替换按整词做。`aMe`→`αMe` 会把 Parameter、Name、same、pentamethyl 打坏。改完再搜这些词。`PV`→`PPQ` 会把 Cpv 打成 CPPQ。
- 版本说明里写明的有意改动不报。11-03 用户确认的四项不再当错误：活化时间 20±15 min；原来的 20±5°C 写成 15–25°C；IPC 写明阴性/阳性颜色；部分偶联目标改为 23°C。同一文件里还留着旧写法（例如仍写 `20±5 oC`）要报。
- 标准 De-Fmoc 写成 100–140 min（target 2 h）还没有被确认为有意放宽。先前口径是 110–130 min。列入「需要你确认」，不要擅自改回。
- 同一股物料写了两次 Transfer the mixture，且第二次在通氮气并开搅拌之后，要报。氯冉酸或茚三酮放行句停在 `and then`，或停在 `If the color is lighter`，要报。
- 20% 哌啶/DMF 锁定为 10.4 V、9.68 kg/kg。写回 10.39 V，或写成 9.68 kg 而没有 /kg，是回退，报高。密度 0.931，10.4 × 0.931 = 9.68 kg/kg，× 48.00 kg 树脂 = 464.64 kg。

### 换算时用这些密度

DMF 0.948；IPAc 0.873；甲苯 0.872；MTBE 0.741；20% 哌啶/DMF 0.931。单位是 g/mL。V × 密度 = kg/kg。kg/kg × 48.00 kg = kg。

## 中英对照

工艺表、批记录这一类执行稿：

- 以英文原件为底稿复制，再加中文。页、栏宽、边框、字体、字号、对齐和原有编号保持不动。`0020工艺.docx` 是纵向、Times New Roman 9 磅、表宽 100%、步骤用原有自动编号。
- 中文放在同一单元格、同一段落里，紧接英文另起一行。编号仍只打在英文上，不要给中文再编一条。
- 短标签可以写成 `English/中文`，清洁 SOP 中英对照是这种。工艺长句不要挤在斜杠后面。
- 稿子里只留英文和中文。不要标题、读法、蓝色字、术语表、待核实、批注或评论。批注仍是单独的 Excel。
- 数字、单位、设备位号、物料代码与英文一致。英文里的笔误留在英文里，中文照该句翻译，不在译文里改对。`blue to back` 就保留 back。
- 工艺执行稿的中文用英文半角符号，覆盖翻译技能里的全角标点。中文与英文、数字之间留空格。
- 设备叫法：合成釜 synthesis reactor，活化罐 activation tank（Sy15-R01）。封端液釜是 TJ4S-1213-R22，不要译成活化罐。DMF、IPAc、MTBE、DIPEA、DIC、OxymaPure、PyBOP 保留代码。哌啶、乙酸异丙酯、甲醇、吡啶、乙酸酐、甲苯写出中文。

邮件、纪要、幻灯仍按 `cn-en-regulatory-translation`：可以有文末术语和待核实，中文可以用全角标点。

## 还不能自己定的数

写进 Excel 的「需要你确认」，不要在正文里选边。

- 脱保护后 IPAc：10.0 V / 8.73 kg/kg，还是 MOR 的 10.64 V / 9.29 kg/kg。
- 1.0 V / 2.0 V DMF 冲洗：0.95 / 1.90 kg/kg，还是 0.94 / 1.89 kg/kg。
- 标准脱保护：100–140 min，还是 110–130 min。
- 第二次 Transfer the mixture 是淋洗液，还是应删除。
- `inner grade LDPE` 内袋按哪一级别。外袋已写 pharmaceutical-grade。

其余未决项见项目事实「待用户确认」。用户答复后，把结论写进已锁定的值，并从待确认里删掉。
