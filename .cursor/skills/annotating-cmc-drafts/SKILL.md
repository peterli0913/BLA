---
name: annotating-cmc-drafts
description: Use when reviewing a PPQ protocol (PVP), MOR, ESA, ORA, gap assessment, or process-parameter workbook for 批注, 硬伤, inconsistencies, mistranslation, or typos; when cross-checking that PVP, parameter workbook, and MOR parameters match (三方核对); or when producing a 硬伤清单 / 批注清单 Excel.
---

# 审 CMC 草稿，出硬伤 Excel

只报会改变执行、设备结论或文件引用的问题。不改源 Word，不在 Word 里插批注，结果一律是 Excel，每行带一个可以在 Word 里原样搜索的短句。

## 流程

1. 用 `scripts/docx_lines.py <文件>` 取接受修订后的文本（含表格、页眉页脚）。先读版本说明，再和上一版做差，再读批注。客户按邮箱认：`@lilly.com`、`@network.lilly.com`。
2. 按 [references/check-logic.md](references/check-logic.md) 核对。依据顺序是 PAR / PFD → 参数表 → PVP → MOR。里面有三方核对的步骤、MOR / PVP / ESA / GA 的检查清单、报与不报的界线和交付前自查。
3. 审 MOR 时先跑 `scripts/mor_auto_checks.py <MOR.docx> --resin <树脂kg>`，得到候选清单，逐条人工判断后再写入。
4. 审 L2 Step 1 时读 [references/l2-step1-project-facts.md](references/l2-step1-project-facts.md)：已锁定的数、文件号、密度和待用户确认的问题。用户答复后更新这份文件。
5. 合上清单，再按操作顺序把文件当执行文件读一遍。每一步只问：按中文做会得到什么数，和上一份文件是不是同一个数，这一步的指令号是不是唯一。这一遍新发现的，追加进同一张 Excel，不另起文件。
6. 按 [references/excel-format.md](references/excel-format.md) 写 JSON spec，运行 `scripts/write_review_xlsx.py spec.json`。它会校验每条搜索句、排序并套样式；校验不过就不会写文件。`--example` 打印样例 spec。
7. 把第二遍新发现的检查点写回 [references/check-logic.md](references/check-logic.md) 和 [GROK-PROMPT.md](GROK-PROMPT.md)，两处用同一句话。用户给出人工改过的终版时也做这件事：对接受修订后的文本做差，终版改掉的不再报，终版改回去或改错的单独报，终版实际采纳的规则写进检查清单、项目事实和 GROK-PROMPT。
8. 用中文回复：先说共几条、几条高，再列最严重的几条和“需要你确认”的问题。之前报错的项，被新证据推翻的要删掉并说明。

## 给其他模型用

[GROK-PROMPT.md](GROK-PROMPT.md) 是自包含的提示词，包含格式、核对逻辑和 JSON spec 样例，可以直接粘给 Grok。能运行 Python 时，把 `scripts/` 一起给它；不能运行时，让它输出 JSON spec，用户在本地用 `write_review_xlsx.py` 生成和校验。

脚本依赖：`python-docx`、`openpyxl`、`lxml`。`pyspellchecker` 可选，装了才查拼写。
