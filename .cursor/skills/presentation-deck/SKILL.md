---
name: presentation-deck
description: Use when creating, editing, translating, or visually polishing a PowerPoint deck for BLA retrofit, CDMO client reviews, or internal deployment reviews. Trigger on PPT, PPTX, 幻灯, 汇报材料, slides, pitch deck, or 展示性很好的PPT.
---

# 展示性汇报 PPT

做任何幻灯前先读本 skill；涉及主题色时同时读 `theme-factory`。中英文本走 `cn-en-regulatory-translation`。不要用空白标题+密密麻麻子弹充当“专业”。

## 默认视觉（未指定品牌时）

客户汇报默认 **Ocean Depths**（`theme-factory/themes/ocean-depths.md`）：深海军蓝背景或白底+海军蓝标题，留白充足。内部部署会可用白底 + 同一套强调色。不要混用第三套色。

- 画布：`LAYOUT_WIDE`（13.3" × 7.5"）或 16:9。开写前设定 `pres.layout`。
- 标题区固定，页码+项目名+日期放脚注，每页同一位置。
- 一页一个观点。标题写成结论，不要写成“情况介绍”。
- 正文最多 5–6 行；数字、状态、责任人用卡片或表格，不用段落。
- 中英对照页：左中右英或上结论下对照，不要同一行中英夹杂到无法扫读。

## 推荐页序（客户汇报可裁剪）

1. 封面：项目、场合、日期、密级（内部/客户）
2. 今天要什么（决策清单，3 条以内）
3. 背景与范围（BLA 改造边界）
4. 现状 / 差距（用红黄绿，不要彩虹）
5. 方案与路径（选项对比表）
6. 风险与依赖（客户需拍板的项单独成列）
7. 计划与责任人（里程碑，不要 20 行甘特废话）
8. 附录：术语、详细表、法规索引

内部部署会把第 2 页改成“会议要落地的分工”。

## 技术实现

优先用 **pptxgenjs** 新建；改已有模板再用 `python-pptx` 或拆包编辑。模块在 `.skill-runtime/node_modules`（由 `scripts/install-skill-runtime.sh` 安装，已 gitignore）。

```javascript
const PptxGenJS = require(process.cwd() + "/.skill-runtime/node_modules/pptxgenjs");
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Asymchem";
pres.title = "BLA retrofit review";
```

硬性避坑（违反会生成坏文件）：

- 颜色用 `"1B4F72"`，不要 `#`，不要 8 位带 alpha 的 hex。透明用 `transparency`。
- 每个 `add*` 传入新的 options 对象，不要复用同一个 `shadow` 对象。
- 阴影 `offset >= 0`；向上投影用 `angle: 270`。
- 列表用 `bullet: true`，不要手打 `•`。
- 图表用 `addChart`；堆叠图的 `dataLabelPosition` 只能是 `ctr` / `inEnd` / `inBase`。
- 写完用文件是否能被 `python-pptx` 打开做冒烟：`Presentation("out.pptx")` 能读幻灯数。

依赖不在仓库里。本机未装时再执行 `scripts/install-skill-runtime.sh`。

## 内容规则

- 标题是判断句：“PPQ 批次定义仍待客户确认”，不是“PPQ 讨论”。
- 状态词固定：On track / Watch / Blocked，或 正常 / 关注 / 阻塞。
- 不把未核实法规结论做成大字标题。
- 来源页脚写文件名或会议日期，便于会后对纪要。
- 动画只用于揭示路径，默认不要动画。

## 交付前检查

1. 用 `python -m markitdown out.pptx` 读一遍，看每页是否只讲一件事。
2. 有 LibreOffice 时转 PDF 再看缩略图：对齐、溢出、对比度。
3. 中英术语与 glossary 一致。
4. 把文件放到 `deliverables/pptx/`，文件名：`YYYY-MM-DD-<audience>-<topic>.pptx`。
