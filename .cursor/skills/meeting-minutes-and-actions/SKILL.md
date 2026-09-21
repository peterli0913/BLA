---
name: meeting-minutes-and-actions
description: Use when turning a transcript, recording notes, PDF minutes, or messy call notes into structured meeting minutes, a decision log, and a persistent action-item register. Trigger on 会议纪要, 纪要, 行动项, action items, follow-up, 部署会, client call recap, or after a BLA/CDMO meeting.
---

# 会议纪要与行动项跟踪

默认服务本仓库的 BLA / 凯莱英客户项目。只记录会上已说内容；推断必须标「推断」。不替任何人承诺新交期。

## 落盘位置（固定）

| 文件 | 用途 |
|---|---|
| `meetings/YYYY/YYYY-MM-DD-<slug>.md` | 单场纪要 |
| `meetings/action-items.md` | 跨会行动项总账 |
| `meetings/decision-log.md` | 跨会决策总账 |

已有 PDF/录音文字时，先抽取原文再写纪要，原文与纪要分开存放：`meetings/YYYY/YYYY-MM-DD-<slug>.source.md`。

## 步骤

1. **识别会议类型**：内部部署 / 客户例会 / 技术对齐 / 质量-法规 / 商务。类型只影响话题写法，不改变总结构。
2. **抽事实**：出席、时间、议程、原话级决定、点名任务、未决问题。缺出席名单就写「纪要未列全，待补」。
3. **写单场纪要**（用下面模板，章节顺序不要改）。
4. **回写总账**：新行动项追加到 `meetings/action-items.md`；已有项只更新状态，不另开一行。决策同样回写 `decision-log.md`。
5. **会后确认稿**：用 `client-comms` 出一封可发给客户或内部的短确认（中文，必要时中英对照）。
6. **翻译**：对外稿调用 `cn-en-regulatory-translation`，不要在纪要里即兴改术语。

## 单场纪要模板

```markdown
# <会议名称>

- 日期：YYYY-MM-DD
- 时间：<起–止，时区>
- 类型：内部部署 | 客户例会 | 技术对齐 | 质量法规 | 商务
- 项目：<BLA/产品代码>
- 主持 / 记录：
- 出席：
- 缺席：
- 资料来源：<文件名或“现场笔记”>

## 1. 结论（最多 5 条）

## 2. 决策

| ID | 决策 | 依据 | 影响范围 | 状态 |
|---|---|---|---|---|
| D-YYYYMMDD-01 |  |  |  | 已决 / 待批 |

## 3. 行动项（本场新增或关闭）

| ID | 行动 | 责任人 | 截止日期 | 验收标准 | 状态 | 来源 |
|---|---|---|---|---|---|---|
| A-YYYYMMDD-01 | 动词开头的一件事 | 人名 | YYYY-MM-DD | 可检查的完成定义 | Open | 本场 |

## 4. 分议题纪要

### <议题>
- 现状：
- 讨论要点：
- 决定 / 未决：

## 5. 开放问题

| ID | 问题 | 谁来澄清 | 目标日期 |
|---|---|---|---|

## 6. 会后待确认
- 纪要里标了「推断」或会后补充的条目
```

## 行动项规则

- 必须有**一个人名**责任人。禁止 “项目组”“双方”“TBD 团队”。
- 必须有日期或明确写成 `TBD（待会上指定）`。
- 一条只做一件事，动词开头：提交、修订、确认、召开、提供。
- 关闭时写证据（文件名、版本、邮件日期），状态改为 `Done`。
- 阻塞时写卡点和对哪个里程碑有影响，状态 `Blocked`。
- 更新总账时保留 ID；不要因为重开会就换号。

## 总账列

`meetings/action-items.md` 使用：

`ID | 行动 | 责任人 | 截止日期 | 优先级 | 状态 | 来源会议 | 验收标准 | 最新进展 | 更新日`

状态枚举：`Open | In progress | Blocked | Done | Cancelled`。

每次更新在「最新进展」写一句，并改「更新日」。不要删历史 Done 行，可折到文件底部「已关闭」。

## 红线

- 不要把讨论过的选项写成已决定。
- 不要把内部吐槽写进客户版纪要。
- 客户版先出，再出内部加长版（内部版可含风险和升级建议）。
- 与现有总账冲突时，以更新的会议日期为准，并在进展里写“替代 A-xxxx”。
