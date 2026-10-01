# BLA 工作台 Skills

本目录是 Cursor Agent 会自动发现的项目技能。用于凯莱英侧 BLA 改造：客户沟通、中英翻译、展示性 PPT、会议纪要和行动项跟踪。

## 本项目技能（已纳入仓库）

| 技能 | 何时用 |
|---|---|
| `client-comms` | 对外邮件、即时消息、会议话术、升级口径 |
| `cn-en-regulatory-translation` | 中英翻译、术语统一、中英对照幻灯/纪要 |
| `meeting-minutes-and-actions` | 纪要、决策账、行动项总账 |
| `presentation-deck` | 新建或润色汇报 PPT |
| `internal-comms` | 内部周报、FAQ、3P 更新（Anthropic, Apache-2.0） |
| `theme-factory` | PPT/文档配色主题（Anthropic, Apache-2.0） |
| `doc-coauthoring` | 长文档共创（Anthropic, 示例技能） |

## 本机文档技能（不进 Git）

Anthropic 的 `pptx` / `pdf` / `docx` / `xlsx` 许可不允许再分发。本环境已安装到 `~/.cursor/skills/`，仅供本机 Agent 使用。其他机器可自行从 [anthropics/skills](https://github.com/anthropics/skills) 取用，或打开 Cursor **Settings → Agents → Sync Skills for Cloud Agents**。

## 运行时

```bash
bash scripts/install-skill-runtime.sh
```

会安装 `markitdown`、`python-pptx`、`pptxgenjs` 等。不要把专有 skill 文件拷进本仓库。

## 产出约定

- 纪要与总账：`meetings/`
- PPT：`deliverables/pptx/`
- 不要在对话里粘贴真实密钥或客户未公开数据的完整原文到公共位置
