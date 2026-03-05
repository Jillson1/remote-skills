---
name: user-feedback-analysis
description: "基于Excel源数据，自动过滤、分析并生成结构化的用户反馈周报。"
version: "2.1.0"
author: "yushuyi@wps.cn"
scope: "internal"
triggers:
  - "analyze user feedback"
  - "generate user feedback report"
  - "分析用户反馈"
  - "生成用户反馈周报"
---

# User Feedback Analysis Skill

## Overview

此 Skill 旨在自动化处理 WPS AI 的用户反馈分析流程。它将原始反馈数据转换为结构化的洞察报告，大幅减少人工筛选和总结的时间。

核心能力包括：
1.  **数据获取**：优先使用 `AirsheetFile` 的 `sheet_manager.py` 将 Sheet1 导出到本地 CSV 再解析；若导出失败则退化为在线表格 API 直接读取。
2.  **数据过滤**：自动根据"周报日期"从海量反馈中提取特定时间段的数据。
3.  **上下文组装**：智能结合分析提示词（Prompts）、原始数据和报告模板。
4.  **报告生成**：生成符合周报规范的 Markdown 格式报告，包含核心数据和 Top 问题（注：严格遵循客观陈述原则，无需包含改进建议）。

## Prerequisites (前置条件)

在使用此 Skill 之前，请确保：
1.  已安装 Python 环境及以下依赖：
    ```
    pandas>=2.0.0
    tabulate>=0.9.0
    requests>=2.28.0
    ```
2.  依赖 `AirsheetFile` 模块，确保 `AirsheetFile/config/airsheet.properties` 已正确配置：
    - ACCESS_KEY
    - SECRET_KEY
    - DEFAULT_DRIVE_ID

## Configuration (配置文件)

配置文件位于 `UserFeedbackAnalysis/config/config.properties`：

```properties
[DEFAULT]
# 用户反馈数据源在线表格 URL
FEEDBACK_SOURCE_URL = https://365.kdocs.cn/l/ct3YewJ2rEDn

# 默认工作表 ID
DEFAULT_SHEET_ID = 1
```

配置后，运行脚本时可省略 `--url` 参数。

## Analysis Principles (分析原则)

**CRITICAL**: 在生成分析报告时，必须严格遵守以下核心原则：

1.  **Objectivity First (客观第一)**: 你的角色是观察者而非决策者。报告必须基于数据事实，严禁包含主观臆断、猜测或个人建议。
2.  **Evidence-Based (证据导向)**: 每一个结论或分类都必须有具体的反馈数据作为支撑。严禁编造数据或引用不存在的 Case。
3.  **Accuracy & Fidelity (精准还原)**: 引用用户反馈时，必须保持原话（包括标点和语气），严禁为了美化而篡改用户原意。

## Workflow (使用流程)

当需要生成周报时，请按照以下步骤操作：

### Step 1: 数据准备与上下文构建

运行脚本从在线表格读取数据并构建 AI 分析所需的上下文。

**基础用法（使用配置文件中的 URL）**
```bash
python UserFeedbackAnalysis/scripts/process_feedback.py --date "12.11-12.17"
```

**指定 URL（覆盖配置文件）**
```bash
python UserFeedbackAnalysis/scripts/process_feedback.py \
  --url "https://365.kdocs.cn/l/xxxxx" \
  --date "12.11-12.17"
```

**关注周 + 对比周（导出两周数据用于环比结论）**
```bash
python UserFeedbackAnalysis/scripts/process_feedback.py \
  --date "26-2.1-2.4" \
  --compare-date "26-1.22-1.28"
```
将导出关注周与对比周两个 CSV 到 `UserFeedbackAnalysis/cache`（如 `focus_2.1-2.4.csv`、`compare_1.22-1.28.csv`），上下文会指引 AI 填写「结论总结」与「结论对比表」。

**带关键词筛选（专项话题分析）**
```bash
python UserFeedbackAnalysis/scripts/process_feedback.py \
  --date "12.11-12.17" \
  --keyword "AI讲解" \
  --output feedback_context_topic.md
```

**参数说明**:
| 参数 | 必选 | 说明 |
|------|------|------|
| `--date` | 否 | **关注周**日期区间。示例：`26-2.1-2.4`、`26.2.1-2.4`。不传则用配置 `DEFAULT_WEEK_RANGE` / `DEFAULT_MONTH_PREFIX` |
| `--compare-date` | 否 | **对比周**日期区间。示例：`26-1.22-1.28`。填写后将导出关注周与对比周两个 CSV 到 `UserFeedbackAnalysis/cache`，上下文中会说明以填写报告的「结论总结」与「结论对比表」 |
| `--url` | 否 | 在线表格 URL，不填则从配置文件读取 |
| `--keyword` | 否 | 关键词筛选 (如 `AI讲解`) |
| `--output` | 否 | 输出上下文文件名 (默认写入 `UserFeedbackAnalysis/cache/feedback_context.md`) |
| `--sheet_id` | 否 | 工作表 ID (默认从配置读取或 `1`) |

**输出目录**：生成的上下文文件、关注周/对比周 CSV 等临时文件均写入 **`UserFeedbackAnalysis/cache`** 目录。

**按日期区间导出**：传入 `--date` 时，只导出关注周数据；同时传入 `--compare-date` 时会再导出对比周数据，两个 CSV 会复制到 `cache` 供后续分析使用。此时脚本会**自动**生成 `cache/feedback_compare_stats.txt`（UTF-8），内含按二级分类的统计与环比增长率，Step 2 可直接读取该文件填写「结论总结」与「结论对比表」，无需再跑统计命令或解析 CSV。

### 运行说明（Agent 执行 Step 1 时必读）

为避免多余尝试与错误分析，执行 Step 1 时请严格遵守以下要求：

1. **Shell 与命令语法**
   - 在 **Windows / PowerShell** 环境下，请使用 **PowerShell 兼容**写法，不要使用 `cd /d ... && python ...`（`&&` 在 PowerShell 中无效，会报错）。
   - **正确示例**（工作区根路径为 `e:\codee\aiskills` 时）：
     ```powershell
     Set-Location e:\codee\aiskills; python .claude\skills\UserFeedbackAnalysis\scripts\process_feedback.py --date "26-2.1-2.4" --compare-date "26-1.22-1.28"
     ```
   - 或先 `Set-Location`（或 `cd`）到工作区根目录，再在同一会话中执行 `python <脚本相对路径> ...`。

2. **控制台编码**
   - `process_feedback.py` 已在脚本内部对 Windows 控制台做了 UTF-8 输出处理，**无需**在外部执行 `chcp 65001` 或其它编码切换。
   - 若 Agent 需要运行其它会输出中文的 Python 命令（如统计脚本），可设置环境变量 `PYTHONIOENCODING=utf-8` 或将输出重定向到 UTF-8 文件，避免乱码。

3. **必须等待脚本完成后再进行数据分析**
   - Step 1 应**同步执行**并**等待** `process_feedback.py` 正常退出，建议超时时间为 **10 分钟（600 秒）**。
   - 在脚本**未成功结束**之前，**禁止**基于 `cache` 中的部分结果（例如仅存在 `focus_*.csv` 而尚无 `feedback_context.md` 或 `compare_*.csv`）进行数据分析或进入 Step 2。
   - 仅当脚本**成功退出**且已生成 `cache/feedback_context.md`（若指定了 `--compare-date` 则还应存在 `compare_*.csv`）后，方可执行 Step 2：读取上下文并生成报告。

**Script Actions**:
1.  **Export then Read (优先)**：先调用 `AirsheetFile/scripts/sheet_manager.py get_range` 将 Sheet1 导出到本地 CSV；若导出成功则从该本地文件解析数据。
2.  **Fallback Read**：若导出 Sheet 失败（超时、无权限、脚本缺失等），则退化为**在线表格读取**（通过 WpsClient API 直接拉取数据）。
3.  **Filter**: 
    - 筛选 `周报日期` == 指定日期范围。
    - (可选) 筛选包含指定 `keyword` 的记录（搜索内容、分类等字段）。
3.  **Combine**: 读取 `prompts/analysis_prompt.md` 和 `templates/report.md`。
4.  **Generate**: 生成包含指令、数据和模板的上下文文件。
5.  **Compare Stats（当存在 `--compare-date` 时）**：自动按二级分类统计关注周与对比周数量并计算增长率，写入 `cache/feedback_compare_stats.txt`（UTF-8），避免控制台编码导致的中文乱码，供 Step 2 直接使用。

### Step 2: AI 分析与报告生成

**前置条件**：仅当 Step 1 已在 **10 分钟超时内** 成功完成（脚本正常退出并生成 `cache/feedback_context.md`）后，才执行本步骤。若 Step 1 超时或失败，不得基于部分 cache 文件进行分析，应报错并提示用户。

使用 AI Agent 读取上一步生成的上下文文件，并输出最终报告。

**Agent Action**:
1.  **Read Context**: 读取 `feedback_context.md`。
2.  **Analyze**: 根据文件中的 "ROLE & INSTRUCTIONS" 部分进行深度分析。
3.  **Generate**: 严格按照 "OUTPUT TEMPLATE" 部分的格式输出报告。

**Prompt for Agent**:
> "请读取 feedback_context.md 文件，并根据其中的指令和数据，生成最终的 Markdown 分析报告。**请将生成的报告内容写入名为 `feedback_report.md` 的文件中。**"

### Step 3: 发布为在线文档 (Publish)
报告生成完成后，请立即自动使用 `AirpageFile` Skill 将其一键发布为 WPS 在线智能文档并获取分享链接，无需等待人工确认。

**Command Format**:
```bash
python3 AirpageFile/scripts/publish_doc.py publish --title "2025年Wxx用户反馈周报" --file "feedback_report.md"
```

**Result**:
-   自动创建在线文档。
-   返回访问 URL（支持多人协作与评论）。


## 🛠 File Structure

-   `config/config.properties`: 配置文件（数据源 URL 等）。
-   `scripts/process_feedback.py`: 核心处理脚本。
-   `templates/report.md`: 标准周报输出模板。
-   `prompts/analysis_prompt.md`: 核心分析逻辑与角色定义。
-   `cache/`: 临时输出目录。上下文文件 `feedback_context.md`、关注周 CSV、以及存在对比周时的 `focus_*.csv` / `compare_*.csv` 均生成于此。
-   `reports/`: 生成的报告 Markdown 输出目录。
