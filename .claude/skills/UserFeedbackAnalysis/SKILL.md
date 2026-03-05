---
name: user-feedback-analysis
description: "基于Excel源数据，自动过滤、分析并生成结构化的用户反馈周报。"
version: "2.0.0"
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
FEEDBACK_SOURCE_URL = https://365.kdocs.cn/l/cjh0F6PJiRCB

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
| `--date` | 否 | 周报日期区间。不传则用配置 `DEFAULT_WEEK_RANGE` / `DEFAULT_MONTH_PREFIX`（最近一周）。示例：`26.2.1-2.4`（26年2月、2.1-2.4）、`25.5.3-5.8`、`5.3-5.8`、`2.1-2.4` |
| `--url` | 否 | 在线表格 URL，不填则从配置文件读取 |
| `--keyword` | 否 | 关键词筛选 (如 `AI讲解`) |
| `--output` | 否 | 输出上下文文件 (默认: `feedback_context.md`) |
| `--sheet_id` | 否 | 工作表 ID (默认从配置读取或 `1`) |

**按日期区间导出**：传入 `--date` 时，会优先按「月份」+「周报日期」只导出该区间的数据到 CSV，再基于该 CSV 生成报告，避免全量拉取。

**Script Actions**:
1.  **Export then Read (优先)**：先调用 `AirsheetFile/scripts/sheet_manager.py get_range` 将 Sheet1 导出到本地 CSV；若导出成功则从该本地文件解析数据。
2.  **Fallback Read**：若导出 Sheet 失败（超时、无权限、脚本缺失等），则退化为**在线表格读取**（通过 WpsClient API 直接拉取数据）。
3.  **Filter**: 
    - 筛选 `周报日期` == 指定日期范围。
    - (可选) 筛选包含指定 `keyword` 的记录（搜索内容、分类等字段）。
3.  **Combine**: 读取 `prompts/analysis_prompt.md` 和 `templates/report.md`。
4.  **Generate**: 生成包含指令、数据和模板的上下文文件。

### Step 2: AI 分析与报告生成
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
