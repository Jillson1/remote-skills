---
name: AirsheetFile
description: "自动化生成和读写 WPS 在线表格 (.ksheet/.et/.xlsx)。支持创建表格、上传本地文件到云空间、读取内容、追加数据、更新选区和管理工作表。"
version: "1.2.0"
author: "yushuyi@wps.cn"
scope: "internal"
triggers:
  - "创建表格"
  - "在线表格"
  - "create sheet"
  - "上传表格"
  - "upload file"
  - "读取表格"
  - "read sheet"
  - "更新表格"
  - "update sheet"
  - "airsheet"
---

# AirsheetFile Skill

## Overview

此 Skill 能够管理 WPS 在线表格（包含智能表格 AirSheet 及普通表格）。它封装了 WPS Open Platform API，支持创建在线表格 (.ksheet)、管理工作表、读取和写入数据（支持公式）。

## Capabilities

1.  **创建表格 (`create`)**：创建一个新的在线表格 (.ksheet)。
2.  **上传文件 (`upload`)**：将本地文件（如 .xlsx/.et/.ksheet 等）通过三步上传接口上传到 WPS 云空间，支持指定目录、文件名冲突策略及上传后自动开启分享。
3.  **工作表管理**：
    *   `list_sheets`: 列出所有工作表信息（ID、名称、隐藏状态等），并**自动预览每个表的前 5 行数据**。
    *   `add_sheet`: 添加新的工作表。
3.  **写入数据**：
    *   `add_rows`: 向工作表末尾追加一行或多行数据（支持文本、数值、公式）。
    *   `update_range`: 精准更新/覆盖指定区域（Range）的数据。
4.  **读取数据**：
    *   `get_range`: 读取单元格数据。
        *   支持**精准读取**（指定行列范围）或**全量读取**（不指定范围，自动探测边界）。
        *   **默认导出为 CSV** 文件到 `AirsheetFile/output` 目录，便于后续处理。
        *   可使用 `--console` 强制打印到控制台。

## Workflow

### 场景 1: 创建新表格 (Create)
```bash
# 创建一个新的在线表格
python3 AirsheetFile/scripts/sheet_manager.py create --title "销售统计表"
```

### 场景 2: 管理工作表 (Sheet Management)
```bash
# 列出所有工作表
python3 AirsheetFile/scripts/sheet_manager.py list_sheets --url "https://kdocs.cn/l/xxxx"

# 添加新工作表
python3 AirsheetFile/scripts/sheet_manager.py add_sheet --url "https://kdocs.cn/l/xxxx" --name "2024Q1"
```

### 场景 3: 追加数据 (Add Rows)
```bash
# 追加单行：--values 传入一维数组
# 自动识别类型：文本直接传，数字保持原样，"="开头视为公式
python3 AirsheetFile/scripts/sheet_manager.py add_rows \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1 \
  --values '["产品A", 100, "=B2*0.8"]'

# 追加多行：--values 传入二维数组（每个子数组为一行）
python3 AirsheetFile/scripts/sheet_manager.py add_rows \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1 \
  --values '[["产品A", 100, "=B2*0.8"], ["产品B", 200, "=B3*0.8"]]'
```

### 场景 4: 更新/覆盖数据 (Update Range)
```bash
# 覆盖单行：从 A1 (R0C0) 开始水平写入
python3 AirsheetFile/scripts/sheet_manager.py update_range \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1 \
  --row 0 --col 0 \
  --values '["产品", "单价", "折后价"]'

# 覆盖多行：从 A1 (R0C0) 开始，逐行写入（二维数组）
python3 AirsheetFile/scripts/sheet_manager.py update_range \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1 \
  --row 0 --col 0 \
  --values '[["产品", "单价", "折后价"], ["苹果", 8.5, 6.8], ["香蕉", 3.5, 2.8]]'
```

### 场景 5: 上传本地文件到云空间 (Upload)
```bash
# 上传本地文件到默认 Drive 根目录
python3 AirsheetFile/scripts/sheet_manager.py upload --file "path/to/文件.xlsx"

# 上传并自动开启分享链接
python3 AirsheetFile/scripts/sheet_manager.py upload --file "path/to/文件.xlsx" --share

# 指定父目录或相对路径、文件名冲突策略
python3 AirsheetFile/scripts/sheet_manager.py upload --file "path/to/文件.xlsx" \
  --drive_id "xxx" --parent_id "0" \
  --parent_path '["文档", "报告"]' \
  --on_conflict rename
```

### 场景 6: 读取数据 (Read)
```bash
# 1. 全量读取 (默认模式)
# 自动探测表格边界，分页下载所有数据，并保存为 AirsheetFile/output/{file_id}_sheet_{sheet_id}.csv
python3 AirsheetFile/scripts/sheet_manager.py get_range \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1

# 2. 精准读取 (指定范围)
# 默认打印到控制台，也可以结合 --console 明确指定
python3 AirsheetFile/scripts/sheet_manager.py get_range \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1 \
  --row_from 0 --row_to 10 \
  --col_from 0 --col_to 5

# 3. 全量读取并打印到控制台
python3 AirsheetFile/scripts/sheet_manager.py get_range \
  --url "https://kdocs.cn/l/xxxx" \
  --sheet_id 1 \
  --console
```

## Best Practices

### 数据分析与处理 (Data Analysis)
当需要对表格数据进行**查找、筛选、统计、匹配**等复杂操作时，**请遵循以下规范**，避免依赖 LLM 直接阅读控制台输出：
1.  **导出 CSV**：使用 `get_range` (不带 `--console`) 将数据下载为 CSV 文件。
2.  **编写脚本**：编写 Python 脚本读取生成的 CSV 文件。
3.  **代码处理**：在脚本中实现具体的查找或匹配逻辑，并输出最终结果。
    *   *Why?* 这种方式处理大数据量更稳定、逻辑更精准、可复用且无 Token 限制。

## Configuration
配置文件位于 `AirsheetFile/config/airsheet.properties`。请确保填入正确的 `ACCESS_KEY`, `SECRET_KEY` 和 `DEFAULT_DRIVE_ID`。
