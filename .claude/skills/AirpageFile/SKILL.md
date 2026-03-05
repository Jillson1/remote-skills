---
name: AirpageFile
description: "自动化生成 WPS 在线智能文档 (.otl)，支持 Markdown 内容转换与插入、自动开启分享，以及读取现有文档（支持 kdocs.cn/365.kdocs.cn）的详细内容。"
version: "1.1.0"
author: "yushuyi@wps.cn"
scope: "internal"
triggers:
  - "生成文档"
  - "创建在线文档"
  - "更新文档"
  - "覆盖更新"
  - "替换文档内容"
  - "删除文档内容"
  - "追加内容"
  - "查询文档结构"
  - "create wps doc"
  - "publish to wps"
---

# AirpageFile Skill

## Overview

此 Skill 能够将 Markdown 格式的文本或文件，自动转换为 WPS 在线智能文档 (.otl)，同时也支持**读取和解析**现有在线文档的详细内容（包括 `kdocs.cn` 和 `365.kdocs.cn` 域名）。它封装了 WPS OpenAPI 的复杂调用流程，实现了“一键发布”、“覆盖更新”、“内容获取”和“附件查询”。

核心能力包括：

1.  **自动化创建 (`publish`)**：自动在指定的 Drive（应用盘）中创建文档。
2.  **内容追加 (`append`)**：向已有文档插入新的 Markdown 内容。
3.  **内容覆盖 (`replace`)**：支持**全量**或**局部**覆盖更新文档内容，并提供自动备份机制（Markdown + JSON）。
4.  **内容删除 (`delete`)**：精准删除指定范围的文档块。
5.  **内容获取 (`inspect`)**：透视文档结构，支持关键词搜索定位（获取 Index），支持导出为 Markdown。
6.  **附件查询 (`attachment`)**：获取文档内图片等附件的下载链接。
7.  **附件上传 (`upload`)**：支持上传图片或其他文件作为文档附件，并返回 sourceKey 供插入使用。
8.  **即时分享**：创建后自动开启“任何人可读”权限，并返回访问链接。

## Usage Principles

1.  **URL 优先**: 操作已有文档时，优先使用文档 URL（如 `https://kdocs.cn/l/xxx`），脚本会自动提取 ID。
2.  **安全更新**: 使用 `replace` 命令进行覆盖更新时，默认会自动备份被删除的内容，无需担心数据丢失。
3.  **精准定位**: 使用 `inspect --find` 查找关键词，获取精确的 `start_index`，再进行后续的编辑操作。

## Workflow (操作流程)

### 场景 1: 创建新文档 (Publish)
```bash
python3 AirpageFile/scripts/publish_doc.py publish --title "文档标题" --content "# Markdown内容"
# 或者
python3 AirpageFile/scripts/publish_doc.py publish --title "文档标题" --file "文件路径"
```

### 场景 2: 覆盖更新 / 替换内容 (Replace)
最常用的更新模式。支持**全量替换**（默认）或**局部替换**。

**全量覆盖（自动保留标题，替换剩余所有内容，并自动备份旧内容）：**
```bash
python3 AirpageFile/scripts/publish_doc.py replace --url "https://kdocs.cn/l/xxxx" --content "# 新的全文内容"
```

**局部替换（将指定位置的 N 个块替换为新内容）：**
```bash
# 例如：将 Index=2 开始的 1 个块替换为新文本
python3 AirpageFile/scripts/publish_doc.py replace \
  --url "https://kdocs.cn/l/xxxx" \
  --start_index 2 \
  --delete_count 1 \
  --content "这是替换后的新段落"
```

### 场景 3: 追加内容 (Append)
```bash
# 默认插入到文档开头 (index=1, 紧接标题之后)
python3 AirpageFile/scripts/publish_doc.py append --url "https://kdocs.cn/l/xxxx" --content "追加的内容"

# 追加内容并同时上传插入本地图片 (一键操作)
python3 AirpageFile/scripts/publish_doc.py append \
  --url "https://kdocs.cn/l/xxxx" \
  --content "这是图片的描述" \
  --attachment "/path/to/image.png"
```

### 场景 4: 删除内容 (Delete)
```bash
# 删除从 Index 2 开始的 3 个块
python3 AirpageFile/scripts/publish_doc.py delete --url "https://kdocs.cn/l/xxxx" --start_index 2 --delete_count 3
```

### 场景 5: 获取信息 (Inspect & Attachment)
```bash
# 获取 Markdown 内容
python3 AirpageFile/scripts/publish_doc.py inspect --url "https://kdocs.cn/l/xxxx" --format markdown

# 搜索关键词位置 (获取 start_index 用于后续操作)
python3 AirpageFile/scripts/publish_doc.py inspect --url "https://kdocs.cn/l/xxxx" --find "关键词"

# 获取图片附件链接 (需先通过 inspect 找到图片的 attachment_id/sourceKey)
python3 AirpageFile/scripts/publish_doc.py attachment --url "https://kdocs.cn/l/xxxx" --attachment_id "sourceKey_xxx"
```

### 场景 6: 上传附件 (Upload)
```bash
# 上传本地图片到文档，并获取 sourceKey (attachment_id)
# 成功后会提示：![image](sourceKey:xxxx)
python3 AirpageFile/scripts/publish_doc.py upload --url "https://kdocs.cn/l/xxxx" --file "/path/to/image.png"
```

## Output Format (输出结果)
脚本执行成功后，通常会输出操作日志或 JSON 结果。
对于 `replace` 操作，如果触发了备份，会显示：
```text
⚠️  已自动备份将被删除的内容到:
   - Markdown: backup_xxxxx.md
   - JSON    : backup_xxxxx.json
✅ 覆盖更新成功！
```

## 🛠 Tools & Resources
- `scripts/publish_doc.py`: 核心执行脚本。
- `config/airpage.properties`: 配置文件（AK/SK/DriveID）。

## Troubleshooting
**Q: 执行脚本时报错 `urllib3.exceptions.SSLError: [Errno 1] Operation not permitted`？**
- **原因**: 脚本需要访问系统 SSL 证书以建立 HTTPS 连接，但默认沙箱环境限制了对此类系统文件的读取。
- **解决**: 运行此 Skill 的相关脚本时，请务必在工具调用中申请 `all` 权限（即 `required_permissions: ['all']`），以绕过沙箱限制。

## Agent Guidelines
执行文档创建、更新或内容追加操作成功后，**必须**在回复末尾显式展示可点击的文档链接。
格式示例：`🔗 [文档标题](URL)`
