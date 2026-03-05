#!/usr/bin/env python3
import sys
import os
# 修复 Windows 环境下中文输出乱码问题
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
# 防止生成 __pycache__ 目录和 .pyc 文件
sys.dont_write_bytecode = True

import argparse
import json
import re
from wps_client import WpsClient
from table_fixer import fix_table_blocks

def block_to_markdown(block):
    """递归将 Block 转换为 Markdown 字符串"""
    b_type = block.get('type')
    content_list = block.get('content', [])
    
    # 提取 text 内容
    text_content = ""
    for item in content_list:
        # 安全检查：确保 item 是字典类型
        if not isinstance(item, dict):
            continue
            
        item_type = item.get('type')
        if item_type == 'text':
            text = item.get('content', '')
            attrs = item.get('attrs', {})
            if attrs.get('bold'):
                text = f"**{text}**"
            if attrs.get('italic'):
                text = f"*{text}*"
            if attrs.get('code'):
                text = f"`{text}`"
            if attrs.get('link'):
                link = attrs.get('link', {}).get('href', '')
                text = f"[{text}]({link})"
            text_content += text
            
        elif item_type == 'WPSUser':
             # 处理 WPSUser 类型，显示为 @Name
             attrs = item.get('attrs', {})
             name = attrs.get('name', 'Unknown')
             text_content += f" @{name} "
             
        elif item_type == 'emoji':
             # 处理 emoji 类型，提取 emoji 字符
             attrs = item.get('attrs', {})
             emoji_char = attrs.get('emoji', '')
             text_content += emoji_char
             
        elif item_type == 'hardBreak':
             # 处理硬换行符，转换为换行
             text_content += '\n'
             
        # 递归处理嵌套结构 (如 list item)
        else:
            try:
                text_content += block_to_markdown(item)
            except (AttributeError, TypeError, KeyError):
                # 如果遇到无法处理的类型，跳过
                pass

    if b_type == 'doc':
        return text_content
    
    elif b_type == 'title':
        # 忽略 Title 块或者作为 H1
        return f"# {text_content}\n\n"
        
    elif b_type == 'heading':
        level = block.get('attrs', {}).get('level', 1)
        return f"{'#' * level} {text_content}\n\n"
        
    elif b_type == 'paragraph':
        # 检查是否为列表项 (WPS 有时用 paragraph + listAttrs 表示列表)
        attrs = block.get('attrs', {})
        list_attrs = attrs.get('listAttrs', {})
        if list_attrs:
            list_type = list_attrs.get('type') # 1: Bullet, 2: Ordered, etc.
            level = list_attrs.get('level', 0)
            
            # 使用 4 个空格作为标准缩进，确保 Markdown 解析器能正确识别嵌套
            indent = "    " * level
            
            if list_type == 1: # Bullet
                return f"{indent}- {text_content}\n"
            elif list_type == 2: # Ordered
                # Markdown 中嵌套的有序列表通常也使用 1. 配合缩进
                # 渲染器会根据层级自动显示为 1. a. i. 等样式
                return f"{indent}1. {text_content}\n"
            else:
                 # 其他类型默认按无序处理
                 return f"{indent}- {text_content}\n"

        return f"{text_content}\n\n"
        
    elif b_type == 'blockQuote':
        lines = text_content.strip().split('\n')
        quoted = '\n'.join([f"> {line}" for line in lines])
        return f"{quoted}\n\n"
        
    elif b_type == 'codeBlock':
        # 尝试获取语言
        lang = block.get('attrs', {}).get('language', '')
        return f"```{lang}\n{text_content}\n```\n\n"
        
    elif b_type == 'bulletList':
        result = ""
        for item in content_list:
            # item is listItem
            item_text = block_to_markdown(item).strip()
            # 简单的处理，假设 listItem 内部是一个 paragraph
            # 实际上 listItem 内部可以有多个 block
            result += f"- {item_text}\n"
        return f"{result}\n"

    elif b_type == 'orderedList':
        result = ""
        for i, item in enumerate(content_list, 1):
            item_text = block_to_markdown(item).strip()
            result += f"{i}. {item_text}\n"
        return f"{result}\n"

    elif b_type == 'listItem':
        # listItem 只是容器，递归内容即可
        # 注意：这里简单的返回内容，外层 list 会处理前缀
        # 如果 listItem 包含多个段落，缩进可能需要处理 (暂简化)
        return text_content
        
    elif b_type == 'table':
        # 表格处理比较复杂，这里做简化处理
        rows = []
        for row_block in content_list: # tableRow
            cells = []
            for cell_block in row_block.get('content', []): # tableCell
                cell_text = block_to_markdown(cell_block).strip().replace('\n', '<br>')
                cells.append(cell_text)
            rows.append(cells)
            
        if not rows:
            return ""
            
        # 生成 Markdown 表格
        # Header
        header = rows[0]
        md_table = f"| {' | '.join(header)} |\n"
        md_table += f"| {' | '.join(['---'] * len(header))} |\n"
        
        # Body
        for row in rows[1:]:
            md_table += f"| {' | '.join(row)} |\n"
            
        return f"{md_table}\n"
        
    elif b_type == 'tableRow':
        # 由 table 处理
        return "" 
    elif b_type == 'tableCell':
        # 由 table 处理，但也可能独立递归
        return text_content
        
    elif b_type == 'picture':
        attrs = block.get('attrs', {})
        source_key = attrs.get('sourceKey', 'unknown')
        return f"![image](sourceKey:{source_key})\n\n"

    elif b_type == 'WPSDocument':
        attrs = block.get('attrs', {})
        name = attrs.get('wpsDocumentName', '未命名文档')
        link = attrs.get('wpsDocumentLink', '')
        # 使用 emoji 标识这是嵌入的文档
        return f"📎 [{name}]({link})\n\n"

    # 其他类型直接返回文本或者空
    return text_content

def search_blocks_by_content(blocks, keyword):
    """
    在 Block 列表中搜索关键词。
    为了适配 append --index 的需求，主要关注第一层 Block 的索引。
    
    Returns:
        matches: list of dict
        [
          {
            "index": 0,          # 顶层索引 (可直接用于 append --index)
            "id": "block_id",
            "type": "heading",
            "preview": "标题内容...",
            "match_level": "top" # top(顶层匹配) or child(子元素匹配)
          },
          ...
        ]
    """
    matches = []
    
    # 确保处理的是列表
    target_list = []
    if isinstance(blocks, list):
        target_list = blocks
    elif isinstance(blocks, dict):
        if 'blocks' in blocks:
            target_list = blocks['blocks']
        elif blocks.get('type') == 'doc' and 'content' in blocks:
             target_list = blocks['content']
        else:
            # 单个 Block，放入列表处理
            target_list = [blocks]

    for idx, block in enumerate(target_list):
        # 获取当前 Block 的 Markdown 文本表示 (递归获取所有子内容)
        # 这样做的好处是：哪怕关键词在表格、列表中，也能被匹配到
        full_text = block_to_markdown(block)
        
        if keyword.lower() in full_text.lower():
            # 提取简短预览 (去掉 Markdown 标记，取前 50 字符)
            preview = full_text.strip().replace('\n', ' ')
            if len(preview) > 60:
                preview = preview[:60] + "..."
                
            # 判断是否为直接匹配 (简单判断：如果预览文字里直接有，通常是顶层或浅层)
            # 实际上对于 append 操作，用户更关心的是 "这个 Block 包含我要找的东西，它的 Index 是多少"
            
            matches.append({
                "index": idx,  # 这里的 Index 是基于当前 block 列表的，如果 block 列表是 doc 的直接子节点，则可用于 insert
                "id": block.get('id'),
                "type": block.get('type'),
                "preview": preview
            })
            
    return matches

def _prepare_images_for_conversion(client, file_id, markdown_content, markdown_file_path=None, debug=False):
    if not markdown_content:
        return markdown_content, 0
    
    def is_local_path(path):
        if not path: return False
        if path.startswith(('sourceKey:', 'http://', 'https://', 'ftp://', 'data:')) or '://' in path:
            return False
        return True
    
    # 改进的正则表达式：支持单层嵌套括号，如 file (1).png
    # (?:[^()]|\([^()]*\))+ 匹配非括号字符或成对的括号
    local_image_pattern = r'!\[([^\]]*)\]\(((?:[^()]|\([^()]*\))+)\)'
    
    local_images = []
    for match in re.finditer(local_image_pattern, markdown_content):
        alt_text, image_path = match.group(1), match.group(2)
        if is_local_path(image_path):
            actual_path = image_path
            if markdown_file_path and not os.path.isabs(image_path):
                base_dir = os.path.dirname(os.path.abspath(markdown_file_path))
                actual_path = os.path.normpath(os.path.join(base_dir, image_path))
            
            if os.path.exists(actual_path):
                local_images.append({
                    'match': match, 'alt': alt_text, 'actual_path': actual_path
                })
            elif debug:
                print(f"[警告] 本地图片文件不存在: {actual_path}")
    
    processed_content = markdown_content
    upload_count = 0
    
    for img_info in reversed(local_images):
        if debug: print(f"[转换] 正在上传本地图片: {img_info['actual_path']}")
        
        upload_res = client.upload_attachment(file_id, img_info['actual_path'])
        
        if upload_res and 'data' in upload_res and (attachment_id := upload_res['data'].get('attachment_id')):
            new_image_syntax = f"![{img_info['alt']}](sourceKey:{attachment_id})"
            start, end = img_info['match'].start(), img_info['match'].end()
            processed_content = processed_content[:start] + new_image_syntax + processed_content[end:]
            upload_count += 1
            if debug: print(f"[转换] ✅ 图片上传成功，已替换为 sourceKey:{attachment_id}")
        elif debug:
            print(f"[转换] ❌ 图片上传失败: {img_info['actual_path']}")
    
    if upload_count > 0:
        print(f"已自动上传并替换 {upload_count} 张本地图片")
    
    return processed_content, upload_count


def _execute_conversion_request(client, file_id, markdown_content):
    print("正在转换 Markdown 内容...")
    converted = client.convert_markdown_to_blocks(file_id, markdown_content)
    return converted.get('blocks') if converted and 'blocks' in converted else None


def _optimize_converted_blocks(client, file_id, blocks, markdown_content, fix_tables=True, debug=False):
    if fix_tables:
        return fix_table_blocks(client, file_id, blocks, markdown_content, debug=debug)
    return blocks


def convert_markdown_to_blocks_pipeline(client, file_id, markdown_content, markdown_file_path=None, fix_tables=True, debug=False):
    """
    统一的 Markdown 转换管道：预处理(本地图片) -> 执行转换 -> 后处理(优化Blocks)。
    """
    if not markdown_content: return None
    
    # 1. 预处理：扫描并上传本地图片，替换 Markdown 中的路径
    processed_content, upload_count = _prepare_images_for_conversion(
        client, file_id, markdown_content, markdown_file_path, debug=debug
    )
    
    # 2. 核心转换：调用 WPS API
    blocks = _execute_conversion_request(client, file_id, processed_content)
    if blocks is None: return None
    
    # 3. 后处理：修复表格等结构
    final_blocks = _optimize_converted_blocks(
        client, file_id, blocks, processed_content, fix_tables=fix_tables, debug=debug
    )
    
    return {
        'blocks': final_blocks,
        'original_content': markdown_content,
        'processed_content': processed_content,
        'uploaded_images_count': upload_count
    }

def unescape_content(content):
    """
    处理命令行传入的内容字符串中的转义字符。
    将字面量 \\n 转换为实际换行符，\\t 转换为制表符等。
    """
    if not content:
        return content
    # 按顺序替换常见的转义序列
    content = content.replace('\\n', '\n')
    content = content.replace('\\t', '\t')
    content = content.replace('\\r', '\r')
    # 处理转义的反斜杠 (必须最后处理，避免影响上面的替换)
    content = content.replace('\\\\', '\\')
    return content


def main():
    parser = argparse.ArgumentParser(description="WPS Airpage 文档发布工具")
    subparsers = parser.add_subparsers(dest="command", help="可用命令")

    # Command: publish
    parser_pub = subparsers.add_parser("publish", help="将内容发布为新的 WPS 文档")
    parser_pub.add_argument("--file", help="要上传的 Markdown 文件路径")
    parser_pub.add_argument("--content", help="直接输入的 Markdown 内容字符串")
    parser_pub.add_argument("--title", required=True, help="文档标题")
    parser_pub.add_argument("--drive_id", help="覆盖默认的 Drive ID")
    parser_pub.add_argument("--parent_id", help="覆盖默认的 Parent ID")
    parser_pub.add_argument("--debug", action="store_true", help="打印详细调试日志")

    # Command: setup
    parser_setup = subparsers.add_parser("setup", help="初始化环境 (创建 Drive)")
    parser_setup.add_argument("--name", default="AI Application Drive", help="新 Drive 的名称")

    # Command: inspect
    parser_inspect = subparsers.add_parser("inspect", help="查询文档的 Block 结构或搜索内容")
    group_inspect = parser_inspect.add_mutually_exclusive_group(required=True)
    group_inspect.add_argument("--file_id", help="文档 ID")
    group_inspect.add_argument("--url", help="文档 URL (将提取末尾 ID)")
    
    parser_inspect.add_argument("--block_id", default="doc", help="要查询的 Block ID (默认为 doc)")
    parser_inspect.add_argument("--format", choices=["json", "markdown"], default="json", help="输出格式 (默认: json)")
    parser_inspect.add_argument("--find", help="搜索关键词，返回匹配 Block 的索引(index)和ID等信息，便于精准插入")

    # Command: append
    parser_append = subparsers.add_parser("append", help="向已有文档追加内容")
    group_append = parser_append.add_mutually_exclusive_group(required=True)
    group_append.add_argument("--file_id", help="目标文档 ID")
    group_append.add_argument("--url", help="目标文档 URL (将提取末尾 ID)")
    
    parser_append.add_argument("--file", help="要上传的 Markdown 文件路径")
    parser_append.add_argument("--content", help="直接输入的 Markdown 内容字符串")
    parser_append.add_argument("--attachment", help="要上传并插入的本地附件(图片)路径")
    parser_append.add_argument("--title", help="追加内容的标题 (可选)")
    parser_append.add_argument("--index", type=int, default=1, help="插入位置索引 (默认: 1)")
    parser_append.add_argument("--debug", action="store_true", help="打印详细调试日志")

    # Command: attachment
    parser_att = subparsers.add_parser("attachment", help="获取附件(图片)下载链接")
    group_att = parser_att.add_mutually_exclusive_group(required=True)
    group_att.add_argument("--file_id", help="文档 ID")
    group_att.add_argument("--url", help="文档 URL")
    parser_att.add_argument("--attachment_id", required=True, help="附件 ID (通常为 Block 的 attrs.sourceKey)")

    # Command: upload
    parser_upload = subparsers.add_parser("upload", help="上传附件(图片/文件)到文档")
    group_upload = parser_upload.add_mutually_exclusive_group(required=True)
    group_upload.add_argument("--file_id", help="文档 ID")
    group_upload.add_argument("--url", help="文档 URL")
    parser_upload.add_argument("--file", required=True, help="要上传的本地文件路径")

    # Command: update
    parser_update = subparsers.add_parser("update", help="更新文档中指定 Block 的内容")
    group_update = parser_update.add_mutually_exclusive_group(required=True)
    group_update.add_argument("--file_id", help="文档 ID")
    group_update.add_argument("--url", help="文档 URL")
    
    parser_update.add_argument("--block_id", required=True, help="要更新的 Block ID")
    parser_update.add_argument("--content", help="新的内容字符串 (Markdown 格式，将被转换)")
    parser_update.add_argument("--file", help="新的内容文件路径 (Markdown 格式)")
    parser_update.add_argument("--debug", action="store_true", help="打印详细调试日志")
    
    # Command: replace (覆盖式更新)
    parser_replace = subparsers.add_parser("replace", help="覆盖式更新: 先删除指定位置的 Block，再插入新内容")
    group_replace = parser_replace.add_mutually_exclusive_group(required=True)
    group_replace.add_argument("--file_id", help="文档 ID")
    group_replace.add_argument("--url", help="文档 URL")
    
    parser_replace.add_argument("--start_index", type=int, default=1, help="要替换的起始索引 (从0开始, 默认: 1, 即保留标题)")
    parser_replace.add_argument("--delete_count", type=int, help="要替换(删除)的 Block 数量 (如果不传，将自动计算从 start_index 到末尾的数量)")
    parser_replace.add_argument("--parent_id", default="doc", help="父 Block ID (默认: doc)")
    parser_replace.add_argument("--no_backup", action="store_true", help="禁用备份 (默认: 自动备份被删除的内容)")
    
    parser_replace.add_argument("--content", help="新的内容字符串 (Markdown 格式)")
    parser_replace.add_argument("--file", help="新的内容文件路径 (Markdown 格式)")
    parser_replace.add_argument("--debug", action="store_true", help="打印详细调试日志")

    # Command: delete
    parser_delete = subparsers.add_parser("delete", help="删除文档中的 Block (按索引范围)")
    group_delete = parser_delete.add_mutually_exclusive_group(required=True)
    group_delete.add_argument("--file_id", help="文档 ID")
    group_delete.add_argument("--url", help="文档 URL")
    
    parser_delete.add_argument("--parent_id", default="doc", help="父 Block ID (默认: doc)")
    parser_delete.add_argument("--start_index", type=int, required=True, help="起始索引 (从0开始)")
    parser_delete.add_argument("--delete_count", type=int, default=1, help="要删除的 Block 数量 (默认: 1)")

    args = parser.parse_args()
    
    client = WpsClient()

    if args.command == "setup":
        print(f"正在创建新的 Drive: {args.name}...")
        res = client.create_drive(args.name)
        if res and 'data' in res:
            new_id = res['data']['id']
            print(f"成功: Drive 已创建。ID: {new_id}")
            print("请使用此 ID 更新您的 airpage.properties 配置。")
        else:
            print("失败: 无法创建 Drive。")
            sys.exit(1)

    elif args.command == "inspect":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        print(f"正在查询文档 {target_id} 的 Block 结构 (Root: {args.block_id})...")
        blocks = client.get_blocks(target_id, args.block_id)
        
        if blocks:
            # 预处理：为了配合 append --index，我们需要在 doc 的直接子节点列表中搜索
            # 目标：找到 doc block 的 content 列表
            search_target_list = []
            
            # 1. 尝试解析 API 返回结构
            raw_list = []
            if isinstance(blocks, dict):
                if 'blocks' in blocks:
                    raw_list = blocks['blocks']
                elif blocks.get('type') == 'doc':
                    raw_list = [blocks]
                else:
                    raw_list = [blocks]
            elif isinstance(blocks, list):
                raw_list = blocks
            
            # 2. 检查列表，如果包含 doc block，则“剥开”它
            # 通常 API 返回的列表中第一个就是 doc
            if raw_list and len(raw_list) > 0 and raw_list[0].get('type') == 'doc':
                # 取 doc 的子节点
                search_target_list = raw_list[0].get('content', [])
            else:
                # 如果不是 doc (比如查询特定 block children)，直接使用
                search_target_list = raw_list

            if args.find:
                print(f"正在搜索关键词: '{args.find}' ...")
                # print(f"搜索范围: 共 {len(search_target_list)} 个顶层 Block")
                matches = search_blocks_by_content(search_target_list, args.find)
                
                if matches:
                    print(f"\n✅ 找到 {len(matches)} 个匹配项 (Index 可用于 append 命令):")
                    for m in matches:
                        print(f"- Index: {m['index']}")
                        print(f"  Type : {m['type']}")
                        print(f"  Text : {m['preview']}")
                        print(f"  ID   : {m['id']}")
                        print("  -------------------")
                else:
                    print(f"\n⚠️ 未找到包含 '{args.find}' 的内容。")
            
            elif args.format == "markdown":
                # print("\n--- 文档 Markdown 内容 ---")
                # 使用 search_target_list，它已经是 doc 的子节点列表
                virtual_doc = { "type": "doc", "content": search_target_list }
                md_content = block_to_markdown(virtual_doc)
                print(md_content)
            else:
                print("\n--- 文档 Block 结构 ---")
                print(json.dumps(blocks, indent=2, ensure_ascii=False))
        else:
            print("查询失败或结果为空。")

    elif args.command == "append":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        # 1. Prepare content
        content = ""
        has_content = False
        
        if args.file:
            if os.path.exists(args.file):
                with open(args.file, 'r', encoding='utf-8') as f:
                    content = f.read()
                has_content = True
            else:
                print(f"错误: 找不到文件: {args.file}")
                sys.exit(1)
        elif args.content:
            content = unescape_content(args.content)
            has_content = True
        elif args.title:
            # 允许只插入标题
            pass
        elif args.attachment:
             # 允许只插入附件
             pass
        else:
            print("错误: 必须提供 --file, --content, --attachment 或 --title 参数其中之一。")
            sys.exit(1)

        print(f"正在处理文档 {target_id} ...")
        
        # 1.5 Handle Attachment
        if args.attachment:
            if not os.path.exists(args.attachment):
                print(f"错误: 找不到附件文件: {args.attachment}")
                sys.exit(1)
                
            print(f"正在上传附件: {args.attachment} ...")
            att_res = client.upload_attachment(target_id, args.attachment)
            
            if att_res and 'data' in att_res:
                att_id = att_res['data'].get('attachment_id')
                kind = att_res['data'].get('kind', '')
                
                # 简单判断是否为图片 (如果服务端没返回 kind，根据后缀判断)
                import mimetypes
                mt, _ = mimetypes.guess_type(args.attachment)
                is_image = False
                if kind == 'image':
                    is_image = True
                elif mt and mt.startswith('image/'):
                    is_image = True
                    
                if is_image and att_id:
                    print(f"附件上传成功 (ID: {att_id})，已追加到内容末尾。")
                    # 构造自定义图片语法并追加到 content
                    file_name = os.path.basename(args.attachment)
                    img_md = f"\n\n![{file_name}](sourceKey:{att_id})"
                    content += img_md
                    has_content = True
                else:
                    print(f"附件上传成功 (ID: {att_id})，但不是图片或无法识别，仅上传不插入正文。")
            else:
                 print("附件上传失败，跳过插入。")

        # 2. Convert
        final_blocks = []
        
        # 如果有正文内容，先转换
        if has_content:
            # 确定 Markdown 文件路径（如果是从文件读取的）
            md_file_path = args.file if args.file and os.path.exists(args.file) else None
            converted = convert_markdown_to_blocks_pipeline(
                client, target_id, content, 
                markdown_file_path=md_file_path,
                fix_tables=True,
                debug=args.debug
            )
            if converted and 'blocks' in converted:
                final_blocks = converted['blocks']
            else:
                print("警告: 内容转换失败或为空。")
        
        # 如果指定了 title，尝试更新现有标题
        if args.title:
            print(f"正在尝试更新文档标题为: {args.title}")
            updated = client.update_title(target_id, args.title)
            if updated:
                print("标题更新成功！")
            else:
                print("标题更新失败或未找到 Title Block。")
            
        if final_blocks:
            print(f"准备插入 (共 {len(final_blocks)} 个块)。正在插入到位置 {args.index}...")
            client.insert_blocks(target_id, final_blocks, index=args.index)
            print("插入成功！")
        else:
            if not args.title:
                print("没有生成任何内容块，且未指定标题更新，跳过。")
            else:
                 print("仅更新了标题，未插入新内容。")

    elif args.command == "attachment":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        res = client.get_attachment_info(target_id, args.attachment_id)
        if res and 'data' in res:
            data = res['data']
            print("\n--- 附件信息 ---")
            print(f"Name: {data.get('name')}")
            print(f"Size: {data.get('size')}")
            print(f"URL : {data.get('download_url')}")
            print("----------------")
        else:
            print("获取附件信息失败。")

    elif args.command == "upload":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        print(f"正在上传附件: {args.file} -> 文档 {target_id}")
        res = client.upload_attachment(target_id, args.file)
        
        if res and 'data' in res:
            data = res['data']
            print("\n--- 附件上传结果 ---")
            att_id = data.get('attachment_id')
            kind = data.get('kind')
            
            # 兜底逻辑：如果服务端未返回 kind，但我们知道是图片后缀，强制视为图片
            if not kind and args.file:
                import mimetypes
                mt, _ = mimetypes.guess_type(args.file)
                if mt and mt.startswith('image/'):
                    kind = 'image'
            
            print(f"Attachment ID: {att_id}")
            print(f"Kind         : {kind}")
            
            # 提示用户如何使用
            print("\n[Usage Hint]")
            if kind == 'image':
                print(f"要在文档中插入此图片，请在 Markdown 中使用: \n![image](sourceKey:{att_id})")
            else:
                print(f"附件已上传。ID: {att_id}")
        else:
            print("上传流程结束，但未获取到有效的返回数据。")

    elif args.command == "update":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        # 1. Prepare content
        content_str = ""
        if args.file:
            if os.path.exists(args.file):
                with open(args.file, 'r', encoding='utf-8') as f:
                    content_str = f.read()
            else:
                print(f"错误: 找不到文件: {args.file}")
                sys.exit(1)
        elif args.content:
            content_str = unescape_content(args.content)
        else:
            print("错误: update 命令必须提供 --content 或 --file 参数")
            sys.exit(1)
        
        # 2. Convert Markdown to Blocks
        md_file_path = args.file if args.file and os.path.exists(args.file) else None
        converted = convert_markdown_to_blocks_pipeline(
            client, target_id, content_str,
            markdown_file_path=md_file_path,
            fix_tables=True,  # update 命令默认也修复表格
            debug=args.debug
        )
        if converted and 'blocks' in converted:
            new_blocks = converted['blocks']
            
            final_content = new_blocks
            # 自动提取子内容逻辑
            if len(new_blocks) == 1 and 'content' in new_blocks[0]:
                print(f"检测到单 Block 转换结果 ({new_blocks[0]['type']})，将提取其子内容进行更新...")
                final_content = new_blocks[0]['content']
            
            res = client.update_block(target_id, args.block_id, content=final_content)
            if res:
                print("更新内容成功！")
            else:
                print("更新内容失败。")
        else:
            print("内容转换失败或为空。")

    elif args.command == "replace":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        # 0. 准备新内容
        content_str = ""
        if args.file:
            if os.path.exists(args.file):
                with open(args.file, 'r', encoding='utf-8') as f:
                    content_str = f.read()
            else:
                print(f"错误: 找不到文件: {args.file}")
                sys.exit(1)
        elif args.content:
            content_str = unescape_content(args.content)
        else:
            print("错误: replace 命令必须提供 --content 或 --file 参数")
            sys.exit(1)

        # 1. 转换 Markdown
        md_file_path = args.file if args.file and os.path.exists(args.file) else None
        converted = convert_markdown_to_blocks_pipeline(
            client, target_id, content_str,
            markdown_file_path=md_file_path,
            fix_tables=True,  # replace 命令默认也修复表格
            debug=args.debug
        )
        if not (converted and 'blocks' in converted):
             print("Markdown 转换失败。")
             sys.exit(1)

        new_blocks = converted['blocks']
        print(f"转换成功，准备插入 {len(new_blocks)} 个新 Block。")

        # 2. 确定删除范围 & 自动备份
        start_idx = args.start_index
        del_count = args.delete_count

        # 如果未指定 delete_count，需要查询文档结构来确定剩余数量
        if del_count is None:
            print("未指定删除数量，正在查询文档结构以计算剩余 Block...")
            doc_struct = client.get_blocks(target_id, args.parent_id)
            
            total_count = 0
            current_blocks = []
            
            # 解析 current blocks
            if doc_struct:
                if isinstance(doc_struct, dict):
                    # 常见的 doc 结构 { blocks: [ { type:doc, content: [...] } ] }
                    # 或者直接返回 { type:doc, content: [...] }
                    root_list = []
                    if 'blocks' in doc_struct:
                         root_list = doc_struct['blocks']
                    elif 'content' in doc_struct:
                         root_list = [doc_struct] # 本身就是个 block
                    elif isinstance(doc_struct, list):
                         root_list = doc_struct

                    # 找到目标 parent 的 content
                    # 如果 parent_id 是 doc，通常是 root_list[0].content
                    if args.parent_id == "doc" and root_list and root_list[0].get('type') == 'doc':
                         current_blocks = root_list[0].get('content', [])
                    else:
                         current_blocks = root_list # 简化处理，假设直接返回了 children
                elif isinstance(doc_struct, list):
                    current_blocks = doc_struct

            total_count = len(current_blocks)
            print(f"当前文档共有 {total_count} 个子 Block。")
            
            # 计算剩余数量
            if start_idx >= total_count:
                del_count = 0
                print(f"起始位置 {start_idx} 已超过文档末尾，不执行删除操作。")
            else:
                del_count = total_count - start_idx
                print(f"自动计算删除数量: {del_count} (从 {start_idx} 到末尾)")
                
                # 执行备份
                if not args.no_backup:
                    import time
                    timestamp = int(time.time())
                    
                    blocks_to_backup = current_blocks[start_idx : start_idx + del_count]

                    # 1. 备份 Markdown
                    backup_filename_md = f"backup_{target_id}_{timestamp}.md"
                    # 构造一个临时 doc 容器方便转换 md
                    virtual_doc = { "type": "doc", "content": blocks_to_backup }
                    backup_md = block_to_markdown(virtual_doc)
                    
                    with open(backup_filename_md, "w", encoding="utf-8") as f:
                        f.write(backup_md)
                    
                    # 2. 备份原始 JSON
                    backup_filename_json = f"backup_{target_id}_{timestamp}.json"
                    with open(backup_filename_json, "w", encoding="utf-8") as f:
                        json.dump(blocks_to_backup, f, indent=2, ensure_ascii=False)

                    print(f"⚠️  已自动备份将被删除的内容到:")
                    print(f"   - Markdown: {backup_filename_md}")
                    print(f"   - JSON    : {backup_filename_json}")

        if del_count > 0:
            print(f"步骤 1/2: 删除旧 Block (Start: {start_idx}, Count: {del_count})...")
            del_res = client.delete_blocks(target_id, args.parent_id, start_idx, start_idx + del_count)
            if not del_res:
                print("删除失败，终止操作。")
                sys.exit(1)
        else:
            print("步骤 1/2: 跳过删除 (删除数量为 0)")
        
        # 3. 插入新 Block
        print(f"步骤 2/2: 插入新 Block 到 Index {start_idx}...")
        ins_res = client.insert_blocks(target_id, new_blocks, index=start_idx, block_id=args.parent_id)
        if ins_res:
            print("✅ 覆盖更新成功！")
        else:
            print("插入新内容失败。")

    elif args.command == "delete":
        target_id = args.file_id
        doc_url = args.url
        if args.url:
            target_id = client.resolve_file_id(args.url)
            print(f"从 URL 提取 ID: {target_id}")

        # 检查文件类型
        is_ap, file_name, err_msg = client.check_ap_document(target_id, doc_url=doc_url)
        if is_ap is False:
            print(err_msg)
            sys.exit(1)

        start_idx = args.start_index
        end_idx = start_idx + args.delete_count
        
        # 为了安全起见，打印确认信息
        print(f"准备删除: 文档 {target_id}, Parent: {args.parent_id}, Range: [{start_idx}, {end_idx})")
        
        res = client.delete_blocks(target_id, args.parent_id, start_idx, end_idx)
        if res:
             print("删除成功！")
        else:
             print("删除失败。")

    elif args.command == "publish":
        # 1. Prepare content
        content = ""
        if args.file:
            if os.path.exists(args.file):
                with open(args.file, 'r', encoding='utf-8') as f:
                    content = f.read()
            else:
                print(f"错误: 找不到文件: {args.file}")
                sys.exit(1)
        elif args.content:
            content = unescape_content(args.content)
        else:
            print("错误: 必须提供 --file 或 --content 参数。")
            sys.exit(1)

        # 2. Create File
        print(f"正在创建文件 '{args.title}'...")
        file_res = client.create_file(args.title, drive_id=args.drive_id, parent_id=args.parent_id)
        file_id = file_res['data']['id']
        drive_id = args.drive_id or client.drive_id
        
        print(f"文件已创建 (ID: {file_id})。正在转换内容...")

        # 3. Convert & Insert
        md_file_path = args.file if args.file and os.path.exists(args.file) else None
        converted = convert_markdown_to_blocks_pipeline(
            client, file_id, content,
            markdown_file_path=md_file_path,
            fix_tables=True,  # publish 命令默认也修复表格
            debug=args.debug
        )
        final_blocks = []
        if converted and 'blocks' in converted:
            final_blocks = converted['blocks']
        
        # 3.1 尝试更新标题 (更稳健的方式)
        print(f"正在设置文档标题: {args.title}")
        client.update_title(file_id, args.title)

        # 3.2 插入正文内容
        if final_blocks:
            print(f"内容转换完成 (共 {len(final_blocks)} 个块)。正在插入...")
            # 注意：因为已经单独更新了标题，这里不需要再在 blocks 里插 title block 了
            client.insert_blocks(file_id, final_blocks)
        else:
             print("Markdown 内容为空或转换失败，仅创建了文档并设置了标题。")

        # 4. Open Share
        print("正在开启分享...")
        share_res = client.open_share(file_id, drive_id=drive_id)
        
        # 5. Output Result
        result = {
            "status": "success",
            "file_id": file_id,
            "drive_id": drive_id,
            "title": args.title,
            "url": share_res['data']['url'] if share_res and 'data' in share_res else None
        }
        
        print("\n--- 发布结果 ---")
        print(json.dumps(result, indent=2, ensure_ascii=False))
        
        # 显式打印纯净链接，确保在终端中可点击
        if result.get('url'):
            print(f"\n👉 文档链接: {result['url']}")

    else:
        parser.print_help()

if __name__ == "__main__":
    main()

