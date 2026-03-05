#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import io
import os
import argparse
import json
import tempfile
import shutil

# Fix Windows console encoding
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from wps_client import WpsClient

# 尝试导入数据转换器
try:
    from data_converter import DataConverter
except ImportError:
    # 如果是在同一目录下运行
    if os.path.exists(os.path.join(os.path.dirname(__file__), 'data_converter.py')):
        sys.path.append(os.path.dirname(__file__))
        from data_converter import DataConverter
    else:
        DataConverter = None

# 确保 pandas 可用
try:
    import pandas as pd
except ImportError:
    print("错误: 需要安装 pandas。请执行: pip install pandas openpyxl")
    sys.exit(1)

def fetch_all_data(client, file_id, sheet_id, file_type='auto'):
    """
    自动探测并获取工作表的所有数据
    """
    print(f"正在探测工作表 {sheet_id} 的数据范围...")
    
    # 1. 探测最大列数 (Max Column)
    # 探测前 10 行的数据，取最大列宽，以应对首行是合并标题的情况
    probe_row_limit = 10
    probe_col_limit = 100
    max_col = 0
    
    res = client.get_range_data(file_id, sheet_id, 0, probe_row_limit, 0, probe_col_limit, file_type=file_type)
    if res and res.get('code') == 0:
        data = res['data'].get('range_data', [])
        if data:
            # 找到最大的非空列索引
            for cell in data:
                if cell.get('cell_text') or str(cell.get('original_cell_value', '')):
                    max_col = max(max_col, cell.get('col_from', 0))
    
            # 简单的启发式：如果探测到了最后一列都有数据，可能还有更多列
            if max_col >= probe_col_limit - 1:
                print(f"警告: 列数可能超过 {probe_col_limit}，目前仅截取前 {max_col+1} 列。")
        else:
            print("工作表头部似乎为空。")
            # 即使头部为空，也不能断定整个表为空，但为了简单起见，这里返回空
            # 或者给个默认值？
            pass
    else:
        print(f"探测列失败: {res}")
        return []
        
    print(f"探测到有效列数: {max_col + 1}")
    
    # 2. 分页读取行数据
    all_cells = []
    batch_size = 1000
    current_row = 0
    empty_batch_count = 0
    
    print("开始分页读取数据...")
    while True:
        # print(f"读取行 {current_row} - {current_row + batch_size} ...", end='\r')
        res = client.get_range_data(file_id, sheet_id, current_row, current_row + batch_size - 1, 0, max_col, file_type=file_type)
        
        if res and res.get('code') == 0:
            batch_data = res['data'].get('range_data', [])
            
            # 检查是否为空批次（根据实际返回的数据量或内容）
            # 注意：API 可能返回空列表，或者包含空值的单元格列表
            # 我们检查这一批次是否有任何有效内容
            has_content = False
            for cell in batch_data:
                if cell.get('cell_text') or str(cell.get('original_cell_value', '')):
                    has_content = True
                    break
            
            if not batch_data or not has_content:
                # 连续空批次检测（防止中间有空行但后面有数据）
                # 这里简化：只要一批次全空，就停止。
                break
                
            all_cells.extend(batch_data)
            current_row += batch_size
            
            # 安全阀：防止无限循环
            if current_row > 100000:
                print("\n达到行数上限 (100,000)，停止读取。")
                break
        else:
            print(f"\n读取出错: {res}")
            break
            
    print(f"\n读取完成，共获取 {len(all_cells)} 个单元格数据。")
    return all_cells

def main():
    parser = argparse.ArgumentParser(description="WPS AirsheetFile 管理器")
    subparsers = parser.add_subparsers(dest="command", help="可用命令")

    # 命令: create
    parser_create = subparsers.add_parser("create", help="创建新的智能表格")
    parser_create.add_argument("--title", required=True, help="文档标题")
    parser_create.add_argument("--file", help="初始内容（CSV 或 Excel 文件）")
    parser_create.add_argument("--drive_id", help="覆盖默认 Drive ID")
    parser_create.add_argument("--parent_id", help="覆盖默认父目录 ID")

    # 命令: add_sheet
    parser_add_sheet = subparsers.add_parser("add_sheet", help="向表格添加新工作表")
    group_add_sheet = parser_add_sheet.add_mutually_exclusive_group(required=True)
    group_add_sheet.add_argument("--file_id", help="文件 ID")
    group_add_sheet.add_argument("--url", help="文件 URL")
    parser_add_sheet.add_argument("--name", required=True, help="工作表名称")
    parser_add_sheet.add_argument("--col_width", type=float, help="默认列宽")
    parser_add_sheet.add_argument("--type", choices=['auto', 'airsheet', 'sheets'], default='auto', help="表格类型：auto(自动检测，默认), airsheet(智能表格) 或 sheets(传统表格)")
    # 位置参数
    group_pos = parser_add_sheet.add_mutually_exclusive_group()
    group_pos.add_argument("--end", action="store_true", help="插入到末尾（默认）")
    group_pos.add_argument("--after_sheet", type=str, help="在指定工作表 ID 之后插入")
    group_pos.add_argument("--before_sheet", type=str, help="在指定工作表 ID 之前插入")

    # 命令: list_sheets
    parser_list_sheets = subparsers.add_parser("list_sheets", help="列出表格中的所有工作表")
    group_list_sheets = parser_list_sheets.add_mutually_exclusive_group(required=True)
    group_list_sheets.add_argument("--file_id", help="文件 ID")
    group_list_sheets.add_argument("--url", help="文件 URL")
    parser_list_sheets.add_argument("--type", choices=['auto', 'airsheet', 'sheets'], default='auto', help="表格类型：auto(默认), airsheet, sheets")

    # 命令: add_rows
    parser_add_rows = subparsers.add_parser("add_rows", help="向工作表添加行")
    group_add_rows = parser_add_rows.add_mutually_exclusive_group(required=True)
    group_add_rows.add_argument("--file_id", help="文件 ID")
    group_add_rows.add_argument("--url", help="文件 URL")
    parser_add_rows.add_argument("--sheet_id", required=True, type=int, help="工作表 ID（整数）")
    parser_add_rows.add_argument("--type", choices=['auto', 'airsheet', 'sheets'], default='auto', help="表格类型")
    
    group_data = parser_add_rows.add_mutually_exclusive_group(required=True)
    group_data.add_argument("--data", help="range_data 的原始 JSON 数据（高级用法）")
    group_data.add_argument("--values", help=(
        "追加行数据（JSON 格式）。"
        "单行: '[\"Name\", 123, \"=A1*2\"]'；"
        "多行: '[[\"A\", 1], [\"B\", 2]]'。"
        "自动识别：一维数组=单行，二维数组=多行"
    ))

    # 命令: update_range
    parser_update_range = subparsers.add_parser("update_range", help="更新选区数据")
    group_update_range = parser_update_range.add_mutually_exclusive_group(required=True)
    group_update_range.add_argument("--file_id", help="文件 ID")
    group_update_range.add_argument("--url", help="文件 URL")
    parser_update_range.add_argument("--sheet_id", required=True, type=int, help="工作表 ID（整数）")
    parser_update_range.add_argument("--row", required=True, type=int, help="起始行索引")
    parser_update_range.add_argument("--col", required=True, type=int, help="起始列索引")
    parser_update_range.add_argument("--values", required=True, help=(
        "覆盖写入数据（JSON 格式）。"
        "单行: '[\"A\", 1]' 从 (row,col) 水平写入；"
        "多行: '[[\"A\",1],[\"B\",2]]' 从 (row,col) 起逐行写入。"
        "自动识别：一维数组=单行，二维数组=多行"
    ))
    parser_update_range.add_argument("--type", choices=['auto', 'airsheet', 'sheets'], default='auto', help="表格类型")

    # 命令: get_range
    parser_get_range = subparsers.add_parser("get_range", help="获取选区数据")
    group_get_range = parser_get_range.add_mutually_exclusive_group(required=True)
    group_get_range.add_argument("--file_id", help="文件 ID")
    group_get_range.add_argument("--url", help="文件 URL")
    parser_get_range.add_argument("--sheet_id", required=True, type=int, help="工作表 ID（整数）")
    # 改为可选参数
    parser_get_range.add_argument("--row_from", type=int, help="起始行索引")
    parser_get_range.add_argument("--row_to", type=int, help="结束行索引")
    parser_get_range.add_argument("--col_from", type=int, help="起始列索引")
    parser_get_range.add_argument("--col_to", type=int, help="结束列索引")
    parser_get_range.add_argument("--type", choices=['auto', 'airsheet', 'sheets'], default='auto', help="表格类型")
    parser_get_range.add_argument("--console", action='store_true', help="将结果打印到控制台（默认保存为 CSV）")
    parser_get_range.add_argument("--filter-date-prefix", metavar="PREFIX", help="仅导出「月份」列以该前缀开头的行（如 26-2 表示 26年2月）")
    parser_get_range.add_argument("--date-column", default="月份", help="用于月份过滤的列名（默认: 月份）")
    parser_get_range.add_argument("--filter-week", metavar="WEEK", help="仅导出「周报日期」等于该值的行（如 2.1-2.4）")
    parser_get_range.add_argument("--date-column-week", default="周报日期", help="用于周报日期过滤的列名（默认: 周报日期）")

    # 命令: upload (上传本地文件到云空间)
    parser_upload = subparsers.add_parser("upload", help="上传本地文件到 WPS 云空间")
    parser_upload.add_argument("--file", required=True, help="本地文件路径")
    parser_upload.add_argument("--drive_id", help="覆盖默认 Drive ID")
    parser_upload.add_argument("--parent_id", help="覆盖默认父目录 ID")
    parser_upload.add_argument("--parent_path", help="相对路径（JSON 数组格式），不存在时自动创建，如: '[\"文档\", \"报告\"]'")
    parser_upload.add_argument("--on_conflict", choices=['rename', 'overwrite', 'fail'], default='rename', help="文件名冲突处理方式（默认: rename）")
    parser_upload.add_argument("--share", action='store_true', help="上传完成后自动开启分享链接")

    # 命令: import_file
    parser_import = subparsers.add_parser("import_file", help="将本地 CSV/Excel 文件导入到表格")
    group_import = parser_import.add_mutually_exclusive_group(required=True)
    group_import.add_argument("--file_id", help="文件 ID")
    group_import.add_argument("--url", help="文件 URL")
    
    parser_import.add_argument("--file", required=True, help="本地文件路径 (.csv, .xlsx)")
    parser_import.add_argument("--type", choices=['auto', 'airsheet', 'sheets'], default='auto', help="表格类型")
    
    # 目标 Sheet 策略
    group_target = parser_import.add_mutually_exclusive_group(required=True)
    group_target.add_argument("--sheet_id", type=int, help="追加到现有工作表 ID")
    group_target.add_argument("--new_sheet", help="创建新工作表并导入（指定新工作表名称）")

    # 命令: append
    parser_append = subparsers.add_parser("append", help="向表格追加行数据")
    group_append = parser_append.add_mutually_exclusive_group(required=True)
    group_append.add_argument("--file_id", help="文件 ID")
    group_append.add_argument("--url", help="文件 URL")
    parser_append.add_argument("--data", required=True, help="要追加的数据（JSON 列表或字典列表）")
    parser_append.add_argument("--sheet", default=0, help="工作表名称或索引（默认: 0）")

    # 命令: update
    parser_update = subparsers.add_parser("update", help="更新指定单元格或选区（通过下载-编辑-上传方式）")
    group_update = parser_update.add_mutually_exclusive_group(required=True)
    group_update.add_argument("--file_id", help="文件 ID")
    group_update.add_argument("--url", help="文件 URL")
    parser_update.add_argument("--cell", help="要更新的单元格（例如 'A1', 'B2'）- 需要 openpyxl/pandas 逻辑（此处简化为行/列更新）")
    parser_update.add_argument("--row", type=int, help="要更新的行索引（从 0 开始）")
    parser_update.add_argument("--col", help="要更新的列名")
    parser_update.add_argument("--value", required=True, help="新值")
    parser_update.add_argument("--sheet", default=0, help="工作表名称或索引（默认: 0）")

    args = parser.parse_args()
    client = WpsClient()

    if args.command == "create":
        # 1. 创建智能表格元数据
        title = args.title
        # 智能表格后缀为 .ksheet
        if not title.endswith('.ksheet'):
             title += '.ksheet'
        
        print(f"正在创建智能表格 '{title}'...")
        # 使用新的 create_airsheet API
        res = client.create_airsheet(title, drive_id=args.drive_id, parent_id=args.parent_id)
        
        if res and 'data' in res:
            file_id = res['data']['id']
            print(f"智能表格创建成功。ID: {file_id}")
            
            # 2. 如果提供了初始内容则上传（通常此 API 不直接支持 ksheet 格式，
            # 但我们可以尝试转换，或者用户可能想要标准 Excel）
            # 注意：新的 Airsheet API 创建的是 .ksheet 文件。
            # 如果用户提供了 CSV/Excel 文件，我们可能需要不同的方法或导入 API。
            # 目前，我们假设 create_airsheet 创建一个空表格，然后返回链接。
            # 如果需要标准 Excel (.xlsx)，我们应该使用 create_file。
            # 但由于我们重命名为 AirsheetFile，让我们优先使用 Airsheet。
            
            if args.file:
                print("⚠️  警告: 此版本尚不支持在创建智能表格 (.ksheet) 时上传初始内容。")
                print("    如果需要上传初始内容，请使用标准 Excel 格式。")

            # 3. 开启分享
            print("正在开启分享链接...")
            share_res = client.open_share(file_id)
            url = share_res['data']['url'] if share_res and 'data' in share_res else "未知"
            
            print(f"\n✅ 创建成功！")
            print(f"🔗 链接: {url}")
            print(f"🆔 文件 ID: {file_id}")
        else:
            print("❌ 创建失败。")
            sys.exit(1)

    elif args.command == "add_sheet":
        file_id = args.file_id or client.resolve_file_id(args.url)
        print(f"正在向文件 {file_id} 添加工作表 '{args.name}'...")
        
        position = {"end": True}
        if args.after_sheet:
            position = {"after_sheet_id": args.after_sheet}
        elif args.before_sheet:
            position = {"before_sheet_id": args.before_sheet}
            
        res = client.create_worksheet(file_id, name=args.name, position=position, col_width=args.col_width, file_type=args.type)
        
        if res and res.get('code') == 0:
            sheet_id = res['data']['sheet_id']
            print(f"✅ 工作表创建成功！工作表 ID: {sheet_id}")
        else:
            print(f"❌ 创建工作表失败: {res}")

    elif args.command == "list_sheets":
        file_id = args.file_id or client.resolve_file_id(args.url)
        print(f"正在获取文件 {file_id} 的工作表列表...")
        
        res = client.get_worksheets(file_id, file_type=args.type)
        
        if res and res.get('code') == 0:
            sheets = res['data'].get('sheets', [])
            print(f"\n✅ 找到 {len(sheets)} 个工作表:")
            
            # 打印表格
            print(f"{'ID':<10} {'名称':<20} {'索引':<5} {'隐藏'}")
            print("-" * 50)
            for s in sheets:
                print(f"{s['sheet_id']:<10} {s['name']:<20} {s['index']:<5} {s.get('hidden', False)}")
                
                # 预览前 5 行
                try:
                    # 预览 5 行, 10 列
                    preview_res = client.get_range_data(file_id, s['sheet_id'], 0, 4, 0, 9, file_type=args.type)
                    if preview_res and preview_res.get('code') == 0:
                        p_data = preview_res['data'].get('range_data', [])
                        if p_data:
                            print("  [预览数据]:")
                            # 简单的行归组
                            rows = {}
                            for cell in p_data:
                                r = cell.get('row_from')
                                c = cell.get('col_from')
                                val = cell.get('cell_text') or cell.get('original_cell_value')
                                if r not in rows: rows[r] = {}
                                rows[r][c] = val
                            
                            for r_idx in sorted(rows.keys()):
                                row_vals = []
                                for c_idx in sorted(rows[r_idx].keys()):
                                    row_vals.append(str(rows[r_idx][c_idx])[:10]) # 截断显示
                                print(f"    R{r_idx}: | {' | '.join(row_vals)} |")
                        else:
                            print("  [空表]")
                except Exception as e:
                    # 预览失败不影响主流程
                    pass
                print("-" * 50)
        else:
            print(f"❌ 获取工作表列表失败: {res}")

    elif args.command == "add_rows":
        file_id = args.file_id or client.resolve_file_id(args.url)
        
        range_data = []
        def _convert_row_values(row_values):
            """将一行简单值列表转换为 range_data 格式（单行）"""
            row_data = []
            for idx, val in enumerate(row_values):
                cell_data = {
                    "col": idx,
                    "op_type": "cell_operation_type_formula"
                }
                
                if isinstance(val, (int, float)):
                    cell_data["formula"] = str(val)
                elif isinstance(val, str):
                    if val.startswith("="):
                        cell_data["formula"] = val
                    else:
                        cell_data["formula"] = val
                elif val is None:
                     continue
                else:
                     cell_data["formula"] = f'"{str(val)}"'
                
                row_data.append(cell_data)
            return row_data

        if args.values:
            try:
                values = json.loads(args.values)
                if not isinstance(values, list):
                    print("错误: --values 必须是 JSON 列表。")
                    sys.exit(1)
                
                # 自动识别：一维数组 = 单行，二维数组 = 多行
                if len(values) > 0 and isinstance(values[0], list):
                    # 二维数组：多行模式
                    all_rows = [_convert_row_values(row) for row in values]
                else:
                    # 一维数组：单行模式
                    all_rows = [_convert_row_values(values)]
                    
            except json.JSONDecodeError:
                print("错误: --values 必须是有效的 JSON。")
                sys.exit(1)
        else:
            try:
                raw_data = json.loads(args.data)
                if not isinstance(raw_data, list):
                    print("错误: --data 必须是 JSON 列表。")
                    sys.exit(1)
                # --data 为高级用法，直接作为单行 range_data
                all_rows = [raw_data]
            except json.JSONDecodeError:
                print("错误: --data 必须是有效的 JSON。")
                sys.exit(1)

        total = len(all_rows)
        # 空 sheet 检测：智能表格新建 sheet 后 R0 有占位空行，add_rows 会追加到 R1。若 R0 无数据则首行写入 R0。
        probe = client.get_range_data(file_id, args.sheet_id, 0, 0, 0, 0, file_type=args.type)
        range_data_r0 = (probe or {}).get("data", {}).get("range_data") if probe and probe.get("code") == 0 else None
        is_empty_sheet = range_data_r0 is None or len(range_data_r0) == 0

        def _row_to_update_payload(row_data, row_index):
            """将 add_rows 风格的一行数据转为 update_range_data 的 range_data（写入指定行）"""
            return [
                {
                    "row_from": row_index,
                    "row_to": row_index,
                    "col_from": c["col"],
                    "col_to": c["col"],
                    "op_type": c["op_type"],
                    "formula": c["formula"],
                }
                for c in row_data
            ]

        success_count = 0
        if is_empty_sheet and total > 0:
            # 首行写入 R0，其余行追加
            print(f"检测到空表，首行将写入第 1 行 (R0)，共 {total} 行...")
            update_payload = _row_to_update_payload(all_rows[0], 0)
            res = client.update_range_data(file_id, args.sheet_id, update_payload, file_type=args.type)
            if res and res.get("code") == 0:
                success_count += 1
            else:
                print(f"❌ 第 1 行（R0）写入失败: {res}")
            for i, row_data in enumerate(all_rows[1:], start=2):
                res = client.add_rows(file_id, args.sheet_id, row_data, file_type=args.type)
                if res and res.get("code") == 0:
                    success_count += 1
                else:
                    print(f"❌ 第 {i} 行添加失败: {res}")
        else:
            print(f"正在向文件 {file_id} 的工作表 {args.sheet_id} 添加 {total} 行...")
            for i, row_data in enumerate(all_rows, start=1):
                res = client.add_rows(file_id, args.sheet_id, row_data, file_type=args.type)
                if res and res.get("code") == 0:
                    success_count += 1
                else:
                    print(f"❌ 第 {i} 行添加失败: {res}")

        if success_count == total:
            print(f"✅ 全部 {total} 行添加成功！")
        else:
            print(f"⚠️ 添加完成: {success_count}/{total} 行成功")

    elif args.command == "update_range":
        file_id = args.file_id or client.resolve_file_id(args.url)
        
        try:
            values = json.loads(args.values)
            if not isinstance(values, list):
                print("错误: --values 必须是 JSON 列表。")
                sys.exit(1)
        except json.JSONDecodeError:
            print("错误: --values 必须是有效的 JSON。")
            sys.exit(1)
            
        # 构造批量更新的选区数据
        start_row = args.row
        start_col = args.col

        def _build_row_range_data(row_values, row_index):
            """将一行值列表转为 update_range_data 的 range_data"""
            cells = []
            for col_offset, val in enumerate(row_values):
                cell_data = {
                    "row_from": row_index,
                    "row_to": row_index,
                    "col_from": start_col + col_offset,
                    "col_to": start_col + col_offset,
                    "op_type": "cell_operation_type_formula"
                }
                if isinstance(val, (int, float)):
                    cell_data["formula"] = str(val)
                elif isinstance(val, str):
                    cell_data["formula"] = val
                elif val is None:
                    continue
                else:
                    cell_data["formula"] = str(val)
                cells.append(cell_data)
            return cells

        # 自动识别：一维数组 = 单行，二维数组 = 多行
        range_data = []
        if len(values) > 0 and isinstance(values[0], list):
            # 多行模式：每个子数组写一行，行号从 start_row 递增
            for row_offset, row_vals in enumerate(values):
                range_data.extend(_build_row_range_data(row_vals, start_row + row_offset))
            row_count = len(values)
        else:
            # 单行模式
            range_data = _build_row_range_data(values, start_row)
            row_count = 1

        print(f"正在更新工作表 {args.sheet_id} 中起始于 R{start_row}C{start_col} 的选区（{row_count} 行）...")
        res = client.update_range_data(file_id, args.sheet_id, range_data, file_type=args.type)
        
        if res and res.get('code') == 0:
            print(f"✅ 选区更新成功！共写入 {row_count} 行。")
        else:
            print(f"❌ 更新选区失败: {res}")

    elif args.command == "get_range":
        file_id = args.file_id or client.resolve_file_id(args.url)
        
        # 判断是否需要全量读取
        # 如果四个范围参数都提供了，则进行精确读取
        if all(x is not None for x in [args.row_from, args.row_to, args.col_from, args.col_to]):
            print(f"正在从工作表 {args.sheet_id} 读取选区 R{args.row_from}C{args.col_from}:R{args.row_to}C{args.col_to}...")
            
            res = client.get_range_data(file_id, args.sheet_id, args.row_from, args.row_to, args.col_from, args.col_to, file_type=args.type)
            
            if res and res.get('code') == 0:
                data = res['data'].get('range_data', [])
                
                # 如果要求 console 输出或者是精确读取模式（保持兼容性），则打印
                # 但根据新需求，默认可能是 CSV。
                # 兼容旧逻辑：如果用户指定了范围，通常希望直接看结果。
                # 但为了统一，我们也可以应用 --console 逻辑。
                # 这里保留旧逻辑：精确读取直接打印到控制台，除非... 
                # 不，用户说"默认写入到一个 .csv的文件里面，除非明确要求打印在控制台"
                # 这应该适用于所有 get_range 调用，还是只针对全量读取？
                # "允许...不填写，不填写 则出目标表的 所有行和所有列... 默认不再是将所有行和列打印出来，而是默认写入到一个 .csv的文件里面"
                # 看起来这个规则主要针对全量模式。对于精确指定范围的，原来的行为是打印。
                # 为了不破坏现有习惯，我们保留精确模式下的打印，但也支持导出。
                
                if args.console or True: # 这里的 True 是为了保持向后兼容，或者我们可以改变行为
                     # 既然用户明确说“常规的思路是...下载到本地”，我倾向于所有 get_range 都统一行为。
                     # 但为了安全起见，我们对精确指定的情况，如果没加 --console 也没加 --output（这里写死output），还是打印吧？
                     # 或者我们可以统一：
                     # 1. 如果指定了范围 -> 默认打印（兼容旧习惯），但也支持导出？
                     # 2. 如果未指定范围 -> 默认导出，console 才打印。
                     # 让代码简单点：统一转成 DataFrame，然后根据参数决定去向。
                     pass # 下面统一处理
            else:
                print(f"❌ 获取选区数据失败: {res}")
                sys.exit(1)
                
        else:
            # 全量/自动读取模式
            print(f"未指定完整范围，进入全量读取模式...")
            data = fetch_all_data(client, file_id, args.sheet_id, file_type=args.type)
            if not data:
                print("未获取到数据。")
                sys.exit(0)

        # 统一处理数据输出
        if not data:
             print("无数据。")
             sys.exit(0)
             
        # 转换为 DataFrame 方便处理
        # 稀疏数据转密集矩阵
        # 1. 找出最大行列
        max_r = 0
        max_c = 0
        for cell in data:
            max_r = max(max_r, cell.get('row_from', 0))
            max_c = max(max_c, cell.get('col_from', 0))
            
        # 2. 填充
        # 注意：如果数据量很大，pandas 可能会慢，但对于一般 Excel 表格还好。
        # 如果是精确范围，我们可以用 args.row_from 做 offset。
        # 这里简化处理，直接映射。
        
        # 预填充空矩阵
        # 考虑到 row_from 可能不为 0 (精确读取模式)
        min_r = min(cell.get('row_from', 0) for cell in data)
        min_c = min(cell.get('col_from', 0) for cell in data)
        
        # 只有在全量模式下 min_r 才是 0。
        rows_count = max_r - min_r + 1
        cols_count = max_c - min_c + 1
        
        matrix = [['' for _ in range(cols_count)] for _ in range(rows_count)]
        
        for cell in data:
            r = cell.get('row_from', 0) - min_r
            c = cell.get('col_from', 0) - min_c
            val = cell.get('cell_text')
            orig = cell.get('original_cell_value')
            # 优先使用 text，如果没有则用 value
            final_val = val if val is not None else orig
            if final_val is None: final_val = ""
            matrix[r][c] = final_val
            
        df = pd.DataFrame(matrix)
        
        # 按日期区间过滤：月份列前缀 + 周报日期精确匹配（仅在全量模式下生效）
        month_prefix = getattr(args, 'filter_date_prefix', None) and getattr(args, 'filter_date_prefix', None).strip()
        week_val = getattr(args, 'filter_week', None) and getattr(args, 'filter_week', None).strip()
        if (month_prefix or week_val) and len(df) > 0:
            # 表头可能在第 0 行或第 1 行（第 0 行常为合并/标题行，无「月份」「周报日期」）
            def _row_has(col_name, row_vals):
                return any(col_name in str(v).strip() for v in row_vals)
            row0 = df.iloc[0].astype(str).str.strip().tolist()
            if not _row_has("月份", row0) and len(df) > 1:
                row1 = df.iloc[1].astype(str).str.strip().tolist()
                if _row_has("月份", row1):
                    header_row = row1
                    data_df = df.iloc[2:].copy()
                else:
                    header_row = row0
                    data_df = df.iloc[1:].copy()
            else:
                header_row = row0
                data_df = df.iloc[1:].copy()
            data_df.columns = header_row
            def _find_col(name, default_name):
                for c in data_df.columns:
                    if c and str(c).strip() == default_name:
                        return c
                for c in data_df.columns:
                    if name in str(c):
                        return c
                return None
            if month_prefix:
                col_month = _find_col((getattr(args, 'date_column', None) or '月份').strip(), '月份')
                if col_month is not None:
                    data_df[col_month] = data_df[col_month].astype(str).str.strip()
                    data_df = data_df[data_df[col_month].str.startswith(month_prefix)]
                    print(f"已按列「{col_month}」过滤前缀「{month_prefix}」，保留 {len(data_df)} 行。")
            if week_val:
                col_week = _find_col((getattr(args, 'date_column_week', None) or '周报日期').strip(), '周报日期')
                if col_week is not None:
                    data_df[col_week] = data_df[col_week].astype(str).str.strip()
                    data_df = data_df[data_df[col_week] == week_val]
                    print(f"已按列「{col_week}」过滤为「{week_val}」，保留 {len(data_df)} 行。")
            df = data_df.copy()
            df.columns = header_row
            if len(df) == 0:
                print("过滤后无数据。")
                sys.exit(0)
        
        # 决定输出方式
        is_explicit_range = all(x is not None for x in [args.row_from, args.row_to, args.col_from, args.col_to])
        
        if args.console:
            # 强制打印到控制台
            # 尝试使用 Markdown，如果缺少 tabulate 库则回退到 string
            try:
                print(df.to_markdown(index=False))
            except ImportError:
                 print("提示: 安装 tabulate 库可获得更好的 Markdown 表格输出 (pip install tabulate)")
                 print(df.to_string(index=False))
        elif is_explicit_range:
            # 精确范围模式：默认保持原有的打印行为 (简单文本表格)
            print(f"\n✅ 读取了 {len(data)} 个单元格:")
            print(df.to_string(index=False, header=False))
        else:
            # 全量模式且未指定 console：默认导出 CSV
            # 计算输出目录：AirsheetFile/output
            # 脚本位于 AirsheetFile/scripts/sheet_manager.py
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            output_dir = os.path.join(base_dir, 'output')
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            filename = f"{file_id}_sheet_{args.sheet_id}.csv"
            if week_val:
                safe_week = week_val.replace("/", "_").replace("\\", "_").replace(" ", "_")
                filename = f"{file_id}_sheet_{args.sheet_id}_{safe_week}.csv"
            filepath = os.path.join(output_dir, filename)
            
            use_header = bool(month_prefix or week_val)
            df.to_csv(filepath, index=False, header=use_header)
            print(f"\n✅ 数据已导出到: {filepath}")

    elif args.command == "upload":
        file_path = args.file
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)

        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            sys.exit(1)

        filename = os.path.basename(file_path)
        filesize = os.path.getsize(file_path)
        print(f"准备上传文件: {filename} ({filesize:,} bytes)")

        # 解析 parent_path
        parent_path = None
        if args.parent_path:
            try:
                parent_path = json.loads(args.parent_path)
                if not isinstance(parent_path, list):
                    print("错误: --parent_path 必须是 JSON 字符串数组。")
                    sys.exit(1)
            except json.JSONDecodeError:
                print("错误: --parent_path 必须是有效的 JSON 数组。")
                sys.exit(1)

        res = client.upload_file(
            file_path=file_path,
            drive_id=args.drive_id,
            parent_id=args.parent_id,
            parent_path=parent_path,
            on_name_conflict=args.on_conflict
        )

        if res and res.get('code', -1) == 0:
            file_data = res.get('data', {})
            file_id = file_data.get('id', '')
            file_name = file_data.get('name', '')

            print(f"\n✅ 上传成功！")
            print(f"   文件名: {file_name}")
            print(f"   文件ID: {file_id}")
            print(f"   大小: {file_data.get('size', 0):,} bytes")

            # 可选：开启分享
            if args.share and file_id:
                print("正在开启分享链接...")
                share_res = client.open_share(file_id, drive_id=args.drive_id)
                if share_res and 'data' in share_res:
                    share_url = share_res['data'].get('url', '未知')
                    print(f"🔗 分享链接: {share_url}")
                else:
                    print(f"⚠️  开启分享失败: {share_res}")
        else:
            print(f"❌ 上传失败。")
            sys.exit(1)

    elif args.command == "import_file":
        if not DataConverter:
            print("错误: 未找到 data_converter 模块。")
            sys.exit(1)
            
        file_id = args.file_id or client.resolve_file_id(args.url)
        
        # 1. 确定目标 Sheet ID
        target_sheet_id = args.sheet_id
        
        if args.new_sheet:
            print(f"正在创建新工作表 '{args.new_sheet}'...")
            res_sheet = client.create_worksheet(file_id, name=args.new_sheet, file_type=args.type)
            if res_sheet and res_sheet.get('code') == 0:
                target_sheet_id = res_sheet['data']['sheet_id']
                print(f"✅ 新工作表创建成功，ID: {target_sheet_id}")
            else:
                print(f"❌ 创建新工作表失败: {res_sheet}")
                sys.exit(1)
        
        # 2. 读取并转换数据
        print(f"正在读取本地文件: {args.file}...")
        try:
            df = DataConverter.load_file(args.file)
            # 默认包含表头
            rows_payloads = DataConverter.to_rows_payloads(df, include_header=True)
            total_rows = len(rows_payloads)
            print(f"解析成功，共 {total_rows} 行数据 (含表头)。")
        except Exception as e:
            print(f"❌ 读取文件失败: {e}")
            sys.exit(1)
            
        # 3. 批量写入
        # 注意：WPS add_rows 接口可能不支持一次性提交所有行（如果行数很多）。
        # 这里我们逐行或小批量提交。为了简单起见，目前实现为逐行提交。
        # TODO: 如果 API 支持批量 List[List[Cell]] 结构，可以优化此处。
        
        success_count = 0
        print(f"开始导入数据到工作表 {target_sheet_id}...")
        
        for i, row_data in enumerate(rows_payloads):
            # 简单的进度显示
            if i % 10 == 0:
                print(f"正在写入第 {i+1}/{total_rows} 行...", end='\r')
                
            res = client.add_rows(file_id, target_sheet_id, row_data, file_type=args.type)
            if res and res.get('code') == 0:
                success_count += 1
            else:
                print(f"\n❌ 第 {i+1} 行写入失败: {res}")
                
        print(f"\n✅ 导入完成！成功写入 {success_count}/{total_rows} 行。")

    elif args.command == "append":
        file_id = args.file_id or client.resolve_file_id(args.url)
        
        # 解析数据
        try:
            new_data = json.loads(args.data)
        except json.JSONDecodeError:
            print("错误: --data 必须是有效的 JSON")
            sys.exit(1)
            
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name

        print(f"正在下载文件 {file_id}...")
        client.download_file(file_id, tmp_path)
        
        try:
            # 读取现有数据
            df = pd.read_excel(tmp_path, sheet_name=args.sheet)
            
            # 从新数据创建 DataFrame
            new_df = pd.DataFrame(new_data)
            
            # 合并
            final_df = pd.concat([df, new_df], ignore_index=True)
            
            # 保存
            final_df.to_excel(tmp_path, index=False)
            
            # 上传
            print("正在上传更新后的内容...")
            client.update_file_content(file_id, tmp_path)
            print("✅ 追加成功。")
            
        except Exception as e:
            print(f"追加操作时出错: {e}")
        finally:
            os.unlink(tmp_path)

    elif args.command == "update":
        # 简化更新：在行/列索引处更新值
        file_id = args.file_id or client.resolve_file_id(args.url)
        
        if args.row is None or args.col is None:
            print("错误: 更新操作需要 --row 和 --col 参数（简化模式）")
            sys.exit(1)
            
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name

        print(f"正在下载文件 {file_id}...")
        client.download_file(file_id, tmp_path)
        
        try:
            df = pd.read_excel(tmp_path, sheet_name=args.sheet)
            
            # 更新值
            # 确保行存在
            if args.row < len(df):
                df.at[args.row, args.col] = args.value
                
                # 保存
                df.to_excel(tmp_path, index=False)
                
                # 上传
                print("正在上传更新后的内容...")
                client.update_file_content(file_id, tmp_path)
                print("✅ 更新成功。")
            else:
                print(f"错误: 行索引 {args.row} 超出范围。")
                
        except Exception as e:
            print(f"更新操作时出错: {e}")
        finally:
            os.unlink(tmp_path)

    else:
        parser.print_help()

if __name__ == "__main__":
    main()
