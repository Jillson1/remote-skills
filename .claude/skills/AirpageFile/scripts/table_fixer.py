import re

def _parse_markdown_table(markdown_table_str):
    """
    简单的 Markdown 表格解析器。
    返回一个二维列表：rows[row_index][col_index] = cell_content
    """
    lines = markdown_table_str.strip().split('\n')
    rows = []
    for line in lines:
        line = line.strip()
        if not line.startswith('|'): continue
        
        # 去掉首尾的 |
        content = line.strip('|')
        
        # 分割单元格
        cells = [c.strip() for c in content.split('|')]
        
        # 跳过分隔行 (e.g. ---|:---:)
        if all(re.match(r'^[\s:\-]+$', c) for c in cells):
            continue
            
        rows.append(cells)
    return rows

def _extract_all_table_mds(markdown_content):
    """提取所有连续的表格块"""
    lines = markdown_content.split('\n')
    tables = []
    current_table = []
    in_table = False
    
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|'):
            if not in_table:
                in_table = True
                current_table = []
            current_table.append(stripped)
        else:
            if in_table:
                # 表格结束，只有当看起来像表格时才添加 (至少有一行)
                if current_table:
                    tables.append('\n'.join(current_table))
                in_table = False
                current_table = []
                
    # 处理文件末尾是表格的情况
    if in_table and current_table:
        tables.append('\n'.join(current_table))
    
    return tables

def fix_table_blocks(client, file_id, blocks, markdown_content, debug=False):
    """
    检查 Blocks 中的表格，根据 Markdown 原文修复带 <br> 的单元格
    支持多表格处理
    """
    # 1. 提取 Markdown 中的所有表格文本
    if debug:
        print("[DEBUG] 正在尝试从 Markdown 提取表格...")
    
    md_table_strs = _extract_all_table_mds(markdown_content)
    
    if not md_table_strs:
        if debug:
            print("[DEBUG] 未提取到表格文本。")
        return blocks

    if debug:
        print(f"[DEBUG] 共提取到 {len(md_table_strs)} 个 Markdown 表格。")

    # 2. 收集 Blocks 中的所有 Table Block
    def collect_all_table_blocks(block_list):
        tables = []
        # 如果传入的是 dict (单个 block)，转为 list
        targets = block_list if isinstance(block_list, list) else [block_list]
        
        for b in targets:
            if not isinstance(b, dict): continue
            
            if b.get('type') == 'table':
                tables.append(b)
            
            # 递归查找 (例如在 columns, quote 等容器中)
            # 注意：不建议在表格里套表格，但在 doc 结构里 table 可能在任何层级
            if 'content' in b and isinstance(b['content'], list):
                tables.extend(collect_all_table_blocks(b['content']))
        return tables

    wps_table_blocks = collect_all_table_blocks(blocks if isinstance(blocks, list) else blocks.get('blocks', []))
    
    if not wps_table_blocks:
        if debug:
            print("[DEBUG] 在转换后的 Blocks 中未找到任何 Table Block。")
        return blocks
    
    if debug:
        print(f"[DEBUG] WPS Blocks 中共找到 {len(wps_table_blocks)} 个表格块。")

    # 3. 按顺序匹配并修复
    # 假设 Markdown 提取顺序与 WPS Block 顺序一致
    count = min(len(md_table_strs), len(wps_table_blocks))
    
    for t_idx in range(count):
        md_table_str = md_table_strs[t_idx]
        wps_table_block = wps_table_blocks[t_idx]
        
        parsed_rows = _parse_markdown_table(md_table_str)
        wps_rows = wps_table_block.get('content', [])
        
        if debug:
            print(f"[DEBUG] >>> 处理第 {t_idx + 1} 个表格: MD行数={len(parsed_rows)}, WPS行数={len(wps_rows)}")

        # 遍历行
        for r_idx, wps_row in enumerate(wps_rows):
            if r_idx >= len(parsed_rows): break
            
            wps_cells = wps_row.get('content', [])
            md_cells = parsed_rows[r_idx]
            
            # 遍历列
            for c_idx, wps_cell in enumerate(wps_cells):
                if c_idx >= len(md_cells): break
                
                md_content = md_cells[c_idx]
                
                # 检查是否存在 <br>
                if '<br>' in md_content or '<br/>' in md_content:
                    if debug:
                        print(f"    [修复] 单元格 ({r_idx}, {c_idx}) 包含 <br>")
                    
                    # 预处理：将 HTML <br> 替换为 Markdown 的分段符 (\n\n)
                    fixed_content = md_content.replace('<br>', '\n\n').replace('<br/>', '\n\n')
                    
                    # 递归调用转换接口
                    fixed_blocks = client.convert_markdown_to_blocks(file_id, fixed_content)
                    
                    if fixed_blocks:
                        new_content = []
                        if isinstance(fixed_blocks, dict) and 'blocks' in fixed_blocks:
                            new_content = fixed_blocks['blocks']
                        elif isinstance(fixed_blocks, list):
                            new_content = fixed_blocks
                        
                        if new_content:
                            wps_cell['content'] = new_content
                
    return blocks
