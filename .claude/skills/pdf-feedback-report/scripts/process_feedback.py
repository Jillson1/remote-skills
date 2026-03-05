import pandas as pd
import argparse
import configparser
import os
import sys
import subprocess
import io
import shutil

# Fix Windows console encoding for emoji/special chars
if sys.platform == "win32" and hasattr(sys.stdout, "buffer"):
    if sys.stdout.encoding != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if sys.stderr.encoding != "utf-8":
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ---------------------------------------------------------
# EMBEDDED PROMPT INSTRUCTIONS
# (Moved to 'prompts/analysis_prompt.md')
# ---------------------------------------------------------


def _parse_file_id_from_url(url):
    """从在线表格 URL 解析出 file_id（如 ct3YewJ2rEDn）。"""
    if not url or not isinstance(url, str):
        return None
    base = url.split("?")[0].split("#")[0].rstrip("/")
    if "/" in base:
        return base.split("/")[-1]
    return base.strip()


def parse_date_range(date_str):
    """
    解析用户传入的日期区间，得到 (月份前缀, 周报日期) 用于导出过滤。
    - "26.2.1-2.4" / "26.2.1-2.4" → ("26-2", "2.1-2.4")  即 26年2月、周报 2.1-2.4
    - "25.5.3-5.8" / "25.5.3-5.8" → ("25-5", "5.3-5.8")  即 25年5月、周报 5.3-5.8
    - "25年5月3号到5月8号" → ("25-5", "5.3-5.8")
    - "5.3-5.8"（仅周报）→ 缺少年份时用配置或当年简写，如 ("26-5", "5.3-5.8")
    - 未传入或空 → 返回 None，表示使用「最近一周」或全量（由调用方用 config 决定）
    返回: (month_prefix, week_range) 或 (None, None)
    """
    import re
    if not date_str or not isinstance(date_str, str):
        return None, None
    s = date_str.strip()
    if not s:
        return None, None
    # 25-1.16-1.22（年-月.日-月.日）→ 25年1月、周报 1.16-1.22
    m = re.match(r"(\d{2})-(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})", s)
    if m:
        yy, mm, d1, m2, d2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4)), m.group(5)
        month_prefix = f"{yy}-{mm}"
        week_range = f"{mm}.{int(d1)}-{m2}.{int(d2)}"
        return month_prefix, week_range
    # 25年5月3号到5月8号 / 25年5月3日-5月8日
    m = re.match(r"(\d{2})年(\d{1,2})月.*?(\d{1,2})[号日].*?(\d{1,2})[号日]", s)
    if m:
        yy, mm, d1, d2 = m.group(1), m.group(2), m.group(3), m.group(4)
        month_prefix = f"{yy}-{int(mm)}"
        week_range = f"{int(mm)}.{int(d1)}-{int(mm)}.{int(d2)}"
        return month_prefix, week_range
    # 26.2.1-2.4 或 26.2.1-2.4（年.月.日-日 或 年.月.日-月.日）
    m = re.match(r"(\d{2})\.(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})", s)
    if m:
        yy, mm, d1, m2, d2 = m.group(1), int(m.group(2)), m.group(3), m.group(4), m.group(5)
        month_prefix = f"{yy}-{mm}"
        week_range = f"{mm}.{int(d1)}-{int(m2)}.{int(d2)}"
        return month_prefix, week_range
    # 26.2.1-2.4 形式（年.月.日-日，同一月）
    m = re.match(r"(\d{2})\.(\d{1,2})\.(\d{1,2})-(\d{1,2})", s)
    if m:
        yy, mm, d1, d2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        month_prefix = f"{yy}-{mm}"
        week_range = f"{mm}.{int(d1)}-{int(d2)}"
        return month_prefix, week_range
    # 5.3-5.8（仅月.日-月.日，缺少年份：用当年，如 26）
    m = re.match(r"(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})", s)
    if m:
        m1, d1, m2, d2 = int(m.group(1)), m.group(2), m.group(3), m.group(4)
        from datetime import datetime
        yy = str(datetime.now().year)[2:]  # 26
        month_prefix = f"{yy}-{m1}"
        week_range = f"{m1}.{int(d1)}-{int(m2)}.{int(d2)}"
        return month_prefix, week_range
    # 2.1-2.4（仅日-日，同一月，缺年月：用当年当月）
    m = re.match(r"(\d{1,2})\.(\d{1,2})-(\d{1,2})", s)
    if m:
        mm, d1, d2 = int(m.group(1)), m.group(2), m.group(3)
        from datetime import datetime
        yy = str(datetime.now().year)[2:]
        month_prefix = f"{yy}-{mm}"
        week_range = f"{mm}.{int(d1)}-{int(d2)}"
        return month_prefix, week_range
    # 无法解析时视为周报日期原文，月份前缀为空（仅按周报日期过滤）
    return None, s


def try_export_sheet_to_local(url, sheet_id, timeout=600, use_existing_if_exists=True,
                              month_prefix=None, week_range=None):
    """
    优先使用 AirsheetFile 的 sheet_manager.py 将 Sheet1 导出到本地 CSV。
    若提供 month_prefix / week_range，则只导出该日期区间的数据（月份前缀 + 周报日期）。
    成功返回 (True, csv_path)，失败返回 (False, None)。
    """
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    airsheet_scripts = os.path.join(project_root, "AirsheetFile", "scripts")
    sheet_manager_py = os.path.join(project_root, "AirsheetFile", "scripts", "sheet_manager.py")
    airsheet_output_dir = os.path.join(project_root, "AirsheetFile", "output")

    if not os.path.isfile(sheet_manager_py):
        return False, None

    file_id = _parse_file_id_from_url(url)
    if not file_id:
        return False, None

    if week_range:
        safe_week = week_range.replace("/", "_").replace("\\", "_").replace(" ", "_")
        expected_csv = os.path.join(airsheet_output_dir, f"{file_id}_sheet_{sheet_id}_{safe_week}.csv")
    else:
        expected_csv = os.path.join(airsheet_output_dir, f"{file_id}_sheet_{sheet_id}.csv")

    if use_existing_if_exists and not (month_prefix or week_range) and os.path.isfile(expected_csv) and os.path.getsize(expected_csv) > 0:
        return True, expected_csv
    if use_existing_if_exists and (month_prefix or week_range) and os.path.isfile(expected_csv) and os.path.getsize(expected_csv) > 0:
        return True, expected_csv

    cmd = [sys.executable, "sheet_manager.py", "get_range", "--url", url, "--sheet_id", str(sheet_id)]
    if month_prefix:
        cmd.extend(["--filter-date-prefix", month_prefix, "--date-column", "月份"])
    if week_range:
        cmd.extend(["--filter-week", week_range, "--date-column-week", "周报日期"])

    try:
        proc = subprocess.run(
            cmd,
            cwd=airsheet_scripts,
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding="utf-8",
            errors="replace",
        )
        if proc.returncode != 0:
            return False, None
        if not os.path.isfile(expected_csv) or os.path.getsize(expected_csv) == 0:
            return False, None
        return True, expected_csv
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return False, None


def load_df_from_local_csv(csv_path):
    """从 sheet_manager 导出的本地 CSV 加载 DataFrame。若首行非表头则尝试第二行作为表头。"""
    df = pd.read_csv(csv_path, header=0, encoding="utf-8-sig", low_memory=False)
    # 若第一行不是有效表头（列为 Unnamed 或缺失「周报日期」），尝试用第二行作表头
    if "周报日期" not in df.columns and len(df) > 1:
        df = pd.read_csv(csv_path, header=1, encoding="utf-8-sig", low_memory=False)
    return df


def load_config():
    """
    加载配置文件
    """
    config = configparser.ConfigParser()
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_path = os.path.join(base_dir, 'config', 'config.properties')
    
    if os.path.exists(config_path):
        config.read(config_path, encoding='utf-8')
        return config
    return None

def fetch_online_sheet(url, sheet_id=1):
    """
    从在线表格读取数据，返回 pandas DataFrame
    
    :param url: 在线表格 URL (如 https://365.kdocs.cn/l/xxxxx)
    :param sheet_id: 工作表 ID (默认为 1，即第一个 Sheet)
    :return: pandas DataFrame
    """
    # 动态导入 AirsheetFile 的 WpsClient
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    airsheet_scripts = os.path.join(project_root, 'AirsheetFile', 'scripts')
    
    if airsheet_scripts not in sys.path:
        sys.path.insert(0, airsheet_scripts)
    
    try:
        from wps_client import WpsClient
    except ImportError as e:
        print(f"❌ 无法导入 AirsheetFile 模块: {e}")
        print("请确保 AirsheetFile/scripts/wps_client.py 存在且 AirsheetFile 已正确配置。")
        sys.exit(1)
    
    client = WpsClient()
    file_id = client.resolve_file_id(url)
    
    if not file_id:
        print(f"❌ 无法从 URL 解析文件 ID: {url}")
        sys.exit(1)
    
    print(f"📡 正在从在线表格读取数据...")
    print(f"   文件 ID: {file_id}")
    print(f"   工作表 ID: {sheet_id}")
    
    # 1. 探测数据范围
    probe_row_limit = 10
    probe_col_limit = 100
    max_col = 0
    
    res = client.get_range_data(file_id, sheet_id, 0, probe_row_limit, 0, probe_col_limit, file_type='auto')
    if res and res.get('code') == 0:
        data = res['data'].get('range_data', [])
        if data:
            for cell in data:
                if cell.get('cell_text') or str(cell.get('original_cell_value', '')):
                    max_col = max(max_col, cell.get('col_from', 0))
    else:
        print(f"❌ 探测数据范围失败: {res}")
        sys.exit(1)
    
    print(f"   探测到有效列数: {max_col + 1}")
    
    # 2. 分页读取所有数据
    all_cells = []
    batch_size = 1000
    current_row = 0
    
    while True:
        res = client.get_range_data(file_id, sheet_id, current_row, current_row + batch_size - 1, 0, max_col, file_type='auto')
        
        if res and res.get('code') == 0:
            batch_data = res['data'].get('range_data', [])
            
            has_content = False
            for cell in batch_data:
                if cell.get('cell_text') or str(cell.get('original_cell_value', '')):
                    has_content = True
                    break
            
            if not batch_data or not has_content:
                break
                
            all_cells.extend(batch_data)
            current_row += batch_size
            
            if current_row > 100000:
                print("⚠️ 达到行数上限 (100,000)，停止读取。")
                break
        else:
            print(f"❌ 读取数据失败: {res}")
            break
    
    print(f"   读取完成，共获取 {len(all_cells)} 个单元格。")
    
    if not all_cells:
        print("❌ 未获取到任何数据")
        sys.exit(1)
    
    # 3. 转换为 DataFrame
    # 找出边界
    max_r = max(cell.get('row_from', 0) for cell in all_cells)
    max_c = max(cell.get('col_from', 0) for cell in all_cells)
    min_r = min(cell.get('row_from', 0) for cell in all_cells)
    min_c = min(cell.get('col_from', 0) for cell in all_cells)
    
    rows_count = max_r - min_r + 1
    cols_count = max_c - min_c + 1
    
    matrix = [['' for _ in range(cols_count)] for _ in range(rows_count)]
    
    for cell in all_cells:
        r = cell.get('row_from', 0) - min_r
        c = cell.get('col_from', 0) - min_c
        val = cell.get('cell_text')
        orig = cell.get('original_cell_value')
        final_val = val if val is not None else orig
        if final_val is None:
            final_val = ""
        matrix[r][c] = final_val
    
    # 第一行作为表头
    if rows_count > 1:
        headers = matrix[0]
        data_rows = matrix[1:]
        df = pd.DataFrame(data_rows, columns=headers)
    else:
        df = pd.DataFrame(matrix)
    
    print(f"✅ 成功加载 {len(df)} 行数据")
    return df


def write_compare_stats(focus_df, compare_csv_path, stats_output_path, focus_total=None, compare_total=None):
    """
    按二级分类统计关注周与对比周数量并计算增长率，写入 UTF-8 文件，供报告生成使用。
    避免控制台编码导致的中文乱码，统计结果统一落盘。
    """
    col = "二级分类"
    if col not in focus_df.columns:
        return False
    try:
        compare_df = pd.read_csv(compare_csv_path, encoding="utf-8")
    except Exception:
        try:
            compare_df = pd.read_csv(compare_csv_path, encoding="utf-8-sig")
        except Exception as e:
            print(f"⚠️ 读取对比周 CSV 失败，跳过统计文件生成: {e}")
            return False
    if col not in compare_df.columns:
        print(f"⚠️ 对比周 CSV 中无「{col}」列，跳过统计文件生成。")
        return False
    fc = focus_df[col].fillna("(未分类)").value_counts().sort_index()
    cc = compare_df[col].fillna("(未分类)").value_counts().sort_index()
    all_cats = sorted(set(fc.index) | set(cc.index))
    n_focus = focus_total if focus_total is not None else len(focus_df)
    n_compare = compare_total if compare_total is not None else len(compare_df)
    total_rate = round((n_focus - n_compare) / n_compare * 100) if n_compare else 0
    lines = [
        "focus_total\t%d" % n_focus,
        "compare_total\t%d" % n_compare,
        "total_rate\t%d" % total_rate,
    ]
    for cat in all_cats:
        f = int(fc.get(cat, 0))
        c = int(cc.get(cat, 0))
        rate = round((f - c) / c * 100) if c else (100 if f else 0)
        lines.append("%s\t%d\t%d\t%d" % (cat, f, c, rate))
    try:
        with open(stats_output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"✅ 已生成对比统计文件（UTF-8）: {stats_output_path}")
        return True
    except Exception as e:
        print(f"⚠️ 写入统计文件失败: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="从在线表格读取用户反馈数据并生成分析上下文。")
    
    # 日期可选：未传则使用配置的「最近一周」或导出全量后再按配置过滤
    parser.add_argument("--date", default=None, help="关注周日期区间。例: 26-2.1-2.4、26.2.1-2.4。不传则用配置 DEFAULT_WEEK_RANGE / DEFAULT_MONTH_PREFIX")
    parser.add_argument("--compare-date", default=None, help="对比周日期区间。例: 26-1.22-1.28。填写后将导出关注周与对比周两个 CSV 到 cache，并在上下文中说明以填写结论对比表")
    
    # 可选参数
    parser.add_argument("--url", default=None, help="在线表格 URL，不填则从配置文件读取")
    parser.add_argument("--output", default="feedback_context.md", help="输出上下文文件名 (默认写入 pdf-feedback-report/cache/feedback_context.md)")
    parser.add_argument("--keyword", default=None, help="关键词筛选 (如 'AI讲解')")
    parser.add_argument("--sheet_id", type=int, default=None, help="工作表 ID，不填则从配置文件读取 (默认: 1)")
    parser.add_argument("--template_file", default=None, help="报告模板文件路径")
    parser.add_argument("--prompt_file", default=None, help="分析提示词文件路径")
    
    args = parser.parse_args()
    
    # Resolve paths：临时文件（上下文、CSV）统一写入 pdf-feedback-report/cache
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    cache_dir = os.path.join(base_dir, "cache")
    os.makedirs(cache_dir, exist_ok=True)
    if not os.path.isabs(args.output):
        args.output = os.path.join(cache_dir, os.path.basename(args.output) or "feedback_context.md")
    
    # Load config
    config = load_config()
    
    # Resolve URL: 命令行参数 > 配置文件
    url = args.url
    if not url:
        if config and 'DEFAULT' in config and 'FEEDBACK_SOURCE_URL' in config['DEFAULT']:
            url = config['DEFAULT']['FEEDBACK_SOURCE_URL']
            print(f"📄 从配置文件读取数据源 URL")
        else:
            print("❌ 未提供 --url 参数，且配置文件中未找到 FEEDBACK_SOURCE_URL")
            sys.exit(1)
    
    # Resolve sheet_id: 命令行参数 > 配置文件 > 默认值
    sheet_id = args.sheet_id
    if sheet_id is None:
        if config and 'DEFAULT' in config and 'DEFAULT_SHEET_ID' in config['DEFAULT']:
            sheet_id = int(config['DEFAULT']['DEFAULT_SHEET_ID'])
        else:
            sheet_id = 1
    
    # Default paths
    if args.template_file is None:
        args.template_file = os.path.join(base_dir, "templates", "report.md")
        
    if args.prompt_file is None:
        args.prompt_file = os.path.join(base_dir, "prompts", "analysis_prompt.md")
    
    # 解析日期区间：用于导出时只拉取该区间数据，以及报告标题
    if args.date:
        month_prefix, week_range = parse_date_range(args.date)
        target_date = week_range or args.date.strip()
    else:
        month_prefix = None
        week_range = None
        if config and "DEFAULT" in config:
            month_prefix = config["DEFAULT"].get("DEFAULT_MONTH_PREFIX", "").strip() or None
            week_range = config["DEFAULT"].get("DEFAULT_WEEK_RANGE", "").strip() or None
        target_date = week_range or (args.date.strip() if args.date else None)
    if not target_date:
        print("❌ 未提供 --date 且配置中无 DEFAULT_WEEK_RANGE，无法确定周报日期。")
        sys.exit(1)
    
    compare_month_prefix = None
    compare_week_range = None
    if args.compare_date:
        compare_month_prefix, compare_week_range = parse_date_range(args.compare_date)
        if not compare_week_range:
            compare_week_range = args.compare_date.strip()
        print(f"对比周: {compare_week_range}" + (f"（月份: {compare_month_prefix}）" if compare_month_prefix else ""))
    
    print(f"关注周: {target_date}" + (f"（月份: {month_prefix}）" if month_prefix else ""))
    
    # 1. 优先使用 sheet_manager 按日期区间导出关注周 Sheet1 到本地；失败则退化为在线读取
    df = None
    exported, csv_path = try_export_sheet_to_local(url, sheet_id, month_prefix=month_prefix, week_range=week_range)
    if exported and csv_path:
        try:
            df = load_df_from_local_csv(csv_path)
            print(f"✅ 已从本地导出文件加载数据: {csv_path}")
        except Exception as e:
            print(f"⚠️ 读取本地导出文件失败: {e}，退化为在线读取...")
            df = None
    if df is None:
        print("⚠️ 本地导出 Sheet 失败或未使用，退化为在线表格读取...")
        df = fetch_online_sheet(url, sheet_id)
    
    # 1.1 若填写了对比周：导出对比周并复制关注周/对比周两个 CSV 到 cache
    focus_csv_in_cache = None
    compare_csv_in_cache = None
    if compare_week_range:
        if exported and csv_path and os.path.isfile(csv_path):
            safe_focus = target_date.replace("/", "_").replace(" ", "_").replace("\\", "_")
            focus_csv_in_cache = os.path.join(cache_dir, f"focus_{safe_focus}.csv")
            shutil.copy2(csv_path, focus_csv_in_cache)
            print(f"已复制关注周 CSV 到: {focus_csv_in_cache}")
        exported_compare, compare_csv_path = try_export_sheet_to_local(
            url, sheet_id, month_prefix=compare_month_prefix, week_range=compare_week_range, use_existing_if_exists=True
        )
        if exported_compare and compare_csv_path and os.path.isfile(compare_csv_path):
            safe_compare = compare_week_range.replace("/", "_").replace(" ", "_").replace("\\", "_")
            compare_csv_in_cache = os.path.join(cache_dir, f"compare_{safe_compare}.csv")
            shutil.copy2(compare_csv_path, compare_csv_in_cache)
            print(f"已导出并复制对比周 CSV 到: {compare_csv_in_cache}")
        
    # 2. Filter Data（若导出时已按周报日期过滤，则不再重复过滤）
    if "周报日期" not in df.columns:
        print("❌ 数据中未找到 '周报日期' 列。")
        print("   可用列名:", df.columns.tolist())
        sys.exit(1)
    
    df["周报日期"] = df["周报日期"].astype(str).str.strip()
    
    formatted_date = target_date.replace("-", "至")
    report_filename = f"PDF增值 用户反馈分析报告（{formatted_date}）.md"
    report_title = f"# PDF增值 用户反馈分析报告 ({formatted_date})"
    csv_output = os.path.splitext(args.output)[0] + ".csv"
    
    if exported and (month_prefix or week_range):
        filtered_df = df
        print(f"已使用导出时的日期区间过滤，共 {len(filtered_df)} 条。")
    else:
        filtered_df = df[df["周报日期"] == target_date]
        if filtered_df.empty:
            print(f"警告: 未找到日期为 {target_date} 的数据")
        
    # 2.1 Filter by Keyword (Optional)
    if args.keyword:
        keyword = args.keyword.strip()
        print(f"筛选关键词: {keyword}")
        
        # Columns to search for keyword
        search_cols = ['内容', '问题类型', '功能点', '二级分类']
        # Filter columns that actually exist
        search_cols = [col for col in search_cols if col in filtered_df.columns]
        
        if search_cols:
            # Create a mask for rows where any of the search columns contain the keyword
            mask = pd.DataFrame(False, index=filtered_df.index, columns=['match'])
            for col in search_cols:
                # Use str.contains with case=False and na=False
                mask['match'] |= filtered_df[col].astype(str).str.contains(keyword, case=False, na=False)
            
            filtered_df = filtered_df[mask['match']]
            print(f"关键词筛选后记录数: {len(filtered_df)}")
            
            # Update report title to reflect keyword
            report_title += f" - {keyword}专项分析"
            report_filename = report_filename.replace(".md", f"_{keyword}.md")
        else:
            print("警告: 未找到可用于关键词搜索的列")

    # 3. Format Data (Markdown Table)
    # Define critical columns for analysis to reduce token usage and noise
    # Based on the analysis needs: ID, User, Content, and Context tags
    target_columns = [
        '序号', '用户名称', '内容', 
        '问题类型', '反馈产品', '端', '用户权益', 
        '功能点', '二级分类'
    ]
    
    # Filter columns that actually exist in the dataframe
    existing_columns = [col for col in target_columns if col in filtered_df.columns]
    
    if not existing_columns:
        print("警告: 未找到任何目标关键列，将使用所有列。")
        final_df = filtered_df
    else:
        final_df = filtered_df[existing_columns].copy()
        
    # Clean up content to prevent markdown table breakage (replace newlines)
    if '内容' in final_df.columns:
        final_df['内容'] = final_df['内容'].astype(str).str.replace('\n', ' ').str.replace('\r', '')

    try:
        # Generate markdown table without truncation
        # pd.set_option is not sufficient for to_markdown in some versions, 
        # but passing string IO or ensuring conversion helps.
        # The key is that to_markdown usually outputs full table unless configured otherwise,
        # but if df is large, pandas display options might interfere if we just str(df).
        # We explicitly use to_markdown which should render all rows by default.
        # However, to be safe, we can manually check or enforce content.
        
        # NOTE: to_markdown() by default does NOT truncate rows. 
        # The truncation usually happens if we use str(df) or print(df).
        # But if the user saw truncation, it might be due to a previous fallback to str(final_df).
        
        # Let's force full display just in case
        with pd.option_context('display.max_rows', None, 'display.max_columns', None):
             data_md = final_df.to_markdown(index=False)
             
    except ImportError:
        # Fallback if tabulate is not installed
        data_md = final_df.to_string(index=False)
    except AttributeError:
        data_md = str(final_df)

    # 4. Read Template
    template_content = ""
    if os.path.exists(args.template_file):
        with open(args.template_file, 'r', encoding='utf-8') as f:
            template_content = f.read()
            
        # Inject dynamic title
        # Replace the first line or known title with the specific date range title
        if "# PDF增值 用户反馈分析报告" in template_content:
            template_content = template_content.replace("# PDF增值 用户反馈分析报告", report_title)
    else:
        print(f"警告: 模板文件 '{args.template_file}' 未找到。")

    # 5. Construct Final Context for AI
    final_output = []
    
    # Section 1: Instructions (From Prompt File)
    system_prompt_content = ""
    if os.path.exists(args.prompt_file):
        with open(args.prompt_file, 'r', encoding='utf-8') as f:
            system_prompt_content = f.read()
    else:
        print(f"警告: 提示词文件 '{args.prompt_file}' 未找到。")
        
    final_output.append("# 1. ROLE & INSTRUCTIONS")
    final_output.append(system_prompt_content)
    
    # Add Filename Instruction
    final_output.append(f"\n[重要输出指令]")
    final_output.append(f"请将生成的报告内容保存到以下特定文件路径：")
    final_output.append(f"`pdf-feedback-report/reports/{report_filename}`")
    
    final_output.append("\n" + "="*50 + "\n")
    
    # Section 2: Data Source
    final_output.append("# 2. 数据源 (按日期筛选)")
    final_output.append(f"**关注周日期**: {target_date}")
    final_output.append(f"**关注周记录数量**: {len(filtered_df)}")
    if compare_csv_in_cache:
        final_output.append(f"**对比周日期**: {compare_week_range}")
        final_output.append(f"**说明**: 请读取关注周与对比周两个 CSV，按二级分类统计数量并计算增长率，填写报告中的「结论总结」与「结论对比表」。")
        stats_file = os.path.join(cache_dir, "feedback_compare_stats.txt")
        final_output.append(f"**统计结果文件（可选）**: 脚本已自动生成 `{stats_file}`（UTF-8），内含 focus_total、compare_total、total_rate 及各二级分类的关注周/对比周数量与增长率。可直接使用该文件填写「结论总结」与「结论对比表」，无需再解析 CSV 计算。")
    
    final_output.append(f"\n**[重要指令] 数据文件位置**:")
    final_output.append(f"> 关注周数据已清洗并保存为 CSV。请读取以下文件以获取关注周全量反馈：")
    final_output.append(f"> 文件路径: `{csv_output}`")
    if compare_csv_in_cache:
        final_output.append(f"\n> **对比周**：已导出对比周数据，请同时读取以下文件以计算增长率并填写「结论总结」与「结论对比表」：")
        final_output.append(f"> 关注周 CSV（可同上）: `{csv_output}`")
        final_output.append(f"> 对比周 CSV: `{compare_csv_in_cache}`")
        final_output.append(f"\n> **结论填写要求**：根据两 CSV 按「二级分类」统计各功能点数量，计算 增长率 = (关注周数量−对比周数量)/对比周数量×100（取整）。正常波动范围默认 ±20%，超出则为异常、可在「关注人」列标注。")
    else:
        final_output.append(f"> \n> **注意**：不要依赖本文件中的预览数据（如果有），必须直接读取上述 CSV 文件进行全量分析。")
    
    # We can still include a small sample (head) just for quick glance, but not the full table
    final_output.append("\n**数据预览 (前 5 行)**:")
    try:
        sample_md = final_df.head(5).to_markdown(index=False)
        final_output.append(sample_md)
    except:
        final_output.append(str(final_df.head(5)))
        
    final_output.append("\n" + "="*50 + "\n")
    
    # Section 3: Output Requirement (Template)
    final_output.append("# 3. 输出模板")
    final_output.append("请严格按照以下模板结构生成最终报告：")
    final_output.append("```markdown")
    final_output.append(template_content)
    final_output.append("```")

    # 6. Write to file
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write("\n".join(final_output))
        
    print(f"成功创建上下文文件: {args.output}")

    # 7. Write to CSV (Optional but requested)
    # csv_output is already defined above
    try:
        # Use final_df (cleaned data) instead of filtered_df (raw data)
        final_df.to_csv(csv_output, index=False, encoding='utf-8-sig')
        print(f"成功创建 CSV 筛选结果: {csv_output}")
    except Exception as e:
        print(f"警告: 创建 CSV 文件失败: {e}")

    # 8. 若存在对比周：自动生成按二级分类的统计结果（UTF-8 落盘，避免控制台中文乱码）
    if compare_csv_in_cache and os.path.isfile(compare_csv_in_cache):
        stats_path = os.path.join(cache_dir, "feedback_compare_stats.txt")
        write_compare_stats(
            final_df,
            compare_csv_in_cache,
            stats_path,
            focus_total=len(filtered_df),
            compare_total=None,
        )

    print(f"下一步: 请使用此文件作为上下文，让 AI Agent 生成报告。")
    sys.stdout.flush()
    sys.stderr.flush()

if __name__ == "__main__":
    main()
