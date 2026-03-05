# -*- coding: utf-8 -*-
import sys
import io
import pandas as pd
import json

# Fix Windows console encoding
if sys.platform == 'win32':
    if hasattr(sys.stdout, 'buffer') and sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'buffer') and sys.stderr.encoding != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

class DataConverter:
    @staticmethod
    def load_file(file_path):
        """读取 CSV 或 Excel 文件为 DataFrame"""
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path)
        elif file_path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path)
        else:
            raise ValueError("不支持的文件格式。仅支持 .csv, .xlsx, .xls")

    @staticmethod
    def to_rows_payloads(df, include_header=True):
        """
        将 DataFrame 转换为 WPS API 的 range_data 格式列表。
        返回: List[List[Dict]]，每个内部列表代表一行的数据。
        """
        rows_payloads = []
        
        # 1. 处理表头 (Header)
        if include_header:
            header_cells = []
            for col_idx, col_name in enumerate(df.columns):
                cell_data = {
                    "col": col_idx,
                    "op_type": "cell_operation_type_formula",
                    "formula": str(col_name) # 表头通常是文本
                }
                header_cells.append(cell_data)
            if header_cells:
                rows_payloads.append(header_cells)
        
        # 处理空值，替换为 None
        df = df.where(pd.notnull(df), None)
        
        for _, row in df.iterrows():
            row_cells = []
            for col_idx, val in enumerate(row):
                if val is None:
                    continue
                    
                cell_data = {
                    "col": col_idx,
                    "op_type": "cell_operation_type_formula"
                }
                
                # 类型转换逻辑
                if isinstance(val, (int, float)):
                    cell_data["formula"] = str(val)
                else:
                    val_str = str(val)
                    if val_str.startswith("="):
                        cell_data["formula"] = val_str
                    else:
                        # 对于字符串，如果不包含特殊字符，可以直接传
                        # 为了安全，非公式字符串最好用引号包裹，或者依赖 API 的智能推断
                        # 参考 sheet_manager 中的逻辑：
                        cell_data["formula"] = val_str
                        
                row_cells.append(cell_data)
            
            # 即使是一行空数据（只有 None），也应该添加一个空列表或者跳过？
            # 只有非空行才添加
            if row_cells:
                rows_payloads.append(row_cells)
                
        return rows_payloads
