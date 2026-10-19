#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
采购商数据表 - 多邮箱拆分为独立行工具 (单邮箱导入版)
=============================================================================
功能说明：
1. 解决部分联系人包含 2~3 个邮箱无法直接导入系统的问题。
2. 将多邮箱单元格（以逗号分隔）按「一个邮箱对应一行记录」进行精准拆解。
3. 自动将「邮箱验证状态」（Valid/Risky/未验证）与具体邮箱按顺序一对一匹配。
4. 复制公司的所有背景与贸易画像字段，序号（#）重新连续编号（1, 2, 3...）。
5. 完美继承原表格全部排版（微软雅黑、暗蓝表头、对齐方式、列宽、边框、冻结窗格）。
"""

import os
import sys
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def clean_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def split_emails_in_workbook(input_file: str, output_file: str = None, filter_empty_emails: bool = False):
    """
    读取源 Excel，展开多邮箱为单行，并保存到新文件。
    :param input_file: 输入文件路径
    :param output_file: 输出文件路径，若为 None 则自动添加后缀 '_单邮箱导入版.xlsx'
    :param filter_empty_emails: 是否过滤掉无邮箱的行（针对纯邮箱导入系统，默认 False 保留完整企业记录）
    """
    if not os.path.exists(input_file):
        print(f"[错误] 输入文件不存在: {input_file}")
        return None

    if output_file is None:
        base_name, ext = os.path.splitext(input_file)
        suffix = "_单邮箱导入版" if not filter_empty_emails else "_单邮箱导入版(仅含有效邮箱)"
        output_file = f"{base_name}{suffix}{ext}"

    print("=" * 80)
    print(f" [开始转换] 多邮箱拆分为独立行")
    print(f" 源文件: {os.path.basename(input_file)}")
    print(f" 目标文件: {os.path.basename(output_file)}")
    print(f" 过滤无邮箱行: {'是' if filter_empty_emails else '否'}")
    print("=" * 80)

    src_wb = openpyxl.load_workbook(input_file, data_only=True)
    src_ws = src_wb.active

    # 读取表头与列定义
    headers = [cell.value for cell in src_ws[1]]
    total_cols = len(headers)
    
    # 定位邮箱列与状态列
    # 默认标准表头中：第 15 列（索引 14）为电子邮箱，第 16 列（索引 15）为邮箱验证状态
    email_col_idx = 14
    status_col_idx = 15
    for idx, h in enumerate(headers):
        if h and "电子邮箱" in str(h):
            email_col_idx = idx
        elif h and "邮箱验证状态" in str(h):
            status_col_idx = idx

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = src_ws.title or "单邮箱导入版"

    # 样式配置
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    data_font = Font(name="Microsoft YaHei", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 写入表头
    out_ws.row_dimensions[1].height = 26
    out_ws.append(headers)
    for col_idx in range(1, total_cols + 1):
        cell = out_ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        
        # 复制列宽
        col_letter = get_column_letter(col_idx)
        if col_letter in src_ws.column_dimensions:
            out_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width

    out_ws.freeze_panes = "A2"

    new_row_idx = 0
    split_count = 0
    original_data_rows = 0

    # 居中对齐列的索引集合（1-indexed）：#、国家、税号、交易日期、单数、角色类型、验证状态、更新时间
    center_col_set = {1, 3, 4, 6, 7, 14, 16, 18}

    for row_cells in src_ws.iter_rows(min_row=2, values_only=True):
        original_data_rows += 1
        row_list = list(row_cells)
        # 补足列数
        if len(row_list) < total_cols:
            row_list.extend([""] * (total_cols - len(row_list)))

        email_str = clean_str(row_list[email_col_idx])
        status_str = clean_str(row_list[status_col_idx])

        # 无邮箱情况
        if not email_str:
            if filter_empty_emails:
                continue
            new_row_idx += 1
            row_list[0] = new_row_idx
            out_ws.append(row_list)
            for c_idx in range(1, total_cols + 1):
                cell = out_ws.cell(row=new_row_idx + 1, column=c_idx)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = center_align if c_idx in center_col_set else left_align
            continue

        # 切分多个邮箱
        raw_emails = [e.strip() for e in email_str.split(",") if e.strip()]
        raw_statuses = [s.strip() for s in status_str.split(",") if s.strip()]

        if len(raw_emails) > 1:
            split_count += 1

        for i, single_email in enumerate(raw_emails):
            new_row_idx += 1
            cur_row = list(row_list)
            cur_row[0] = new_row_idx
            cur_row[email_col_idx] = single_email

            # 匹配对应验证状态
            if i < len(raw_statuses):
                cur_row[status_col_idx] = raw_statuses[i]
            elif len(raw_statuses) == 1:
                cur_row[status_col_idx] = raw_statuses[0]
            else:
                cur_row[status_col_idx] = "未验证"

            out_ws.append(cur_row)
            for c_idx in range(1, total_cols + 1):
                cell = out_ws.cell(row=new_row_idx + 1, column=c_idx)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = center_align if c_idx in center_col_set else left_align

    out_wb.save(output_file)
    print(f" [完成] 原始行数: {original_data_rows} 行 -> 展开后: {new_row_idx} 行 (拆分了 {split_count} 处多邮箱记录)")
    print(f" 生成文件: {os.path.abspath(output_file)}")
    return output_file

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="多邮箱拆解为独立行处理工具")
    parser.add_argument("input", nargs="?", default="欧盟采购商_近1年高频采购_有效采购决策人_(451-500页_500家公司).xlsx", help="输入 Excel 文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出 Excel 文件路径")
    parser.add_argument("--filter-empty", action="store_true", help="过滤掉没有邮箱的行")
    args = parser.parse_args()

    split_emails_in_workbook(args.input, args.output, args.filter_empty)
