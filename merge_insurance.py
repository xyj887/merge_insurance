"""
投保数据连续日期合并工具

功能：
  读取投保数据Excel，按"证件号码+方案名称"分组，
  合并连续的保障日期区间，输出汇总Excel。

用法：
  python3 merge_insurance.py <输入文件路径> [输出文件路径]

示例：
  python3 merge_insurance.py ODS2026-06投保数据.xlsx
  python3 merge_insurance.py ODS2026-06投保数据.xlsx output/汇总结果.xlsx

如果未指定输出路径，默认在输入文件同目录下生成"<原文件名>_汇总.xlsx"。
"""

import sys
import os
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def merge_intervals(dates):
    """将日期列表按连续性分组，返回 [(start, end), ...]"""
    if not dates:
        return []
    dates = sorted(set(dates))
    intervals = []
    start = dates[0]
    prev = dates[0]
    for d in dates[1:]:
        if (d - prev).days == 1:
            prev = d
        else:
            intervals.append((start, prev))
            start = d
            prev = d
    intervals.append((start, prev))
    return intervals


# 多格式列名映射：统一映射为脚本内部使用的标准列名
COLUMN_MAPS = [
    {  # 格式1：ODS明细
        '姓名': '姓名', '证件号码': '证件号码', '方案名称': '方案名称',
        '保障生效时间': '保障生效时间', '保障失效时间': '保障失效时间',
    },
    {  # 格式2：企业投保数据
        '投保姓名': '姓名', '投保身份证号': '证件号码', '保险方案': '方案名称',
        '保障日期开始': '保障生效时间', '保障日期结束': '保障失效时间',
    },
    {  # 格式3：ODS传统渠道
        '姓名': '姓名', '证件号码': '证件号码', '保险方案': '方案名称',
        '增员生效日期': '保障生效时间', '投保止期': '保障失效时间',
    },
]


def normalize_columns(df, log=None):
    """根据文件列名自动匹配并重命名为标准列名"""
    log = log or print
    for mapping in COLUMN_MAPS:
        matched = {k for k in mapping if k in df.columns}
        if len(matched) >= 4:  # 匹配到大部分关键列即认为命中
            rename = {k: v for k, v in mapping.items() if k in df.columns and k != v}
            if rename:
                log(f"  列名映射: {rename}")
                df.rename(columns=rename, inplace=True)
            return df
    return df


# 汇总所需的核心字段
REQUIRED_COLUMNS = ['姓名', '证件号码', '方案名称', '保障生效时间', '保障失效时间']


def build_result_df(src_path, log=None):
    """读取Excel并汇总，返回汇总后的DataFrame。

    log: 可选的日志回调（默认print），GUI通过它把进度显示到界面。
    """
    log = log or print

    log(f"读取文件: {src_path}")
    df = pd.read_excel(src_path)
    log(f"原始行数: {len(df)}")

    # 列名标准化
    df = normalize_columns(df, log=log)

    # 校验必需字段，缺失时给出友好提示
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"文件缺少必需的列: {missing}\n"
            f"当前文件的列: {list(df.columns)}\n"
            f"请确认是否选择了正确的投保数据文件。"
        )

    # 预处理
    df['证件号码'] = df['证件号码'].astype(str)
    df = df[df['保障生效时间'].notna() & df['保障失效时间'].notna()]
    df['保障生效时间'] = pd.to_datetime(df['保障生效时间'])
    df['保障失效时间'] = pd.to_datetime(df['保障失效时间'])
    df['生效日期'] = df['保障生效时间'].dt.date

    results = []

    for (id_num, plan), group in df.groupby(['证件号码', '方案名称']):
        name = group['姓名'].iloc[0]
        dates = sorted(group['生效日期'].unique())
        intervals = merge_intervals(dates)

        for idx, (seg_start, seg_end) in enumerate(intervals):
            seg_mask = (group['生效日期'] >= seg_start) & (group['生效日期'] <= seg_end)
            seg_data = group[seg_mask]
            eff_time = seg_data['保障生效时间'].min()
            exp_time = seg_data['保障失效时间'].max()
            coverage_days = (exp_time - eff_time).days + 1
            results.append({
                '姓名': name,
                '证件号码': id_num,
                '方案名称': plan,
                '保障生效时间': eff_time.strftime('%Y-%m-%d'),
                '保障失效时间': exp_time.strftime('%Y-%m-%d'),
                '保障天数': coverage_days,
                '段数': idx + 1,
            })

    result_df = pd.DataFrame(results)
    result_df = result_df.sort_values(['证件号码', '保障生效时间']).reset_index(drop=True)
    log(f"汇总行数: {len(result_df)}")
    return result_df


def write_excel(result_df, out_path):
    """将汇总结果输出为带格式的Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = '投保数据汇总'

    headers = ['姓名', '证件号码', '方案名称', '保障生效时间', '保障失效时间', '保障天数', '段数']

    header_font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill('solid', fgColor='4472C4')
    header_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    cell_alignment = Alignment(horizontal='center', vertical='center')

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    for row_idx, (_, row) in enumerate(result_df.iterrows()):
        for col_idx, header in enumerate(headers, 1):
            val = row[header]
            if header == '证件号码':
                val = str(val)
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=val)
            cell.font = Font(name='Arial', size=10)
            cell.alignment = cell_alignment
            cell.border = thin_border

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 28
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 8

    ws.auto_filter.ref = f'A1:G{len(result_df) + 1}'

    wb.save(out_path)


def process(src_path, out_path):
    """命令行入口：读取 → 汇总 → 输出"""
    result_df = build_result_df(src_path)
    write_excel(result_df, out_path)
    print(f"输出文件: {out_path}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    src = sys.argv[1]

    if len(sys.argv) >= 3:
        out = sys.argv[2]
    else:
        src_path = Path(src)
        out = str(Path.cwd() / f"{src_path.stem}_汇总.xlsx")

    process(src, out)
