#!/usr/bin/env python3
"""
OS店铺数据大盘分析工具
分析Shopee/Tokopedia店铺的库存偏差、发货时间异常、商品重量/名称异常。
"""
import sys
import io
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Force UTF-8 output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ── 配置 ──────────────────────────────────────────────
SHEET_NAME = 'Vkiau店铺折扣 库存 预售汇总5.18'

# 通用基础列 (所有子表都包含): D, E, G, H, I, Z + 重量/长/宽/高
BASE_COLS = [3, 4, 6, 7, 8, 25]
BASE_HEADERS = [
    "Product ID",
    "Product Name(Optional)",
    "Variation ID",
    "Variation name(Optional)",
    "SKU Ref. No.(Optional)",
    "Shipping time 发货时间",
]
TAIL_COLS = [34, 35, 36, 37]
TAIL_HEADERS = ["重量", "长", "宽", "高"]

# IDR001 列组
IDR001_COLS = [29, 30]
IDR001_HEADERS = [
    "店铺（可用量-待审订单预占）IDR001",
    "Inventory on the platform（IDR001）店铺后台库存",
]

# SBY001 列组: AF(31), AG(32)
SBY001_COLS = [31, 32]
SBY001_HEADERS = [
    "店铺（可用量-待审订单预占）SBY001",
    "Inventory on the platform（SBY001）店铺后台库存",
]


# ── 工具函数 ──────────────────────────────────────────
def safe_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def is_true(v):
    return str(v).upper() == 'TRUE'


def extract_rows(ws, filter_func, out_cols):
    """遍历工作表，筛选符合条件的行，返回输出列的值列表。"""
    results = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if filter_func(row):
            results.append(tuple(row[i] for i in out_cols))
    return results


# ── 过滤条件 ──────────────────────────────────────────
# 需求一
def gap_increase_200(r):
    """店铺库存比ERP少200以上"""
    return is_true(r[26])  # AA


def gap_decrease_200(r):
    """店铺库存比ERP多200以上"""
    return is_true(r[27])  # AB


def gap_increase_1000(r):
    """店铺库存比ERP少1000以上"""
    return is_true(r[28])  # AC


# 需求二
def shipping_anomaly(r):
    """可用量>300 但发货时间>2"""
    ad = safe_float(r[29])  # AD 店铺（可用量-待审订单预占）IDR001
    st = safe_float(r[25])  # Z 发货时间
    if ad is None or st is None:
        return False
    return ad > 300 and st > 2


# 需求三 情况1
def weight_anomaly(r):
    """可用量>300 且重量>10000"""
    ad = safe_float(r[29])  # AD
    wt = safe_float(r[34])  # AI
    if ad is None or wt is None:
        return False
    return ad > 300 and wt > 10000


# 需求三 情况2
def habis_anomaly(r):
    """可用量>300 且Variation name含habis"""
    ad = safe_float(r[29])
    if ad is None:
        return False
    vname = str(r[7]) if r[7] is not None else ''
    return ad > 300 and 'habis' in vname.lower()


# 需求四
def sby001_adjust(r):
    """SBY001可用量>500"""
    af = safe_float(r[31])  # AF
    return af is not None and af > 500


# 子表配置: (子表名, 过滤函数, 输出列, 表头)
SHEETS_CONFIG = [
    ("增加库存Gap200", gap_increase_200, BASE_COLS + IDR001_COLS + TAIL_COLS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("减少库存Gap200", gap_decrease_200, BASE_COLS + IDR001_COLS + TAIL_COLS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("增加库存Gap1000", gap_increase_1000, BASE_COLS + IDR001_COLS + TAIL_COLS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("发货时间异常", shipping_anomaly, BASE_COLS + IDR001_COLS + TAIL_COLS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("重量异常_大于10000", weight_anomaly, BASE_COLS + IDR001_COLS + TAIL_COLS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("名称含Habis异常", habis_anomaly, BASE_COLS + IDR001_COLS + TAIL_COLS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("SBY001库存调整", sby001_adjust, BASE_COLS + SBY001_COLS + TAIL_COLS, BASE_HEADERS + SBY001_HEADERS + TAIL_HEADERS),
]


# ── 输出 ──────────────────────────────────────────────
def write_sheet(ws_out, data_rows, headers):
    """写入一个子表的数据。"""
    hfont = Font(bold=True, size=11, color='FFFFFF')
    hfill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')

    # 表头
    for ci, h in enumerate(headers, 1):
        cell = ws_out.cell(row=1, column=ci, value=h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    # 数据
    for ri, row_data in enumerate(data_rows, 2):
        for ci, val in enumerate(row_data, 1):
            ws_out.cell(row=ri, column=ci, value=val)

    # 列宽自适应
    for ci in range(1, len(headers) + 1):
        ml = max(len(str(headers[ci - 1])), 12)
        for ri in range(2, min(len(data_rows) + 2, 50)):
            cv = ws_out.cell(row=ri, column=ci).value
            if cv:
                ml = max(ml, min(len(str(cv)), 40))
        ws_out.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = ml + 3


def build_workbook(input_path):
    """读取输入文件，执行所有分析，返回输出工作簿。"""
    wb = openpyxl.load_workbook(input_path, data_only=True)

    # 验证工作表存在
    if SHEET_NAME not in wb.sheetnames:
        available = '\n  - '.join(wb.sheetnames)
        raise ValueError(
            f'找不到工作表 "{SHEET_NAME}"。\n'
            f'可用的工作表:\n  - {available}'
        )

    ws = wb[SHEET_NAME]
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)  # 删除默认Sheet

    total = 0
    for sheet_name, filter_func, out_cols, headers in SHEETS_CONFIG:
        rows = extract_rows(ws, filter_func, out_cols)
        total += len(rows)
        ws_out = out_wb.create_sheet(title=sheet_name)
        write_sheet(ws_out, rows, headers)
        print(f'  [{sheet_name}] {len(rows)} 行')

    return out_wb, total


# ── 主入口 ────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='OS店铺数据大盘分析工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            '示例:\n'
            '  %(prog)s                                          # 默认文件\n'
            '  %(prog)s --input "OS店铺基本数据大盘.xlsx"          # 指定输入\n'
            '  %(prog)s -i "数据.xlsx" -o "结果.xlsx"             # 指定输入输出\n'
            '  %(prog)s --output "结果.xlsx"                      # 仅指定输出\n'
        ),
    )
    parser.add_argument(
        '--input', '-i',
        default='OS店铺基本数据大盘.xlsx',
        help='输入Excel文件路径 (默认: OS店铺基本数据大盘.xlsx)',
    )
    parser.add_argument(
        '--output', '-o',
        default='OS店铺数据大盘_分析结果.xlsx',
        help='输出Excel文件路径 (默认: OS店铺数据大盘_分析结果.xlsx)',
    )
    args = parser.parse_args()

    print(f'🔍 读取: {args.input}')
    print(f'📊 分析工作表: {SHEET_NAME}')

    try:
        out_wb, total = build_workbook(args.input)
    except Exception as e:
        print(f'❌ 错误: {e}', file=sys.stderr)
        sys.exit(1)

    out_wb.save(args.output)
    print(f'\n✅ 完成！共 {total} 行结果')
    print(f'📁 输出: {args.output}')
    print(f'📑 子表 ({len(out_wb.sheetnames)}个): {", ".join(out_wb.sheetnames)}')


if __name__ == '__main__':
    main()
