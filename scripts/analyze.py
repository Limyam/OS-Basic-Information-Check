#!/usr/bin/env python3
"""
OS店铺数据大盘分析工具 (v4 - 列名匹配 + 新旧列名兼容)
分析Shopee/Tokopedia店铺的库存偏差、发货时间异常、商品重量/名称异常。
通过表头名称匹配列位置，兼容旧列名与新列名。
"""
import sys
import io
import re
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict

# Force UTF-8 output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ── 配置 ──────────────────────────────────────────────
SHEET_NAME = 'Vkiau店铺折扣 库存 预售汇总9.18'

# 列名规格: key → [候选列名...]（新列名在前，旧列名作为兼容回退）
COL_SPECS = {
    'product_id':      ["Product ID"],
    'product_name':    ["Product Name(Optional)"],
    'variation_id':    ["Variation ID"],
    'variation_name':  ["Variation name(Optional)"],
    'sku_ref':         ["SKU Ref. No.(Optional)"],
    'shipping_time':   ["Shipping time 发货时间"],
    'gap200_inc':      ["判断是否增加库存Gap200"],
    'gap200_dec':      ["判断是否减少库存Gap200"],
    'gap1000_inc':     ["判断是否增加库存Gap1000"],
    'idr001_avail':    [
        "(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)",
        "店铺（可用量-待审订单预占）IDR001",
    ],
    'idr001_platform': [
        "Store IDR001 Inventory / 店铺IDR001库存",
        "Inventory on the platform（IDR001）店铺后台库存",
    ],
    'sby001_avail':    [
        "(SBY001 Available Stock - Pending Order Quantity) / (SBY001 可用量-待审订单量)",
        "店铺（可用量-待审订单预占）SBY001",
    ],
    'sby001_platform': [
        "Store SBY001 Inventory / 店铺SBY001库存",
        "Inventory on the platform（SBY001）店铺后台库存",
    ],
    'weight':          ["重量"],
    'length':          ["长"],
    'width':           ["宽"],
    'height':          ["高"],
    'sell_days':       [
        "Days of Stock Available / 可售天数",
        "Available selling days 可售天数",
    ],
}

# 输出表头（使用新列名）
BASE_COL_KEYS = ['product_id', 'product_name', 'variation_id', 'variation_name', 'sku_ref', 'shipping_time']
BASE_HEADERS = [
    "Product ID", "Product Name(Optional)", "Variation ID",
    "Variation name(Optional)", "SKU Ref. No.(Optional)",
    "Shipping time 发货时间",
]

TAIL_COL_KEYS = ['weight', 'length', 'width', 'height']
TAIL_HEADERS = ["重量", "长", "宽", "高"]

IDR001_COL_KEYS = ['idr001_avail', 'idr001_platform']
IDR001_HEADERS = [
    "(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)",
    "Store IDR001 Inventory / 店铺IDR001库存",
]

SBY001_COL_KEYS = ['sby001_avail', 'sby001_platform']
SBY001_HEADERS = [
    "(SBY001 Available Stock - Pending Order Quantity) / (SBY001 可用量-待审订单量)",
    "Store SBY001 Inventory / 店铺SBY001库存",
]


# ── 列名映射 ──────────────────────────────────────────
def normalize_name(s):
    """规范化列名: 斜杠视为分隔符, 空白字符统一为单个空格"""
    s = str(s).replace('/', ' ')
    return re.sub(r'\s+', ' ', s).strip()


def build_column_map(ws):
    """读取第一行表头，建立 规范化列名→索引(0-based) 的映射"""
    col_map = {}
    for i, cell in enumerate(ws[1]):
        if cell.value:
            col_map[normalize_name(cell.value)] = i
    return col_map


def resolve_columns(col_map):
    """按候选列名解析出 key→索引 的映射，缺失的列会报错提示"""
    resolved = {}
    missing = []
    for key, candidates in COL_SPECS.items():
        for name in candidates:
            norm = normalize_name(name)
            if norm in col_map:
                resolved[key] = col_map[norm]
                break
        else:
            missing.append(f'{key} (候选列名: {candidates})')

    if missing:
        available = '\n  - '.join(sorted(col_map.keys()))
        raise ValueError(
            '找不到以下列:\n  - ' + '\n  - '.join(missing) +
            f'\n\n表中实际可用列:\n  - {available}'
        )
    return resolved


# ── 工具函数 ──────────────────────────────────────────
def safe_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def is_true(v):
    return str(v).upper() == 'TRUE'


def extract_rows(ws, filter_func, cols, out_col_keys):
    """遍历工作表，筛选符合条件的行，返回输出列的值列表。"""
    out_indices = [cols[k] for k in out_col_keys]
    results = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if filter_func(row, cols):
            results.append(tuple(row[i] for i in out_indices))
    return results


# ── 过滤条件 ──────────────────────────────────────────
def gap_increase_200(r, cols):
    """店铺库存比ERP少200以上"""
    return is_true(r[cols['gap200_inc']])


def gap_decrease_200(r, cols):
    """店铺库存比ERP多200以上"""
    return is_true(r[cols['gap200_dec']])


def gap_increase_1000(r, cols):
    """店铺库存比ERP少1000以上"""
    return is_true(r[cols['gap1000_inc']])


def shipping_anomaly(r, cols):
    """可用量>300 但发货时间>2"""
    ad = safe_float(r[cols['idr001_avail']])
    st = safe_float(r[cols['shipping_time']])
    if ad is None or st is None:
        return False
    return ad > 300 and st > 2


def weight_anomaly(r, cols):
    """可用量>300 且重量>10000"""
    ad = safe_float(r[cols['idr001_avail']])
    wt = safe_float(r[cols['weight']])
    if ad is None or wt is None:
        return False
    return ad > 300 and wt > 10000


def habis_anomaly(r, cols):
    """可用量>300 且Variation name含habis"""
    ad = safe_float(r[cols['idr001_avail']])
    if ad is None:
        return False
    vname = str(r[cols['variation_name']]) if r[cols['variation_name']] is not None else ''
    return ad > 300 and 'habis' in vname.lower()


def sby001_adjust(r, cols):
    """SBY001可用量>500"""
    af = safe_float(r[cols['sby001_avail']])
    return af is not None and af > 500


# 子表配置: (子表名, 过滤函数, 输出列key列表, 表头列表)
SHEETS_CONFIG = [
    ("增加库存Gap200", gap_increase_200, BASE_COL_KEYS + IDR001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("减少库存Gap200", gap_decrease_200, BASE_COL_KEYS + IDR001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("增加库存Gap1000", gap_increase_1000, BASE_COL_KEYS + IDR001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("发货时间异常", shipping_anomaly, BASE_COL_KEYS + IDR001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("重量异常_大于10000", weight_anomaly, BASE_COL_KEYS + IDR001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("名称含Habis异常", habis_anomaly, BASE_COL_KEYS + IDR001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + IDR001_HEADERS + TAIL_HEADERS),
    ("SBY001库存调整", sby001_adjust, BASE_COL_KEYS + SBY001_COL_KEYS + TAIL_COL_KEYS, BASE_HEADERS + SBY001_HEADERS + TAIL_HEADERS),
]


# ── 广告操作判断（计算型）────────────────────────────────
def compute_ad_judgment(ws, cols):
    """按PID分组，统计MID数量和可售天数<=7的MID数量"""
    data = defaultdict(lambda: {'mids': set(), 'lte7': set()})
    idx_pid = cols['product_id']
    idx_mid = cols['variation_id']
    idx_sell_days = cols['sell_days']

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        pid = row[idx_pid]
        mid = row[idx_mid]
        sell_days = safe_float(row[idx_sell_days])

        if pid is None:
            continue
        pid = str(pid)
        if mid is not None:
            data[pid]['mids'].add(mid)
            if sell_days is not None and sell_days <= 7:
                data[pid]['lte7'].add(mid)

    results = [(pid, len(d['mids']), len(d['lte7']))
               for pid, d in sorted(data.items())]
    return results


def compute_mid_sell_days(ws, cols):
    """列出每个PID下每个MID及其可售天数"""
    idx_pid = cols['product_id']
    idx_mid = cols['variation_id']
    idx_sell_days = cols['sell_days']

    results = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        pid = row[idx_pid]
        mid = row[idx_mid]
        sell_days = safe_float(row[idx_sell_days])

        if pid is not None and mid is not None:
            results.append((str(pid), str(mid), sell_days if sell_days is not None else ''))

    results.sort(key=lambda x: (x[0], x[1]))
    return results


COMPUTED_SHEETS = [
    ("广告操作判断", compute_ad_judgment,
     ["PID", "MID数量", "可售天数小于等于7的MID数量"]),
    ("MID可售天数明细", compute_mid_sell_days,
     ["PID", "MID", "MID可售天数"]),
]


# ── 输出 ──────────────────────────────────────────────
def write_sheet(ws_out, data_rows, headers):
    """写入一个子表的数据。"""
    hfont = Font(bold=True, size=11, color='FFFFFF')
    hfill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')

    for ci, h in enumerate(headers, 1):
        cell = ws_out.cell(row=1, column=ci, value=h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    for ri, row_data in enumerate(data_rows, 2):
        for ci, val in enumerate(row_data, 1):
            ws_out.cell(row=ri, column=ci, value=val)

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

    if SHEET_NAME not in wb.sheetnames:
        available = '\n  - '.join(wb.sheetnames)
        raise ValueError(
            f'找不到工作表 "{SHEET_NAME}"。\n'
            f'可用的工作表:\n  - {available}'
        )

    ws = wb[SHEET_NAME]
    col_map = build_column_map(ws)
    cols = resolve_columns(col_map)
    print(f'  📋 匹配到 {len(cols)}/{len(COL_SPECS)} 个所需列')
    used_old = [
        k for k, specs in COL_SPECS.items()
        if len(specs) > 1 and normalize_name(specs[0]) not in col_map
    ]
    if used_old:
        print(f'  ℹ️ 这些字段使用旧列名回退: {", ".join(used_old)}')

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    total = 0
    for sheet_name, filter_func, out_col_keys, headers in SHEETS_CONFIG:
        rows = extract_rows(ws, filter_func, cols, out_col_keys)
        total += len(rows)
        ws_out = out_wb.create_sheet(title=sheet_name)
        write_sheet(ws_out, rows, headers)
        print(f'  [{sheet_name}] {len(rows)} 行')

    for sheet_name, compute_func, headers in COMPUTED_SHEETS:
        rows = compute_func(ws, cols)
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
    parser.add_argument('--input', '-i', default='OS店铺基本数据大盘.xlsx',
                        help='输入Excel文件路径 (默认: OS店铺基本数据大盘.xlsx)')
    parser.add_argument('--output', '-o', default='OS店铺数据大盘_分析结果.xlsx',
                        help='输出Excel文件路径 (默认: OS店铺数据大盘_分析结果.xlsx)')
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
