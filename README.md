# OS Shop Data Dashboard Analyzer

An automated analysis tool for Shopee/Tokopedia online stores. Detects inventory discrepancies, shipping time anomalies, product weight/name issues, and generates ad operation recommendations.

> **Data update frequency:** Every Monday, Wednesday, and Friday.
> The input Excel file and its worksheet name follow the pattern `Vkiau店铺折扣 库存 预售汇总{MM.DD}` (e.g., `6.01`, `6.05`), depending on the data date.

## Features

- **Inventory Gap Detection (Gap200/Gap1000)** — Compares store platform inventory against ERP actual availability, identifying items with discrepancies exceeding 200 or 1000 units
- **Shipping Time Anomaly Check** — Detects products with sufficient stock but abnormally long shipping times
- **Weight Anomaly Detection** — Flags listings that have stock but unreasonable weight values (>10,000g)
- **"Habis" (Sold Out) Label Anomaly** — Identifies products with adequate stock whose variation names still contain "Habis" (Indonesian for "sold out"), indicating outdated listings
- **SBY001 Warehouse Stock Review** — Analyzes stock levels in the SBY001 warehouse
- **Ad Operation Judgment** — Groups by Product ID (PID), counts total variations (MIDs) and variations with available selling days ≤ 7, helping identify products needing advertising adjustments

## Usage

```bash
pip install openpyxl

python scripts/analyze.py -i "input_data.xlsx" -o "output_results.xlsx"
```

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-i / --input` | Input Excel file path (default: `OS店铺基本数据大盘.xlsx`) |
| `-o / --output` | Output results file path (default: `OS店铺数据大盘_分析结果.xlsx`) |

> **Note:** The script's `SHEET_NAME` variable may need to be updated to match the worksheet name in your input file (e.g., `Vkiau店铺折扣 库存 预售汇总6.01`, `Vkiau店铺折扣 库存 预售汇总6.05`).

## Output

The generated Excel workbook contains 8 sheets:

| Sheet | Description |
|-------|-------------|
| 增加库存Gap200 | Store inventory is 200+ less than ERP → needs restocking |
| 减少库存Gap200 | Store inventory is 200+ more than ERP → needs removal |
| 增加库存Gap1000 | Store inventory is 1000+ less than ERP → severely understocked |
| 发货时间异常 | Available qty > 300 but shipping time > 2 days |
| 重量异常_大于10000 | Available qty > 300 but weight > 10,000g |
| 名称含Habis异常 | Available qty > 300 but name still labeled "Habis" |
| SBY001库存调整 | SBY001 available qty > 500 |
| **广告操作判断** (v2) | PID, MID count, and count of MIDs with available selling days ≤ 7 |

## Data Source Requirements

The Excel file must contain a worksheet named `Vkiau店铺折扣 库存 预售汇总{MM.DD}` (versioned by date) with the following columns: Product ID, Variation ID, SKU Ref No, Shipping time, store availability/inventory, weight/dimensions, Available selling days, etc.

## Version History

### v2 (Current)
- **New sheet:** "广告操作判断" — Groups data by PID, counts total MIDs and MIDs with available selling days ≤ 7
- **New architecture:** Introduced `COMPUTED_SHEETS` mechanism to support computed/aggregated analysis sheets beyond simple filtering
- **Updated default worksheet:** `Vkiau店铺折扣 库存 预售汇总6.01`

### v1
- Initial release with 7 analysis sheets (Gap200/1000, shipping anomalies, weight anomalies, Habis anomalies, SBY001 adjustments)
- Basic row-filtering analysis engine

## Install as Claude Code Skill

Clone this repository to `~/.claude/skills/os-shop-analyzer/` to use it as a Claude Code skill.

```bash
git clone https://github.com/Limyam/OS-Basic-Information-Check.git ~/.claude/skills/os-shop-analyzer/
```

## Dependencies

- Python 3.8+
- openpyxl
