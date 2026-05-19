# OS Shop Data Dashboard Analyzer

An automated analysis tool for Shopee/Tokopedia online stores. Detects inventory discrepancies, shipping time anomalies, and product weight/name issues by comparing store platform data with ERP system data.

## Features

- **Inventory Gap Detection (Gap200/Gap1000)** — Compares store platform inventory against ERP actual availability, identifying items with discrepancies exceeding 200 or 1000 units
- **Shipping Time Anomaly Check** — Detects products with sufficient stock but abnormally long shipping times
- **Weight Anomaly Detection** — Flags listings that have stock but unreasonable weight values (>10,000g)
- **"Habis" (Sold Out) Label Anomaly** — Identifies products with adequate stock whose variation names still contain "Habis" (Indonesian for "sold out"), indicating outdated listings
- **SBY001 Warehouse Stock Review** — Analyzes stock levels in the SBY001 warehouse

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

## Output

The generated Excel workbook contains 7 sheets:

| Sheet | Description |
|-------|-------------|
| 增加库存Gap200 | Store inventory is 200+ less than ERP → needs restocking |
| 减少库存Gap200 | Store inventory is 200+ more than ERP → needs removal |
| 增加库存Gap1000 | Store inventory is 1000+ less than ERP → severely understocked |
| 发货时间异常 | Available qty > 300 but shipping time > 2 days |
| 重量异常_大于10000 | Available qty > 300 but weight > 10,000g |
| 名称含Habis异常 | Available qty > 300 but name still labeled "Habis" |
| SBY001库存调整 | SBY001 available qty > 500 |

## Data Source Requirements

The Excel file must contain a worksheet named `Vkiau店铺折扣 库存 预售汇总5.18` with the following columns: Product ID, Variation ID, SKU Ref No, Shipping time, store availability/inventory, weight/dimensions, etc.

## Install as Claude Code Skill

Clone this repository to `~/.claude/skills/os-shop-analyzer/` to use it as a Claude Code skill.

## Dependencies

- Python 3.8+
- openpyxl
