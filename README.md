# OS 店铺数据大盘分析工具 (OS Basic Information Check)

分析 Shopee/Tokopedia 等电商店铺的库存偏差、发货时间异常、商品重量/名称异常的自动化分析工具。

## 功能

- **库存偏差检测 (Gap200/Gap1000)** — 对比店铺后台库存与 ERP 实际可用量，找出偏差超过 200 或 1000 的商品
- **发货时间异常检查** — 检测有足够库存但发货时间过长的商品
- **重量异常检测** — 检查有库存但重量异常的 listing
- **已售罄标记异常** — 检测库存充足但名称仍标 "Habis"（印尼语"已售罄"）的商品
- **SBY001 仓库库存调整** — 分析 SBY001 仓库的库存状况

## 使用方法

```bash
pip install openpyxl

python scripts/analyze.py -i "数据文件.xlsx" -o "结果文件.xlsx"
```

### 参数

| 参数 | 说明 |
|------|------|
| `-i / --input` | 输入 Excel 文件路径（默认：`OS店铺基本数据大盘.xlsx`） |
| `-o / --output` | 输出结果文件路径（默认：`OS店铺数据大盘_分析结果.xlsx`） |

## 输出说明

生成的结果 Excel 包含 7 个子表：

| 子表 | 说明 |
|------|------|
| 增加库存Gap200 | 店铺库存比ERP少200以上，需补货 |
| 减少库存Gap200 | 店铺库存比ERP多200以上，需下架 |
| 增加库存Gap1000 | 店铺库存比ERP少1000以上，严重缺货 |
| 发货时间异常 | 可用量>300 但发货时间>2天的商品 |
| 重量异常_大于10000 | 可用量>300 但重量>10000g的商品 |
| 名称含Habis异常 | 可用量>300 但名称仍标"Habis"的商品 |
| SBY001库存调整 | SBY001可用量>500的商品 |

## 数据源要求

Excel 文件必须包含名为 `Vkiau店铺折扣 库存 预售汇总5.18` 的工作表，且包含以下关键列：Product ID、Variation ID、SKU Ref No、Shipping time、店铺可用量/库存、重量/尺寸等。

## 安装为 Claude Code Skill

将本仓库克隆到 `~/.claude/skills/os-shop-analyzer/` 目录即可作为 Claude Code 的 skill 使用。

## 依赖

- Python 3.8+
- openpyxl
