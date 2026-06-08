# OS 店铺数据大盘分析工具

分析 Shopee/Tokopedia 等电商店铺的库存偏差、发货时间异常、商品重量/名称异常的自动化分析工具，同时支持广告操作数据判断。

> **数据更新频率：** 每周一、三、五更新。
> 输入的 Excel 文件及内部工作表名称为 `Vkiau店铺折扣 库存 预售汇总{MM.DD}` 格式（如 `6.01`、`6.05`），根据数据日期动态变化。

## 功能

- **库存偏差检测 (Gap200/Gap1000)** — 对比店铺后台库存与 ERP 实际可用量，找出偏差超过 200 或 1000 的商品
- **发货时间异常检查** — 检测有足够库存但发货时间过长的商品
- **重量异常检测** — 检查有库存但重量异常的 listing
- **已售罄标记异常** — 检测库存充足但名称仍标 "Habis"（印尼语"已售罄"）的商品
- **SBY001 仓库库存调整** — 分析 SBY001 仓库的库存状况
- **广告操作判断** — 按 PID 分组统计 MID 数量和可售天数 ≤ 7 的 MID 数量，辅助广告投放决策

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

> **注意：** 脚本中的 `SHEET_NAME` 需要根据输入文件的实际工作表名称进行更新（例如 `Vkiau店铺折扣 库存 预售汇总6.01`、`Vkiau店铺折扣 库存 预售汇总6.05`）。

## 输出说明

生成的结果 Excel 包含 8 个子表：

| 子表 | 说明 |
|------|------|
| 增加库存Gap200 | 店铺库存比ERP少200以上，需补货 |
| 减少库存Gap200 | 店铺库存比ERP多200以上，需下架 |
| 增加库存Gap1000 | 店铺库存比ERP少1000以上，严重缺货 |
| 发货时间异常 | 可用量>300 但发货时间>2天的商品 |
| 重量异常_大于10000 | 可用量>300 但重量>10000g的商品 |
| 名称含Habis异常 | 可用量>300 但名称仍标"Habis"的商品 |
| SBY001库存调整 | SBY001可用量>500的商品 |
| **广告操作判断** (v2新增) | PID、MID数量、可售天数≤7的MID数量 |

## 数据源要求

Excel 文件必须包含名为 `Vkiau店铺折扣 库存 预售汇总{MM.DD}`（按日期版本命名）的工作表，且包含以下关键列：Product ID、Variation ID、SKU Ref No、Shipping time、店铺可用量/库存、重量/尺寸、Available selling days 可售天数等。

## 版本历史

### v2（当前版本）
- **新增子表：** "广告操作判断" — 按 PID 分组统计 MID 总数和可售天数 ≤ 7 的 MID 数量
- **新架构：** 引入 `COMPUTED_SHEETS` 机制，支持聚合计算类分析子表（不局限于行筛选）
- **更新默认工作表：** `Vkiau店铺折扣 库存 预售汇总6.01`

### v1
- 初始版本，7 个分析子表（Gap200/1000、发货时间异常、重量异常、Habis异常、SBY001调整）
- 基础行筛选分析引擎

## 安装为 Claude Code Skill

将本仓库克隆到 `~/.claude/skills/os-shop-analyzer/` 目录即可作为 Claude Code 的 skill 使用。

```bash
git clone https://github.com/Limyam/OS-Basic-Information-Check.git ~/.claude/skills/os-shop-analyzer/
```

## 依赖

- Python 3.8+
- openpyxl
