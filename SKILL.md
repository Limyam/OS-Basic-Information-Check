---
name: OS-Basic-Information-Check
description: OS店铺数据大盘分析工具。分析Shopee/Tokopedia等店铺的库存偏差(Gap200/Gap1000)、发货时间异常、重量异常、名称含"Habis"已售罄标记异常、广告操作判断、MID可售天数。读取"Vkiau店铺折扣 库存 预售汇总{MM.DD}"工作表（每周一三五更新），按列名匹配（兼容新旧列名），输出结构化的分析结果Excel。
triggers:
  - "OS店铺分析"
  - "店铺数据大盘"
  - "库存偏差分析"
  - "Gap200"
  - "Gap1000"
  - "发货时间异常"
  - "Habis异常"
  - "店铺库存核对"
  - "广告操作判断"
  - "MID可售天数"
---

# OS 店铺数据大盘分析 Skill

分析 Shopee/Tokopedia 店铺的库存偏差、发货时间异常、商品重量/名称异常，以及广告操作数据判断。

> **数据更新频率：** 每周一、三、五更新。
> 工作表名称格式：`Vkiau店铺折扣 库存 预售汇总{MM.DD}`（如 `6.01`、`9.18`），随数据日期变化。
> 使用前需确认脚本中的 `SHEET_NAME` 与输入文件一致。

## 数据源要求

Excel 文件中必须包含名为 **`Vkiau店铺折扣 库存 预售汇总{MM.DD}`** 的工作表。

**脚本按列名匹配（不依赖列号）**，因此增删列、调整列顺序都不影响。部分字段支持新旧列名兼容（新名优先，找不到则回退旧名）：

| 字段 | 新列名 | 旧列名（兼容回退） |
|------|--------|-------------------|
| Product ID | `Product ID` | — |
| Product Name | `Product Name(Optional)` | — |
| Variation ID | `Variation ID` | — |
| Variation name | `Variation name(Optional)` | — |
| SKU Ref. No. | `SKU Ref. No.(Optional)` | — |
| 发货时间 | `Shipping time 发货时间` | — |
| Gap200 增 | `判断是否增加库存Gap200` | — |
| Gap200 减 | `判断是否减少库存Gap200` | — |
| Gap1000 增 | `判断是否增加库存Gap1000` | — |
| **IDR001 可用量** | `(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)` | `店铺（可用量-待审订单预占）IDR001` |
| **IDR001 后台库存** | `Store IDR001 Inventory / 店铺IDR001库存` | `Inventory on the platform（IDR001）店铺后台库存` |
| **SBY001 可用量** | `(SBY001 Available Stock - Pending Order Quantity) / (SBY001 可用量-待审订单量)` | `店铺（可用量-待审订单预占）SBY001` |
| **SBY001 后台库存** | `Store SBY001 Inventory / 店铺SBY001库存` | `Inventory on the platform（SBY001）店铺后台库存` |
| 重量 / 长 / 宽 / 高 | `重量` / `长` / `宽` / `高` | — |
| **可售天数** | `Days of Stock Available / 可售天数` | `Available selling days 可售天数` |

> 列名匹配会自动规范化：斜杠 `/` 视为分隔符，换行符、连续空格统一为单个空格。因此 `A / B`、`A\nB`、`A  B` 都能正确匹配。

## 分析功能

### 需求一：库存偏差分析 (Gap检测)
筛选三个 Gap 标记列中为 TRUE 的行，分别输出到3个子表：
- **增加库存Gap200** — 店铺库存比ERP少200以上
- **减少库存Gap200** — 店铺库存比ERP多200以上
- **增加库存Gap1000** — 店铺库存比ERP少1000以上

### 需求二：发货时间异常
筛选条件：`IDR001 可用量` > 300 且 `发货时间` > 2

### 需求三：商品异常检测
- **情况1**：`IDR001 可用量` > 300 且 `重量` > 10000
- **情况2**：`IDR001 可用量` > 300 且 `Variation name` 包含 "habis"（不分大小写）

### 需求四：SBY001库存调整
筛选条件：`SBY001 可用量` > 500

### 需求五：广告操作判断
按 Product ID（PID）分组，统计每个 PID 的 Variation ID（MID）总数，以及可售天数 ≤ 7 的 MID 数量。
输出字段：PID、MID数量、可售天数小于等于7的MID数量

### 需求六：MID可售天数明细
逐条列出每个 PID 下的每个 MID 及其可售天数。
输出字段：PID、MID、MID可售天数

## 输出字段

| 子表 | 库存字段 |
|------|---------|
| 增加/减少Gap200、Gap1000、发货时间异常、重量异常、Habis异常 | IDR001字段 |
| SBY001库存调整 | SBY001字段 |
| 广告操作判断、MID可售天数明细 | (计算型，非筛选) |

## 使用方法

```bash
# 使用当前目录下的默认文件
python scripts/analyze.py

# 指定输入文件
python scripts/analyze.py --input "路径/到/数据文件.xlsx"

# 指定输出文件
python scripts/analyze.py --output "自定义输出文件名.xlsx"

# 完整参数
python scripts/analyze.py --input "数据.xlsx" --output "结果.xlsx"
```

## 输出结果

生成的结果 Excel 包含 **9 个子表**：

| 子表 | 说明 |
|------|------|
| 增加库存Gap200 | 店铺库存比ERP少200以上，需补货 |
| 减少库存Gap200 | 店铺库存比ERP多200以上，需下架 |
| 增加库存Gap1000 | 店铺库存比ERP少1000以上，严重缺货 |
| 发货时间异常 | 可用量>300 但发货时间>2天 |
| 重量异常_大于10000 | 可用量>300 但重量>10000g |
| 名称含Habis异常 | 可用量>300 但名称仍标"Habis" |
| SBY001库存调整 | SBY001可用量>500 |
| 广告操作判断 | PID、MID数量、可售天数≤7的MID数量 |
| MID可售天数明细 | PID、MID、MID可售天数 |

## 文件结构

```
OS-Basic-Information-Check/
├── SKILL.md              # 本文件
├── scripts/
│   └── analyze.py        # 核心分析脚本
└── workflows/
    ├── gap-analysis.md   # 库存偏差分析流程
    ├── shipping-check.md # 发货时间检查流程
    └── product-check.md  # 商品异常检查流程
```

## 版本历史

### v4（当前）
- **列名兼容**：6 个字段支持新列名，同时保留旧列名作为回退，新旧表格都能处理
- 列名规范化增强：斜杠 `/`、换行符、连续空格统一处理
- 列缺失时报错会列出表中实际可用列，便于排查

### v3
- 改用**列名匹配**代替固定列索引，支持删除/隐藏列、调整列顺序

### v2
- 新增"广告操作判断"子表（PID、MID数量、可售天数≤7的MID数量）
- 引入 COMPUTED_SHEETS 机制支持计算型分析
- 新增"MID可售天数明细"子表（v2.1）

### v1
- 初始版本，7 个分析子表（Gap200/1000、发货时间异常、重量异常、Habis异常、SBY001调整）
- 基础行筛选分析引擎

## 依赖

- Python 3.8+
- openpyxl
