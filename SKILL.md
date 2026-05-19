---
name: os-shop-analyzer
description: OS店铺数据大盘分析工具。分析Shopee/Tokopedia等店铺的库存偏差(Gap200/Gap1000)、发货时间异常、重量异常、名称含"Habis"已售罄标记异常。读取"Vkiau店铺折扣 库存 预售汇总5.18"工作表，输出结构化的分析结果Excel。
triggers:
  - "OS店铺分析"
  - "店铺数据大盘"
  - "库存偏差分析"
  - "Gap200"
  - "Gap1000"
  - "发货时间异常"
  - "Habis异常"
  - "店铺库存核对"
---

# OS 店铺数据大盘分析 Skill

分析 Shopee/Tokopedia 店铺的库存偏差、发货时间异常、商品重量/名称异常。

## 数据源要求

Excel 文件中必须包含名为 **`Vkiau店铺折扣 库存 预售汇总5.18`** 的工作表，包含以下列：

| 列 | 字段名 |
|----|--------|
| D | Product ID |
| E | Product Name(Optional) |
| G | Variation ID |
| H | Variation name(Optional) |
| I | SKU Ref. No.(Optional) |
| Z | Shipping time 发货时间 |
| AA | 判断是否增加库存Gap200 |
| AB | 判断是否减少库存Gap200 |
| AC | 判断是否增加库存Gap1000 |
| AD | 店铺（可用量-待审订单预占）IDR001 |
| AE | Inventory on the platform（IDR001）店铺后台库存 |
| AI | 重量 |
| AJ | 长 |
| AK | 宽 |
| AL | 高 |

## 分析功能

### 需求一：库存偏差分析 (Gap检测)
筛选 AA/AB/AC 三列中任意为 TRUE 的行，分别输出到3个子表：
- **增加库存Gap200** — 店铺库存比ERP少200以上
- **减少库存Gap200** — 店铺库存比ERP多200以上
- **增加库存Gap1000** — 店铺库存比ERP少1000以上

### 需求二：发货时间异常
筛选条件：`店铺（可用量-待审订单预占）IDR001` > 300 且 `Shipping time 发货时间` > 2

### 需求三：商品异常检测
- **情况1**：`店铺（可用量-待审订单预占）IDR001` > 300 且 `重量` > 10000
- **情况2**：`店铺（可用量-待审订单预占）IDR001` > 300 且 `Variation name` 包含 "habis"（不分大小写）

### 需求四：SBY001库存调整
筛选条件：`店铺（可用量-待审订单预占）SBY001` > 500

## 输出字段

各子表输出字段均为12个，但根据分析目标不同，库存列会对应 IDR001 或 SBY001：

| 子表 | 库存字段 |
|------|---------|
| 增加/减少Gap200、Gap1000、发货时间异常、重量异常、Habis异常 | IDR001字段 |
| **SBY001库存调整** | **SBY001字段** |

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

## 文件结构

```
OS店铺分析skill/
├── SKILL.md              # 本文件
├── scripts/
│   └── analyze.py        # 核心分析脚本
└── workflows/
    ├── gap-analysis.md   # 库存偏差分析流程
    ├── shipping-check.md # 发货时间检查流程
    └── product-check.md  # 商品异常检查流程
```

## 依赖

- Python 3.8+
- openpyxl
