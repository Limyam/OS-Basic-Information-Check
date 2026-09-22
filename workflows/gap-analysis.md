# 库存偏差分析流程 (Gap检测)

## 目的
判断店铺后台库存是否与ERP实际可用量相符，找出偏差超过阈值的商品。

## 数据来源

读取源文件中名为 `Vkiau店铺折扣 库存 预售汇总{MM.DD}` 的工作表。
工作表名中的日期随导出日期变化（如 `9.21`、`6.05`），数据每周一、三、五更新。
脚本中的 `SHEET_NAME` 常量需要跟随最新文件更新。

## 判断逻辑

以下三列**按列名匹配**（不是固定列索引，删除/隐藏列不影响）：

| 列名 | 含义 |
|------|------|
| `判断是否增加库存Gap200` | 店铺库存比ERP少200以上 → 需补货 |
| `判断是否减少库存Gap200` | 店铺库存比ERP多200以上 → 需下架 |
| `判断是否增加库存Gap1000` | 店铺库存比ERP少1000以上 → 严重缺货 |

判断方式：该单元格的值（不区分大小写）等于 `TRUE`。

## 输出

分别输出到3个子表：

1. **增加库存Gap200** — `判断是否增加库存Gap200` = TRUE 的所有行
2. **减少库存Gap200** — `判断是否减少库存Gap200` = TRUE 的所有行
3. **增加库存Gap1000** — `判断是否增加库存Gap1000` = TRUE 的所有行

每个子表输出的列：

- 基础列：`Product ID`、`Product Name(Optional)`、`Variation ID`、`Variation name(Optional)`、`SKU Ref. No.(Optional)`、`Shipping time 发货时间`
- 库存列：`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)`、`Store IDR001 Inventory / 店铺IDR001库存`
- 尾部列：`重量`、`长`、`宽`、`高`

## 使用

```bash
python scripts/analyze.py --input "文件.xlsx" --output "结果.xlsx"
```
