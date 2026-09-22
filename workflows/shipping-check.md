# 发货时间检查流程

## 目的
找出有足够库存（>300）但发货时间不是标准2天的商品，可能是设置异常。

## 数据来源

读取源文件中名为 `Vkiau店铺折扣 库存 预售汇总{MM.DD}` 的工作表。
工作表名中的日期随导出日期变化，数据每周一、三、五更新。

## 判断逻辑

所有列**按列名匹配**（不是固定列索引，删除/隐藏列不影响）。

- 条件1：`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)` > 300
- 条件2：`Shipping time 发货时间` > 2

## 输出

结果写入 **发货时间异常** 子表，输出的列：

- 基础列：`Product ID`、`Product Name(Optional)`、`Variation ID`、`Variation name(Optional)`、`SKU Ref. No.(Optional)`、`Shipping time 发货时间`
- 库存列：`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)`、`Store IDR001 Inventory / 店铺IDR001库存`
- 尾部列：`重量`、`长`、`宽`、`高`

## 列名兼容说明

`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)` 的旧列名为 `店铺（可用量-待审订单预占）IDR001`，脚本会自动回退匹配。
