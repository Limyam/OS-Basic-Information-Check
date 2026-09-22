# 商品异常检测流程

## 目的
检查有库存的商品是否存在重量或名称标注异常。

## 数据来源

读取源文件中名为 `Vkiau店铺折扣 库存 预售汇总{MM.DD}` 的工作表。
工作表名中的日期随导出日期变化，数据每周一、三、五更新。

## 判断逻辑

所有列**按列名匹配**（不是固定列索引，删除/隐藏列不影响）。

### 情况1：重量异常

- 条件1：`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)` > 300
- 条件2：`重量` > 10000

### 情况2：名称含"Habis"异常

"Habis" 是印尼语"已售罄"的意思，如果商品有库存但名称还标记"Habis"，可能是Listing未及时更新。

- 条件1：`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)` > 300
- 条件2：`Variation name(Optional)` 包含 "habis"（不区分大小写）

## 输出

结果分别写入 **重量异常_大于10000** 和 **名称含Habis异常** 子表。

每个子表输出的列：

- 基础列：`Product ID`、`Product Name(Optional)`、`Variation ID`、`Variation name(Optional)`、`SKU Ref. No.(Optional)`、`Shipping time 发货时间`
- 库存列：`(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)`、`Store IDR001 Inventory / 店铺IDR001库存`
- 尾部列：`重量`、`长`、`宽`、`高`

## 列名兼容说明

6个字段在最新文件中已改名，脚本同时兼容旧列名作为回退：

| 新列名 | 旧列名（兼容回退） |
|--------|-------------------|
| `(IDR001 Available Stock - Pending Order Quantity) / (IDR001 可用量-待审订单量)` | `店铺（可用量-待审订单预占）IDR001` |
| `Store IDR001 Inventory / 店铺IDR001库存` | `Inventory on the platform（IDR001）店铺后台库存` |
| `Days of Stock Available / 可售天数` | `Available selling days 可售天数` |
