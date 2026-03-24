# Field Mapping / 字段映射

This document maps the 22 Excel headers used by `chat-order-excel-mcp` to the JSON fields accepted by the MCP tools.

本文档说明 `chat-order-excel-mcp` 使用的 22 列 Excel 表头与 MCP 工具接收的 JSON 字段之间的对应关系。

## Core Rule / 核心规则

- Excel writing is header-based, not column-index-based.
- JSON keys should match the field names below.
- `extra_fields` is only for non-standard columns outside the 22 fixed headers.

- Excel 写入按表头名映射，不按列位置映射。
- JSON key 应与下表中的字段名保持一致。
- `extra_fields` 只用于 22 个固定表头之外的扩展列。

## Header Mapping / 表头映射

| Excel Header | JSON Field | Meaning |
| --- | --- | --- |
| 备注 | `备注` | Payment note, remark, special note / 收款说明、备注、特殊要求 |
| 发货厂家 | `发货厂家` | Shipping manufacturer / 发货厂家 |
| 产品供应商 | `产品供应商` | Product supplier / 产品供应商 |
| 日期 | `日期` | Order date in `YYYY-MM-DD` / 订单日期 |
| 单号 | `单号` | Order number / 订单编号 |
| 销售员 | `销售员` | Salesperson, usually from sender / 业务员，优先取消息发送者 |
| 客户 | `客户` | Customer name / 客户名称 |
| 货品名称 | `货品名称` | Product name / 货品名称 |
| 数量 | `数量` | Quantity / 数量 |
| 数量单位 | `数量单位` | Quantity unit / 数量单位 |
| 销售单价 | `销售单价` | Sales unit price / 销售单价 |
| 销售金额 | `销售金额` | Sales amount / 销售金额 |
| 成本单价 | `成本单价` | Cost unit price / 成本单价 |
| 成本金额 | `成本金额` | Cost amount / 成本金额 |
| 运费 | `运费` | Shipping fee / 运费 |
| 利润 | `利润` | Profit / 利润 |
| 总货款 | `总货款` | Total payment / 总货款 |
| 已收 | `已收` | Received payment / 已收金额 |
| 未收 | `未收` | Unpaid amount / 未收金额 |
| 收货联系人 | `收货联系人` | Receiver contact name / 收货联系人 |
| 收货人电话 | `收货人电话` | Receiver phone number / 收货人电话 |
| 收货地址 | `收货地址` | Receiver address / 收货地址 |

## Internal Matching Fields / 内部匹配字段

These fields may appear in MCP tool payloads but are not written to the 22 Excel headers directly.

以下字段可能会出现在 MCP tool 的请求体中，但不会直接写入这 22 个表头。

| Field | Meaning |
| --- | --- |
| `匹配客户别名` | Customer alias extracted from order number suffix such as `260313-05（B&D）` |
| `extra_fields` | Extra Excel columns not included in the fixed 22 headers |

## Example JSON / JSON 示例

```json
{
  "单号": "26.3.13-7",
  "销售员": "业务员A",
  "客户": "示例客户",
  "货品名称": "示例产品",
  "数量": "2000",
  "数量单位": "支",
  "销售金额": "105",
  "总货款": "105",
  "已收": "105",
  "收货联系人": "示例联系人",
  "收货人电话": "13800000000",
  "收货地址": "示例省示例市示例区示例路88号",
  "备注": "示例收款"
}
```

## Notes / 说明

- `OC_OD_TABLE_NAME` must be the Excel table object name, not the worksheet name.
- Empty values do not overwrite existing values during updates.
- `process_excel_order(dry_run=true)` can be used to preview matching and row content without writing.

- `OC_OD_TABLE_NAME` 必须是 Excel 表格对象名，不是工作表名称。
- 更新时空值不会覆盖旧值。
- 可以用 `process_excel_order(dry_run=true)` 先预览匹配结果和写入行内容，而不真正写入 Excel。
