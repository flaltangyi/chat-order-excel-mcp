# cy-excel-mcp

MCP server for parsing WeChat order messages, merging follow-up updates, and writing structured orders into a OneDrive Excel workbook.

[中文说明](README.zh-CN.md)

Field mapping: [docs/FIELD_MAPPING.md](docs/FIELD_MAPPING.md)
OpenClaw integration: [docs/OPENCLAW_INTEGRATION.md](docs/OPENCLAW_INTEGRATION.md)
Project roadmap: [TODO.md](TODO.md)

## Overview

`cy-excel-mcp` is built for order-entry workflows where salespeople send order details through chat, sometimes as plain text, sometimes as follow-up updates after an image or partial draft.

The server converts those messages into structured JSON, merges updates into the same order, and writes the final result into a OneDrive Excel table through Microsoft Graph.

## Typical Use Cases

- A salesperson sends a full text order and it should be recorded directly
- A salesperson sends part of an order first, then later sends the address, phone number, or payment info
- An OpenClaw Agent needs a stable MCP endpoint instead of relying on free-form chat reasoning

## Example Input

```text
单号:26.3.13-7
测试客户A
示例产品2000支105元
收件人: 测试联系人
手机号码: 13800000000
收货地址：测试省测试市测试区示例路88号
```

## Example Output

```json
{
  "日期": "2026-03-13",
  "单号": "26.3.13-7",
  "销售员": "业务员A",
  "客户": "测试客户A",
  "货品名称": "示例产品",
  "数量": "2000",
  "数量单位": "支",
  "销售金额": "105",
  "总货款": "105",
  "已收": "105",
  "未收": "0",
  "收货联系人": "测试联系人",
  "收货人电话": "13800000000",
  "收货地址": "测试省测试市测试区示例路88号"
}
```

## Flow

1. OpenClaw sends chat text into `ingest_order_message`
2. The server parses the text into a normalized order object
3. If this is a follow-up message, it merges the update into the existing draft
4. The final order is matched against recent Excel rows
5. The server updates an existing row or creates a new one

## Privacy Note

This repository should only contain desensitized sample text and JSON.
Do not commit real customer names, phone numbers, addresses, order screenshots, or payment details.

## Features

- Parse order text from chat messages
- Merge follow-up text into an existing draft order
- Match recent orders by salesperson and customer from bottom to top
- Standardize product names against a OneDrive product catalog with local caching
- Write or update OneDrive Excel rows through Microsoft Graph
- Expose tools over Streamable HTTP for OpenClaw

## Tools

- `ingest_order_message`
- `parse_wechat_order_message`
- `merge_order_update`
- `process_excel_order`
- `check_product_catalog_status`
- `refresh_product_catalog`
- `resolve_product_name`
- `analyze_product_catalog_patterns`

## Requirements

- Python 3.10+
- A Microsoft Azure app registration with OneDrive permission
- An Excel workbook stored in OneDrive with a named table

## Quick Start

```bash
git clone <your-repo-url> cy-excel-mcp
cd cy-excel-mcp
./bootstrap.sh
cp .env.example .env
```

Edit `.env`:

```env
OC_OD_TENANT_ID=consumers
OC_OD_CLIENT_ID=your_microsoft_app_client_id
OC_OD_FILE_PATH=YourFolder/订单汇总.xlsx
OC_OD_TABLE_NAME=表1
OC_OD_CACHE_FILE=onedrive_token_cache.bin
OC_OD_PRODUCT_FILE_PATH=众一/2026诚亿报表.xlsx
OC_OD_PRODUCT_SHEET_NAME=产品明细
OC_OD_PRODUCT_NAME_COLUMN=B
OC_OD_PRODUCT_CATEGORY_COLUMN=C
CY_PRODUCT_CACHE_FILE=product_catalog_cache.json
CY_PRODUCT_ALIAS_FILE=product_aliases.json

CY_EXCEL_MCP_HOST=127.0.0.1
CY_EXCEL_MCP_PORT=18061
CY_EXCEL_MCP_TRANSPORT=streamable-http
```

Start the server:

```bash
./start_cy_excel_mcp_http.sh
```

The default MCP endpoint is:

```text
http://127.0.0.1:18061/mcp
```

## Manual Installation

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install --upgrade pip
pip install -e .
```

Or run directly after installation:

```bash
cy-excel-mcp --transport streamable-http --host 127.0.0.1 --port 18061
```

## OpenClaw MCP Config

Example `mcporter.json`:

```json
{
  "mcpServers": {
    "cy-excel-mcp": {
      "baseUrl": "http://127.0.0.1:18061/mcp"
    }
  },
  "imports": []
}
```

An example file is also included at `config/mcporter.json.example`.

Detailed runtime notes are available in `docs/RUNNING.md`.
Chinese runtime notes are available in `docs/RUNNING.zh-CN.md`.
OpenClaw integration notes are available in `docs/OPENCLAW_INTEGRATION.md`.

## OpenClaw Agent Flow

- Full text order: `ingest_order_message`
- Follow-up text for an existing draft: `ingest_order_message` with `existing_order`
- Parse only: `parse_wechat_order_message`
- Merge only: `merge_order_update`
- Write only: `process_excel_order`

## Why This Project

- Keeps Excel write logic outside prompt-only workflows
- Makes OpenClaw behavior more stable through explicit MCP tools
- Supports salesperson-first matching and bottom-up order updates
- Works well for chat-driven order entry teams using OneDrive Excel

## Matching Rules

- Prefer matching within the same salesperson
- Search from the bottom of Excel upwards
- Match priority:
  1. `单号`
  2. `客户`
  3. `匹配客户别名`

## Repository Layout

- `cy_excel_mcp.py`: MCP server implementation
- `start_cy_excel_mcp_http.sh`: local HTTP startup script
- `bootstrap.sh`: one-command local setup
- `.env.example`: environment variable template
- `config/mcporter.json.example`: OpenClaw MCP config template

## Notes

- Personal OneDrive has been validated with `OC_OD_TENANT_ID=consumers`.
- `OC_OD_FILE_PATH` can include a nested folder path such as `YourFolder/订单汇总.xlsx`.
- `OC_OD_TABLE_NAME` must be the Excel table object name, not the worksheet name.
- Product name standardization reads the product-name column and category column from `OC_OD_PRODUCT_SHEET_NAME` in `OC_OD_PRODUCT_FILE_PATH`.
- Before each order write, the server checks product workbook metadata and uses `product_catalog_cache.json` when the file has not changed.
- The first Microsoft login may require device-flow authorization in the terminal.
- After the first successful sign-in, the token is cached in `onedrive_token_cache.bin` and reused on later runs.
- Empty values do not overwrite existing Excel values.
- Keep `.env` and `onedrive_token_cache.bin` out of version control.
- Temporary files such as swap files and `__pycache__` are already ignored by `.gitignore`.

## License

MIT. See `LICENSE`.
