# Running Guide

This document describes the final recommended local runtime flow for `cy-excel-mcp`.

## Prerequisites

- Python virtual environment has been created with `./bootstrap.sh`
- `.env` has been filled with valid values
- OpenClaw MCP config points to `http://127.0.0.1:18061/mcp`

## Recommended `.env` pattern for personal OneDrive

```env
OC_OD_TENANT_ID=consumers
OC_OD_CLIENT_ID=your_microsoft_app_client_id
OC_OD_FILE_PATH=YourFolder/订单汇总.xlsx
OC_OD_TABLE_NAME=表1
OC_OD_CACHE_FILE=onedrive_token_cache.bin

CY_EXCEL_MCP_HOST=127.0.0.1
CY_EXCEL_MCP_PORT=18061
CY_EXCEL_MCP_TRANSPORT=streamable-http
```

## Start

```bash
cd /opt/ctrlExcel
./start_cy_excel_mcp_http.sh
```

What the startup script does:

- Activates `.venv` if present
- Cleans up old log files in `logs/`
- Kills a stale `cy_excel_mcp.py` process if the configured port is already in use
- Starts the MCP HTTP server

## Logs

Logs are written to:

```text
/opt/ctrlExcel/logs
```

Filename pattern:

```text
YYYY-MM-DD-HHMMSS-SEQ.log
```

Examples:

- `2026-03-21-101530-001.log`
- `2026-03-21-143000-002.log`

Retention:

- By default, only the latest 7 days are kept
- Override with:

```bash
LOG_RETENTION_DAYS=14 ./start_cy_excel_mcp_http.sh
```

## First Login

If no cached Microsoft token exists yet, you may see a device login prompt.

For personal OneDrive, complete the login with your personal Microsoft account.

After the first successful sign-in, the token is stored in:

```text
/opt/ctrlExcel/onedrive_token_cache.bin
```

## Stop

Press:

```text
Ctrl+C
```

The service should stop cleanly.

## OpenClaw Verification

Recommended verification order:

1. Start the MCP server locally
2. Confirm OpenClaw can see `ingest_order_message`
3. Send a desensitized text order
4. Complete first-time device login if prompted
5. Confirm the order is inserted into Excel
6. Send another test order and confirm token cache avoids another login
