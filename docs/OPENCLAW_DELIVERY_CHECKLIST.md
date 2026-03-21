# OpenClaw Delivery Checklist

Use this checklist before handing the project over for daily use.

## Code And Runtime

- `cy_excel_mcp.py` starts successfully through `./start_cy_excel_mcp_http.sh`
- `.venv` is present and dependencies are installed
- `.env` is filled with valid values
- `OC_OD_TENANT_ID=consumers` is used for personal OneDrive
- `OC_OD_FILE_PATH` points to the correct Excel file path
- `OC_OD_TABLE_NAME` matches the real Excel table object name

## Logging

- `logs/` is created automatically on startup
- Each startup creates a new log file
- Log filename format is `YYYY-MM-DD-HHMMSS-SEQ.log`
- Old logs are cleaned up according to `LOG_RETENTION_DAYS`

## Microsoft Login

- First login completes successfully through device flow
- `onedrive_token_cache.bin` is created after login
- A second run can reuse the cached token without another login prompt

## OpenClaw Integration

- OpenClaw `mcporter.json` points to `http://127.0.0.1:18061/mcp`
- OpenClaw can discover `ingest_order_message`
- OpenClaw can discover `parse_wechat_order_message`
- OpenClaw can discover `merge_order_update`
- OpenClaw can discover `process_excel_order`

## Business Validation

- A desensitized text order can be inserted into Excel
- A follow-up message can update an existing order
- Matching follows the expected salesperson-first, bottom-up behavior
- Empty values do not overwrite existing Excel values

## Repository Hygiene

- No real customer data is committed
- `.env` is not committed
- `onedrive_token_cache.bin` is not committed
- `logs/` is ignored
- Temporary files such as `__pycache__`, swap files, and `.egg-info` are ignored

## Release

- `CHANGELOG.md` includes the latest release notes
- The repository is pushed to GitHub
- The latest git tag exists locally and remotely
