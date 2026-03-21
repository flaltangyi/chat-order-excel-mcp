# 运行说明

本文档用于说明 `cy-excel-mcp` 当前推荐的本地运行方式。

## 前置条件

- 已执行 `./bootstrap.sh`
- `.env` 已填写正确
- OpenClaw 的 MCP 配置已指向 `http://127.0.0.1:18061/mcp`

## 个人 OneDrive 推荐配置

```env
OC_OD_TENANT_ID=consumers
OC_OD_CLIENT_ID=your_microsoft_app_client_id
OC_OD_FILE_PATH=你的文件夹/订单汇总.xlsx
OC_OD_TABLE_NAME=表1
OC_OD_CACHE_FILE=onedrive_token_cache.bin

CY_EXCEL_MCP_HOST=127.0.0.1
CY_EXCEL_MCP_PORT=18061
CY_EXCEL_MCP_TRANSPORT=streamable-http
```

## 启动

```bash
cd /opt/ctrlExcel
./start_cy_excel_mcp_http.sh
```

启动脚本会自动完成：

- 激活 `.venv`
- 清理超过保留天数的旧日志
- 如果端口被旧的 `cy_excel_mcp.py` 进程占用，则先清理
- 启动 MCP HTTP 服务

## 日志

日志目录：

```text
/opt/ctrlExcel/logs
```

文件命名规则：

```text
YYYY-MM-DD-HHMMSS-SEQ.log
```

示例：

- `2026-03-21-101530-001.log`
- `2026-03-21-143000-002.log`

默认保留策略：

- 默认只保留最近 7 天日志

如需覆盖：

```bash
LOG_RETENTION_DAYS=14 ./start_cy_excel_mcp_http.sh
```

## 首次登录

如果当前还没有可用 token 缓存，程序会提示微软 device flow 登录。

对于个人 OneDrive：

- 请使用你的个人 Microsoft 账号登录

首次成功后，token 会缓存到：

```text
/opt/ctrlExcel/onedrive_token_cache.bin
```

后续运行会优先复用缓存，不需要每次重新登录。

## 停止

按：

```text
Ctrl+C
```

服务应正常停止。

## OpenClaw 验证顺序

建议按下面顺序验证：

1. 本地启动 MCP 服务
2. 在 OpenClaw 确认能看到 `ingest_order_message`
3. 发送一条脱敏测试订单
4. 如提示登录，则完成首次 device flow 授权
5. 确认 Excel 成功新增或更新
6. 再发一条测试订单，确认 token cache 生效，不再要求重新登录
