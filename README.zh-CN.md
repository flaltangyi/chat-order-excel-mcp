# cy-excel-mcp

用于解析微信群订单消息、合并补充信息，并将结构化订单写入 OneDrive 在线 Excel 的 MCP 服务。

[English README](README.md)

## 项目简介

`cy-excel-mcp` 适合微信群录单场景：业务员有时发完整文字订单，有时先发一部分，再补地址、电话、收款信息，甚至先发图片、后发补充文字。

这个服务的职责是把这些消息转成统一的结构化 JSON，必要时合并成同一笔订单，再通过 Microsoft Graph 写入 OneDrive 里的 Excel 表格。

## 适用场景

- 业务员发送完整文字订单，需要直接录入 Excel
- 同一笔订单分多次发送，需要自动合并更新
- OpenClaw 需要稳定调用 MCP，而不是仅靠聊天推理完成录单

## 输入示例

```text
单号:26.3.13-7
测试客户A
示例产品2000支105元
收件人: 测试联系人
手机号码: 13800000000
收货地址：测试省测试市测试区示例路88号
```

## 输出示例

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

## 工作流程

1. OpenClaw 把聊天消息传给 `ingest_order_message`
2. 服务将消息解析成统一订单对象
3. 如果是补充消息，则合并到已有草稿订单
4. 按最近 Excel 记录进行匹配
5. 最终执行更新或新增

## 隐私说明

仓库中只应保留脱敏后的示例文本和 JSON。
不要提交真实客户姓名、电话、地址、订单截图或收款信息。

## 功能

- 解析微信群里的订单文字
- 将后续补充文字合并到已有草稿订单
- 按业务员和客户信息从 Excel 底部倒序匹配最近订单
- 通过 Microsoft Graph 新增或更新 Excel 行
- 通过 Streamable HTTP 暴露给 OpenClaw 使用

## 可用工具

- `ingest_order_message`
- `parse_wechat_order_message`
- `merge_order_update`
- `process_excel_order`

## 环境要求

- Python 3.10+
- 已在 Azure 注册微软应用，并具备 OneDrive 访问权限
- OneDrive 中已有一个包含命名表格的 Excel 文件

## 快速开始

```bash
git clone <你的仓库地址> cy-excel-mcp
cd cy-excel-mcp
./bootstrap.sh
cp .env.example .env
```

编辑 `.env`：

```env
OC_OD_TENANT_ID=consumers
OC_OD_CLIENT_ID=你的微软应用client_id
OC_OD_FILE_PATH=你的文件夹/订单汇总.xlsx
OC_OD_TABLE_NAME=表1
OC_OD_CACHE_FILE=onedrive_token_cache.bin

CY_EXCEL_MCP_HOST=127.0.0.1
CY_EXCEL_MCP_PORT=18061
CY_EXCEL_MCP_TRANSPORT=streamable-http
```

启动服务：

```bash
./start_cy_excel_mcp_http.sh
```

默认 MCP 地址：

```text
http://127.0.0.1:18061/mcp
```

## 手动安装

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install --upgrade pip
pip install -e .
```

安装后也可以直接执行：

```bash
cy-excel-mcp --transport streamable-http --host 127.0.0.1 --port 18061
```

## OpenClaw 的 MCP 配置

`mcporter.json` 示例：

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

仓库内也附带了模板文件：`config/mcporter.json.example`

## OpenClaw Agent 调用流程

- 完整文本订单：`ingest_order_message`
- 对已有草稿的补充消息：`ingest_order_message` 并传 `existing_order`
- 只解析：`parse_wechat_order_message`
- 只合并：`merge_order_update`
- 只写入：`process_excel_order`

## 为什么适合做成 MCP

- 把 Excel 写入逻辑从提示词里剥离出来
- 用显式工具替代纯聊天推理，稳定性更高
- 支持按业务员、客户、倒序记录进行合并更新
- 适合团队长期使用 OneDrive Excel 做录单归档

## 匹配规则

- 优先限制在同一个业务员范围内匹配
- 从 Excel 底部向上倒序查找最近记录
- 匹配优先级：
  1. `单号`
  2. `客户`
  3. `匹配客户别名`

## 目录说明

- `cy_excel_mcp.py`：MCP 服务主体
- `start_cy_excel_mcp_http.sh`：本地 HTTP 启动脚本
- `bootstrap.sh`：一键初始化脚本
- `.env.example`：环境变量模板
- `config/mcporter.json.example`：OpenClaw 的 MCP 配置模板

## 注意事项

- 个人 OneDrive 已验证可用，建议 `OC_OD_TENANT_ID=consumers`
- `OC_OD_FILE_PATH` 可以填写带文件夹的相对路径，例如 `你的文件夹/订单汇总.xlsx`
- `OC_OD_TABLE_NAME` 必须是 Excel 里的表格对象名，不是工作表名称
- 首次连接微软账号时，终端里可能会出现 device flow 登录提示
- 首次登录成功后，token 会缓存到 `onedrive_token_cache.bin`，后续运行会优先复用
- 空值不会覆盖 Excel 里已有值
- `.env` 和 `onedrive_token_cache.bin` 不应提交到仓库
- `swap`、`__pycache__` 等临时文件已在 `.gitignore` 中忽略

## License

MIT，见 `LICENSE`
