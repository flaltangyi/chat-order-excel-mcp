# Contributing / 贡献指南

Thank you for contributing to `chat-order-excel-mcp`.

感谢你为 `chat-order-excel-mcp` 做贡献。

This document provides a practical contribution guide in both English and Chinese.

本文档提供一份中英文对照的实际贡献说明。

## Scope / 适用范围

This project is an MCP server for OpenClaw order-entry workflows.

本项目是一个面向 OpenClaw 录单流程的 MCP 服务。

Typical contribution areas:

常见贡献范围包括：

- MCP tool behavior
- WeChat order parsing logic
- Excel write/update logic
- OpenClaw integration docs
- Startup/runtime scripts
- Documentation and examples

## Before You Start / 开始之前

Please make sure you understand the current runtime model:

开始之前，请先理解当前运行模型：

- Personal OneDrive flow is currently validated with `OC_OD_TENANT_ID=consumers`
- The MCP server is exposed over Streamable HTTP
- OpenClaw connects through `mcporter.json`
- Tokens are cached locally in `onedrive_token_cache.bin`

## Local Setup / 本地环境准备

Recommended setup:

推荐的初始化方式：

```bash
cd /opt/ctrlExcel
./bootstrap.sh
cp .env.example .env
```

Then fill in `.env` with your own values.

然后根据你自己的环境填写 `.env`。

Start the MCP server with:

启动服务：

```bash
./start_cy_excel_mcp_http.sh
```

## Privacy And Data Safety / 隐私与数据安全

Do not commit any real customer data.

不要提交任何真实客户数据。

Never commit:

严禁提交以下内容：

- Real customer names
- Real phone numbers
- Real addresses
- Real payment details
- Real order screenshots
- `.env`
- `onedrive_token_cache.bin`
- `logs/`

Only desensitized sample text and JSON may appear in the repository.

仓库中只允许出现脱敏后的示例文本和 JSON。

## Code Changes / 代码修改要求

Please keep changes focused and practical.

请保持修改聚焦且实用。

When changing code:

修改代码时建议遵循：

- Preserve the current MCP tool interfaces unless there is a strong reason to change them
- Prefer explicit business rules over vague prompt-only behavior
- Keep OpenClaw-facing behavior predictable
- Avoid adding features that depend on hardcoded private environment details

## Validation / 验证要求

Before submitting changes, verify as many of these as possible:

提交前尽量完成以下验证：

```bash
python3 -m py_compile /opt/ctrlExcel/cy_excel_mcp.py
bash -n /opt/ctrlExcel/start_cy_excel_mcp_http.sh
```

If your change affects runtime behavior, also verify:

如果修改影响运行逻辑，还应验证：

- The MCP server starts successfully
- OpenClaw can discover the MCP tools
- A desensitized test order can be parsed
- If applicable, Excel write/update still works

## Documentation / 文档要求

If you change behavior, update documentation with it.

如果你修改了行为逻辑，请同步更新文档。

Relevant files may include:

常见需要同步更新的文件包括：

- `README.md`
- `README.zh-CN.md`
- `docs/RUNNING.md`
- `docs/RUNNING.zh-CN.md`
- `CHANGELOG.md`

## Commits / 提交规范

Prefer short, clear commit messages.

建议使用简短且明确的提交信息。

Examples:

示例：

- `fix: improve OneDrive token handling`
- `docs: update OpenClaw runtime guide`
- `chore: add startup log rotation`

## Pull Requests / 合并请求

A good pull request should include:

一个好的合并请求应尽量包含：

- What changed
- Why it changed
- What was verified
- Any known limitations

## Questions / 问题反馈

If a change touches Microsoft Graph auth, OneDrive behavior, or OpenClaw integration, describe the exact runtime assumptions clearly.

如果改动涉及 Microsoft Graph 认证、OneDrive 行为或 OpenClaw 集成，请明确写清楚运行前提和假设。
