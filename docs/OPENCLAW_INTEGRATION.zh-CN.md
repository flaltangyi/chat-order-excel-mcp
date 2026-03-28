# OpenClaw 接入说明

这份文档说明如何以公开、可复现的方式，把 `chat-order-excel-mcp` 接到 OpenClaw Agent，用于企业微信的订单录入流程。

文档刻意不包含以下敏感信息：

- 真实 bot 凭证
- 真实企业 ID
- 真实回调密钥
- 真实用户 ID（除示例占位符外）
- 本地 token / cache 内容

## 推荐架构

订单录入场景建议使用一个固定 Agent，而不是为每个用户生成一个动态 workspace。

- 渠道：企业微信 `wecom`
- Agent ID：`orderentry`
- MCP 服务名：`chat-order-excel-mcp`
- Agent 调用 MCP 的方式：`stdio`

这样更适合共享录单场景，因为：

- 所有业务员共用同一套订单规则
- 所有消息都进入同一个订单 Agent
- Agent 可以直接调用稳定的 MCP tools
- 不会因为动态 Agent workspace 导致工具暴露不一致

## 关键限制

OpenClaw 当前把 MCP 挂成 Agent tool 时，优先支持的是 `stdio` 启动方式。

如果你只配置 HTTP MCP，例如：

```json
{
  "mcp": {
    "servers": {
      "chat-order-excel-mcp": {
        "url": "http://127.0.0.1:18061/mcp",
        "enabled": true,
        "transport": "http"
      }
    }
  }
}
```

那么 MCP 服务本身可能可访问，但 Agent 侧仍然可能看不到 `health`、`ingest_order_message` 等工具。

要让 Agent 直接调用 MCP tool，优先用 `stdio`。

## OpenClaw 配置示例

在 `~/.openclaw/openclaw.json` 中加入或更新 MCP 配置：

```json
{
  "mcp": {
    "servers": {
      "chat-order-excel-mcp": {
        "command": "/path/to/project/.venv/bin/python",
        "args": [
          "/path/to/project/cy_excel_mcp.py",
          "--transport",
          "stdio"
        ],
        "cwd": "/path/to/project",
        "enabled": true
      }
    }
  }
}
```

## Agent 配置示例

把订单 Agent 绑定到企业微信，并在 `tools.alsoAllow` 中放行 MCP 服务名：

```json
{
  "agents": {
    "list": [
      {
        "id": "orderentry",
        "name": "订单小助手",
        "workspace": "/root/.openclaw/workspace-orderEntry",
        "tools": {
          "profile": "messaging",
          "alsoAllow": [
            "memory_forget",
            "memory_recall",
            "memory_store",
            "wecom_mcp",
            "tts",
            "chat-order-excel-mcp"
          ]
        }
      }
    ]
  },
  "bindings": [
    {
      "agentId": "orderentry",
      "match": {
        "channel": "wecom",
        "accountId": "default"
      }
    }
  ]
}
```

## 为什么用 `profile: "messaging"`

对于共享业务机器人，`messaging` 比 `full` 更安全。

`full` 可能会把宿主机读写、执行类工具一起暴露出来，这对录单场景并不必要。

`messaging` 配合显式 `alsoAllow`，可以让 Agent 使用：

- 企业微信消息相关工具
- memory 工具
- `chat-order-excel-mcp`

同时避免把宿主机操作能力直接暴露给普通用户。

## 企业微信路由建议

推荐的 WeCom 配置：

```json
{
  "channels": {
    "wecom": {
      "enabled": true,
      "dmPolicy": "pairing",
      "groupPolicy": "open",
      "groupChat": {
        "enabled": true,
        "requireMention": true,
        "mentionPatterns": ["@"]
      },
      "dynamicAgents": {
        "enabled": false,
        "adminBypass": false
      },
      "dm": {
        "createAgentOnFirstMessage": false
      }
    }
  }
}
```

这样企业微信订单消息会固定进入 `orderentry`，而不是给每个用户创建动态 workspace。

## 管理员可操作，普通用户不可操作

如果你需要企业微信管理员具备更高权限，而普通用户不具备，建议把职责分开：

1. 共享订单 bot 固定走 `orderentry`，并保持 `profile: "messaging"`
2. 在 wecom 渠道里配置 `adminUsers`
3. 在 `commands.ownerAllowFrom` 里配置 owner 身份
4. 如果确实要做宿主机操作，建议单独再配一个管理员专用 Agent 或单独 bot/account

示例：

```json
{
  "channels": {
    "wecom": {
      "adminUsers": ["admin-userid"]
    }
  },
  "commands": {
    "ownerAllowFrom": [
      "wecom:admin-userid"
    ]
  }
}
```

这比在共享订单助手上直接开放宿主机控制更安全。

## MCP 项目自身配置

项目本身仍然需要自己的 `.env`，例如：

```env
OC_OD_TENANT_ID=consumers
OC_OD_CLIENT_ID=your_client_id
OC_OD_FILE_PATH=YourFolder/订单汇总.xlsx
OC_OD_TABLE_NAME=表1
OC_OD_CACHE_FILE=onedrive_token_cache.bin
```

这些都是本地部署配置，不应该提交到仓库。

## 验证步骤

修改 OpenClaw 配置后：

1. 重启 OpenClaw gateway
2. 在企业微信发送 `调用 health`
3. 确认 Agent 返回的是 MCP 健康检查结果，而不是提示工具不可用
4. 再执行 `check_login_status`
5. 再执行 `list_excel_tables`
6. 真实写入前，先跑一次 `dry_run` 测试

## 预期的健康检查结果

健康状态下，应该能返回类似这些信息：

- MCP 服务可用
- Excel 文件已配置
- 表格名已配置
- Microsoft Client ID 已配置
- token cache 存在，或提示需要登录

## 安全注意事项

- 不要提交 `.env`
- 不要提交 token cache 文件
- 不要提交真实企业微信 bot secret
- 不要提交真实用户 ID，除非它们是你明确允许公开的占位符
- 不要把共享订单 bot 和无限制宿主机控制工具放在同一个 Agent profile 里
