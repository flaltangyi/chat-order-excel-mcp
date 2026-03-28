# OpenClaw Integration

This document describes a public, reproducible way to connect `chat-order-excel-mcp` to an OpenClaw Agent for Enterprise WeChat order-entry workflows.

It intentionally excludes all sensitive data such as:

- real bot credentials
- real company IDs
- real callback secrets
- real user IDs beyond example placeholders
- local private token/cache contents

## Recommended Topology

Use one fixed Agent for order entry instead of per-user dynamic workspaces.

- Channel: Enterprise WeChat (`wecom`)
- Agent ID: `orderentry`
- MCP server name: `chat-order-excel-mcp`
- MCP transport for Agent tools: `stdio`

This topology is better for shared order-entry workflows because:

- all salespeople use the same order rules
- all messages land in the same order Agent
- the Agent can call stable MCP tools directly
- there is no per-user tool drift caused by dynamic Agent workspaces

## Important Limitation

OpenClaw Agent runtime currently exposes bundled/configured MCP tools as Agent tools through `stdio` launch config.

If you only configure an HTTP MCP endpoint like:

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

the MCP server may be reachable externally, but the Agent may still fail to see `health`, `ingest_order_message`, and the other MCP tools.

For Agent tool exposure, prefer `stdio`.

## OpenClaw Config Example

Add or update the MCP config in `~/.openclaw/openclaw.json`:

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

## Agent Config Example

Bind the order Agent to WeCom and allow the MCP server name in `tools.alsoAllow`.

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

## Why `profile: \"messaging\"`

For a shared business bot, `messaging` is safer than `full`.

`full` may expose host-level read/write/exec style tools that are unnecessary for order entry.

`messaging` plus explicit `alsoAllow` lets the Agent use:

- WeCom messaging tools
- memory tools
- `chat-order-excel-mcp`

without broadly exposing host-computer operation to ordinary users.

## Enterprise WeChat Routing

Recommended WeCom settings:

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

This keeps WeCom order traffic on the fixed `orderentry` Agent instead of generating per-user dynamic workspaces.

## Admin-Only Host Control

If you need WeCom admins to have elevated command/operator privileges while ordinary users do not, separate the concerns:

1. Keep the shared order bot on `orderentry` with `profile: "messaging"`.
2. Configure WeCom admin users in channel config.
3. Configure owner identity in `commands.ownerAllowFrom`.
4. If host-computer operation is required, prefer a separate admin-facing Agent or a separate bot/account.

Example:

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

This is safer than exposing computer-control tools on the shared order-entry Agent.

## MCP Project Requirements

The project itself still needs its own `.env` values, for example:

```env
OC_OD_TENANT_ID=consumers
OC_OD_CLIENT_ID=your_client_id
OC_OD_FILE_PATH=YourFolder/订单汇总.xlsx
OC_OD_TABLE_NAME=表1
OC_OD_CACHE_FILE=onedrive_token_cache.bin
```

These values are local deployment secrets/configuration and should not be committed.

## Validation Checklist

After updating OpenClaw config:

1. Restart the OpenClaw gateway.
2. Send `调用 health` in Enterprise WeChat.
3. Confirm the Agent returns MCP health data rather than saying the tool is unavailable.
4. Run `check_login_status`.
5. Run `list_excel_tables`.
6. Run one `dry_run` order test before real write.

## Expected Health Result

A healthy integration should return something equivalent to:

- MCP service available
- Excel file configured
- table name configured
- Microsoft Client ID configured
- token cache present or login required

## Security Notes

- Do not commit `.env`
- Do not commit token cache files
- Do not commit real WeCom bot secrets
- Do not commit real user IDs unless they are intentionally public placeholders
- Do not mix shared order-entry bots with unrestricted host-control tool profiles
