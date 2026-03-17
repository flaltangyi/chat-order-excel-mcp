#!/usr/bin/env bash
set -euo pipefail

cd /opt/ctrlExcel

export CY_EXCEL_MCP_HOST="${CY_EXCEL_MCP_HOST:-127.0.0.1}"
export CY_EXCEL_MCP_PORT="${CY_EXCEL_MCP_PORT:-18061}"
export CY_EXCEL_MCP_TRANSPORT="${CY_EXCEL_MCP_TRANSPORT:-streamable-http}"

exec python3 /opt/ctrlExcel/cy_excel_mcp.py --transport "${CY_EXCEL_MCP_TRANSPORT}"
