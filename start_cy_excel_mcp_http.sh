#!/usr/bin/env bash
set -euo pipefail

cd /opt/ctrlExcel

LOG_DIR="/opt/ctrlExcel/logs"
TODAY="$(date +%F)"
mkdir -p "${LOG_DIR}"

count=0
for path in "${LOG_DIR}/${TODAY}"-*.log; do
  if [ -e "${path}" ]; then
    count=$((count + 1))
  fi
done

sequence=$(printf "%03d" $((count + 1)))
LOG_FILE="${LOG_DIR}/${TODAY}-${sequence}.log"

exec > >(tee -a "${LOG_FILE}") 2>&1

echo "Logging to ${LOG_FILE}"

if [ -f .venv/bin/activate ]; then
  # shellcheck disable=SC1091
  . .venv/bin/activate
fi

export CY_EXCEL_MCP_HOST="${CY_EXCEL_MCP_HOST:-127.0.0.1}"
export CY_EXCEL_MCP_PORT="${CY_EXCEL_MCP_PORT:-18061}"
export CY_EXCEL_MCP_TRANSPORT="${CY_EXCEL_MCP_TRANSPORT:-streamable-http}"

if command -v fuser >/dev/null 2>&1 && fuser "${CY_EXCEL_MCP_PORT}/tcp" >/dev/null 2>&1; then
  pkill -f "cy_excel_mcp.py" || true
  sleep 1
fi

exec python3 /opt/ctrlExcel/cy_excel_mcp.py --transport "${CY_EXCEL_MCP_TRANSPORT}"
