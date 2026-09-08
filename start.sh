#!/bin/zsh

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$PROJECT_DIR/work/server.pid"
LOG_FILE="$PROJECT_DIR/work/server.log"
BIND_HOST="${BIND_HOST:-127.0.0.1}"
PORT="${PORT:-8765}"

DEFAULT_PYTHON="/Users/yyt/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
if [[ -n "${PYTHON_BIN:-}" ]]; then
  PYTHON="$PYTHON_BIN"
elif [[ -x "$DEFAULT_PYTHON" ]]; then
  PYTHON="$DEFAULT_PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON="$(command -v python3)"
else
  echo "未找到 Python 3，请设置 PYTHON_BIN 后重试。"
  exit 1
fi

mkdir -p "$PROJECT_DIR/work"

if [[ -f "$PID_FILE" ]]; then
  PID="$(cat "$PID_FILE")"
  if [[ "$PID" =~ '^[0-9]+$' ]] && kill -0 "$PID" 2>/dev/null; then
    echo "服务已经在运行：PID=$PID，地址：http://$BIND_HOST:$PORT/"
    exit 0
  fi
  rm -f "$PID_FILE"
fi

if ! "$PYTHON" -c 'import openpyxl' >/dev/null 2>&1; then
  echo "当前 Python 缺少 openpyxl。"
  echo "可执行：$PYTHON -m pip install openpyxl"
  exit 1
fi

nohup "$PYTHON" "$PROJECT_DIR/work/matching_review_tool.py" \
  --host "$BIND_HOST" \
  --port "$PORT" \
  >"$LOG_FILE" 2>&1 &

PID=$!
echo "$PID" > "$PID_FILE"
sleep 0.5

if ! kill -0 "$PID" 2>/dev/null; then
  echo "服务启动失败，日志如下："
  cat "$LOG_FILE"
  rm -f "$PID_FILE"
  exit 1
fi

echo "服务已启动：http://$BIND_HOST:$PORT/"
echo "PID：$PID"
echo "日志：$LOG_FILE"
