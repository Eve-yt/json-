#!/bin/zsh

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$PROJECT_DIR/work/server.pid"

if [[ ! -f "$PID_FILE" ]]; then
  echo "没有找到由本项目脚本启动的服务。"
  exit 0
fi

PID="$(cat "$PID_FILE")"
if [[ ! "$PID" =~ '^[0-9]+$' ]]; then
  echo "PID 文件无效，已清理。"
  rm -f "$PID_FILE"
  exit 1
fi

if kill -0 "$PID" 2>/dev/null; then
  kill "$PID"
  for _ in {1..20}; do
    if ! kill -0 "$PID" 2>/dev/null; then
      break
    fi
    sleep 0.1
  done
  if kill -0 "$PID" 2>/dev/null; then
    echo "服务未能在规定时间内停止：PID=$PID"
    exit 1
  fi
  echo "服务已停止：PID=$PID"
else
  echo "服务进程已不存在：PID=$PID"
fi

rm -f "$PID_FILE"
