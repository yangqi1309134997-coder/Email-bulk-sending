#!/usr/bin/env bash
set -euo pipefail

echo "=== 停止开发服务 ==="

if [ -f ".backend.pid" ]; then
    kill $(cat .backend.pid) 2>/dev/null || echo "后端已停止"
    rm -f .backend.pid
fi

if [ -f ".frontend.pid" ]; then
    kill $(cat .frontend.pid) 2>/dev/null || echo "前端已停止"
    rm -f .frontend.pid
fi

echo "已停止"
