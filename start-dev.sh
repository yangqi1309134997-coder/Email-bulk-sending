#!/usr/bin/env bash
set -euo pipefail

echo "=== 启动邮箱群发系统（开发模式）==="

# 确保后端代理端口一致
if [ -f "frontend/vite.config.js" ]; then
    sed -i 's|http://localhost:[0-9]\+/|http://localhost:8000|g' frontend/vite.config.js || true
    # 也修复 .env.example / .env 中可能被改动的数据库 URL（保留 SQLite 开发默认）
fi

echo "[1/4] 启动后端 (FastAPI) 在 8000 端口..."
cd backend
nohup python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload > ../logs/backend.log 2>&1 &
echo $! > ../.backend.pid
cd ..
mkdir -p logs

echo "[2/4] 启动前端 (Vite) 在 5173 端口..."
nohup npm --prefix ./frontend run dev -- --host 0.0.0.0 --port 5173 > logs/frontend.log 2>&1 &
echo $! > .frontend.pid

echo "[3/4] 等待服务就绪..."
sleep 3

echo "[4/4] 检查状态..."
if curl -sf http://localhost:8000/health/ready > /dev/null 2>&1; then
    echo "✅ 后端就绪: http://localhost:8000"
else
    echo "⚠️  后端未就绪，请检查 logs/backend.log"
fi

if curl -sf -I http://localhost:5173/ > /dev/null 2>&1; then
    echo "✅ 前端就绪: http://localhost:5173"
else
    echo "�️  前端未就绪，请检查 logs/frontend.log"
fi

echo "=== 启动完成 ==="
echo "登录账号: admin / admin123"
echo "如要停止: ./stop-dev.sh"
