#!/usr/bin/env bash
# 一键部署脚本（生产环境）
# 用法：
#   ./deploy.sh                      # 交互式，引导生成 .env 并部署
#   ./deploy.sh --yes                # 全自动，使用默认值
#   ./deploy.sh --domain=https://mail.example.com --password='Your!Pass1'
#   ./deploy.sh --force              # 即使 .env 已存在也重新生成密钥
#
# 依赖：docker、docker compose 插件（v2）、python3（仅用于安全改写 .env）
set -euo pipefail
cd "$(dirname "$0")"

DOMAIN="http://localhost:80"
ADMIN_PASSWORD=""
ASSUME_YES=false
FORCE=false

while [ $# -gt 0 ]; do
  case "$1" in
    --yes) ASSUME_YES=true; shift ;;
    --force) FORCE=true; shift ;;
    --domain=*) DOMAIN="${1#*=}"; shift ;;
    --password=*) ADMIN_PASSWORD="${1#*=}"; shift ;;
    *) echo "未知参数: $1" >&2; exit 1 ;;
  esac
done

command -v docker >/dev/null || { echo "错误：未找到 docker，请先安装 Docker。" >&2; exit 1; }
docker compose version >/dev/null 2>&1 || { echo "错误：需要 docker compose 插件（v2）。" >&2; exit 1; }

random_secret() {
  python3 -c "import secrets; print(secrets.token_hex(32))"
}

# 安全改写 .env 中的键值（值中的任意字符都不会破坏格式）
update_env() {
  python3 - "$1" "$2" <<'PY'
import sys
key, value = sys.argv[1], sys.argv[2]
with open(".env", encoding="utf-8") as f:
    lines = f.readlines()
with open(".env", "w", encoding="utf-8") as f:
    found = False
    for line in lines:
        if line.startswith(key + "="):
            f.write(f"{key}={value}\n")
            found = True
        else:
            f.write(line)
    if not found:
        f.write(f"{key}={value}\n")
PY
}

NEED_GENERATE=false
if [ ! -f .env ]; then
  NEED_GENERATE=true
elif [ "$FORCE" = true ]; then
  NEED_GENERATE=true
  cp .env .env.bak.$(date +%s)
  echo "已备份原 .env 为 .env.bak.*"
fi

if [ "$NEED_GENERATE" = true ]; then
  cp .env.example .env
  update_env SECRET_KEY "$(random_secret)"
  update_env AES_KEY "$(random_secret)"
  update_env POSTGRES_PASSWORD "$(random_secret)"
  echo "已生成随机 SECRET_KEY / AES_KEY / POSTGRES_PASSWORD"
elif [ "$FORCE" = false ]; then
  echo "检测到已有 .env，保留现有配置。使用 --force 可重新生成密钥。"
fi

# 管理员密码
if [ -z "$ADMIN_PASSWORD" ]; then
  if [ "$ASSUME_YES" = true ]; then
    ADMIN_PASSWORD="admin123"
    echo "警告：--yes 模式使用默认管理员密码 admin123，生产环境请尽快修改！"
  else
    read -rsp "设置管理员密码（回车使用默认 admin123，生产请务必修改）: " ADMIN_PASSWORD
    echo
    [ -z "$ADMIN_PASSWORD" ] && ADMIN_PASSWORD="admin123"
  fi
fi
update_env DEFAULT_ADMIN_PASSWORD "$ADMIN_PASSWORD"

# 访问域名
if [ "$ASSUME_YES" = false ] && [ "$DOMAIN" = "http://localhost:80" ]; then
  read -rp "访问域名（用于追踪链接和 CORS，回车默认 ${DOMAIN}）: " INPUT_DOMAIN
  [ -n "$INPUT_DOMAIN" ] && DOMAIN="$INPUT_DOMAIN"
fi
update_env TRACKING_DOMAIN "$DOMAIN"
update_env CORS_ORIGINS "$DOMAIN"

# 部署
echo "===== 构建并启动服务 ====="
docker compose up -d --build

echo "===== 等待服务就绪 ====="
READY=false
for i in $(seq 1 60); do
  # 经本机 nginx(80) 检查，避免外部域名解析/证书问题
  if curl -fsS --max-time 3 "http://localhost/health/ready" >/dev/null 2>&1; then
    READY=true
    break
  fi
  printf "."
  sleep 2
done
echo

if [ "$READY" = true ]; then
  echo "=========================================="
  echo " 部署成功！访问地址: $DOMAIN"
  echo " 管理员账号: $(grep '^DEFAULT_ADMIN_USERNAME=' .env | cut -d= -f2)"
  echo " 管理员密码: $ADMIN_PASSWORD"
  echo "=========================================="
  echo "首次登录后请在「系统设置」中修改默认密码；如无需开放注册，将 .env 中 ALLOW_REGISTER=false 后重启：docker compose up -d"
else
  echo "服务未在 120 秒内就绪，请查看日志：" >&2
  docker compose logs --tail=50 api || true
  exit 1
fi
