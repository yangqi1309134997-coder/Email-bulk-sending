# 邮箱群发系统 4.0

面向合法通知、客户触达和内部运营的批量邮件发送系统。后端使用 FastAPI、SQLModel、Celery、Redis 和 PostgreSQL，前端使用 Vue 3、Element Plus 与 ECharts。

## 主要能力

- 主流 SMTP 服务商预设、自定义 SMTP、SSL/STARTTLS、独立 SMTP 用户名
- 阿里云 DirectMail API 与阿里云邮件推送 SMTP
- 轮询、权重和智能负载均衡，发件人配额原子预占
- SMTP 连接池、每账号并发限制、进程全局并发限制
- 失败退避、熔断、风控自动暂停、可配置冷却和自动恢复
- 立即、定时和智能时段发送，后台任务自动恢复
- 发件人配置模板、历史配置复用、邮件内容模板
- TXT、CSV、Excel 收件人导入，附件上传和用户目录隔离
- 实时 WebSocket 进度、发送日志、打开/点击追踪、流式 CSV 报告
- 任务租约和心跳，避免多进程重复执行同一发送任务

## 本地开发

后端默认可使用 SQLite，适合单机开发：

```bash
chmod +x start-dev.sh
./start-dev.sh
```

默认前端地址为 `http://localhost:5173`。开发代理默认连接 `http://localhost:8000`，可通过 `VITE_API_TARGET` 修改。

## 生产部署

### 一键部署（推荐）

服务器安装 Docker 后，一条命令完成环境引导、密钥生成和启动：

```bash
./deploy.sh                      # 交互式：设置管理员密码、访问域名
./deploy.sh --yes                # 全自动部署（默认管理员密码 admin123）
./deploy.sh --domain=https://mail.example.com --password='Your!Pass1'
```

脚本会：

1. 首次运行时从 `.env.example` 生成 `.env`，自动生成随机 `SECRET_KEY`、`AES_KEY`、`POSTGRES_PASSWORD`；
2. 写入管理员账号与访问域名（同时配置 `TRACKING_DOMAIN` 和 `CORS_ORIGINS`）；
3. `docker compose up -d --build` 启动并等待健康检查通过。

部署成功后访问 `http://localhost`（或你配置的域名）。**默认管理员账号 `admin`**，密码为部署时设置的值；`--yes` 模式为 `admin123`。首次登录后请在「用户管理」中修改密码。

如需关闭自助注册，将 `.env` 中 `ALLOW_REGISTER=false` 后执行 `docker compose up -d`。

### 手动部署

1. 将 `.env.example` 复制为项目根目录的 `.env`。
2. 设置随机的 `SECRET_KEY`、`AES_KEY` 和 `POSTGRES_PASSWORD`（长度至少 32 位，生产环境校验不通过会拒绝启动）。
3. 按实际域名设置 `TRACKING_DOMAIN` 与 `CORS_ORIGINS`。
4. 启动服务：

```bash
docker compose up -d --build
```

首次启动会自动创建管理员账号（默认 `admin` / `admin123`，请通过 `.env` 的 `DEFAULT_ADMIN_*` 配置并登录后修改）。

生产 Compose 使用 PostgreSQL 和 Redis；Nginx 对外提供前端及 API/WebSocket 反向代理。API 的 8000 端口只绑定本机回环地址。

从旧 SQLite 部署升级时，先备份 `backend/data` 或原 Docker 数据卷。Compose 切换到 PostgreSQL 不会自动导入旧 SQLite 数据，应在切换前完成数据迁移并核对任务、发件人和日志数量。

## 验证

```bash
cd backend
python -m compileall -q app
python -m pytest -q

cd ../frontend
npm run build
npm audit --audit-level=low --registry=https://registry.npmjs.org
```

## 使用边界

请仅向已授权或已订阅的收件人发送邮件，并遵守服务商配额、退订要求及适用的反垃圾邮件法规。系统的并发、代理和恢复能力不应用于规避服务商限制。

## 安全与隐私

- 敏感配置（`SECRET_KEY`、`AES_KEY`、`POSTGRES_PASSWORD`、管理员密码）只存在于 `.env`，该文件已被 `.gitignore` 排除，**切勿提交到仓库**。所有推送前请确认 `.env`、日志、数据库文件未进入版本控制。
- 发件人 SMTP 密码与阿里云密钥在生产使用 AES-GCM 加密存储（v3 格式，兼容旧版本）。
- 生产环境建议在外层再加反向代理提供 HTTPS（并同步更新 `TRACKING_DOMAIN` 与 `CORS_ORIGINS`）。

## 环境要求与快速上手

- **部署**：仅需 Docker（含 `docker compose` v2 插件），一条命令即可。
- **开发**：本地需 Python 3.11+ 与 Node.js 18+。

## 目录结构

```
backend/                  # FastAPI 后端
  app/
    api/                  # 路由：认证、发件人、任务、模板、收件人、统计、上传、追踪
    models/               # SQLModel 数据模型
    services/             # 发送引擎、负载均衡、阿里云 DM、SMTP 池
    tasks/                # Celery worker / beat
    websocket/            # WebSocket 实时进度
    utils/                # 安全、SMTP 预设、时间工具
  tests/                  # pytest 测试套件
frontend/                 # Vue 3 + Element Plus + ECharts 前端
  src/views/              # 页面：登录/注册、仪表盘、发送任务、模板、发件人、用户
  src/components/         # 邮件编辑器、收件人上传、计划配置、统计图
  nginx.conf              # 生产 Nginx 反向代理配置（前端 Docker 镜像内使用）
deploy.sh                 # 一键部署脚本
docker-compose.yml        # Redis + PostgreSQL + API + Worker + Beat + Flower + Nginx
.env.example              # 配置模板（复制为 .env）
start-dev.sh / stop-dev.sh# 本地开发启停
```

## 常见问题

- **默认管理员无法登录**：确认 `.env` 中 `DEFAULT_ADMIN_*` 未被改动，且该数据库为空账号表（首次启动才会自动创建）。
- **WebSocket 无法连接（Nginx 生产环境）**：前端的 WebSocket 通过 URL query 参数携带 token，不再依赖 `Sec-WebSocket-Protocol`，已适配 Nginx 反代；如仍异常请检查 Nginx 日志中是否升级失败。
- **改了 `.env` 不生效**：容器内 `.env` 通过 `env_file` 注入，修改后需 `docker compose up -d` 重建（`--build` 仅在有需要时）。
- **端口占用**：API 绑定 `127.0.0.1:8000`（仅本机），Nginx 对外 80 端口；如需改端口请调整 `docker-compose.yml` 的 `ports`。

请仅向已授权或已订阅的收件人发送邮件，并遵守服务商配额、退订要求及适用的反垃圾邮件法规。系统的并发、代理和恢复能力不应用于规避服务商限制。
