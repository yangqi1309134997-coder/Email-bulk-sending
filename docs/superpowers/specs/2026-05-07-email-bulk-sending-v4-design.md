# 邮箱群发系统 4.0 设计文档

## 概述

将现有的 tkinter 桌面版邮件群发工具升级为基于 Web 的企业级群发系统，支持十万+封/天的大规模发送，具备分布式任务队列、多策略负载均衡、富文本 WebUI 编辑器和完整的发送监控能力。

## 技术栈

| 层级 | 技术 |
|------|------|
| 前端 | Vue 3 + Element Plus + TinyMCE + ECharts |
| 后端 API | FastAPI + SQLModel + Pydantic |
| 任务队列 | Celery + Redis |
| 数据库 | SQLite（可迁移 PostgreSQL） |
| 容器化 | Docker Compose + Nginx |
| 认证 | JWT（access 2h + refresh 7d） |

## 系统架构

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│   Vue.js     │────▶│   FastAPI     │────▶│   Redis     │
│   WebUI      │◀────│   API 层      │◀────│   消息队列   │
└─────────────┘     └──────┬───────┘     └──────┬──────┘
       │                   │                     │
       │            ┌──────▼───────┐     ┌──────▼──────┐
       │            │   Celery      │◀────│  Flower     │
       │            │   Workers xN  │     │  任务监控    │
       │            └──────┬───────┘     └─────────────┘
       │                   │
       │            ┌──────▼───────┐
       │            │   SMTP 发送   │
       │            │   引擎        │
       │            └──────────────┘
       │
 ┌─────▼──────┐
 │  SQLite     │
 │  用户/配置   │
 └────────────┘
```

### 核心组件

| 组件 | 职责 | 技术 |
|------|------|------|
| WebUI | 用户界面、邮件编辑、配置管理 | Vue 3 + Element Plus + TinyMCE |
| API 层 | REST API、WebSocket 实时推送、用户认证 | FastAPI + JWT |
| 任务队列 | 发送任务调度、重试、负载均衡 | Celery + Redis |
| Workers | 实际 SMTP 发送，多实例并行 | Celery Worker 容器 |
| 数据存储 | 用户、发件人配置、发送记录 | SQLite |
| 任务监控 | Worker 状态、任务进度可视化 | Flower |

## 负载均衡

### 策略

1. **轮询（Round Robin）**：发件人依次轮流
2. **权重分配（Weighted）**：按每个发件人设置的 weight 比例分配
3. **智能模式（Smart）**：综合考虑配额余量（40%）和成功率（60%）动态打分

### 智能降级与恢复

- 连续 3 次失败 → 暂停该发件人 5 分钟
- 暂停到期后试探性发送 1 封，成功则恢复
- 配额耗尽 → 自动跳过，切换下一个可用发件人

### 发送引擎核心流程

1. 用户创建发送任务
2. 根据 schedule_type 决定投递方式（immediate/scheduled/smart）
3. Celery Worker 取任务
4. 负载均衡器选择发件人
5. 构建邮件（变量替换 + 附件 + 追踪像素）
6. SMTP 发送（重试机制：3 次，指数退避 2s/4s/8s）
7. 重试时自动切换发件人
8. WebSocket 推送实时进度

## WebUI 功能模块

### 仪表盘

- 今日发送量 / 成功率 / 打开率 / 点击率
- 实时发送进度 + WebSocket 日志流
- Worker 状态 / 队列积压

### 发送任务向导（6 步）

1. 选择发件人（多选 + 负载均衡策略）
2. 导入收件人（CSV/Excel/TXT 拖拽上传）
3. 编辑邮件（TinyMCE 富文本 + 变量插入）
4. 附件管理（多文件上传）
5. 发送策略（立即/定时/智能分时段）
6. 确认发送

### 模板管理

- 模板列表（缩略图预览）
- 创建/编辑/复制/删除
- 变量管理（{name}, {email}, 自定义变量）

### 设置

- 发件人管理（CRUD + 连接测试）
- 负载均衡配置（策略选择 + 权重设置）
- 发送参数（延迟范围、重试次数、并发数）
- 代理配置

### 用户管理（管理员）

- 用户列表 + 角色分配（admin/operator）

### 邮件编辑器

- TinyMCE 富文本：粗体/斜体/颜色/字体/表格/图片/链接
- 变量插入按钮：一键插入 {name}、{email} 等
- HTML 源码编辑模式切换
- 实时预览（变量替换后的效果）

## 数据模型

### users

| 字段 | 类型 | 说明 |
|------|------|------|
| id | PK | 主键 |
| username | str（唯一） | 用户名 |
| password_hash | str | bcrypt 加密 |
| role | str | admin / operator |
| email | str | 邮箱 |
| created_at | datetime | 创建时间 |
| last_login | datetime | 最后登录 |

### senders

| 字段 | 类型 | 说明 |
|------|------|------|
| id | PK | 主键 |
| user_id | FK | 所属用户 |
| email | str | 发件邮箱 |
| password | str | AES 加密存储 |
| smtp_server | str | SMTP 服务器 |
| smtp_port | int | 端口 |
| use_tls | bool | 是否 TLS |
| sender_type | str | QQ/163/Gmail/自定义 |
| enabled | bool | 是否启用 |
| weight | int | 权重（1-100） |
| daily_quota | int | 每日配额 |
| daily_sent | int | 今日已发 |
| success_rate | float | 成功率 |
| status | str | active/paused/banned |
| created_at | datetime | |
| updated_at | datetime | |

### send_tasks

| 字段 | 类型 | 说明 |
|------|------|------|
| id | PK | 主键 |
| user_id | FK | 所属用户 |
| name | str | 任务名称 |
| status | str | pending/running/paused/completed/cancelled |
| sender_ids | JSON | 发件人列表 |
| recipient_count | int | 收件人总数 |
| success_count | int | 成功数 |
| fail_count | int | 失败数 |
| open_count | int | 打开数 |
| click_count | int | 点击数 |
| schedule_type | str | immediate/scheduled/smart |
| schedule_time | datetime | 计划发送时间 |
| smart_config | JSON | 智能发送配置 |
| created_at | datetime | |
| completed_at | datetime | |

### templates

| 字段 | 类型 | 说明 |
|------|------|------|
| id | PK | 主键 |
| user_id | FK | 所属用户 |
| name | str | 模板名称 |
| subject | str | 邮件主题 |
| body | str | 邮件正文（HTML） |
| variables | JSON | 支持的变量 |
| created_at | datetime | |
| updated_at | datetime | |

### send_logs

| 字段 | 类型 | 说明 |
|------|------|------|
| id | PK | 主键 |
| task_id | FK | 任务 ID |
| sender_id | FK | 发件人 ID |
| recipient_email | str | 收件人 |
| recipient_name | str | 收件人姓名 |
| subject | str | 邮件主题 |
| status | str | success/failed/bounced |
| error_message | str | 错误信息 |
| sent_at | datetime | 发送时间 |
| opened_at | datetime | 打开时间 |
| clicked_at | datetime | 点击时间 |

## API 路由

```
认证
POST   /api/auth/login
POST   /api/auth/register
GET    /api/auth/me

发件人
GET    /api/senders
POST   /api/senders
PUT    /api/senders/{id}
DELETE /api/senders/{id}
POST   /api/senders/{id}/test
POST   /api/senders/{id}/toggle

发送任务
POST   /api/tasks
GET    /api/tasks
GET    /api/tasks/{id}
POST   /api/tasks/{id}/pause
POST   /api/tasks/{id}/resume
POST   /api/tasks/{id}/cancel
GET    /api/tasks/{id}/logs
GET    /api/tasks/{id}/export

模板
GET    /api/templates
POST   /api/templates
PUT    /api/templates/{id}
DELETE /api/templates/{id}
POST   /api/templates/{id}/duplicate

收件人
POST   /api/recipients/parse
POST   /api/recipients/validate

追踪
GET    /track/open/{log_id}
GET    /track/click/{log_id}

仪表盘
GET    /api/dashboard/stats
GET    /api/dashboard/realtime

WebSocket
WS     /ws/tasks/{id}

用户管理（管理员）
GET    /api/users
PUT    /api/users/{id}
DELETE /api/users/{id}

文件上传
POST   /api/upload/attachment
POST   /api/upload/image
```

## 智能分时段发送

- 根据收件人邮箱域名推断时区（.cn → UTC+8 等）
- 按时区分组，每组在当地时间 9:00-11:00 / 14:00-16:00 发送
- 自动控制发送速率，避免触发 SMTP 限流

## 打开率/点击率追踪

- 打开追踪：邮件底部 1px 追踪像素 `<img src="{domain}/track/open/{log_id}">`
- 点击追踪：链接替换为 `{domain}/track/click/{log_id}?url=原始URL`，302 跳转
- log_id 使用 UUID，不可猜测
- 点击追踪 URL 做域名白名单校验，防止开放重定向

## 安全设计

### 认证与授权

- JWT Token（access 2h + refresh 7d）
- admin 可管理用户和所有任务，operator 只能操作自己的资源
- 所有 API 依赖注入校验身份和权限

### 数据安全

- 发件人密码 AES 加密存储，密钥从环境变量读取
- 用户密码 bcrypt（salt rounds=12）
- JWT 密钥从环境变量读取
- .env.example 提供模板，.env 加入 .gitignore

### API 安全

- 文件上传：类型白名单 + 大小上限（附件 25MB，图片 5MB）
- 速率限制：登录 5次/分钟，任务创建 10次/分钟
- CORS 仅允许前端域名
- Pydantic 模型严格校验

## 错误处理

### API 层

- 统一异常处理中间件：`{"code": 400, "message": "错误描述", "detail": null}`
- 401/403/404/422/500 标准状态码

### 发送引擎

- SMTP 连接超时（30s）→ 重试
- SMTP 认证失败 → 标记发件人 paused，通知用户
- SMTP 限流（552/554）→ 自动降速
- 网络中断 → Celery 自动重派
- 附件读取失败 → 跳过该附件，记录警告

### Worker 容错

- Worker 崩溃 → Celery 自动重派
- 单封邮件超时（60s）→ 强制终止，标记失败
- Redis 断开 → Worker 自动重连，任务不丢失

## 项目目录结构

```
email-bulk-sending/
├── docker-compose.yml
├── .env.example
├── README.md
├── backend/
│   ├── Dockerfile
│   ├── requirements.txt
│   ├── alembic/
│   │   └── versions/
│   ├── app/
│   │   ├── main.py
│   │   ├── config.py
│   │   ├── database.py
│   │   ├── models/
│   │   │   ├── user.py
│   │   │   ├── sender.py
│   │   │   ├── task.py
│   │   │   ├── template.py
│   │   │   └── send_log.py
│   │   ├── api/
│   │   │   ├── auth.py
│   │   │   ├── senders.py
│   │   │   ├── tasks.py
│   │   │   ├── templates.py
│   │   │   ├── recipients.py
│   │   │   ├── tracking.py
│   │   │   ├── dashboard.py
│   │   │   ├── users.py
│   │   │   └── upload.py
│   │   ├── services/
│   │   │   ├── email_sender.py
│   │   │   ├── load_balancer.py
│   │   │   ├── scheduler.py
│   │   │   └── tracker.py
│   │   ├── tasks/
│   │   │   ├── celery_app.py
│   │   │   └── send_email.py
│   │   ├── websocket/
│   │   │   └── manager.py
│   │   └── utils/
│   │       ├── security.py
│   │       └── file_parser.py
│   └── uploads/
├── frontend/
│   ├── Dockerfile
│   ├── package.json
│   ├── vite.config.js
│   ├── src/
│   │   ├── main.js
│   │   ├── App.vue
│   │   ├── router/
│   │   │   └── index.js
│   │   ├── stores/
│   │   │   ├── auth.js
│   │   │   ├── task.js
│   │   │   └── sender.js
│   │   ├── api/
│   │   │   └── index.js
│   │   ├── views/
│   │   │   ├── Login.vue
│   │   │   ├── Dashboard.vue
│   │   │   ├── SendTask.vue
│   │   │   ├── Templates.vue
│   │   │   ├── Settings.vue
│   │   │   └── Users.vue
│   │   └── components/
│   │       ├── EmailEditor.vue
│   │       ├── RecipientUpload.vue
│   │       ├── SenderManager.vue
│   │       ├── ScheduleConfig.vue
│   │       ├── TaskProgress.vue
│   │       └── StatsChart.vue
│   └── public/
└── nginx/
    └── nginx.conf
```

## Docker Compose

```yaml
services:
  api:        # FastAPI
  worker:     # Celery Worker（docker compose up --scale worker=N）
  flower:     # 任务监控
  redis:      # 消息队列
  nginx:      # 反向代理 + 前端静态文件
```

SQLite 数据文件挂载 volume，无需额外数据库容器。
