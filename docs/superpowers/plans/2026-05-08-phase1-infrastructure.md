# 邮箱群发系统 4.0 - 实施计划 1：基础设施与后端核心

> **目标：** 搭建项目结构、数据库模型、用户认证系统、基础配置

---

## 任务 1：项目初始化

**Files:**
- Create: `docker-compose.yml`
- Create: `.env.example`
- Create: `backend/requirements.txt`
- Create: `backend/Dockerfile`
- Create: `backend/app/__init__.py`

### 步骤

1. 创建 docker-compose.yml（包含 redis/api/worker/flower/nginx 服务）
2. 创建 .env.example（环境变量模板）
3. 创建 backend/requirements.txt（Python 依赖）
4. 创建 backend/Dockerfile
5. 创建 backend/app/__init__.py
6. 提交

---

## 任务 2：配置与数据库

**Files:**
- Create: `backend/app/config.py`
- Create: `backend/app/database.py`
- Create: `backend/app/models/__init__.py`

### 步骤

1. 创建 backend/app/config.py（Settings 配置类）
2. 创建 backend/app/database.py（SQLModel 引擎和会话）
3. 创建 backend/app/models/__init__.py
4. 提交

---

## 任务 3：用户模型与认证工具

**Files:**
- Create: `backend/app/models/user.py`
- Create: `backend/app/utils/security.py`
- Create: `backend/app/api/deps.py`

### 步骤

1. 创建 backend/app/models/user.py（User 模型）
2. 创建 backend/app/utils/security.py（JWT/AES/密码工具）
3. 创建 backend/app/api/deps.py（依赖注入：get_current_user, require_admin）
4. 提交

---

## 任务 4：认证 API

**Files:**
- Create: `backend/app/api/auth.py`
- Modify: `backend/app/main.py`

### 步骤

1. 创建 backend/app/api/auth.py（login/register/refresh/me）
2. 创建 backend/app/utils/__init__.py
3. 创建 backend/app/api/__init__.py
4. 修改 backend/app/main.py（FastAPI 入口、CORS、路由注册）
5. 提交

---

## 任务 5：发件人模型

**Files:**
- Create: `backend/app/models/sender.py`

### 步骤

1. 创建 backend/app/models/sender.py（Sender 模型，包含 is_available 方法）
2. 提交

---

## 任务 6：创建初始 .env 文件

**Files:**
- Create: `backend/.env`

### 步骤

1. 基于 .env.example 创建 backend/.env，填入实际密钥
2. 提交

---

## 验证步骤

完成所有任务后，运行以下验证：

```bash
cd E:/Users/huancheng/Desktop/code/Email-bulk-sending

# 1. 检查 Docker 是否运行
docker ps

# 2. 构建并启动服务
docker-compose up --build -d

# 3. 检查 API 是否启动
curl http://localhost:8000/health

# 4. 访问 API 文档
# 浏览器打开 http://localhost:8000/docs
```

预期结果：
- health 接口返回 `{"status": "ok"}`
- 可以访问 Swagger 文档