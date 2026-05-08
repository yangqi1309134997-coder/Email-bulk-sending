# 邮箱群发系统 4.0 - 实施计划 2：任务、模板与日志模型

> **目标：** 完成发送任务、模板、发送日志数据模型，以及基础 API

---

## 任务 7：任务与模板模型

**Files:**
- Create: `backend/app/models/task.py`
- Create: `backend/app/models/template.py`
- Create: `backend/app/models/send_log.py`

### 步骤

1. 创建 backend/app/models/task.py（SendTask 模型）
2. 创建 backend/app/models/template.py（Template 模型）
3. 创建 backend/app/models/send_log.py（SendLog 模型）
4. 提交

---

## 任务 8：发件人 API

**Files:**
- Create: `backend/app/api/senders.py`

### 步骤

1. 创建 backend/app/api/senders.py（CRUD + test + toggle）
2. 提交

---

## 任务 9：模板 API

**Files:**
- Create: `backend/app/api/templates.py`

### 步骤

1. 创建 backend/app/api/templates.py（CRUD + duplicate）
2. 提交

---

## 任务 10：用户管理 API

**Files:**
- Create: `backend/app/api/users.py`

### 步骤

1. 创建 backend/app/api/users.py（admin 管理用户）
2. 提交

---

## 任务 11：收件人解析 API

**Files:**
- Create: `backend/app/utils/file_parser.py`
- Create: `backend/app/api/recipients.py`

### 步骤

1. 创建 backend/app/utils/file_parser.py（解析 CSV/Excel/TXT）
2. 创建 backend/app/api/recipients.py（parse + validate）
3. 提交

---

## 验证步骤

```bash
# 重启服务
docker-compose down && docker-compose up --build -d

# 测试发件人 API（需要先登录获取 token）
# 1. 注册用户
curl -X POST http://localhost:8000/api/auth/register \
  -H "Content-Type: application/json" \
  -d '{"username":"admin","password":"admin123","email":"admin@test.com","role":"admin"}'

# 2. 登录
curl -X POST http://localhost:8000/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"username":"admin","password":"admin123"}'
# 获取 access_token

# 3. 创建发件人（使用上一步的 token）
curl -X POST http://localhost:8000/api/senders \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"email":"test@qq.com","password":"authcode","smtp_server":"smtp.qq.com","smtp_port":587,"use_tls":true,"sender_type":"QQ邮箱","enabled":true,"weight":50,"daily_quota":500}'

# 4. 查看发件人列表
curl -X GET http://localhost:8000/api/senders \
  -H "Authorization: Bearer YOUR_TOKEN"
```

预期结果：API 正常响应，数据写入 SQLite