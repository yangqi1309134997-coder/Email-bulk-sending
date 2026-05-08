# 邮箱群发系统 4.0 - 实施计划 3：发送引擎与任务队列

> **目标：** 实现 Celery 任务队列、负载均衡、SMTP 发送引擎、WebSocket 实时推送

---

## 任务 12：Celery 配置

**Files:**
- Create: `backend/app/tasks/__init__.py`
- Create: `backend/app/tasks/celery_app.py`

### 步骤

1. 创建 backend/app/tasks/celery_app.py（Celery 实例配置）
2. 创建 backend/app/tasks/__init__.py
3. 提交

---

## 任务 13：负载均衡器

**Files:**
- Create: `backend/app/services/__init__.py`
- Create: `backend/app/services/load_balancer.py`

### 步骤

1. 创建 backend/app/services/load_balancer.py（三种策略 + 智能降级）
2. 创建 backend/app/services/__init__.py
3. 提交

---

## 任务 14：SMTP 发送引擎

**Files:**
- Create: `backend/app/services/email_sender.py`

### 步骤

1. 创建 backend/app/services/email_sender.py（构建邮件 + SMTP 发送 + 重试）
2. 提交

---

## 任务 15：追踪服务

**Files:**
- Create: `backend/app/services/tracker.py`
- Create: `backend/app/api/tracking.py`

### 步骤

1. 创建 backend/app/services/tracker.py（追踪像素注入 + 链接替换）
2. 创建 backend/app/api/tracker.py（/track/open + /track/click 端点）
3. 提交

---

## 任务 16：Celery 发送任务

**Files:**
- Create: `backend/app/tasks/send_email.py`

### 步骤

1. 创建 backend/app/tasks/send_email.py（批量发送任务 + 实时推送）
2. 提交

---

## 任务 17：WebSocket 管理

**Files:**
- Create: `backend/app/websocket/__init__.py`
- Create: `backend/app/websocket/manager.py`

### 步骤

1. 创建 backend/app/websocket/manager.py（连接管理 + 广播）
2. 创建 backend/app/websocket/__init__.py
3. 提交

---

## 任务 18：发送任务 API

**Files:**
- Create: `backend/app/api/tasks.py`

### 步骤

1. 创建 backend/app/api/tasks.py（创建/列表/详情/暂停/恢复/取消/日志/导出）
2. 提交

---

## 任务 19：仪表盘 API

**Files:**
- Create: `backend/app/api/dashboard.py`

### 步骤

1. 创建 backend/app/api/dashboard.py（stats + realtime）
2. 提交

---

## 任务 20：文件上传 API

**Files:**
- Create: `backend/app/api/upload.py`

### 步骤

1. 创建 backend/app/api/upload.py（attachment + image 上传）
2. 提交

---

## 任务 21：智能调度器

**Files:**
- Create: `backend/app/services/scheduler.py`

### 步骤

1. 创建 backend/app/services/scheduler.py（时区推断 + 分时段调度）
2. 提交

---

## 验证步骤

```bash
# 重启所有服务
docker-compose down && docker-compose up --build -d

# 扩展 worker 数量
docker-compose up --scale worker=3 -d

# 测试完整发送流程
# 1. 登录获取 token
# 2. 创建发件人
# 3. 上传导收件人文件
# 4. 创建发送任务
# 5. 观察 Worker 日志
docker-compose logs -f worker

# 6. 检查 Flower 监控
# 浏览器打开 http://localhost:5555
```