# 邮箱群发系统 4.0 - 实施计划 5：Nginx 与集成测试

> **目标：** Nginx 反向代理配置、前后端联调、端到端集成测试

---

## 任务 30：Nginx 配置

**Files:**
- Create: `nginx/nginx.conf`

### 步骤

1. 创建 nginx/nginx.conf（前端静态文件 + API 反向代理 + WebSocket）
2. 提交

---

## 任务 31：前后端联调

**Files:**
- Modify: `frontend/vite.config.js`
- Modify: `frontend/src/api/index.js`

### 步骤

1. 修改 vite.config.js（开发代理到后端）
2. 修改 api/index.js（baseURL 适配）
3. 提交

---

## 任务 32：端到端集成测试

### 步骤

1. 启动完整 Docker Compose 环境
2. 测试登录 → 创建发件人 → 上传导收件人 → 创建发送任务 → 查看进度
3. 测试模板创建与加载
4. 测试用户管理
5. 测试仪表盘数据
6. 修复发现的问题
7. 提交

---

## 任务 33：README 与收尾

**Files:**
- Modify: `README.md`

### 步骤

1. 更新 README.md（项目说明 + 部署指南 + 使用说明）
2. 最终提交