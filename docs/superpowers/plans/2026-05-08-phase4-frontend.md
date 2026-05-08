# 邮箱群发系统 4.0 - 实施计划 4：前端 WebUI

> **目标：** 构建完整的 Vue 3 前端，包括登录、仪表盘、发送向导、模板管理、设置

---

## 任务 22：前端项目初始化

**Files:**
- Create: `frontend/package.json`
- Create: `frontend/vite.config.js`
- Create: `frontend/index.html`
- Create: `frontend/src/main.js`
- Create: `frontend/src/App.vue`
- Create: `frontend/Dockerfile`

### 步骤

1. 初始化 Vue 3 项目（Vite + Vue Router + Pinia + Element Plus + Axios）
2. 创建 frontend/Dockerfile
3. 提交

---

## 任务 23：API 封装与路由

**Files:**
- Create: `frontend/src/api/index.js`
- Create: `frontend/src/router/index.js`
- Create: `frontend/src/stores/auth.js`

### 步骤

1. 创建 frontend/src/api/index.js（Axios 实例 + 拦截器 + 所有 API 调用）
2. 创建 frontend/src/router/index.js（路由 + 导航守卫）
3. 创建 frontend/src/stores/auth.js（认证状态管理）
4. 提交

---

## 任务 24：登录页面

**Files:**
- Create: `frontend/src/views/Login.vue`

### 步骤

1. 创建 Login.vue（用户名/密码表单 + JWT 存储）
2. 提交

---

## 任务 25：布局与仪表盘

**Files:**
- Create: `frontend/src/views/Dashboard.vue`
- Create: `frontend/src/components/StatsChart.vue`

### 步骤

1. 创建 Dashboard.vue（统计卡片 + ECharts 图表）
2. 创建 StatsChart.vue（ECharts 封装组件）
3. 提交

---

## 任务 26：发送任务向导

**Files:**
- Create: `frontend/src/views/SendTask.vue`
- Create: `frontend/src/components/SenderManager.vue`
- Create: `frontend/src/components/RecipientUpload.vue`
- Create: `frontend/src/components/EmailEditor.vue`
- Create: `frontend/src/components/ScheduleConfig.vue`
- Create: `frontend/src/components/TaskProgress.vue`

### 步骤

1. 创建 SenderManager.vue（发件人多选 + 负载均衡策略）
2. 创建 RecipientUpload.vue（拖拽上传 + 预览表格）
3. 创建 EmailEditor.vue（TinyMCE + 变量插入按钮）
4. 创建 ScheduleConfig.vue（立即/定时/智能三种模式）
5. 创建 TaskProgress.vue（WebSocket 实时进度 + 日志）
6. 创建 SendTask.vue（6 步向导整合）
7. 提交

---

## 任务 27：模板管理

**Files:**
- Create: `frontend/src/views/Templates.vue`

### 步骤

1. 创建 Templates.vue（模板列表 + 创建/编辑/复制/删除）
2. 提交

---

## 任务 28：设置页面

**Files:**
- Create: `frontend/src/views/Settings.vue`

### 步骤

1. 创建 Settings.vue（发件人管理 + 负载均衡配置 + 发送参数 + 代理）
2. 提交

---

## 任务 29：用户管理页面

**Files:**
- Create: `frontend/src/views/Users.vue`

### 步骤

1. 创建 Users.vue（管理员：用户列表 + 角色分配）
2. 提交

---

## 验证步骤

```bash
# 前端开发模式
cd frontend && npm run dev
# 浏览器打开 http://localhost:5173

# 或 Docker 模式
docker-compose up --build -d
# 浏览器打开 http://localhost:80

# 测试流程：
# 1. 登录页面 → 输入用户名密码 → 跳转仪表盘
# 2. 仪表盘 → 查看统计数据
# 3. 发送任务 → 6 步向导走完
# 4. 模板管理 → 创建/编辑模板
# 5. 设置 → 添加发件人、配置负载均衡
# 6. 用户管理（管理员）→ 创建操作员账号
```