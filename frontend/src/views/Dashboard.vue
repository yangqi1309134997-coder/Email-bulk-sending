<template>
  <el-container class="app-shell">
    <el-aside width="232px" class="sidebar">
      <div class="brand">
        <div class="brand-mark">EB</div>
        <div>
          <div class="brand-title">邮箱群发</div>
          <div class="brand-sub">Enterprise 4.0</div>
        </div>
      </div>
      <el-menu
        :default-active="$route.path.startsWith('/tasks/') ? '/dashboard' : $route.path"
        class="side-menu"
        @select="handleMenuSelect"
      >
        <el-menu-item index="/dashboard">
          <el-icon><Odometer /></el-icon>
          <span>仪表盘</span>
        </el-menu-item>
        <el-menu-item index="/send">
          <el-icon><Promotion /></el-icon>
          <span>发送任务</span>
        </el-menu-item>
        <el-menu-item index="/templates">
          <el-icon><Document /></el-icon>
          <span>模板管理</span>
        </el-menu-item>
        <el-menu-item index="/settings">
          <el-icon><Setting /></el-icon>
          <span>发件人设置</span>
        </el-menu-item>
        <el-menu-item v-if="authStore.user?.role === 'admin'" index="/users">
          <el-icon><User /></el-icon>
          <span>用户管理</span>
        </el-menu-item>
      </el-menu>
      <div class="sidebar-footer">
        <div class="user-chip">
          <el-avatar :size="32">{{ avatarText }}</el-avatar>
          <div>
            <div class="user-name">{{ authStore.user?.username || '用户' }}</div>
            <div class="user-role">{{ authStore.user?.role === 'admin' ? '管理员' : '操作员' }}</div>
          </div>
        </div>
        <el-button class="logout-btn" @click="handleLogout">退出登录</el-button>
      </div>
    </el-aside>

    <el-container>
      <el-header class="topbar">
        <div>
          <div class="page-title">{{ pageTitle }}</div>
          <div class="page-desc">{{ pageDesc }}</div>
        </div>
        <div class="top-actions">
          <el-button type="primary" @click="$router.push('/send')">
            <el-icon><Plus /></el-icon>
            <span class="new-send-label">新建发送</span>
          </el-button>
        </div>
      </el-header>
      <el-main class="main-area">
        <router-view />
      </el-main>
    </el-container>
  </el-container>
</template>

<script setup>
import { computed, onMounted } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { useAuthStore } from '../stores/auth'

const router = useRouter()
const route = useRoute()
const authStore = useAuthStore()

onMounted(() => {
  authStore.fetchUser()
})

const avatarText = computed(() => (authStore.user?.username || 'U').slice(0, 1).toUpperCase())

const pageMeta = {
  '/dashboard': ['仪表盘', '总览发送表现、任务状态与系统负载'],
  '/send': ['发送任务', '配置发件人、收件人、策略并启动群发'],
  '/templates': ['模板管理', '复用邮件主题与正文模板'],
  '/settings': ['发件人设置', '管理 SMTP / 阿里云推送账号与历史配置'],
  '/users': ['用户管理', '管理系统账号与权限'],
}

const pageTitle = computed(() => {
  if (route.path.startsWith('/tasks/')) return '任务详情'
  return pageMeta[route.path]?.[0] || '工作台'
})
const pageDesc = computed(() => {
  if (route.path.startsWith('/tasks/')) return '实时查看发送进度、日志与风控状态'
  return pageMeta[route.path]?.[1] || ''
})

const handleMenuSelect = (path) => router.push(path)
const handleLogout = () => {
  authStore.logout()
  router.push('/login')
}

// 路由切换时回到页面顶部，避免在移动端停留在上一页的滚动位置
router.afterEach(() => {
  const main = document.querySelector('.main-area')
  if (main) main.scrollTop = 0
  window.scrollTo(0, 0)
})
</script>

<style scoped>
.app-shell {
  min-height: 100vh;
  background: var(--app-bg);
}
.sidebar {
  background: #0f172a;
  color: #fff;
  display: flex;
  flex-direction: column;
  border-right: 1px solid rgba(255,255,255,0.06);
}
.brand {
  display: flex;
  gap: 12px;
  align-items: center;
  padding: 22px 18px 16px;
}
.brand-mark {
  width: 40px;
  height: 40px;
  border-radius: 12px;
  display: grid;
  place-items: center;
  font-weight: 700;
  background: linear-gradient(135deg, #3b82f6, #06b6d4);
}
.brand-title { font-size: 16px; font-weight: 700; }
.brand-sub { font-size: 12px; color: #94a3b8; margin-top: 2px; }
.side-menu {
  border-right: none;
  background: transparent;
  flex: 1;
  padding: 8px 10px;
}
.side-menu :deep(.el-menu-item) {
  border-radius: 10px;
  margin-bottom: 6px;
  color: #cbd5e1;
}
.side-menu :deep(.el-menu-item.is-active) {
  background: rgba(59, 130, 246, 0.18) !important;
  color: #fff !important;
}
.side-menu :deep(.el-menu-item:hover) {
  background: rgba(148, 163, 184, 0.12) !important;
}
.sidebar-footer {
  padding: 14px;
  border-top: 1px solid rgba(255,255,255,0.08);
}
.user-chip {
  display: flex;
  gap: 10px;
  align-items: center;
  margin-bottom: 12px;
}
.user-name { font-size: 14px; font-weight: 600; }
.user-role { font-size: 12px; color: #94a3b8; }
.logout-btn { width: 100%; }
.topbar {
  height: auto !important;
  min-height: 84px;
  background: transparent;
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 18px 24px 8px;
}
.page-title {
  font-size: 24px;
  font-weight: 700;
  letter-spacing: -0.02em;
}
.page-desc {
  margin-top: 4px;
  color: var(--app-muted);
  font-size: 13px;
}
.top-actions { display: flex; gap: 10px; }
.main-area {
  padding: 8px 24px 28px;
}
@media (max-width: 640px) {
  .new-send-label { display: none; }
}
@media (max-width: 960px) {
  .sidebar { width: 72px !important; }
  .brand-title, .brand-sub, .side-menu span, .user-chip, .logout-btn { display: none; }
}
@media (max-width: 640px) {
  .sidebar { width: 0 !important; display: none; }
  .topbar { padding: 14px 16px 8px; }
  .page-title { font-size: 20px; }
  .page-desc { display: none; }
  .main-area { padding: 8px 12px 20px; }
}
</style>
