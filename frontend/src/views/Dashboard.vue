<template>
  <el-container style="min-height: 100vh">
    <el-header style="background: #545c64; display: flex; align-items: center; justify-content: space-between">
      <div style="display: flex; align-items: center; gap: 20px">
        <h3 style="color: #fff; margin: 0">邮箱群发系统 4.0</h3>
        <el-menu :default-active="$route.path" mode="horizontal" background-color="#545c64" text-color="#fff" active-text-color="#ffd04b" @select="handleMenuSelect">
          <el-menu-item index="/dashboard">仪表盘</el-menu-item>
          <el-menu-item index="/send">发送任务</el-menu-item>
          <el-menu-item index="/templates">模板管理</el-menu-item>
          <el-menu-item index="/settings">设置</el-menu-item>
          <el-menu-item v-if="authStore.user?.role === 'admin'" index="/users">用户管理</el-menu-item>
        </el-menu>
      </div>
      <div style="display: flex; align-items: center; gap: 10px; color: #fff">
        <span>{{ authStore.user?.username }}</span>
        <el-button type="text" style="color: #fff" @click="handleLogout">退出</el-button>
      </div>
    </el-header>
    <el-main>
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
import { useRouter } from 'vue-router'
import { useAuthStore } from '../stores/auth'

const router = useRouter()
const authStore = useAuthStore()

authStore.fetchUser()

const handleMenuSelect = (path) => {
  router.push(path)
}

const handleLogout = () => {
  authStore.logout()
  router.push('/login')
}
</script>

<style scoped>
.el-menu--horizontal {
  border-bottom: none;
}
</style>
