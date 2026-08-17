<template>
  <div class="login-page">
    <main class="login-shell">
      <header class="brand-row">
        <div class="brand-mark">EB</div>
        <div>
          <h1>邮箱群发</h1>
          <p>Enterprise 4.0</p>
        </div>
      </header>
      <el-card class="login-card" shadow="never">
        <h2>登录工作台</h2>
        <p class="sub">使用系统账号继续</p>
        <el-form :model="form" :rules="rules" ref="formRef" @submit.prevent="handleLogin">
          <el-form-item prop="username">
            <el-input v-model="form.username" placeholder="用户名" prefix-icon="User" size="large" />
          </el-form-item>
          <el-form-item prop="password">
            <el-input v-model="form.password" type="password" placeholder="密码" prefix-icon="Lock" show-password size="large" />
          </el-form-item>
          <el-form-item>
            <el-button type="primary" native-type="submit" :loading="loading" size="large" style="width: 100%">
              登录
            </el-button>
          </el-form-item>
        </el-form>
        <p class="switch-link">还没有账号？<router-link to="/register">立即注册</router-link></p>
      </el-card>
    </main>
  </div>
</template>

<script setup>
import { ref, reactive } from 'vue'
import { useRouter } from 'vue-router'
import { useAuthStore } from '../stores/auth'
import { ElMessage } from 'element-plus'

const router = useRouter()
const authStore = useAuthStore()
const formRef = ref()
const loading = ref(false)

const form = reactive({
  username: '',
  password: '',
})

const rules = {
  username: [{ required: true, message: '请输入用户名', trigger: 'blur' }],
  password: [{ required: true, message: '请输入密码', trigger: 'blur' }],
}

const handleLogin = async () => {
  const valid = await formRef.value.validate().catch(() => false)
  if (!valid) return

  loading.value = true
  try {
    await authStore.login(form.username, form.password)
    ElMessage.success('登录成功')
    router.push('/dashboard')
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '登录失败')
  } finally {
    loading.value = false
  }
}
</script>

<style scoped>
.login-page {
  min-height: 100vh;
  display: grid;
  place-items: center;
  padding: 24px;
  position: relative;
  background: #f2f4f7;
}
.login-shell {
  width: min(420px, 100%);
  position: relative;
}
.brand-row {
  display: flex;
  align-items: center;
  gap: 14px;
  margin: 0 0 20px 2px;
  color: #111827;
}
.brand-mark {
  display: grid;
  place-items: center;
  width: 42px;
  height: 42px;
  border-radius: 8px;
  background: #2494e8;
  font-weight: 700;
}
.brand-row h1 {
  margin: 0;
  font-size: 22px;
}
.brand-row p {
  margin: 2px 0 0;
  color: #667085;
  font-size: 12px;
}
.login-card {
  border-radius: 8px !important;
  border: 1px solid #d8dde6;
  padding: 18px 10px 8px;
  box-shadow: 0 12px 30px rgba(17, 24, 39, 0.12) !important;
}
.login-card h2 {
  margin: 0;
  font-size: 24px;
}
.sub {
  margin: 8px 0 22px;
  color: #6b7280;
}
.switch-link {
  margin: 14px 0 6px;
  text-align: center;
  color: #6b7280;
  font-size: 14px;
}
.switch-link a {
  color: #2494e8;
  text-decoration: none;
}
@media (max-width: 480px) {
  .login-page { padding: 18px; }
  .login-card { padding: 12px 4px 4px; }
}
</style>
