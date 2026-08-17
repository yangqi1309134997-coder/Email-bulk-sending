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
        <h2>注册账号</h2>
        <p class="sub">创建操作员账号后即可使用</p>
        <el-form :model="form" :rules="rules" ref="formRef" @submit.prevent="handleRegister">
          <el-form-item prop="username">
            <el-input v-model="form.username" placeholder="用户名（3-64 位，字母数字_.-）" prefix-icon="User" size="large" />
          </el-form-item>
          <el-form-item prop="email">
            <el-input v-model="form.email" placeholder="邮箱" prefix-icon="Message" size="large" />
          </el-form-item>
          <el-form-item prop="password">
            <el-input v-model="form.password" type="password" placeholder="密码（至少 10 位）" prefix-icon="Lock" show-password size="large" />
          </el-form-item>
          <el-form-item prop="confirm">
            <el-input v-model="form.confirm" type="password" placeholder="确认密码" prefix-icon="Lock" show-password size="large" />
          </el-form-item>
          <el-form-item>
            <el-button type="primary" native-type="submit" :loading="loading" size="large" style="width: 100%">
              注册
            </el-button>
          </el-form-item>
        </el-form>
        <p class="switch-link">已有账号？<router-link to="/login">去登录</router-link></p>
      </el-card>
    </main>
  </div>
</template>

<script setup>
import { ref, reactive } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import api from '../api'

const router = useRouter()
const formRef = ref()
const loading = ref(false)

const form = reactive({
  username: '',
  email: '',
  password: '',
  confirm: '',
})

const rules = {
  username: [
    { required: true, message: '请输入用户名', trigger: 'blur' },
    { min: 3, max: 64, message: '长度 3-64 位', trigger: 'blur' },
    { pattern: /^[A-Za-z0-9_.-]+$/, message: '仅支持字母、数字、_ . -', trigger: 'blur' },
  ],
  email: [
    { required: true, message: '请输入邮箱', trigger: 'blur' },
    { type: 'email', message: '邮箱格式不正确', trigger: 'blur' },
  ],
  password: [
    { required: true, message: '请输入密码', trigger: 'blur' },
    { min: 10, max: 128, message: '长度至少 10 位', trigger: 'blur' },
  ],
  confirm: [
    { required: true, message: '请再次输入密码', trigger: 'blur' },
    {
      validator: (rule, value, callback) => {
        if (value !== form.password) callback(new Error('两次输入的密码不一致'))
        else callback()
      },
      trigger: 'blur',
    },
  ],
}

const handleRegister = async () => {
  const valid = await formRef.value.validate().catch(() => false)
  if (!valid) return

  loading.value = true
  try {
    await api.post('/api/auth/register', {
      username: form.username,
      email: form.email,
      password: form.password,
    })
    ElMessage.success('注册成功，请登录')
    router.push('/login')
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '注册失败')
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
  background: #f2f4f7;
}
.login-shell {
  width: min(420px, 100%);
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
