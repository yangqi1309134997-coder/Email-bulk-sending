<template>
  <div>
    <el-card>
      <template #header>
        <div style="display: flex; justify-content: space-between">
          <span>用户管理</span>
          <el-button type="primary" @click="showCreateDialog">创建用户</el-button>
        </div>
      </template>
      <el-table :data="users" stripe>
        <el-table-column prop="username" label="用户名" />
        <el-table-column prop="email" label="邮箱" />
        <el-table-column prop="role" label="角色" width="100">
          <template #default="{ row }">
            <el-tag :type="row.role === 'admin' ? 'danger' : 'info'">{{ row.role === 'admin' ? '管理员' : '操作员' }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="created_at" label="创建时间" />
        <el-table-column prop="last_login" label="最后登录" />
        <el-table-column label="操作" width="200">
          <template #default="{ row }">
            <el-button size="small" @click="changeRole(row)">{{ row.role === 'admin' ? '降为操作员' : '升为管理员' }}</el-button>
            <el-button size="small" type="danger" @click="deleteUser(row.id)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <el-dialog v-model="dialogVisible" title="创建用户" width="400px">
      <el-form :model="userForm" label-width="80px">
        <el-form-item label="用户名">
          <el-input v-model="userForm.username" />
        </el-form-item>
        <el-form-item label="密码">
          <el-input v-model="userForm.password" type="password" show-password />
        </el-form-item>
        <el-form-item label="邮箱">
          <el-input v-model="userForm.email" />
        </el-form-item>
        <el-form-item label="角色">
          <el-radio-group v-model="userForm.role">
            <el-radio value="admin">管理员</el-radio>
            <el-radio value="operator">操作员</el-radio>
          </el-radio-group>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="createUser">创建</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import api from '../api'
import { ElMessage, ElMessageBox } from 'element-plus'

const users = ref([])
const dialogVisible = ref(false)

const userForm = reactive({
  username: '',
  password: '',
  email: '',
  role: 'operator',
})

const loadUsers = async () => {
  try {
    const res = await api.get('/api/users')
    users.value = res.data
  } catch {
    ElMessage.error('加载用户失败')
  }
}

const showCreateDialog = () => {
  userForm.username = ''
  userForm.password = ''
  userForm.email = ''
  userForm.role = 'operator'
  dialogVisible.value = true
}

const createUser = async () => {
  try {
    await api.post('/api/users', userForm)
    ElMessage.success('用户创建成功')
    dialogVisible.value = false
    loadUsers()
  } catch {
    ElMessage.error('创建用户失败')
  }
}

const changeRole = async (user) => {
  const newRole = user.role === 'admin' ? 'operator' : 'admin'
  try {
    await api.put(`/api/users/${user.id}`, { role: newRole })
    ElMessage.success('角色更新成功')
    loadUsers()
  } catch {
    ElMessage.error('更新角色失败')
  }
}

const deleteUser = async (id) => {
  try {
    await ElMessageBox.confirm('确定删除此用户？', '确认')
    await api.delete(`/api/users/${id}`)
    ElMessage.success('删除成功')
    loadUsers()
  } catch {}
}

onMounted(() => {
  loadUsers()
})
</script>