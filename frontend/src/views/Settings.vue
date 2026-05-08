<template>
  <div>
    <el-card>
      <template #header>
        <div style="display: flex; justify-content: space-between">
          <span>发件人管理</span>
          <el-button type="primary" @click="showCreateDialog">添加发件人</el-button>
        </div>
      </template>
      <el-table :data="senders" stripe>
        <el-table-column prop="email" label="邮箱" />
        <el-table-column prop="sender_type" label="类型" width="100" />
        <el-table-column prop="smtp_server" label="SMTP服务器" />
        <el-table-column prop="smtp_port" label="端口" width="80" />
        <el-table-column label="权重" width="80">
          <template #default="{ row }">
            <el-tag>{{ row.weight }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="配额" width="100">
          <template #default="{ row }">
            {{ row.daily_sent }} / {{ row.daily_quota }}
          </template>
        </el-table-column>
        <el-table-column label="状态" width="80">
          <template #default="{ row }">
            <el-tag :type="row.enabled ? 'success' : 'danger'">{{ row.enabled ? '启用' : '禁用' }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="200">
          <template #default="{ row }">
            <el-button size="small" @click="testSender(row.id)">测试</el-button>
            <el-button size="small" @click="toggleSender(row.id)">{{ row.enabled ? '禁用' : '启用' }}</el-button>
            <el-button size="small" type="danger" @click="deleteSender(row.id)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <el-dialog v-model="dialogVisible" title="添加发件人" width="500px">
      <el-form :model="senderForm" label-width="100px">
        <el-form-item label="邮箱类型">
          <el-select v-model="senderForm.sender_type" placeholder="选择类型">
            <el-option value="QQ邮箱" label="QQ邮箱" />
            <el-option value="163邮箱" label="163邮箱" />
            <el-option value="Gmail" label="Gmail" />
            <el-option value="Outlook" label="Outlook" />
            <el-option value="自定义SMTP" label="自定义SMTP" />
          </el-select>
        </el-form-item>
        <el-form-item label="发件邮箱">
          <el-input v-model="senderForm.email" />
        </el-form-item>
        <el-form-item label="授权码/密码">
          <el-input v-model="senderForm.password" type="password" show-password />
        </el-form-item>
        <el-form-item label="SMTP服务器">
          <el-input v-model="senderForm.smtp_server" />
        </el-form-item>
        <el-form-item label="端口">
          <el-input-number v-model="senderForm.smtp_port" :min="1" :max="65535" />
        </el-form-item>
        <el-form-item>
          <el-checkbox v-model="senderForm.use_tls">使用TLS</el-checkbox>
        </el-form-item>
        <el-form-item label="权重">
          <el-slider v-model="senderForm.weight" :min="1" :max="100" show-input />
        </el-form-item>
        <el-form-item label="每日配额">
          <el-input-number v-model="senderForm.daily_quota" :min="1" :max="10000" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="createSender">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import api from '../api'
import { ElMessage, ElMessageBox } from 'element-plus'

const senders = ref([])
const dialogVisible = ref(false)

const senderForm = reactive({
  sender_type: 'QQ邮箱',
  email: '',
  password: '',
  smtp_server: 'smtp.qq.com',
  smtp_port: 587,
  use_tls: true,
  weight: 50,
  daily_quota: 500,
})

const loadSenders = async () => {
  try {
    const res = await api.get('/api/senders')
    senders.value = res.data
  } catch {
    ElMessage.error('加载发件人失败')
  }
}

const showCreateDialog = () => {
  senderForm.sender_type = 'QQ邮箱'
  senderForm.email = ''
  senderForm.password = ''
  senderForm.smtp_server = 'smtp.qq.com'
  senderForm.smtp_port = 587
  senderForm.use_tls = true
  senderForm.weight = 50
  senderForm.daily_quota = 500
  dialogVisible.value = true
}

const createSender = async () => {
  try {
    await api.post('/api/senders', senderForm)
    ElMessage.success('发件人添加成功')
    dialogVisible.value = false
    loadSenders()
  } catch {
    ElMessage.error('添加发件人失败')
  }
}

const testSender = async (id) => {
  try {
    const res = await api.post(`/api/senders/${id}/test`)
    if (res.data.success) {
      ElMessage.success('SMTP连接成功')
    } else {
      ElMessage.warning(res.data.message)
    }
  } catch {
    ElMessage.error('测试连接失败')
  }
}

const toggleSender = async (id) => {
  try {
    await api.post(`/api/senders/${id}/toggle`)
    ElMessage.success('状态更新成功')
    loadSenders()
  } catch {
    ElMessage.error('更新失败')
  }
}

const deleteSender = async (id) => {
  try {
    await ElMessageBox.confirm('确定删除此发件人？', '确认')
    await api.delete(`/api/senders/${id}`)
    ElMessage.success('删除成功')
    loadSenders()
  } catch {}
}

onMounted(() => {
  loadSenders()
})
</script>