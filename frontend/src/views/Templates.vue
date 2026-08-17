<template>
  <div>
    <el-card>
      <template #header>
        <div style="display: flex; justify-content: space-between">
          <span>模板管理</span>
          <el-button type="primary" @click="showCreateDialog">创建模板</el-button>
        </div>
      </template>
      <el-table :data="templates" stripe>
        <el-table-column prop="name" label="模板名称" />
        <el-table-column prop="subject" label="邮件主题" />
        <el-table-column prop="updated_at" label="更新时间">
          <template #default="{ row }">
            {{ formatTime(row.updated_at) }}
          </template>
        </el-table-column>
        <el-table-column label="操作" width="300">
          <template #default="{ row }">
            <el-button size="small" type="primary" @click="useTemplate(row)">使用</el-button>
            <el-button size="small" @click="editTemplate(row)">编辑</el-button>
            <el-button size="small" @click="duplicateTemplate(row.id)">复制</el-button>
            <el-button size="small" type="danger" @click="deleteTemplate(row.id)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <el-dialog v-model="dialogVisible" :title="isEdit ? '编辑模板' : '创建模板'" width="700px">
      <el-form :model="templateForm" label-width="80px">
        <el-form-item label="模板名称">
          <el-input v-model="templateForm.name" placeholder="请输入模板名称" />
        </el-form-item>
        <el-form-item label="邮件主题">
          <el-input v-model="templateForm.subject" placeholder="支持 {name}, {email} 变量" />
        </el-form-item>
        <el-form-item label="邮件正文">
          <EmailEditor v-model="templateForm.body" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="saveTemplate">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import api from '../api'
import { ElMessage, ElMessageBox } from 'element-plus'
import EmailEditor from '../components/EmailEditor.vue'
import { formatApiDateTime } from '../utils/time'

const router = useRouter()
const templates = ref([])
const dialogVisible = ref(false)
const isEdit = ref(false)
const editId = ref(null)

const templateForm = reactive({
  name: '',
  subject: '',
  body: '',
})

const formatTime = formatApiDateTime

const loadTemplates = async () => {
  try {
    const res = await api.get('/api/templates')
    templates.value = res.data
  } catch {
    ElMessage.error('加载模板失败')
  }
}

const showCreateDialog = () => {
  isEdit.value = false
  editId.value = null
  templateForm.name = ''
  templateForm.subject = ''
  templateForm.body = ''
  dialogVisible.value = true
}

const editTemplate = (row) => {
  isEdit.value = true
  editId.value = row.id
  templateForm.name = row.name
  templateForm.subject = row.subject
  templateForm.body = row.body
  dialogVisible.value = true
}

const saveTemplate = async () => {
  if (!templateForm.name.trim()) {
    ElMessage.warning('请输入模板名称')
    return
  }
  try {
    if (isEdit.value) {
      await api.put(`/api/templates/${editId.value}`, templateForm)
      ElMessage.success('模板更新成功')
    } else {
      await api.post('/api/templates', templateForm)
      ElMessage.success('模板创建成功')
    }
    dialogVisible.value = false
    loadTemplates()
  } catch {
    ElMessage.error('保存模板失败')
  }
}

const useTemplate = (row) => {
  // Keep potentially large/sensitive HTML out of browser history and URLs.
  router.push({ path: '/send', query: { template_id: row.id } })
}

const duplicateTemplate = async (id) => {
  try {
    await api.post(`/api/templates/${id}/duplicate`)
    ElMessage.success('模板复制成功')
    loadTemplates()
  } catch {
    ElMessage.error('复制模板失败')
  }
}

const deleteTemplate = async (id) => {
  try {
    await ElMessageBox.confirm('确定删除此模板？', '确认')
    await api.delete(`/api/templates/${id}`)
    ElMessage.success('模板删除成功')
    loadTemplates()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '删除模板失败')
  }
}

onMounted(() => {
  loadTemplates()
})
</script>
