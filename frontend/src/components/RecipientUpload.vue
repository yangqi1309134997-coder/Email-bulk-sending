<template>
  <div class="recipient-upload">
    <el-tabs v-model="inputMode">
      <el-tab-pane label="手动输入" name="manual">
        <el-input
          v-model="manualInput"
          type="textarea"
          :rows="4"
          placeholder="每行一个邮箱，或 邮箱,姓名 格式"
        />
        <el-button type="primary" style="margin-top: 10px" @click="parseManualInput">解析</el-button>
      </el-tab-pane>
      <el-tab-pane label="文件上传" name="file">
        <el-upload drag multiple :auto-upload="false" :on-change="handleFileChange" :file-list="fileList" :accept="'.txt,.csv,.xlsx,.xls'">
          <el-icon class="el-icon--upload"><Upload /></el-icon>
          <div class="el-upload__text">拖拽文件到此处或点击上传</div>
          <template #tip>
            <div class="el-upload__tip">支持 .txt, .csv, .xlsx 格式，每行一个邮箱或 邮箱,姓名；重复邮箱自动去重</div>
          </template>
        </el-upload>
      </el-tab-pane>
    </el-tabs>

    <el-divider v-if="recipients.length > 0" />

    <div v-if="recipients.length > 0" class="preview-section">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px">
        <h4 style="margin: 0">预览 (前10条)</h4>
        <el-tag type="success">共 {{ recipients.length }} 个有效收件人</el-tag>
      </div>
      <el-table :data="recipients.slice(0, 10)" max-height="300" size="small">
        <el-table-column type="index" width="50" label="#" />
        <el-table-column prop="email" label="邮箱" />
        <el-table-column prop="name" label="姓名" />
      </el-table>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import api from '../api'
import { ElMessage } from 'element-plus'
import { Upload } from '@element-plus/icons-vue'

const recipients = ref([])
const inputMode = ref('manual')
const manualInput = ref('')
const fileList = ref([])

const emit = defineEmits(['update:recipients'])

const isValidEmail = (email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)

const parseManualInput = () => {
  const lines = manualInput.value.trim().split('\n')
  const result = []
  const seen = new Set()
  for (const line of lines) {
    const trimmed = line.trim()
    if (!trimmed) continue
    const parts = trimmed.split(',', 2)
    const email = parts[0].trim().toLowerCase()
    const name = parts.length > 1 ? parts[1].trim() : ''
    if (isValidEmail(email) && !seen.has(email)) {
      seen.add(email)
      result.push({ email, name })
    }
  }
  if (result.length === 0) {
    ElMessage.warning('未找到有效邮箱地址')
    return
  }
  recipients.value = result
  emit('update:recipients', result)
  ElMessage.success(`解析成功，共 ${result.length} 个收件人`)
}

const handleFileChange = async (file, fileListData) => {
  fileList.value = fileListData
  const formData = new FormData()
  formData.append('file', file.raw)
  try {
    const res = await api.post('/api/recipients/parse', formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    })
    const parsed = res.data.recipients || []
    const seen = new Set()
    const unique = parsed.filter((item) => {
      const email = String(item.email || '').trim().toLowerCase()
      if (!email || seen.has(email)) return false
      seen.add(email)
      return true
    })
    recipients.value = unique
    emit('update:recipients', unique)
    if (res.data.invalid > 0) {
      ElMessage.warning(`解析完成：有效 ${unique.length} 条，无效 ${res.data.invalid} 条（已跳过）`)
    } else {
      ElMessage.success(`解析成功，共 ${unique.length} 个收件人`)
    }
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '解析文件失败')
  } finally {
    // 移除已处理的文件，避免重复解析同一文件
    fileList.value = []
  }
}

const reset = () => {
  recipients.value = []
  manualInput.value = ''
  fileList.value = []
  emit('update:recipients', [])
}

defineExpose({ reset })
</script>

<style scoped>
.recipient-upload {
  width: 100%;
}
.preview-section {
  margin-top: 10px;
}
</style>
