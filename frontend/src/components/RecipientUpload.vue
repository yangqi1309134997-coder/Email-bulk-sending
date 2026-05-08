<template>
  <div class="recipient-upload">
    <el-upload drag multiple :auto-upload="false" :on-change="handleFileChange" :accept="'.txt,.csv,.xlsx,.xls'">
      <el-icon class="el-icon--upload"><Upload /></el-icon>
      <div class="el-upload__text">拖拽文件到此处或点击上传</div>
      <template #tip>
        <div class="el-upload__tip">支持 .txt, .csv, .xlsx 格式，每行一个邮箱或 邮箱,姓名</div>
      </template>
    </el-upload>

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

const emit = defineEmits(['update:recipients'])

const handleFileChange = async (file) => {
  const formData = new FormData()
  formData.append('file', file.raw)
  try {
    const res = await api.post('/api/recipients/parse', formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    })
    recipients.value = res.data.recipients
    emit('update:recipients', res.data.recipients)
    ElMessage.success(`解析成功，共 ${res.data.total} 个收件人`)
  } catch (err) {
    ElMessage.error('解析文件失败')
  }
}

const reset = () => {
  recipients.value = []
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
