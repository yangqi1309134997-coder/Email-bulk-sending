<template>
  <div class="send-task">
    <el-steps :active="step" finish-status="success" align-center>
      <el-step title="选择发件人" />
      <el-step title="导入收件人" />
      <el-step title="编辑邮件" />
      <el-step title="附件管理" />
      <el-step title="发送策略" />
      <el-step title="确认发送" />
    </el-steps>

    <div class="step-content">
      <!-- Step 1: 发件人选择 -->
      <div v-show="step === 0">
        <el-card>
          <template #header>选择发件人</template>
          <SenderManager @update="onSenderUpdate" />
        </el-card>
      </div>

      <!-- Step 2: 收件人导入 -->
      <div v-show="step === 1">
        <el-card>
          <template #header>导入收件人</template>
          <RecipientUpload @update:recipients="onRecipientsUpdate" />
        </el-card>
      </div>

      <!-- Step 3: 邮件编辑 -->
      <div v-show="step === 2">
        <el-card>
          <template #header>编辑邮件</template>
          <el-form label-width="80px">
            <el-form-item label="邮件主题">
              <el-input v-model="form.subject" placeholder="支持 {name}, {email} 变量" />
            </el-form-item>
            <el-form-item label="邮件正文">
              <EmailEditor v-model="form.body" />
            </el-form-item>
          </el-form>
        </el-card>
      </div>

      <!-- Step 4: 附件管理 -->
      <div v-show="step === 3">
        <el-card>
          <template #header>附件管理</template>
          <el-upload v-model:file-list="form.attachments" action="/api/upload/attachment" multiple :headers="uploadHeaders">
            <el-button type="primary">上传附件</el-button>
          </el-upload>
        </el-card>
      </div>

      <!-- Step 5: 发送策略 -->
      <div v-show="step === 4">
        <el-card>
          <template #header>发送策略</template>
          <ScheduleConfig @update="onScheduleUpdate" />
        </el-card>
      </div>

      <!-- Step 6: 确认发送 -->
      <div v-show="step === 5">
        <el-card>
          <template #header>确认发送</template>
          <el-descriptions :column="2" border>
            <el-descriptions-item label="发件人数">{{ form.senderIds.length }}</el-descriptions-item>
            <el-descriptions-item label="收件人数">{{ recipients.length }}</el-descriptions-item>
            <el-descriptions-item label="邮件主题">{{ form.subject }}</el-descriptions-item>
            <el-descriptions-item label="发送方式">{{ scheduleConfig.scheduleType || 'immediate' }}</el-descriptions-item>
            <el-descriptions-item label="延迟范围">{{ scheduleConfig.delayMin || 5 }}-{{ scheduleConfig.delayMax || 15 }}秒</el-descriptions-item>
            <el-descriptions-item label="负载策略">{{ form.senderStrategy }}</el-descriptions-item>
          </el-descriptions>
          <el-divider />
          <el-button type="primary" size="large" :loading="sending" @click="handleSubmit">确认发送</el-button>
        </el-card>
      </div>
    </div>

    <div class="step-actions">
      <el-button v-if="step > 0" @click="step--">上一步</el-button>
      <el-button v-if="step < 5" type="primary" @click="step++" :disabled="!canNext">下一步</el-button>
    </div>
  </div>
</template>

<script setup>
import { ref, reactive, computed } from 'vue'
import { useRouter } from 'vue-router'
import api from '../api'
import { ElMessage } from 'element-plus'
import SenderManager from '../components/SenderManager.vue'
import RecipientUpload from '../components/RecipientUpload.vue'
import EmailEditor from '../components/EmailEditor.vue'
import ScheduleConfig from '../components/ScheduleConfig.vue'

const router = useRouter()
const step = ref(0)
const sending = ref(false)
const recipients = ref([])

const form = reactive({
  senderStrategy: 'round_robin',
  senderIds: [],
  subject: '',
  body: '',
  attachments: [],
})

const scheduleConfig = reactive({
  scheduleType: 'immediate',
  scheduleTime: null,
  delayMin: 5,
  delayMax: 15,
})

const uploadHeaders = computed(() => ({
  Authorization: `Bearer ${localStorage.getItem('access_token')}`
}))

const canNext = computed(() => {
  if (step.value === 0) return form.senderIds.length > 0
  if (step.value === 1) return recipients.value.length > 0
  if (step.value === 2) return form.subject && form.body
  return true
})

const onSenderUpdate = ({ strategy, senderIds }) => {
  form.senderStrategy = strategy
  form.senderIds = senderIds
}

const onRecipientsUpdate = (list) => {
  recipients.value = list
}

const onScheduleUpdate = (config) => {
  Object.assign(scheduleConfig, config)
}

const handleSubmit = async () => {
  sending.value = true
  try {
    await api.post('/api/tasks', {
      name: `发送任务-${Date.now()}`,
      sender_ids: form.senderIds,
      subject: form.subject,
      body: form.body,
      recipients: recipients.value,
      attachments: form.attachments.map(a => a.response?.path || a.url).filter(Boolean),
      schedule_type: scheduleConfig.scheduleType,
      schedule_time: scheduleConfig.scheduleTime,
      delay_min: scheduleConfig.delayMin,
      delay_max: scheduleConfig.delayMax,
      load_balance_strategy: form.senderStrategy,
    })
    ElMessage.success('任务创建成功')
    router.push('/dashboard')
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '创建任务失败')
  } finally {
    sending.value = false
  }
}
</script>

<style scoped>
.send-task {
  max-width: 900px;
  margin: 0 auto;
}
.step-content {
  margin: 40px 0;
}
.step-actions {
  display: flex;
  justify-content: center;
  gap: 20px;
}
</style>
