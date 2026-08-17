<template>
  <div class="send-task">
    <el-steps :active="step" finish-status="success" align-center>
      <el-step title="选择发件人" />
      <el-step title="导入收件人" />
      <el-step title="编辑邮件" />
      <el-step title="附件管理" />
      <el-step title="发送策略" />
      <el-step title="容错设置" />
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
          <template #header>
            <div style="display: flex; justify-content: space-between; align-items: center">
              <span>编辑邮件</span>
              <div style="display: flex; gap: 10px; align-items: center">
                <el-select v-model="selectedTemplateId" placeholder="选择模板填充" clearable style="width: 200px" @change="onTemplateSelect">
                  <el-option v-for="t in templateList" :key="t.id" :label="t.name" :value="t.id" />
                </el-select>
                <el-button @click="showSaveTemplateDialog">保存为模板</el-button>
              </div>
            </div>
          </template>
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

      <!-- Step 6: 容错设置 -->
      <div v-show="step === 5">
        <el-card>
          <template #header>容错设置</template>
          <el-form :model="faultConfig" label-width="180px" label-position="right">
            <el-form-item label="最大重试次数">
              <el-input-number v-model="faultConfig.maxRetries" :min="0" :max="10" controls-position="right" />
            </el-form-item>
            <el-form-item label="重试退避基数 (秒)">
              <el-input-number v-model="faultConfig.retryBackoffBase" :min="1" :max="60" controls-position="right" />
            </el-form-item>
            <el-form-item label="速率限制冷却时间 (秒)">
              <el-input-number v-model="faultConfig.rateLimitCooldown" :min="10" :max="3600" controls-position="right" />
            </el-form-item>
            <el-form-item label="最大连续速率限制次数">
              <el-input-number v-model="faultConfig.maxConsecutiveRateLimits" :min="1" :max="20" controls-position="right" />
            </el-form-item>
            <el-form-item label="冷却后自动恢复发送">
              <el-switch v-model="faultConfig.autoResumeAfterCooldown" />
            </el-form-item>
            <el-form-item label="风控自动暂停任务">
              <el-switch v-model="faultConfig.riskAutoPauseTask" />
            </el-form-item>
            <el-form-item label="风控暂停后等待(秒)">
              <el-input-number v-model="faultConfig.riskPauseSeconds" :min="30" :max="7200" controls-position="right" />
            </el-form-item>
            <el-form-item label="每账号并发数">
              <el-input-number v-model="faultConfig.concurrencyPerSender" :min="1" :max="20" controls-position="right" />
            </el-form-item>
            <el-form-item label="批次大小">
              <el-input-number v-model="faultConfig.batchSize" :min="10" :max="500" controls-position="right" />
            </el-form-item>
            <el-form-item label="速率限制错误匹配模式">
              <el-tag v-for="p in faultConfig.rateLimitPatterns" :key="p" closable @close="removeRateLimitPattern(p)">{{ p }}</el-tag>
              <el-input v-model="newRateLimitPattern" placeholder="添加新模式" style="width: 300px; margin-left: 10px" @keyup.enter="addRateLimitPattern" />
            </el-form-item>
          </el-form>
        </el-card>
      </div>

      <!-- Step 7: 确认发送 -->
      <div v-show="step === 6">
        <el-card>
          <template #header>确认发送</template>
          <el-descriptions :column="2" border>
            <el-descriptions-item label="发件人数">{{ form.senderIds.length }}</el-descriptions-item>
            <el-descriptions-item label="收件人数">{{ recipients.length }}</el-descriptions-item>
            <el-descriptions-item label="邮件主题">{{ form.subject }}</el-descriptions-item>
            <el-descriptions-item label="发送方式">{{ scheduleConfig.scheduleType || 'immediate' }}</el-descriptions-item>
            <el-descriptions-item label="延迟范围">{{ scheduleConfig.delayMin ?? 5 }}-{{ scheduleConfig.delayMax ?? 15 }}秒</el-descriptions-item>
            <el-descriptions-item label="负载策略">{{ form.senderStrategy }}</el-descriptions-item>
            <el-descriptions-item label="最大重试">{{ faultConfig.maxRetries }}</el-descriptions-item>
            <el-descriptions-item label="冷却时间">{{ faultConfig.rateLimitCooldown }}秒</el-descriptions-item>
            <el-descriptions-item label="风控等待">{{ faultConfig.riskPauseSeconds }}秒</el-descriptions-item>
            <el-descriptions-item label="并发数">{{ faultConfig.concurrencyPerSender }}</el-descriptions-item>
            <el-descriptions-item label="自动恢复">{{ faultConfig.autoResumeAfterCooldown ? '是' : '否' }}</el-descriptions-item>
            <el-descriptions-item label="代理数量">{{ (scheduleConfig.proxies || []).length }}</el-descriptions-item>
          </el-descriptions>
          <el-divider />
          <el-button type="primary" size="large" :loading="sending" @click="handleSubmit">确认发送</el-button>
        </el-card>
      </div>
    </div>

    <div class="step-actions">
      <el-button v-if="step > 0" @click="step--">上一步</el-button>
      <el-button v-if="step < 6" type="primary" @click="step++" :disabled="!canNext">下一步</el-button>
    </div>

    <!-- 保存为模板对话框 -->
    <el-dialog v-model="saveTemplateDialogVisible" title="保存为模板" width="400px">
      <el-form label-width="80px">
        <el-form-item label="模板名称">
          <el-input v-model="saveTemplateName" placeholder="请输入模板名称" />
        </el-form-item>
        <el-form-item label="邮件主题">
          <el-input :model-value="form.subject" disabled />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="saveTemplateDialogVisible = false">取消</el-button>
        <el-button type="primary" @click="saveAsTemplate">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, computed, onMounted } from 'vue'
import { useRouter, useRoute } from 'vue-router'
import api from '../api'
import { ElMessage } from 'element-plus'
import SenderManager from '../components/SenderManager.vue'
import RecipientUpload from '../components/RecipientUpload.vue'
import EmailEditor from '../components/EmailEditor.vue'
import ScheduleConfig from '../components/ScheduleConfig.vue'

const router = useRouter()
const route = useRoute()
const step = ref(0)
const sending = ref(false)
const recipients = ref([])
const templateList = ref([])
const selectedTemplateId = ref(null)
const saveTemplateDialogVisible = ref(false)
const saveTemplateName = ref('')

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
  proxies: [],
})

// 故障容错配置
const faultConfig = reactive({
  maxRetries: 3,
  retryBackoffBase: 2,
  rateLimitCooldown: 60,
  maxConsecutiveRateLimits: 5,
  autoResumeAfterCooldown: true,
  riskAutoPauseTask: true,
  riskPauseSeconds: 300,
  concurrencyPerSender: 3,
  batchSize: 100,
  rateLimitPatterns: [
    'Too many attempts',
    'rate limit',
    'spam',
    'blocked',
    'too many',
    'quota',
    'limit exceeded',
  ],
})

const newRateLimitPattern = ref('')

const uploadHeaders = computed(() => {
  const token = localStorage.getItem('access_token') || ''
  return {
    Authorization: token ? `Bearer ${token}` : '',
  }
})

const canNext = computed(() => {
  if (step.value === 0) return form.senderIds.length > 0
  if (step.value === 1) return recipients.value.length > 0
  if (step.value === 2) return form.subject && form.body
  if (step.value === 4) {
    const validDelay = scheduleConfig.delayMin >= 0 && scheduleConfig.delayMax >= scheduleConfig.delayMin
    return validDelay && (scheduleConfig.scheduleType !== 'scheduled' || !!scheduleConfig.scheduleTime)
  }
  if (step.value === 5) return true // 容错设置可选
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
    const smartConfig = {
      max_retries: faultConfig.maxRetries,
      retry_backoff_base: faultConfig.retryBackoffBase,
      rate_limit_cooldown: faultConfig.rateLimitCooldown,
      max_consecutive_rate_limits: faultConfig.maxConsecutiveRateLimits,
      auto_resume_after_cooldown: faultConfig.autoResumeAfterCooldown,
      risk_auto_pause_task: faultConfig.riskAutoPauseTask,
      risk_pause_seconds: faultConfig.riskPauseSeconds,
      concurrency_per_sender: faultConfig.concurrencyPerSender,
      batch_size: faultConfig.batchSize,
      rate_limit_patterns: faultConfig.rateLimitPatterns,
    }

    const res = await api.post('/api/tasks', {
      name: "发送任务-" + Date.now(),
      sender_ids: form.senderIds,
      subject: form.subject,
      body: form.body,
      recipients: recipients.value,
      attachments: form.attachments.map(a => a.response?.path || a.url).filter(Boolean),
      schedule_type: scheduleConfig.scheduleType,
      schedule_time: scheduleConfig.scheduleTime,
      delay_min: scheduleConfig.delayMin,
      delay_max: scheduleConfig.delayMax,
      proxies: scheduleConfig.proxies || [],
      load_balance_strategy: form.senderStrategy,
      smart_config: smartConfig,
    })
    ElMessage.success('任务创建成功')
    const taskId = res.data?.id
    if (taskId) router.push(`/tasks/${taskId}`)
    else router.push('/dashboard')
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '创建任务失败')
  } finally {
    sending.value = false
  }
}

const loadTemplates = async () => {
  try {
    const res = await api.get('/api/templates')
    templateList.value = res.data
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '加载模板失败')
  }
}

const onTemplateSelect = (id) => {
  if (!id) return
  const tpl = templateList.value.find(t => t.id === id)
  if (tpl) {
    form.subject = tpl.subject
    form.body = tpl.body
    ElMessage.success("已加载模板: " + tpl.name)
  }
}

const showSaveTemplateDialog = () => {
  if (!form.subject && !form.body) {
    ElMessage.warning('请先填写邮件主题或正文')
    return
  }
  saveTemplateName.value = ''
  saveTemplateDialogVisible.value = true
}

const saveAsTemplate = async () => {
  if (!saveTemplateName.value.trim()) {
    ElMessage.warning('请输入模板名称')
    return
  }
  try {
    await api.post('/api/templates', {
      name: saveTemplateName.value.trim(),
      subject: form.subject,
      body: form.body,
    })
    ElMessage.success('模板保存成功')
    saveTemplateDialogVisible.value = false
    loadTemplates()
  } catch {
    ElMessage.error('保存模板失败')
  }
}

const addRateLimitPattern = () => {
  const pattern = newRateLimitPattern.value.trim()
  if (!pattern) return
  if (pattern.length > 200 || faultConfig.rateLimitPatterns.length >= 50) {
    ElMessage.warning('匹配模式数量或长度超过限制')
    return
  }
  if (!faultConfig.rateLimitPatterns.includes(pattern)) faultConfig.rateLimitPatterns.push(pattern)
  newRateLimitPattern.value = ''
}

const removeRateLimitPattern = (pattern) => {
  if (faultConfig.rateLimitPatterns.length <= 1) {
    ElMessage.warning('至少保留一个匹配模式')
    return
  }
  const idx = faultConfig.rateLimitPatterns.indexOf(pattern)
  if (idx > -1) faultConfig.rateLimitPatterns.splice(idx, 1)
}

onMounted(async () => {
  await loadTemplates()
  // If navigated from template "use" button, pre-fill subject and body
  if (route.query.template_id) {
    const tpl = templateList.value.find(t => t.id === Number(route.query.template_id))
    if (tpl) {
      form.subject = tpl.subject
      form.body = tpl.body
      selectedTemplateId.value = tpl.id
      step.value = 2
      ElMessage.success("已加载模板: " + tpl.name)
    }
  } else if (route.query.subject || route.query.body) {
    if (route.query.subject) form.subject = route.query.subject
    if (route.query.body) form.body = route.query.body
    step.value = 2
  }
})
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
