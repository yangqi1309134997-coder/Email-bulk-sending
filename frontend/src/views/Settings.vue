<template>
  <div>
    <el-tabs v-model="activeTab" type="card">
      <el-tab-pane label="发件人管理" name="senders">
        <el-card>
          <template #header>
            <div style="display: flex; justify-content: space-between">
              <span>发件人管理</span>
              <div style="display: flex; gap: 10px">
                <el-button type="primary" @click="showCreateDialog">添加发件人</el-button>
              <el-button @click="showCreateTemplateDialog">配置模板</el-button>
              </div>
            </div>
          </template>
          <el-table :data="senders" stripe class="responsive-table">
            <el-table-column prop="email" label="邮箱" />
            <el-table-column prop="sender_type" label="类型" width="120" />
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
            <el-table-column label="状态" width="110">
              <template #default="{ row }">
                <el-tag v-if="row.status === 'banned'" type="danger">封禁</el-tag>
                <el-tag v-else-if="row.status === 'paused'" type="warning">冷却中</el-tag>
                <el-tag v-else-if="row.enabled" type="success">启用</el-tag>
                <el-tag v-else type="info">禁用</el-tag>
              </template>
            </el-table-column>
            <el-table-column label="操作" width="230">
              <template #default="{ row }">
                <el-button size="small" @click="testSender(row.id)">测试</el-button>
                <el-button size="small" @click="toggleSender(row.id)">{{ row.enabled ? '禁用' : '启用' }}</el-button>
                <el-button v-if="row.status === 'paused'" size="small" type="success" @click="unpauseSender(row.id)">恢复</el-button>
                <el-button size="small" type="danger" @click="deleteSender(row.id)">删除</el-button>
              </template>
            </el-table-column>
          </el-table>
        </el-card>
      </el-tab-pane>

      <el-tab-pane label="发件人模板" name="templates">
        <el-card>
          <template #header>
            <div style="display: flex; justify-content: space-between">
              <span>发件人配置模板</span>
              <el-button type="primary" @click="showCreateTemplateDialog">保存为模板</el-button>
            </div>
          </template>
          <el-table :data="templates" stripe class="responsive-table">
            <el-table-column prop="name" label="模板名称" />
            <el-table-column prop="description" label="描述" />
            <el-table-column prop="sender_type" label="发件人类型" width="120" />
            <el-table-column prop="smtp_server" label="SMTP服务器" />
            <el-table-column prop="smtp_port" label="端口" width="80" />
            <el-table-column label="权重" width="80">
              <template #default="{ row }">
                <el-tag>{{ row.weight }}</el-tag>
              </template>
            </el-table-column>
            <el-table-column label="每日配额" width="100">
              <template #default="{ row }">
                {{ row.daily_quota }}
              </template>
            </el-table-column>
            <el-table-column label="操作" width="150">
              <template #default="{ row }">
                <el-button size="small" @click="applyTemplate(row)">应用</el-button>
                <el-button size="small" type="danger" @click="deleteTemplate(row.id)">删除</el-button>
              </template>
            </el-table-column>
          </el-table>
        </el-card>
      </el-tab-pane>
    </el-tabs>

    <!-- 添加发件人对话框 -->
    <el-dialog v-model="dialogVisible" title="添加发件人" width="550px">
      <el-form :model="senderForm" label-width="120px">
        <el-form-item label="邮箱类型">
          <el-select v-model="senderForm.sender_type" placeholder="选择类型" filterable @change="onSenderTypeChange">
            <el-option
              v-for="preset in presetOptions"
              :key="preset.key"
              :value="preset.sender_type"
              :label="preset.name"
            />
          </el-select>
        </el-form-item>

        <!-- 阿里云邮箱推送专属配置 -->
        <template v-if="senderForm.sender_type === '阿里云邮箱推送'">
          <el-form-item label="发信地址">
            <el-input v-model="senderForm.email" placeholder="例如: noreply@yourdomain.com" />
          </el-form-item>
          <el-form-item label="AccessKey ID">
            <el-input v-model="senderForm.aliyun_access_key" placeholder="阿里云 AccessKey ID" />
          </el-form-item>
          <el-form-item label="AccessKey Secret">
            <el-input v-model="senderForm.aliyun_access_secret" type="password" show-password placeholder="阿里云 AccessKey Secret" />
          </el-form-item>
          <el-form-item label="区域">
            <el-select v-model="senderForm.aliyun_region" placeholder="选择区域">
              <el-option value="cn-hangzhou" label="华东1（杭州）" />
              <el-option value="cn-shanghai" label="华东2（上海）" />
              <el-option value="cn-beijing" label="华北2（北京）" />
              <el-option value="cn-shenzhen" label="华南1（深圳）" />
              <el-option value="ap-southeast-1" label="亚太东南1（新加坡）" />
            </el-select>
          </el-form-item>
          <el-form-item label="发信人昵称">
            <el-input v-model="senderForm.aliyun_from_name" placeholder="收件人看到的发件人名称" />
          </el-form-item>
        </template>

        <!-- 常规 SMTP 配置 -->
        <template v-else>
          <el-form-item label="发件邮箱">
            <el-input v-model="senderForm.email" :placeholder="emailPlaceholder" />
          </el-form-item>
          <el-form-item label="授权码/密码">
            <el-input v-model="senderForm.password" type="password" show-password :placeholder="passwordPlaceholder" />
          </el-form-item>
          <el-form-item label="SMTP服务器">
            <el-input v-model="senderForm.smtp_server" :disabled="senderForm.sender_type !== '自定义SMTP'" />
          </el-form-item>
          <el-form-item label="SMTP用户名">
            <el-input v-model="senderForm.smtp_username" placeholder="默认使用发件邮箱，可填中继账号" />
          </el-form-item>
          <el-form-item label="端口">
            <el-input-number v-model="senderForm.smtp_port" :min="1" :max="65535" :disabled="senderForm.sender_type !== '自定义SMTP'" />
          </el-form-item>
          <el-form-item>
            <el-checkbox v-model="senderForm.use_tls" :disabled="senderForm.sender_type !== '自定义SMTP'">使用TLS</el-checkbox>
          </el-form-item>
          <el-form-item label="安全模式">
            <el-select v-model="senderForm.smtp_security">
              <el-option value="" label="自动（按端口判断）" />
              <el-option value="ssl" label="SSL/TLS（465）" />
              <el-option value="starttls" label="STARTTLS（587）" />
              <el-option value="none" label="无加密（仅内网）" />
            </el-select>
          </el-form-item>
        </template>

        <el-form-item label="权重">
          <el-slider v-model="senderForm.weight" :min="1" :max="100" show-input />
        </el-form-item>
        <el-form-item label="每日配额">
          <el-input-number v-model="senderForm.daily_quota" :min="0" :max="100000" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="createSender">保存</el-button>
      </template>
    </el-dialog>

    <!-- 创建模板对话框 -->
    <el-dialog v-model="templateDialogVisible" title="保存为发件人模板" width="500px">
      <el-form :model="templateForm" label-width="100px">
        <el-form-item label="模板名称">
          <el-input v-model="templateForm.name" placeholder="请输入模板名称" />
        </el-form-item>
        <el-form-item label="描述">
          <el-input v-model="templateForm.description" placeholder="请输入描述" />
        </el-form-item>
        <el-form-item label="发件人类型">
          <el-select v-model="templateForm.sender_type" placeholder="选择类型" filterable @change="onTemplateTypeChange">
            <el-option
              v-for="preset in presetOptions"
              :key="`template-${preset.key}`"
              :value="preset.sender_type"
              :label="preset.name"
            />
          </el-select>
        </el-form-item>

        <template v-if="templateForm.sender_type === '阿里云邮箱推送'">
          <el-form-item label="AccessKey ID">
            <el-input v-model="templateForm.aliyun_access_key" placeholder="阿里云 AccessKey ID" />
          </el-form-item>
          <el-form-item label="AccessKey Secret">
            <el-input v-model="templateForm.aliyun_access_secret" type="password" show-password placeholder="阿里云 AccessKey Secret" />
          </el-form-item>
          <el-form-item label="区域">
            <el-select v-model="templateForm.aliyun_region" placeholder="选择区域">
              <el-option value="cn-hangzhou" label="华东1（杭州）" />
              <el-option value="cn-shanghai" label="华东2（上海）" />
              <el-option value="cn-beijing" label="华北2（北京）" />
              <el-option value="cn-shenzhen" label="华南1（深圳）" />
              <el-option value="ap-southeast-1" label="亚太东南1（新加坡）" />
            </el-select>
          </el-form-item>
          <el-form-item label="发信人昵称">
            <el-input v-model="templateForm.aliyun_from_name" placeholder="收件人看到的发件人名称" />
          </el-form-item>
        </template>

        <template v-else>
          <el-form-item label="SMTP服务器">
            <el-input v-model="templateForm.smtp_server" :disabled="templateForm.sender_type !== '自定义SMTP'" />
          </el-form-item>
          <el-form-item label="SMTP用户名">
            <el-input v-model="templateForm.smtp_username" placeholder="可选中继登录名" />
          </el-form-item>
          <el-form-item label="端口">
            <el-input-number v-model="templateForm.smtp_port" :min="1" :max="65535" :disabled="templateForm.sender_type !== '自定义SMTP'" />
          </el-form-item>
          <el-form-item>
            <el-checkbox v-model="templateForm.use_tls" :disabled="templateForm.sender_type !== '自定义SMTP'">使用TLS</el-checkbox>
          </el-form-item>
          <el-form-item label="安全模式">
            <el-select v-model="templateForm.smtp_security">
              <el-option value="" label="自动（按端口判断）" />
              <el-option value="ssl" label="SSL/TLS（465）" />
              <el-option value="starttls" label="STARTTLS（587）" />
              <el-option value="none" label="无加密（仅内网）" />
            </el-select>
          </el-form-item>
        </template>

        <el-form-item label="权重">
          <el-slider v-model="templateForm.weight" :min="1" :max="100" show-input />
        </el-form-item>
        <el-form-item label="每日配额">
          <el-input-number v-model="templateForm.daily_quota" :min="0" :max="100000" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="templateDialogVisible = false">取消</el-button>
        <el-button type="primary" @click="createTemplate">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, computed, onMounted } from 'vue'
import api from '../api'
import { ElMessage, ElMessageBox } from 'element-plus'

const activeTab = ref('senders')
const senders = ref([])
const templates = ref([])
const dialogVisible = ref(false)
const templateDialogVisible = ref(false)
const presetOptions = ref([])

const SMTP_PRESETS = {
  'QQ邮箱': { smtp_server: 'smtp.qq.com', smtp_port: 587, use_tls: true, daily_quota: 200 },
  '163邮箱': { smtp_server: 'smtp.163.com', smtp_port: 465, use_tls: true, daily_quota: 300 },
  '126邮箱': { smtp_server: 'smtp.126.com', smtp_port: 465, use_tls: true, daily_quota: 300 },
  'yeah邮箱': { smtp_server: 'smtp.yeah.net', smtp_port: 465, use_tls: true, daily_quota: 300 },
  '新浪邮箱': { smtp_server: 'smtp.sina.com', smtp_port: 587, use_tls: true, daily_quota: 200 },
  '搜狐邮箱': { smtp_server: 'smtp.sohu.com', smtp_port: 465, use_tls: true, daily_quota: 200 },
  '139邮箱': { smtp_server: 'smtp.139.com', smtp_port: 465, use_tls: true, daily_quota: 200 },
  '189邮箱': { smtp_server: 'smtp.189.cn', smtp_port: 465, use_tls: true, daily_quota: 200 },
  'AOL': { smtp_server: 'smtp.aol.com', smtp_port: 587, use_tls: true, daily_quota: 500 },
  'Fastmail': { smtp_server: 'smtp.fastmail.com', smtp_port: 587, use_tls: true, daily_quota: 2000 },
  'Yandex': { smtp_server: 'smtp.yandex.com', smtp_port: 465, use_tls: true, daily_quota: 500 },
  'Gmail': { smtp_server: 'smtp.gmail.com', smtp_port: 587, use_tls: true, daily_quota: 500 },
  'Outlook': { smtp_server: 'smtp-mail.outlook.com', smtp_port: 587, use_tls: true, daily_quota: 300 },
  'Yahoo': { smtp_server: 'smtp.mail.yahoo.com', smtp_port: 587, use_tls: true, daily_quota: 500 },
  'iCloud': { smtp_server: 'smtp.mail.me.com', smtp_port: 587, use_tls: true, daily_quota: 1000 },
  'Zoho': { smtp_server: 'smtp.zoho.com', smtp_port: 587, use_tls: true, daily_quota: 500 },
  '阿里企业邮箱': { smtp_server: 'smtp.mxhichina.com', smtp_port: 465, use_tls: true, daily_quota: 2000 },
  '腾讯企业邮箱': { smtp_server: 'smtp.exmail.qq.com', smtp_port: 465, use_tls: true, daily_quota: 2000 },
  '华为企业邮箱': { smtp_server: 'smtp.sparkspace.huaweicloud.com', smtp_port: 465, use_tls: true, daily_quota: 2000 },
  '网易企业邮箱': { smtp_server: 'smtp.qiye.163.com', smtp_port: 465, use_tls: true, daily_quota: 2000 },
  '飞书邮箱': { smtp_server: 'smtp.feishu.cn', smtp_port: 465, use_tls: true, daily_quota: 1000 },
  '钉钉邮箱': { smtp_server: 'smtp.qiye.aliyun.com', smtp_port: 465, use_tls: true, daily_quota: 1000 },
  '阿里云邮箱推送SMTP': { smtp_server: 'smtpdm.aliyun.com', smtp_port: 465, use_tls: true, daily_quota: 10000 },
  'SendGrid': { smtp_server: 'smtp.sendgrid.net', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'Mailgun': { smtp_server: 'smtp.mailgun.org', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'Amazon SES': { smtp_server: 'email-smtp.us-east-1.amazonaws.com', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'Postmark': { smtp_server: 'smtp.postmarkapp.com', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'SparkPost': { smtp_server: 'smtp.sparkpostmail.com', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'Brevo': { smtp_server: 'smtp-relay.brevo.com', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'Elastic Email': { smtp_server: 'smtp.elasticemail.com', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  'Mailjet': { smtp_server: 'in-v3.mailjet.com', smtp_port: 587, use_tls: true, daily_quota: 100000 },
  '自定义SMTP': { smtp_server: '', smtp_port: 587, use_tls: true, daily_quota: 500 },
}

const senderForm = reactive({
  sender_type: 'QQ邮箱',
  email: '',
  password: '',
  smtp_server: 'smtp.qq.com',
  smtp_port: 587,
  use_tls: true,
  smtp_username: '',
  smtp_security: '',
  weight: 50,
  daily_quota: 500,
  aliyun_access_key: '',
  aliyun_access_secret: '',
  aliyun_region: 'cn-hangzhou',
  aliyun_from_name: '',
})

const templateForm = reactive({
  name: '',
  description: '',
  sender_type: 'QQ邮箱',
  smtp_server: 'smtp.qq.com',
  smtp_port: 587,
  use_tls: true,
  smtp_username: '',
  smtp_security: '',
  weight: 50,
  daily_quota: 500,
  aliyun_access_key: '',
  aliyun_access_secret: '',
  aliyun_region: 'cn-hangzhou',
  aliyun_from_name: '',
})

const emailPlaceholder = computed(() => {
  const t = senderForm.sender_type
  if (t === 'QQ邮箱') return '例如: 123456789@qq.com'
  if (t === '163邮箱') return '例如: yourname@163.com'
  if (t === '126邮箱') return '例如: yourname@126.com'
  if (t === 'Gmail') return '例如: yourname@gmail.com'
  if (t === 'Outlook') return '例如: yourname@outlook.com'
  return '请输入邮箱地址'
})

const passwordPlaceholder = computed(() => {
  const t = senderForm.sender_type
  if (['QQ邮箱', '163邮箱', '126邮箱', 'yeah邮箱', '新浪邮箱', '搜狐邮箱', '腾讯企业邮箱'].includes(t)) {
    return '请输入授权码（非登录密码）'
  }
  return '请输入邮箱密码或应用专用密码'
})

const fallbackPresetOptions = Object.entries(SMTP_PRESETS).map(([sender_type, preset], index) => ({
  key: `fallback-${index}-${sender_type}`,
  sender_type,
  name: sender_type,
  ...preset,
}))

const presetForType = (type) => (
  presetOptions.value.find((preset) => preset.sender_type === type)
  || fallbackPresetOptions.find((preset) => preset.sender_type === type)
)

const normalizePreset = (preset) => ({
  ...preset,
  // The API uses compact names (server/port/tls), while the form model uses
  // explicit SMTP names. Normalize once so selecting a provider never clears
  // the server or port after presets are loaded.
  smtp_server: preset.smtp_server ?? preset.server ?? '',
  smtp_port: preset.smtp_port ?? preset.port ?? 587,
  use_tls: preset.use_tls ?? preset.tls ?? true,
  security: preset.security ?? '',
  daily_quota: preset.daily_quota ?? preset.daily_limit ?? 500,
})

const loadPresets = async () => {
  try {
    const res = await api.get('/api/senders/presets')
    const unique = new Map()
    for (const preset of res.data || []) {
      if (!preset.sender_type) continue
      const normalized = normalizePreset(preset)
      const existing = unique.get(preset.sender_type)
      const securityRank = { starttls: 3, ssl: 2, none: 0, '': 1 }
      if (!existing || (securityRank[normalized.security] ?? 1) > (securityRank[existing.security] ?? 1)) {
        unique.set(preset.sender_type, normalized)
      }
    }
    presetOptions.value = unique.size ? [...unique.values()] : fallbackPresetOptions
  } catch {
    presetOptions.value = fallbackPresetOptions
  }
}

const loadSenders = async () => {
  try {
    const res = await api.get('/api/senders')
    senders.value = res.data
  } catch {
    ElMessage.error('加载发件人失败')
  }
}

const loadTemplates = async () => {
  try {
    const res = await api.get('/api/senders/templates')
    templates.value = res.data
  } catch {
    ElMessage.error('加载模板失败')
  }
}

const showCreateDialog = () => {
  senderForm.sender_type = 'QQ邮箱'
  senderForm.email = ''
  senderForm.password = ''
  senderForm.smtp_server = 'smtp.qq.com'
  senderForm.smtp_port = 587
  senderForm.use_tls = true
  senderForm.smtp_username = ''
  senderForm.smtp_security = ''
  senderForm.weight = 50
  senderForm.daily_quota = 500
  senderForm.aliyun_access_key = ''
  senderForm.aliyun_access_secret = ''
  senderForm.aliyun_region = 'cn-hangzhou'
  senderForm.aliyun_from_name = ''
  dialogVisible.value = true
}

const showCreateTemplateDialog = () => {
  templateForm.name = ''
  templateForm.description = ''
  templateForm.sender_type = 'QQ邮箱'
  templateForm.smtp_server = 'smtp.qq.com'
  templateForm.smtp_port = 587
  templateForm.use_tls = true
  templateForm.smtp_username = ''
  templateForm.smtp_security = ''
  templateForm.weight = 50
  templateForm.daily_quota = 500
  templateForm.aliyun_access_key = ''
  templateForm.aliyun_access_secret = ''
  templateForm.aliyun_region = 'cn-hangzhou'
  templateForm.aliyun_from_name = ''
  templateDialogVisible.value = true
}

const onSenderTypeChange = (type) => {
  const preset = presetForType(type)
  if (preset) {
    senderForm.smtp_server = preset.smtp_server
    senderForm.smtp_port = preset.smtp_port
    senderForm.use_tls = preset.use_tls
    senderForm.smtp_security = preset.security || ''
    if (preset.daily_quota !== undefined && preset.daily_quota !== null) {
      senderForm.daily_quota = preset.daily_quota
    }
  }
  if (type === '阿里云邮箱推送') {
    senderForm.smtp_server = ''
    senderForm.smtp_port = 0
    senderForm.password = ''
  }
  senderForm.email = ''
  if (type !== '阿里云邮箱推送') senderForm.password = ''
}

const onTemplateTypeChange = (type) => {
  const preset = presetForType(type)
  if (preset) {
    templateForm.smtp_server = preset.smtp_server
    templateForm.smtp_port = preset.smtp_port
    templateForm.use_tls = preset.use_tls
    templateForm.smtp_security = preset.security || ''
    if (preset.daily_quota !== undefined && preset.daily_quota !== null) {
      templateForm.daily_quota = preset.daily_quota
    }
  }
  if (type === '阿里云邮箱推送') {
    templateForm.smtp_server = ''
    templateForm.smtp_port = 0
  }
}

const createSender = async () => {
  try {
    const payload = {
      sender_type: senderForm.sender_type,
      email: senderForm.email,
      password: senderForm.password,
      smtp_server: senderForm.smtp_server,
      smtp_port: senderForm.smtp_port,
      use_tls: senderForm.use_tls,
      smtp_username: senderForm.smtp_username,
      smtp_security: senderForm.smtp_security,
      weight: senderForm.weight,
      daily_quota: senderForm.daily_quota,
    }
    if (senderForm.sender_type === '阿里云邮箱推送') {
      payload.aliyun_access_key = senderForm.aliyun_access_key
      payload.aliyun_access_secret = senderForm.aliyun_access_secret
      payload.aliyun_region = senderForm.aliyun_region
      payload.aliyun_from_name = senderForm.aliyun_from_name
    }
    await api.post('/api/senders', payload)
    ElMessage.success('发件人添加成功')
    dialogVisible.value = false
    loadSenders()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '添加发件人失败')
  }
}

const createTemplate = async () => {
  if (!templateForm.name.trim()) {
    ElMessage.warning('请输入模板名称')
    return
  }
  try {
    const payload = {
      name: templateForm.name.trim(),
      description: templateForm.description,
      sender_type: templateForm.sender_type,
      smtp_server: templateForm.smtp_server,
      smtp_port: templateForm.smtp_port,
      use_tls: templateForm.use_tls,
      smtp_username: templateForm.smtp_username,
      smtp_security: templateForm.smtp_security,
      weight: templateForm.weight,
      daily_quota: templateForm.daily_quota,
      aliyun_access_key: templateForm.aliyun_access_key,
      aliyun_access_secret: templateForm.aliyun_access_secret,
      aliyun_region: templateForm.aliyun_region,
      aliyun_from_name: templateForm.aliyun_from_name,
    }
    await api.post('/api/senders/templates', payload)
    ElMessage.success('模板保存成功')
    templateDialogVisible.value = false
    loadTemplates()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '保存模板失败')
  }
}

const testSender = async (id) => {
  try {
    const res = await api.post(`/api/senders/${id}/test`)
    if (res.data.success) {
      ElMessage.success(res.data.message || 'SMTP连接成功')
    } else {
      ElMessage.warning(res.data.message || '连接失败')
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

const unpauseSender = async (id) => {
  try {
    await api.put(`/api/senders/${id}`, { status: 'active' })
    ElMessage.success('发件人已恢复')
    loadSenders()
  } catch {
    ElMessage.error('恢复失败')
  }
}

const deleteSender = async (id) => {
  try {
    await ElMessageBox.confirm('确定删除此发件人？', '确认')
    await api.delete(`/api/senders/${id}`)
    ElMessage.success('删除成功')
    loadSenders()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '删除发件人失败')
  }
}

const applyTemplate = async (template) => {
  const email = window.prompt('请输入新发件人的邮箱地址：')
  if (!email) return
  let password = ''
  if (template.sender_type !== '阿里云邮箱推送') {
    password = window.prompt('请输入密码/授权码：') || ''
    if (!password) return
  }

  try {
    // Prefer apply endpoint for history reuse
    const applyPayload = {
      email,
      password,
      aliyun_access_key: template.aliyun_access_key || '',
      aliyun_access_secret: '',
    }
    if (template.sender_type === '阿里云邮箱推送') {
      const secret = window.prompt('请输入 AccessKey Secret（留空则复用模板已保存密钥）：') || ''
      applyPayload.aliyun_access_secret = secret
    }
    await api.post(`/api/senders/templates/${template.id}/apply`, applyPayload)
    ElMessage.success('应用模板创建发件人成功')
    loadSenders()
    activeTab.value = 'senders'
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '应用模板失败')
  }
}

const deleteTemplate = async (id) => {
  try {
    await ElMessageBox.confirm('确定删除此模板？', '确认')
    await api.delete(`/api/senders/templates/${id}`)
    ElMessage.success('删除成功')
    loadTemplates()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '删除模板失败')
  }
}

onMounted(() => {
  loadPresets()
  loadSenders()
  loadTemplates()
})
</script>

<style scoped>
.responsive-table { width: 100%; }
@media (max-width: 720px) {
  :deep(.el-card__body) { padding: 12px; }
  :deep(.el-table) { min-width: 760px; }
  :deep(.el-table__body-wrapper), :deep(.el-table__header-wrapper) { overflow-x: auto; }
  :deep(.el-dialog) { width: calc(100vw - 24px) !important; margin-top: 5vh !important; }
  :deep(.el-form-item__label) { width: 96px !important; }
  :deep(.el-form-item__content) { margin-left: 96px !important; }
}
</style>
