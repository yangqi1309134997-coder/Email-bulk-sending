<template>
  <div class="task-detail">
    <div class="toolbar">
      <el-button @click="$router.push('/dashboard')">返回仪表盘</el-button>
      <div class="title-wrap">
        <h2>{{ task?.name || '任务详情' }}</h2>
        <el-tag v-if="task" :type="statusTagType" effect="dark">{{ statusLabel }}</el-tag>
      </div>
      <div class="actions">
        <el-button v-if="task?.status === 'running'" type="warning" @click="pauseTask">暂停</el-button>
        <el-button v-if="task?.status === 'paused' || task?.status === 'failed'" type="success" @click="resumeTask">继续</el-button>
        <el-button v-if="['pending','running','paused'].includes(task?.status)" type="danger" @click="cancelTask">取消</el-button>
        <el-button v-if="task" @click="exportReport">导出报告</el-button>
      </div>
    </div>

    <el-row :gutter="16">
      <el-col :span="16">
        <el-card shadow="never" class="panel">
          <template #header>
            <div class="panel-header">
              <span>实时进度</span>
              <el-tag size="small" :type="wsConnected ? 'success' : 'info'">{{ wsConnected ? 'WebSocket 已连接' : '轮询模式' }}</el-tag>
            </div>
          </template>
          <el-progress :percentage="progressPercent" :status="progressStatus" :stroke-width="18" striped :striped-flow="task?.status==='running'" />
          <el-row :gutter="12" class="stats-row">
            <el-col :span="6"><el-statistic title="收件人" :value="task?.recipient_count || 0" /></el-col>
            <el-col :span="6"><el-statistic title="成功" :value="task?.success_count || 0" /></el-col>
            <el-col :span="6"><el-statistic title="失败" :value="task?.fail_count || 0" /></el-col>
            <el-col :span="6">
              <el-statistic title="打开/点击" :value="task?.open_count || 0">
                <template #suffix>/{{ task?.click_count || 0 }}</template>
              </el-statistic>
            </el-col>
          </el-row>
          <el-alert v-if="lastEvent" :title="lastEvent" type="info" show-icon :closable="false" class="event-alert" />
        </el-card>

        <el-card shadow="never" class="panel">
          <template #header>发送明细</template>
          <el-table :data="logs" stripe size="small" max-height="420">
            <el-table-column prop="recipient_email" label="收件人" min-width="180" show-overflow-tooltip />
            <el-table-column prop="recipient_name" label="姓名" width="100" />
            <el-table-column prop="status" label="状态" width="100">
              <template #default="{ row }">
                <el-tag size="small" :type="logTag(row.status)">{{ row.status }}</el-tag>
              </template>
            </el-table-column>
            <el-table-column prop="error_message" label="失败原因" min-width="180" show-overflow-tooltip />
            <el-table-column prop="sent_at" label="发送时间" width="170">
              <template #default="{ row }">{{ formatTime(row.sent_at) }}</template>
            </el-table-column>
          </el-table>
        </el-card>
      </el-col>

      <el-col :span="8">
        <el-card shadow="never" class="panel">
          <template #header>任务信息</template>
          <el-descriptions :column="1" border size="small">
            <el-descriptions-item label="任务ID">{{ task?.id }}</el-descriptions-item>
            <el-descriptions-item label="发送方式">{{ task?.schedule_type }}</el-descriptions-item>
            <el-descriptions-item label="创建时间">{{ formatTime(task?.created_at) }}</el-descriptions-item>
            <el-descriptions-item label="完成时间">{{ formatTime(task?.completed_at) }}</el-descriptions-item>
            <el-descriptions-item v-if="task?.status === 'paused'" label="暂停原因">{{ pauseReasonLabel }}</el-descriptions-item>
            <el-descriptions-item v-if="task?.status === 'paused'" label="预计恢复">{{ formatTime(task?.next_run_at) }}</el-descriptions-item>
            <el-descriptions-item label="进度">{{ progressPercent }}%</el-descriptions-item>
          </el-descriptions>
        </el-card>
        <el-card shadow="never" class="panel tips">
          <template #header>操作提示</template>
          <ul>
            <li>运行中任务可暂停 / 取消</li>
            <li>风控暂停后会按设置自动恢复</li>
            <li>失败明细可导出 CSV 排查</li>
          </ul>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script setup>
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import api from '../api'
import { formatApiDateTime } from '../utils/time'

const route = useRoute()
const router = useRouter()
const task = ref(null)
const logs = ref([])
const wsConnected = ref(false)
const lastEvent = ref('')
let ws = null
let pollTimer = null
let wsRetryTimer = null
let logRefreshTimer = null
let mounted = false

const statusMap = {
  pending: { label: '等待中', type: 'info' },
  running: { label: '发送中', type: '' },
  paused: { label: '已暂停', type: 'warning' },
  completed: { label: '已完成', type: 'success' },
  cancelled: { label: '已取消', type: 'danger' },
  failed: { label: '发送失败', type: 'danger' },
}

const statusLabel = computed(() => statusMap[task.value?.status]?.label || '未知')
const statusTagType = computed(() => statusMap[task.value?.status]?.type || 'info')
const pauseReasonLabel = computed(() => ({
  manual: '手动暂停',
  rate_limit: '检测到风控，冷却中',
  retry_backoff: '失败重试退避中',
  smart_window_wait: '等待收件人时区窗口',
  no_available_sender: '暂无可用发件人',
  engine_crash: '发送进程恢复中',
}[task.value?.pause_reason] || '系统暂停'))
const progressPercent = computed(() => {
  if (!task.value?.recipient_count) return 0
  return Math.round(((task.value.success_count + task.value.fail_count) / task.value.recipient_count) * 100)
})
const progressStatus = computed(() => {
  if (task.value?.status === 'completed') return 'success'
  if (task.value?.status === 'cancelled' || task.value?.status === 'failed') return 'exception'
  return undefined
})

const formatTime = formatApiDateTime
const logTag = (s) => ({ success: 'success', failed: 'danger', pending: 'info' }[s] || 'info')

const loadTask = async () => {
  const id = route.params.id
  const res = await api.get(`/api/tasks/${id}`)
  task.value = res.data
}

const loadLogs = async () => {
  const id = route.params.id
  const res = await api.get(`/api/tasks/${id}/logs`, { params: { limit: 500 } })
  logs.value = res.data
}

const refresh = async () => {
  try {
    await Promise.all([loadTask(), loadLogs()])
  } catch (e) {
    ElMessage.error('加载任务失败')
  }
}

const scheduleLogRefresh = () => {
  if (logRefreshTimer) return
  logRefreshTimer = setTimeout(async () => {
    logRefreshTimer = null
    if (!mounted) return
    try {
      await loadLogs()
    } catch (err) {
      ElMessage.error(err.response?.data?.detail || '操作失败')
    }
  }, 500)
}

const connectWs = () => {
  const id = route.params.id
  const proto = location.protocol === 'https:' ? 'wss' : 'ws'
  const token = localStorage.getItem('access_token')
  if (!token) return
  try {
    ws = new WebSocket(`${proto}://${location.host}/ws/tasks/${id}`)
    ws.onopen = () => { wsConnected.value = true }
    ws.onclose = (event) => {
      wsConnected.value = false
      if (mounted && ![4401, 4403, 4404].includes(event.code)) {
        clearTimeout(wsRetryTimer)
        wsRetryTimer = setTimeout(connectWs, 3000)
      }
    }
    ws.onerror = () => { wsConnected.value = false }
    ws.onmessage = (ev) => {
      try {
        const data = JSON.parse(ev.data)
        lastEvent.value = data.message || `${data.type || 'event'}: ${data.last_result || data.status || ''}`
        if (typeof data.success_count === 'number') {
          task.value = {
            ...(task.value || {}),
            status: data.status || task.value?.status,
            success_count: data.success_count,
            fail_count: data.fail_count,
            recipient_count: data.recipient_count || task.value?.recipient_count,
          }
        }
        if (data.type === 'progress' || data.type === 'status') {
          scheduleLogRefresh()
        }
      } catch {
        // Ignore malformed frames; the polling timer keeps the page fresh.
      }
    }
  } catch {
    wsConnected.value = false
  }
}

const pauseTask = async () => {
  try {
    await api.post(`/api/tasks/${route.params.id}/pause`)
    ElMessage.success('已暂停')
    refresh()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '暂停失败')
  }
}
const resumeTask = async () => {
  try {
    await api.post(`/api/tasks/${route.params.id}/resume`)
    ElMessage.success('已继续')
    refresh()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '继续失败')
  }
}
const cancelTask = async () => {
  try {
    await ElMessageBox.confirm('确定取消此任务？', '确认')
    await api.post(`/api/tasks/${route.params.id}/cancel`)
    ElMessage.success('已取消')
    refresh()
  } catch (err) {
    if (err !== 'cancel') ElMessage.error(err.response?.data?.detail || '取消失败')
  }
}
const exportReport = async () => {
  try {
    const res = await api.get(`/api/tasks/${route.params.id}/export`, { responseType: 'blob' })
    const url = window.URL.createObjectURL(new Blob([res.data]))
    const a = document.createElement('a')
    a.href = url
    a.download = `task_${route.params.id}_report.csv`
    a.click()
    window.URL.revokeObjectURL(url)
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '导出报告失败')
  }
}

onMounted(async () => {
  mounted = true
  await refresh()
  connectWs()
  pollTimer = setInterval(refresh, 4000)
})
onUnmounted(() => {
  mounted = false
  if (pollTimer) clearInterval(pollTimer)
  if (wsRetryTimer) clearTimeout(wsRetryTimer)
  if (logRefreshTimer) clearTimeout(logRefreshTimer)
  if (ws) ws.close()
})
</script>

<style scoped>
.task-detail { max-width: 1200px; margin: 0 auto; }
.toolbar { display:flex; align-items:center; justify-content:space-between; gap:12px; margin-bottom:16px; }
.title-wrap { display:flex; align-items:center; gap:10px; flex:1; }
.title-wrap h2 { margin:0; font-size:20px; }
.actions { display:flex; gap:8px; flex-wrap:wrap; }
.panel { margin-bottom:16px; border-radius:12px; }
.panel-header { display:flex; justify-content:space-between; align-items:center; }
.stats-row { margin-top:18px; }
.event-alert { margin-top:14px; }
.tips ul { margin:0; padding-left:18px; color:#606266; line-height:1.8; }
</style>
