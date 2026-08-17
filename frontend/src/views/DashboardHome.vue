<template>
  <div class="dashboard-home">
    <el-row :gutter="20" style="margin-bottom: 20px">
      <el-col :xs="24" :sm="12" :md="6">
        <el-card shadow="hover">
          <el-statistic title="今日发送" :value="stats.today_sent">
            <template #suffix>
              <span style="font-size: 12px; color: #909399">封</span>
            </template>
          </el-statistic>
        </el-card>
      </el-col>
      <el-col :xs="24" :sm="12" :md="6">
        <el-card shadow="hover">
          <el-statistic title="今日成功率" :value="stats.today_success_rate">
            <template #suffix>
              <span style="font-size: 12px">%</span>
            </template>
          </el-statistic>
        </el-card>
      </el-col>
      <el-col :xs="24" :sm="12" :md="6">
        <el-card shadow="hover">
          <el-statistic title="本周发送" :value="stats.week_sent">
            <template #suffix>
              <span style="font-size: 12px; color: #909399">封</span>
            </template>
          </el-statistic>
        </el-card>
      </el-col>
      <el-col :xs="24" :sm="12" :md="6">
        <el-card shadow="hover">
          <el-statistic title="活跃任务" :value="stats.active_tasks" />
        </el-card>
      </el-col>
    </el-row>

    <el-row :gutter="20" style="margin-bottom: 20px">
      <el-col :xs="24" :md="12">
        <el-card>
          <template #header>发送趋势</template>
          <StatsChart :option="trendChartOption" height="300px" />
        </el-card>
      </el-col>
      <el-col :xs="24" :md="12">
        <el-card>
          <template #header>发送结果分布</template>
          <StatsChart :option="pieChartOption" height="300px" />
        </el-card>
      </el-col>
    </el-row>

    <el-row :gutter="20">
      <el-col :xs="24" :lg="16">
        <el-card>
          <template #header>
            <div style="display: flex; justify-content: space-between; align-items: center">
              <span>最近任务</span>
              <el-button text type="primary" @click="$router.push('/send')">创建新任务</el-button>
            </div>
          </template>
          <el-table class="desktop-task-table" :data="recentTasks" stripe size="small">
            <el-table-column prop="name" label="任务名称" show-overflow-tooltip>
              <template #default="{ row }">
                <el-button link type="primary" @click="$router.push(`/tasks/${row.id}`)">{{ row.name }}</el-button>
              </template>
            </el-table-column>
            <el-table-column prop="status" label="状态" width="120">
              <template #default="{ row }">
                <el-tag :type="statusTagType(row.status)" size="small">{{ statusLabel(row.status) }}</el-tag>
                <el-tag v-if="isPartiallyCompleted(row)" type="warning" size="small" style="margin-left: 4px">未发完</el-tag>
              </template>
            </el-table-column>
            <el-table-column prop="recipient_count" label="收件人" width="80" />
            <el-table-column label="进度" width="120">
              <template #default="{ row }">
                <el-progress
                  :percentage="taskProgress(row)"
                  :status="progressStatus(row)"
                  :stroke-width="14"
                  :text-inside="true"
                  size="small"
                />
              </template>
            </el-table-column>
            <el-table-column label="成功/失败" width="100">
              <template #default="{ row }">
                <span style="color: #67c23a">{{ row.success_count }}</span> /
                <span style="color: #f56c6c">{{ row.fail_count }}</span>
              </template>
            </el-table-column>
            <el-table-column prop="created_at" label="创建时间" width="160">
              <template #default="{ row }">
                {{ formatTime(row.created_at) }}
              </template>
            </el-table-column>
            <el-table-column label="操作" width="360" fixed="right">
              <template #default="{ row }">
                <el-button size="small" @click="$router.push(`/tasks/${row.id}`)">详情</el-button>
                <el-button v-if="row.status === 'running'" size="small" type="warning" @click="pauseTask(row.id)">暂停</el-button>
                <el-button v-if="row.status === 'paused' || row.status === 'failed'" size="small" type="success" @click="resumeTask(row.id)">继续</el-button>
                <el-button v-if="row.status === 'running' || row.status === 'paused'" size="small" type="danger" @click="cancelTask(row.id)">取消</el-button>
                <el-button v-if="hasFailedLogs(row)" size="small" type="primary" @click="retryFailed(row.id)">重试失败</el-button>
                <el-button v-if="hasUnsent(row) && !hasFailedLogs(row)" size="small" type="primary" @click="retryFailed(row.id)">继续发送</el-button>
              </template>
            </el-table-column>
          </el-table>
          <div class="mobile-task-list">
            <div v-for="row in recentTasks" :key="row.id" class="mobile-task-row">
              <div class="mobile-task-head">
                <el-button link type="primary" @click="$router.push(`/tasks/${row.id}`)">{{ row.name }}</el-button>
                <el-tag :type="statusTagType(row.status)" size="small">{{ statusLabel(row.status) }}</el-tag>
              </div>
              <div class="mobile-task-progress">
                <el-progress
                  :percentage="taskProgress(row)"
                  :status="progressStatus(row)"
                  :stroke-width="12"
                  :text-inside="true"
                  size="small"
                />
                <span>{{ row.success_count || 0 }}/{{ row.recipient_count || 0 }} 成功</span>
              </div>
              <div class="mobile-task-actions">
                <el-button size="small" @click="$router.push(`/tasks/${row.id}`)">详情</el-button>
                <el-button v-if="row.status === 'running'" size="small" type="warning" @click="pauseTask(row.id)">暂停</el-button>
                <el-button v-if="row.status === 'paused' || row.status === 'failed'" size="small" type="success" @click="resumeTask(row.id)">继续</el-button>
                <el-button v-if="row.status === 'running' || row.status === 'paused'" size="small" type="danger" @click="cancelTask(row.id)">取消</el-button>
                <el-button v-if="hasFailedLogs(row) || hasUnsent(row)" size="small" type="primary" @click="retryFailed(row.id)">重试</el-button>
              </div>
            </div>
            <el-empty v-if="!recentTasks.length" description="暂无任务" :image-size="60" />
          </div>
        </el-card>
      </el-col>
      <el-col :xs="24" :lg="8">
        <el-card>
          <template #header>实时状态</template>
          <div class="realtime-info">
            <div class="realtime-item">
              <span class="label">运行中任务</span>
              <span class="value">{{ realtime.running_tasks }}</span>
            </div>
            <div class="realtime-item">
              <span class="label">排队中任务</span>
              <span class="value">{{ realtime.queued_tasks }}</span>
            </div>
            <div class="realtime-item">
              <span class="label">Worker数量</span>
              <span class="value">{{ realtime.worker_count }}</span>
            </div>
          </div>
          <el-divider />
          <div class="quick-stats">
            <h4>本周概览</h4>
            <div class="stat-row">
              <span>发送总量</span>
              <el-tag>{{ stats.week_sent }}</el-tag>
            </div>
            <div class="stat-row">
              <span>成功量</span>
              <el-tag type="success">{{ stats.week_success }}</el-tag>
            </div>
            <div class="stat-row">
              <span>成功率</span>
              <el-tag :type="stats.week_success_rate > 90 ? 'success' : stats.week_success_rate > 70 ? 'warning' : 'danger'">
                {{ stats.week_success_rate }}%
              </el-tag>
            </div>
            <div class="stat-row">
              <span>今日失败</span>
              <el-tag type="danger">{{ stats.today_fail }}</el-tag>
            </div>
          </div>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script setup>
import { ref, reactive, computed, onMounted, onUnmounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import api from '../api'
import StatsChart from '../components/StatsChart.vue'
import { formatApiDateTime } from '../utils/time'

const stats = reactive({
  today_sent: 0, today_success: 0, today_fail: 0, today_success_rate: 0,
  week_sent: 0, week_success: 0, week_success_rate: 0,
  total_tasks: 0, active_tasks: 0,
})

const realtime = reactive({
  running_tasks: 0, queued_tasks: 0, worker_count: 0,
})

const recentTasks = ref([])
const trendData = ref([])

const statusMap = {
  pending: { label: '等待中', type: 'info' },
  running: { label: '发送中', type: '' },
  paused: { label: '已暂停', type: 'warning' },
  completed: { label: '已完成', type: 'success' },
  cancelled: { label: '已取消', type: 'danger' },
  failed: { label: '发送失败', type: 'danger' },
}

const statusLabel = (s) => statusMap[s]?.label || '未知'
const statusTagType = (s) => statusMap[s]?.type || 'info'

const taskProgress = (task) => {
  if (!task.recipient_count) return 0
  return Math.round(((task.success_count + task.fail_count) / task.recipient_count) * 100)
}

const isPartiallyCompleted = (row) => {
  return (row.status === 'completed' || row.status === 'failed' || row.status === 'cancelled')
    && (row.success_count + row.fail_count) < row.recipient_count
}

const progressStatus = (row) => {
  if (row.status === 'cancelled') return 'exception'
  if (row.status === 'completed' && isPartiallyCompleted(row)) return 'warning'
  if (row.status === 'completed') return 'success'
  return ''
}

const formatTime = formatApiDateTime

const trendChartOption = computed(() => ({
  tooltip: { trigger: 'axis' },
  legend: { data: ['发送量', '成功量', '失败量'] },
  xAxis: {
    type: 'category',
    data: trendData.map((item) => {
      const [, month, day] = item.date.split('-')
      return `${month}/${day}`
    }),
  },
  yAxis: { type: 'value', minInterval: 1 },
  series: [
    {
      name: '发送量',
      type: 'line',
      smooth: true,
      data: trendData.map((item) => item.sent),
      areaStyle: { opacity: 0.3 },
      itemStyle: { color: '#409eff' },
    },
    {
      name: '成功量',
      type: 'line',
      smooth: true,
      data: trendData.map((item) => item.success),
      areaStyle: { opacity: 0.3 },
      itemStyle: { color: '#67c23a' },
    },
    {
      name: '失败量',
      type: 'line',
      smooth: true,
      data: trendData.map((item) => item.failed),
      itemStyle: { color: '#f56c6c' },
    },
  ],
}))

const pieChartOption = computed(() => {
  const pending = Math.max(0, stats.today_sent - stats.today_success - stats.today_fail)
  const data = [
    { value: stats.today_success, name: '成功', itemStyle: { color: '#67c23a' } },
    { value: stats.today_fail, name: '失败', itemStyle: { color: '#f56c6c' } },
  ]
  if (pending > 0) data.push({ value: pending, name: '待发送', itemStyle: { color: '#e6a23c' } })
  return {
    tooltip: { trigger: 'item' },
    legend: { bottom: 0 },
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      avoidLabelOverlap: false,
      itemStyle: { borderRadius: 10, borderColor: '#fff', borderWidth: 2 },
      label: { show: false },
      emphasis: { label: { show: true, fontSize: 16, fontWeight: 'bold' } },
      data,
    }],
  }
})

const loadStats = async () => {
  try {
    const res = await api.get('/api/dashboard/stats')
    Object.assign(stats, res.data)
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '加载失败')
  }
}

const loadRealtime = async () => {
  try {
    const res = await api.get('/api/dashboard/realtime')
    Object.assign(realtime, res.data)
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '加载失败')
  }
}

const loadTrend = async () => {
  try {
    const res = await api.get('/api/dashboard/trend')
    trendData.value = res.data
  } catch {
    // 趋势图数据缺失不打断仪表盘主体
  }
}

const loadTasks = async () => {
  try {
    const res = await api.get('/api/tasks', { params: { limit: 10 } })
    recentTasks.value = res.data
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '加载失败')
  }
}

const pauseTask = async (id) => {
  try {
    await api.post(`/api/tasks/${id}/pause`)
    ElMessage.success('任务已暂停')
    loadTasks()
    loadStats()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '暂停失败')
  }
}

const resumeTask = async (id) => {
  try {
    await api.post(`/api/tasks/${id}/resume`)
    ElMessage.success('任务已继续')
    loadTasks()
    loadStats()
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || '继续失败')
  }
}

const cancelTask = async (id) => {
  try {
    await ElMessageBox.confirm('确定取消此任务？取消后不可恢复', '确认取消')
    await api.post(`/api/tasks/${id}/cancel`)
    ElMessage.success('任务已取消')
    loadTasks()
    loadStats()
  } catch (err) {
    if (err !== 'cancel') {
      ElMessage.error(err.response?.data?.detail || '取消任务失败')
    }
  }
}

const hasFailedLogs = (row) => {
  return (row.status === 'completed' || row.status === 'failed' || row.status === 'cancelled') && row.fail_count > 0
}

const hasUnsent = (row) => {
  return (row.status === 'completed' || row.status === 'failed' || row.status === 'cancelled')
    && (row.success_count + row.fail_count) < row.recipient_count
}

const retryFailed = async (id) => {
  try {
    await ElMessageBox.confirm('将重试所有失败的邮件，确认继续？', '重试失败邮件')
    const res = await api.post(`/api/tasks/${id}/retry-failed`)
    ElMessage.success(res.data.message)
    loadTasks()
    loadStats()
  } catch (err) {
    if (err !== 'cancel') {
      ElMessage.error(err.response?.data?.detail || '重试失败')
    }
  }
}

let refreshTimer = null
onMounted(() => {
  loadStats()
  loadRealtime()
  loadTasks()
  loadTrend()
  refreshTimer = setInterval(() => {
    loadStats()
    loadRealtime()
    loadTasks()
  }, 5000)
})
onUnmounted(() => {
  if (refreshTimer) clearInterval(refreshTimer)
})
</script>

<style scoped>
.dashboard-home {
  max-width: 1240px;
  margin: 0 auto;
}
.dashboard-home :deep(.el-card) {
  transition: transform .15s ease, box-shadow .15s ease;
}
.dashboard-home :deep(.el-card:hover) {
  transform: translateY(-1px);
  box-shadow: var(--app-shadow) !important;
}
.realtime-info {
  display: flex;
  flex-direction: column;
  gap: 16px;
}
.realtime-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 8px 0;
  border-bottom: 1px solid #f0f0f0;
}
.realtime-item .label {
  color: #606266;
}
.realtime-item .value {
  font-size: 20px;
  font-weight: bold;
  color: #409eff;
}
.quick-stats {
  display: flex;
  flex-direction: column;
  gap: 12px;
}
.stat-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.mobile-task-list {
  display: none;
}
@media (max-width: 767px) {
  .desktop-task-table {
    display: none;
  }
  .mobile-task-list {
    display: flex;
    flex-direction: column;
  }
  .mobile-task-row {
    padding: 12px 0;
    border-bottom: 1px solid var(--el-border-color-lighter);
  }
  .mobile-task-row:last-child {
    border-bottom: 0;
  }
  .mobile-task-head,
  .mobile-task-progress,
  .mobile-task-actions {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .mobile-task-head {
    justify-content: space-between;
    min-width: 0;
  }
  .mobile-task-head .el-button {
    min-width: 0;
    max-width: 70%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .mobile-task-progress {
    margin: 10px 0;
  }
  .mobile-task-progress .el-progress {
    flex: 1;
    min-width: 0;
  }
  .mobile-task-progress > span {
    flex: 0 0 auto;
    color: var(--el-text-color-secondary);
    font-size: 12px;
  }
  .mobile-task-actions {
    flex-wrap: wrap;
  }
}
</style>
