<template>
  <div class="dashboard-home">
    <el-row :gutter="20" style="margin-bottom: 20px">
      <el-col :span="6">
        <el-card shadow="hover">
          <el-statistic title="今日发送" :value="stats.today_sent">
            <template #suffix>
              <span style="font-size: 12px; color: #909399">封</span>
            </template>
          </el-statistic>
        </el-card>
      </el-col>
      <el-col :span="6">
        <el-card shadow="hover">
          <el-statistic title="今日成功率" :value="stats.today_success_rate">
            <template #suffix>
              <span style="font-size: 12px">%</span>
            </template>
          </el-statistic>
        </el-card>
      </el-col>
      <el-col :span="6">
        <el-card shadow="hover">
          <el-statistic title="本周发送" :value="stats.week_sent">
            <template #suffix>
              <span style="font-size: 12px; color: #909399">封</span>
            </template>
          </el-statistic>
        </el-card>
      </el-col>
      <el-col :span="6">
        <el-card shadow="hover">
          <el-statistic title="活跃任务" :value="stats.active_tasks" />
        </el-card>
      </el-col>
    </el-row>

    <el-row :gutter="20" style="margin-bottom: 20px">
      <el-col :span="12">
        <el-card>
          <template #header>发送趋势</template>
          <StatsChart :option="trendChartOption" height="300px" />
        </el-card>
      </el-col>
      <el-col :span="12">
        <el-card>
          <template #header>发送结果分布</template>
          <StatsChart :option="pieChartOption" height="300px" />
        </el-card>
      </el-col>
    </el-row>

    <el-row :gutter="20">
      <el-col :span="16">
        <el-card>
          <template #header>
            <div style="display: flex; justify-content: space-between; align-items: center">
              <span>最近任务</span>
              <el-button text type="primary" @click="$router.push('/send')">创建新任务</el-button>
            </div>
          </template>
          <el-table :data="recentTasks" stripe size="small">
            <el-table-column prop="name" label="任务名称" show-overflow-tooltip />
            <el-table-column prop="status" label="状态" width="100">
              <template #default="{ row }">
                <el-tag :type="statusTagType(row.status)" size="small">{{ statusLabel(row.status) }}</el-tag>
              </template>
            </el-table-column>
            <el-table-column prop="recipient_count" label="收件人" width="80" />
            <el-table-column label="进度" width="120">
              <template #default="{ row }">
                <el-progress
                  :percentage="taskProgress(row)"
                  :status="row.status === 'completed' ? 'success' : row.status === 'cancelled' ? 'exception' : ''"
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
          </el-table>
        </el-card>
      </el-col>
      <el-col :span="8">
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
import { ref, reactive, computed, onMounted } from 'vue'
import api from '../api'
import StatsChart from '../components/StatsChart.vue'

const stats = reactive({
  today_sent: 0, today_success: 0, today_fail: 0, today_success_rate: 0,
  week_sent: 0, week_success: 0, week_success_rate: 0,
  total_tasks: 0, active_tasks: 0,
})

const realtime = reactive({
  running_tasks: 0, queued_tasks: 0, worker_count: 0,
})

const recentTasks = ref([])

const statusMap = {
  pending: { label: '等待中', type: 'info' },
  running: { label: '发送中', type: '' },
  paused: { label: '已暂停', type: 'warning' },
  completed: { label: '已完成', type: 'success' },
  cancelled: { label: '已取消', type: 'danger' },
}

const statusLabel = (s) => statusMap[s]?.label || '未知'
const statusTagType = (s) => statusMap[s]?.type || 'info'

const taskProgress = (task) => {
  if (!task.recipient_count) return 0
  return Math.round(((task.success_count + task.fail_count) / task.recipient_count) * 100)
}

const formatTime = (t) => {
  if (!t) return '-'
  return new Date(t).toLocaleString('zh-CN')
}

const trendChartOption = computed(() => ({
  tooltip: { trigger: 'axis' },
  xAxis: {
    type: 'category',
    data: ['周一', '周二', '周三', '周四', '周五', '周六', '周日'],
  },
  yAxis: { type: 'value' },
  series: [
    {
      name: '发送量',
      type: 'line',
      smooth: true,
      data: [0, 0, 0, 0, 0, 0, 0],
      areaStyle: { opacity: 0.3 },
      itemStyle: { color: '#409eff' },
    },
    {
      name: '成功量',
      type: 'line',
      smooth: true,
      data: [0, 0, 0, 0, 0, 0, 0],
      areaStyle: { opacity: 0.3 },
      itemStyle: { color: '#67c23a' },
    },
  ],
}))

const pieChartOption = computed(() => ({
  tooltip: { trigger: 'item' },
  legend: { bottom: 0 },
  series: [{
    type: 'pie',
    radius: ['40%', '70%'],
    avoidLabelOverlap: false,
    itemStyle: { borderRadius: 10, borderColor: '#fff', borderWidth: 2 },
    label: { show: false },
    emphasis: { label: { show: true, fontSize: 16, fontWeight: 'bold' } },
    data: [
      { value: stats.today_success, name: '成功', itemStyle: { color: '#67c23a' } },
      { value: stats.today_fail, name: '失败', itemStyle: { color: '#f56c6c' } },
      { value: Math.max(0, stats.today_sent - stats.today_success - stats.today_fail), name: '待发送', itemStyle: { color: '#e6a23c' } },
    ],
  }],
}))

const loadStats = async () => {
  try {
    const res = await api.get('/api/dashboard/stats')
    Object.assign(stats, res.data)
  } catch {}
}

const loadRealtime = async () => {
  try {
    const res = await api.get('/api/dashboard/realtime')
    Object.assign(realtime, res.data)
  } catch {}
}

const loadTasks = async () => {
  try {
    const res = await api.get('/api/tasks')
    recentTasks.value = res.data.slice(0, 10)
  } catch {}
}

onMounted(() => {
  loadStats()
  loadRealtime()
  loadTasks()
})
</script>

<style scoped>
.dashboard-home {
  max-width: 1200px;
  margin: 0 auto;
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
</style>
