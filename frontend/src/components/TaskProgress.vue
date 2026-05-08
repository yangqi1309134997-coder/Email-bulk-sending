<template>
  <div class="task-progress">
    <el-card v-if="!task">
      <el-empty description="暂无任务进度" />
    </el-card>
    <el-card v-else>
      <template #header>
        <div style="display: flex; justify-content: space-between; align-items: center">
          <span>{{ task.name }}</span>
          <el-tag :type="statusTagType">{{ statusLabel }}</el-tag>
        </div>
      </template>

      <el-progress
        :percentage="progressPercent"
        :status="progressStatus"
        :stroke-width="20"
        :text-inside="true"
        style="margin-bottom: 20px"
      />

      <el-row :gutter="20">
        <el-col :span="6">
          <el-statistic title="总数" :value="task.recipient_count" />
        </el-col>
        <el-col :span="6">
          <el-statistic title="成功" :value="task.success_count">
            <template #suffix>
              <span style="color: #67c23a; font-size: 12px">
                ({{ task.recipient_count ? ((task.success_count / task.recipient_count) * 100).toFixed(1) : 0 }}%)
              </span>
            </template>
          </el-statistic>
        </el-col>
        <el-col :span="6">
          <el-statistic title="失败" :value="task.fail_count">
            <template #suffix>
              <span style="color: #f56c6c; font-size: 12px">
                ({{ task.recipient_count ? ((task.fail_count / task.recipient_count) * 100).toFixed(1) : 0 }}%)
              </span>
            </template>
          </el-statistic>
        </el-col>
        <el-col :span="6">
          <el-statistic title="打开/点击" :value="`${task.open_count}/${task.click_count}`" />
        </el-col>
      </el-row>

      <el-divider />

      <div style="display: flex; gap: 10px">
        <el-button v-if="task.status === 'running'" type="warning" @click="$emit('pause', task.id)">暂停</el-button>
        <el-button v-if="task.status === 'paused'" type="success" @click="$emit('resume', task.id)">继续</el-button>
        <el-button v-if="['pending', 'running', 'paused'].includes(task.status)" type="danger" @click="$emit('cancel', task.id)">取消</el-button>
      </div>
    </el-card>
  </div>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  task: { type: Object, default: null },
})

defineEmits(['pause', 'resume', 'cancel'])

const statusMap = {
  pending: { label: '等待中', tag: 'info' },
  running: { label: '发送中', tag: '' },
  paused: { label: '已暂停', tag: 'warning' },
  completed: { label: '已完成', tag: 'success' },
  cancelled: { label: '已取消', tag: 'danger' },
}

const statusLabel = computed(() => statusMap[props.task?.status]?.label || '未知')
const statusTagType = computed(() => statusMap[props.task?.status]?.tag || 'info')

const progressPercent = computed(() => {
  if (!props.task || props.task.recipient_count === 0) return 0
  return Math.round(((props.task.success_count + props.task.fail_count) / props.task.recipient_count) * 100)
})

const progressStatus = computed(() => {
  if (props.task?.status === 'completed') return 'success'
  if (props.task?.status === 'cancelled') return 'exception'
  return null
})
</script>

<style scoped>
.task-progress {
  margin-bottom: 20px;
}
</style>
