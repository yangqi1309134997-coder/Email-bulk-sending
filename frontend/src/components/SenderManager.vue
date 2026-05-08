<template>
  <div class="sender-manager">
    <el-radio-group v-model="strategy" style="margin-bottom: 20px" @change="emitChange">
      <el-radio value="round_robin">轮询</el-radio>
      <el-radio value="weighted">权重分配</el-radio>
      <el-radio value="smart">智能模式</el-radio>
    </el-radio-group>

    <el-checkbox-group v-model="selectedIds" @change="emitChange">
      <div v-for="s in senders" :key="s.id" class="sender-item">
        <el-checkbox :value="s.id">
          <span>{{ s.email }}</span>
          <el-tag size="small" style="margin-left: 8px">{{ s.sender_type }}</el-tag>
          <span style="margin-left: 8px; color: #909399">权重: {{ s.weight }}</span>
          <span style="margin-left: 8px; color: #909399">配额: {{ s.daily_sent }}/{{ s.daily_quota }}</span>
        </el-checkbox>
      </div>
    </el-checkbox-group>

    <el-empty v-if="senders.length === 0" description="暂无发件人，请先在设置中添加" />
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import api from '../api'
import { ElMessage } from 'element-plus'

const strategy = ref('round_robin')
const selectedIds = ref([])
const senders = ref([])

const emit = defineEmits(['update'])

const emitChange = () => {
  emit('update', {
    strategy: strategy.value,
    senderIds: selectedIds.value,
  })
}

const loadSenders = async () => {
  try {
    const res = await api.get('/api/senders')
    senders.value = res.data
  } catch {
    ElMessage.error('加载发件人失败')
  }
}

onMounted(() => {
  loadSenders()
})
</script>

<style scoped>
.sender-manager {
  width: 100%;
}
.sender-item {
  margin: 8px 0;
  padding: 8px;
  border-radius: 4px;
}
.sender-item:hover {
  background: #f5f7fa;
}
</style>
