<template>
  <div class="schedule-config">
    <el-form label-width="110px">
      <el-form-item label="发送方式">
        <el-radio-group v-model="form.scheduleType" @change="emitChange">
          <el-radio value="immediate">立即发送</el-radio>
          <el-radio value="scheduled">定时发送</el-radio>
          <el-radio value="smart">智能分时段</el-radio>
        </el-radio-group>
      </el-form-item>

      <el-form-item v-if="form.scheduleType === 'scheduled'" label="发送时间">
        <el-date-picker
          v-model="form.scheduleTime"
          type="datetime"
          placeholder="选择发送时间"
          :disabled-date="disablePastDates"
          @change="emitChange"
        />
      </el-form-item>

      <el-form-item v-if="form.scheduleType === 'smart'" label="智能策略">
        <el-alert type="success" :closable="false" show-icon>
          按收件人邮箱域名推断大致时区，优先在对方本地 9-12 / 14-18 点窗口发送；窗口外自动等待后继续。
        </el-alert>
      </el-form-item>

      <el-form-item label="邮件间隔(秒)">
        <div class="range-row">
          <el-input-number v-model="form.delayMin" :min="0" :max="120" @change="emitChange" />
          <span class="range-sep">~</span>
          <el-input-number v-model="form.delayMax" :min="0" :max="300" @change="emitChange" />
        </div>
      </el-form-item>

      <el-form-item label="代理列表">
        <el-input
          v-model="form.proxiesText"
          type="textarea"
          :rows="4"
          placeholder="可选，每行一个：http://user:pass@ip:port 或 socks5://ip:port"
          @change="emitChange"
          @input="emitChange"
        />
        <div class="hint">启用后按轮询方式分配给每封邮件的 SMTP 连接。</div>
      </el-form-item>
    </el-form>
  </div>
</template>

<script setup>
import { reactive, watch } from 'vue'

const props = defineProps({
  scheduleType: { type: String, default: 'immediate' },
})

const emit = defineEmits(['update'])

const form = reactive({
  scheduleType: props.scheduleType || 'immediate',
  scheduleTime: null,
  delayMin: 5,
  delayMax: 15,
  proxiesText: '',
})

const disablePastDates = (date) => date.getTime() < Date.now() - 24 * 3600 * 1000

const emitChange = () => {
  const proxies = form.proxiesText
    .split(/\n+/)
    .map((x) => x.trim())
    .filter(Boolean)
  emit('update', {
    scheduleType: form.scheduleType,
    scheduleTime: form.scheduleTime,
    delayMin: form.delayMin,
    delayMax: form.delayMax,
    proxies,
  })
}

watch(
  () => props.scheduleType,
  (v) => {
    if (v) form.scheduleType = v
  }
)

emitChange()
</script>

<style scoped>
.schedule-config {
  width: 100%;
}
.range-row {
  display: flex;
  align-items: center;
  gap: 10px;
}
.range-sep {
  color: #909399;
}
.hint {
  margin-top: 6px;
  color: #909399;
  font-size: 12px;
  line-height: 1.4;
}
</style>
