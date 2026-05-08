<template>
  <div class="schedule-config">
    <el-form label-width="100px">
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

      <el-form-item v-if="form.scheduleType === 'smart'" label="智能说明">
        <el-alert type="info" :closable="false" show-icon>
          系统将根据收件人邮箱域名推断时区，在最佳时段自动发送
        </el-alert>
      </el-form-item>

      <el-form-item label="邮件间隔(秒)">
        <el-input-number v-model="form.delayMin" :min="1" :max="60" @change="emitChange" /> ~
        <el-input-number v-model="form.delayMax" :min="1" :max="120" @change="emitChange" />
      </el-form-item>
    </el-form>
  </div>
</template>

<script setup>
import { reactive } from 'vue'

const props = defineProps({
  scheduleType: { type: String, default: 'immediate' },
  scheduleTime: { type: [String, Date], default: null },
  delayMin: { type: Number, default: 5 },
  delayMax: { type: Number, default: 15 },
})

const form = reactive({
  scheduleType: props.scheduleType,
  scheduleTime: props.scheduleTime,
  delayMin: props.delayMin,
  delayMax: props.delayMax,
})

const emit = defineEmits(['update'])

const emitChange = () => {
  emit('update', { ...form })
}

const disablePastDates = (date) => {
  return date.getTime() < Date.now() - 86400000
}
</script>

<style scoped>
.schedule-config {
  width: 100%;
}
</style>
