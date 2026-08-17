<template>
  <div ref="chartRef" :style="{ width: width, height: height }" />
</template>

<script setup>
import { ref, onMounted, onUnmounted, watch } from 'vue'

const props = defineProps({
  option: { type: Object, required: true },
  width: { type: String, default: '100%' },
  height: { type: String, default: '350px' },
})

const chartRef = ref(null)
let chartInstance = null
let unmounted = false

const loadEcharts = async () => {
  const [echarts, charts, components, renderers] = await Promise.all([
    import('echarts/core'),
    import('echarts/charts'),
    import('echarts/components'),
    import('echarts/renderers'),
  ])
  echarts.use([
    charts.LineChart,
    charts.PieChart,
    components.GridComponent,
    components.LegendComponent,
    components.TooltipComponent,
    renderers.CanvasRenderer,
  ])
  return echarts
}

const initChart = async () => {
  const echarts = await loadEcharts()
  if (!unmounted && chartRef.value) {
    chartInstance = echarts.init(chartRef.value)
    chartInstance.setOption(props.option)
  }
}

const handleResize = () => {
  chartInstance?.resize()
}

watch(() => props.option, (newOption) => {
  chartInstance?.setOption(newOption, true)
}, { deep: true })

onMounted(() => {
  unmounted = false
  void initChart()
  window.addEventListener('resize', handleResize)
})

onUnmounted(() => {
  unmounted = true
  window.removeEventListener('resize', handleResize)
  chartInstance?.dispose()
  chartInstance = null
})
</script>
