<script setup name="AiExcelAnalysisChart" lang="ts">
import { ref, onMounted, onBeforeUnmount, watch } from 'vue'
import * as echarts from 'echarts'
import { DataLine } from '@element-plus/icons-vue'
import type { ChatAnalysisModule } from '../../lib/chat-analysis'

const props = defineProps({
  module: {
    type: Object as () => ChatAnalysisModule,
    required: true,
  },
})

const PALETTE = [
  '#19734a',
  '#3977a8',
  '#d3922e',
  '#b85151',
  '#6b63a6',
  '#4f8f8a',
  '#87944c',
  '#c26f3a',
]

const chartRef = ref<HTMLDivElement | null>(null)
let chart: echarts.ECharts | null = null

function truncateLabel(value: string): string {
  const str = String(value ?? '')
  return str.length > 10 ? `${str.slice(0, 9)}…` : str
}

function buildOption(): echarts.EChartsCoreOption {
  const module = props.module
  const data = module.data || []
  const names = data.map((d) => truncateLabel(d.name))
  const values = data.map((d) => d.value)
  const secondaryValues = data.map((d) => d.secondary || 0)
  const hasSecondary = data.some((d) => d.secondary != null)
  const chartType = module.chartType

  const baseGrid = {
    left: 8,
    right: 16,
    top: 16,
    bottom: 24,
    containLabel: true,
  }

  switch (chartType) {
    case 'pie': {
      return {
        color: PALETTE,
        tooltip: { trigger: 'item' },
        legend: { bottom: 6 },
        series: [
          {
            type: 'pie',
            radius: ['0%', '72%'],
            center: ['50%', '48%'],
            label: { formatter: '{b} {d}%' },
            data: data.map((d, i) => ({
              name: d.name,
              value: d.value,
              itemStyle: { color: PALETTE[i % PALETTE.length] },
            })),
          },
        ],
      }
    }

    case 'doughnut': {
      return {
        color: PALETTE,
        tooltip: { trigger: 'item' },
        legend: { bottom: 6 },
        series: [
          {
            type: 'pie',
            radius: ['48%', '76%'],
            center: ['50%', '48%'],
            label: { formatter: '{b} {d}%' },
            data: data.map((d, i) => ({
              name: d.name,
              value: d.value,
              itemStyle: { color: PALETTE[i % PALETTE.length] },
            })),
          },
        ],
      }
    }

    case 'bar':
    case 'stacked-bar':
    case 'grouped-bar': {
      return {
        color: [PALETTE[1]],
        tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
        legend: {},
        grid: baseGrid,
        xAxis: {
          type: 'category',
          data: names,
          axisLabel: { fontSize: 10 },
        },
        yAxis: { type: 'value', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'bar',
            name: '指标值',
            stack: chartType === 'stacked-bar' ? 'bar' : undefined,
            barMaxWidth: 32,
            itemStyle: { borderRadius: [3, 3, 0, 0] },
            data: data.map((d, i) => ({
              value: d.value,
              itemStyle: { color: PALETTE[i % PALETTE.length] },
            })),
          },
          ...(hasSecondary
            ? [
                {
                  type: 'bar' as const,
                  name: '指标二',
                  stack: chartType === 'stacked-bar' ? 'bar' : undefined,
                  barMaxWidth: 32,
                  itemStyle: { color: PALETTE[1], borderRadius: [3, 3, 0, 0] },
                  data: secondaryValues,
                },
              ]
            : []),
        ],
      }
    }

    case 'line': {
      return {
        color: [PALETTE[0], PALETTE[2]],
        tooltip: { trigger: 'axis' },
        legend: {},
        grid: baseGrid,
        xAxis: {
          type: 'category',
          boundaryGap: false,
          data: names,
          axisLabel: { fontSize: 10 },
        },
        yAxis: { type: 'value', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'line',
            name: '指标值',
            smooth: true,
            symbol: 'circle',
            symbolSize: 6,
            lineStyle: { width: 2.5 },
            data: values,
          },
          ...(hasSecondary
            ? [
                {
                  type: 'line' as const,
                  name: '指标二',
                  smooth: true,
                  symbol: 'circle',
                  symbolSize: 5,
                  lineStyle: { width: 2 },
                  data: secondaryValues,
                },
              ]
            : []),
        ],
      }
    }

    case 'area':
    case 'stacked-area': {
      const stacked = chartType === 'stacked-area'
      return {
        color: [PALETTE[0], PALETTE[2]],
        tooltip: { trigger: 'axis' },
        legend: {},
        grid: baseGrid,
        xAxis: {
          type: 'category',
          boundaryGap: false,
          data: names,
          axisLabel: { fontSize: 10 },
        },
        yAxis: { type: 'value', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'line',
            name: '指标值',
            stack: stacked ? 'area' : undefined,
            smooth: true,
            symbol: 'circle',
            symbolSize: 6,
            lineStyle: { width: 2.5 },
            areaStyle: { opacity: 0.35 },
            data: values,
          },
          ...(hasSecondary
            ? [
                {
                  type: 'line' as const,
                  name: '指标二',
                  stack: stacked ? 'area' : undefined,
                  smooth: true,
                  symbol: 'circle',
                  symbolSize: 5,
                  lineStyle: { width: 2 },
                  areaStyle: { opacity: 0.35 },
                  data: secondaryValues,
                },
              ]
            : []),
        ],
      }
    }

    case 'combo': {
      return {
        color: [PALETTE[1], PALETTE[2]],
        tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
        legend: {},
        grid: baseGrid,
        xAxis: {
          type: 'category',
          data: names,
          axisLabel: { fontSize: 10 },
        },
        yAxis: { type: 'value', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'bar',
            name: '实际值',
            barMaxWidth: 32,
            itemStyle: { color: PALETTE[1], borderRadius: [3, 3, 0, 0] },
            data: values,
          },
          {
            type: 'line',
            name: '对比/预测值',
            smooth: true,
            symbol: 'circle',
            symbolSize: 6,
            lineStyle: { width: 2.5 },
            itemStyle: { color: PALETTE[2] },
            data: secondaryValues,
          },
        ],
      }
    }

    case 'scatter': {
      return {
        color: [PALETTE[1]],
        tooltip: { trigger: 'item' },
        grid: baseGrid,
        xAxis: { type: 'value', name: '指标一', axisLabel: { fontSize: 10 } },
        yAxis: { type: 'value', name: '指标二', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'scatter',
            symbolSize: 9,
            data: data.map((d) => [d.secondary || 0, d.value]),
          },
        ],
      }
    }

    case 'bubble': {
      return {
        color: [PALETTE[1]],
        tooltip: { trigger: 'item' },
        grid: baseGrid,
        xAxis: { type: 'value', name: '指标一', axisLabel: { fontSize: 10 } },
        yAxis: { type: 'value', name: '指标二', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'scatter',
            symbolSize: (val: number[]) => val[2],
            data: data.map((d) => [
              d.secondary || 0,
              d.value,
              Math.max(8, Math.min(64, d.value)),
            ]),
          },
        ],
      }
    }

    case 'histogram': {
      const maximum = Math.max(...values)
      const minimum = Math.min(...values)
      const bucketCount = Math.max(
        1,
        Math.min(10, Math.ceil(Math.sqrt(values.length || 1))),
      )
      const bucketSize = Math.max(
        1e-9,
        (maximum - minimum) / bucketCount || 1,
      )
      const histogram = Array.from({ length: bucketCount }, (_, index) => ({
        name: `${Math.round(minimum + index * bucketSize)}-${Math.round(
          minimum + (index + 1) * bucketSize,
        )}`,
        value: 0,
      }))
      values.forEach((value) => {
        const idx = Math.min(
          bucketCount - 1,
          Math.floor((value - minimum) / bucketSize),
        )
        if (histogram[idx]) histogram[idx].value += 1
      })
      return {
        color: [PALETTE[1]],
        tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
        grid: baseGrid,
        xAxis: {
          type: 'category',
          data: histogram.map((b) => b.name),
          axisLabel: { fontSize: 9 },
        },
        yAxis: {
          type: 'value',
          minInterval: 1,
          axisLabel: { fontSize: 10 },
        },
        series: [
          {
            type: 'bar',
            name: '频数',
            barMaxWidth: 32,
            itemStyle: { color: PALETTE[1], borderRadius: [3, 3, 0, 0] },
            data: histogram.map((b) => b.value),
          },
        ],
      }
    }

    case 'radar': {
      const maxVal = Math.max(...values) || 1
      const indicators = data.map((d) => ({ name: d.name, max: maxVal }))
      const series: Record<string, any>[] = [
        {
          type: 'radar' as const,
          areaStyle: { opacity: 0.42 },
          data: [{ value: values, name: '指标值' }],
        },
      ]
      if (hasSecondary) {
        series.push({
          type: 'radar' as const,
          areaStyle: { opacity: 0.28 },
          itemStyle: { color: PALETTE[1] },
          lineStyle: { color: PALETTE[1] },
          data: [{ value: secondaryValues, name: '指标二' }],
        })
      }
      return {
        color: [PALETTE[0], PALETTE[1]],
        tooltip: {},
        legend: {},
        radar: {
          radius: '66%',
          indicator: indicators,
        },
        series,
      }
    }

    case 'waterfall': {
      let running = 0
      const rows = data.map((d) => {
        const base = running
        running += d.value
        return { name: d.name, base, value: d.value }
      })
      return {
        tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
        grid: baseGrid,
        xAxis: {
          type: 'category',
          data: rows.map((r) => r.name),
          axisLabel: { fontSize: 9 },
        },
        yAxis: { type: 'value', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'bar',
            stack: 'waterfall',
            itemStyle: { color: 'transparent' },
            emphasis: { itemStyle: { color: 'transparent' } },
            data: rows.map((r) => r.base),
          },
          {
            type: 'bar',
            stack: 'waterfall',
            name: '变化值',
            barMaxWidth: 32,
            itemStyle: { color: PALETTE[0] },
            label: { show: true, position: 'top', fontSize: 10 },
            data: rows.map((r) => r.value),
          },
        ],
      }
    }

    case 'funnel': {
      const sorted = [...data].sort((a, b) => b.value - a.value)
      return {
        color: PALETTE,
        tooltip: { trigger: 'item', formatter: '{b} : {c}' },
        series: [
          {
            type: 'funnel',
            left: '10%',
            width: '80%',
            gap: 2,
            label: { show: true, position: 'inside', fontSize: 11 },
            data: sorted.map((d, i) => ({
              name: d.name,
              value: Math.max(0, d.value),
              itemStyle: { color: PALETTE[i % PALETTE.length] },
            })),
          },
        ],
      }
    }

    case 'treemap': {
      return {
        tooltip: { formatter: '{b}: {c}' },
        series: [
          {
            type: 'treemap',
            roam: false,
            nodeClick: false,
            breadcrumb: { show: false },
            label: { fontSize: 11 },
            itemStyle: {
              borderColor: '#fff',
              borderWidth: 2,
              gapWidth: 2,
            },
            data: data.map((d, i) => ({
              name: d.name,
              value: Math.max(0, d.value),
              itemStyle: { color: PALETTE[i % PALETTE.length] },
            })),
          },
        ],
      }
    }

    default: {
      // Fallback: treat as bar.
      return {
        color: [PALETTE[1]],
        tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
        legend: {},
        grid: baseGrid,
        xAxis: {
          type: 'category',
          data: names,
          axisLabel: { fontSize: 10 },
        },
        yAxis: { type: 'value', axisLabel: { fontSize: 10 } },
        series: [
          {
            type: 'bar',
            name: '指标值',
            barMaxWidth: 32,
            itemStyle: { borderRadius: [3, 3, 0, 0] },
            data: data.map((d, i) => ({
              value: d.value,
              itemStyle: { color: PALETTE[i % PALETTE.length] },
            })),
          },
        ],
      }
    }
  }
}

function handleResize() {
  if (chart) chart.resize()
}

onMounted(() => {
  if (!chartRef.value) return
  chart = echarts.init(chartRef.value)
  chart.setOption(buildOption())
  window.addEventListener('resize', handleResize)
})

onBeforeUnmount(() => {
  window.removeEventListener('resize', handleResize)
  if (chart) {
    chart.dispose()
    chart = null
  }
})

watch(
  () => props.module,
  () => {
    if (chart) chart.setOption(buildOption(), true)
  },
  { deep: true },
)
</script>

<template>
  <div v-if="!module.data || module.data.length === 0" class="chat-chart-empty">
    <el-icon><DataLine /></el-icon>
    <span>当前口径没有可绘制数据</span>
  </div>
  <div
    v-else
    ref="chartRef"
    class="analysis-chart"
    style="width: 100%; height: 100%; min-height: 280px"
  />
</template>

<style scoped>
.analysis-chart {
  width: 100%;
  height: 100%;
  min-height: 280px;
}
.chat-chart-empty {
  display: grid;
  place-content: center;
  justify-items: center;
  gap: 8px;
  height: 100%;
  color: #718078;
  font-size: 13px;
}
.chat-chart-empty :deep(svg) {
  width: 22px;
  height: 22px;
}
</style>
