<template>
  <div class="new-view-panel" role="region" aria-label="新建数据视图">
    <div class="new-view-header">
      <button class="icon-button inspector-back" type="button" aria-label="返回" title="返回" @click="emit('close')">
        <ArrowLeft :size="18" />
      </button>
      <div class="object-icon"><Sparkles :size="18" /></div>
      <div class="panel-title"><span>新建数据视图</span><h2>创建一个数据视图</h2></div>
      <button class="icon-button close-button" type="button" aria-label="关闭新建视图" title="关闭" @click="emit('close')">
        <X :size="18" />
      </button>
    </div>

    <form class="new-view-form" @submit.prevent="handleCreate">
      <div class="new-view-scroll">
        <section class="form-stack">
          <div class="field-label">
            视图类型
            <div class="chart-type-grid">
              <button
                v-for="option in chartOptions"
                :key="option.id"
                type="button"
                :class="{ active: chartType === option.id }"
                @click="selectChartType(option.id)"
              >
                <component :is="option.icon" :size="17" />
                <span><strong>{{ option.name }}</strong><small>{{ option.description }}</small></span>
              </button>
            </div>
          </div>

          <div class="field-label">
            数据来源
            <div class="sheet-source-picker">
              <div class="sheet-source-list">
                <label
                  v-for="sheet in sheets"
                  :key="sheet.id"
                  class="sheet-source-option"
                  :class="{ selected: sourceSheetIds.includes(sheet.id) }"
                >
                  <input
                    type="checkbox"
                    :checked="sourceSheetIds.includes(sheet.id)"
                    @change="toggleSheet(sheet.id)"
                  />
                  <span class="sheet-source-copy">
                    <strong>{{ sheet.name }}</strong>
                    <small>{{ sheet.rows.length }} 行 · {{ sheet.columns.length }} 列{{ sheet.id === activeSheetId ? ' · 当前预览' : '' }}</small>
                  </span>
                </label>
              </div>
              <div class="sheet-source-summary">
                <Database :size="13" />已选择 {{ sourceSheetIds.length }} 个工作表，AI 会自动识别公共键并关联
              </div>
            </div>
          </div>

          <label class="field-label">
            视图名称
            <input v-model="name" placeholder="例如：销售趋势" />
          </label>

          <label class="field-label">
            你希望看到什么
            <textarea
              v-model="description"
              rows="4"
              placeholder="例如：按进度分组的可拖拽看板；或按负责人展示甘特日历；也可以是基于表格卡牌数据的互动游戏。"
            />
          </label>

          <div v-if="chartType === 'custom'" class="field-label">
            参考图片（{{ images.length }}/{{ REFERENCE_IMAGE_LIMIT }}）
            <div class="reference-grid">
              <div v-for="image in images" :key="image.id" class="reference-image">
                <img :src="image.dataUrl" :alt="image.name" />
                <button type="button" :aria-label="`删除 ${image.name}`" @click="removeImage(image.id)"><X :size="13" /></button>
              </div>
              <label v-if="images.length < REFERENCE_IMAGE_LIMIT" class="image-upload">
                <LoaderCircle v-if="processingImages" :size="18" class="spin" />
                <ImagePlus v-else :size="18" />
                <span>{{ processingImages ? '处理中' : '添加图片' }}</span>
                <input type="file" accept="image/*" multiple :disabled="processingImages" @change="handleImages" />
              </label>
            </div>
            <span v-if="imageError" class="field-error">{{ imageError }}</span>
          </div>

          <div class="inline-note"><WandSparkles :size="15" />单表预制图表在本地生成；选择多表时由 AI 识别公共键并生成跨表视图。</div>
        </section>
      </div>

      <div class="new-view-footer">
        <button type="button" class="button secondary" @click="emit('close')">取消</button>
        <button class="button primary" :disabled="!name.trim() || !sourceSheetIds.length || processingImages">
          <Sparkles :size="16" />创建视图
        </button>
      </div>
    </form>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import {
  ArrowLeft, AreaChart, BarChart3, CircleDot, Database, Filter, GitBranch, ImagePlus,
  LineChart, LoaderCircle, PieChart, Radar, ScatterChart, Sparkles, WandSparkles, X,
} from 'lucide-vue-next'
import { CHART_OPTIONS, chartDescription } from '../lib/chart-templates'
import { createId } from '../lib/core'
import { REFERENCE_IMAGE_LIMIT } from '../lib/generated-app-runtime'
import type { ChartType, ReferenceImage, SheetDefinition } from '../lib/types'

const props = defineProps<{
  sheets: SheetDefinition[]
  activeSheetId?: string
}>()

const emit = defineEmits<{
  close: []
  create: [payload: { name: string; description: string; chartType: ChartType; images: ReferenceImage[]; sourceSheetIds: string[] }]
}>()

const iconMap = {
  custom: Sparkles,
  pie: PieChart,
  bar: BarChart3,
  line: LineChart,
  area: AreaChart,
  scatter: ScatterChart,
  bubble: CircleDot,
  histogram: BarChart3,
  radar: Radar,
  waterfall: BarChart3,
  funnel: Filter,
  treemap: GitBranch,
}
const chartOptions = CHART_OPTIONS.map((option) => ({ ...option, icon: iconMap[option.id] }))
const name = ref('')
const description = ref('')
const chartType = ref<ChartType>('custom')
const images = ref<ReferenceImage[]>([])
const imageError = ref('')
const processingImages = ref(false)
const sourceSheetIds = ref(props.activeSheetId ? [props.activeSheetId] : props.sheets[0] ? [props.sheets[0].id] : [])

function selectChartType(type: ChartType) {
  chartType.value = type
  if (type !== 'custom') {
    const option = CHART_OPTIONS.find((item) => item.id === type)
    name.value = option?.name || ''
    description.value = chartDescription(type)
  }
}

function toggleSheet(sheetId: string) {
  if (sourceSheetIds.value.includes(sheetId)) {
    if (sourceSheetIds.value.length > 1) sourceSheetIds.value = sourceSheetIds.value.filter((id) => id !== sheetId)
  } else {
    sourceSheetIds.value = [...sourceSheetIds.value, sheetId]
  }
}

async function imageToDataUrl(file: File) {
  const source = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(new Error('图片读取失败'))
    reader.readAsDataURL(file)
  })
  const image = await new Promise<HTMLImageElement>((resolve, reject) => {
    const element = new Image()
    element.onload = () => resolve(element)
    element.onerror = () => reject(new Error('图片格式无法识别'))
    element.src = source
  })
  const canvas = document.createElement('canvas')
  let maxEdge = 1400
  let quality = 0.82
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const scale = Math.min(1, maxEdge / Math.max(image.width, image.height))
    canvas.width = Math.max(1, Math.round(image.width * scale))
    canvas.height = Math.max(1, Math.round(image.height * scale))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('浏览器无法处理图片')
    context.drawImage(image, 0, 0, canvas.width, canvas.height)
    const dataUrl = canvas.toDataURL('image/jpeg', quality)
    if (Math.ceil(dataUrl.length * 0.75) <= 480 * 1024) return dataUrl
    maxEdge *= 0.82
    quality = Math.max(0.5, quality - 0.05)
  }
  throw new Error('图片压缩后仍超过 480KB，请选择更小的图片')
}

async function handleImages(event: Event) {
  const input = event.target as HTMLInputElement
  const files = Array.from(input.files || []).slice(0, REFERENCE_IMAGE_LIMIT - images.value.length)
  input.value = ''
  if (!files.length) return
  processingImages.value = true
  imageError.value = ''
  try {
    for (const file of files) {
      if (!file.type.startsWith('image/')) continue
      images.value.push({ id: createId('image'), name: file.name, dataUrl: await imageToDataUrl(file) })
    }
  } catch (error) {
    imageError.value = error instanceof Error ? error.message : '图片处理失败'
  } finally {
    processingImages.value = false
  }
}

function removeImage(imageId: string) {
  images.value = images.value.filter((image) => image.id !== imageId)
}

function handleCreate() {
  if (!name.value.trim() || !sourceSheetIds.value.length || processingImages.value) return
  emit('create', {
    name: name.value.trim(),
    description: description.value.trim(),
    chartType: chartType.value,
    images: images.value,
    sourceSheetIds: sourceSheetIds.value,
  })
}
</script>

<style scoped>
.new-view-panel { --green: #19734a; --line: #dfe4e1; min-height: 0; height: 100%; display: flex; flex-direction: column; background: #fff; color: #243129; container-type: inline-size; }
.new-view-header { min-height: 74px; display: flex; align-items: center; gap: 11px; padding: 13px 16px; border-bottom: 1px solid var(--line); }
.icon-button { width: 30px; height: 30px; display: grid; place-items: center; flex: 0 0 auto; border: 1px solid #d7ded9; border-radius: 6px; background: #fff; color: #3f4e46; cursor: pointer; }
.icon-button:hover { border-color: #aebdb4; background: #edf4f0; color: #176b46; }
.object-icon { width: 34px; height: 34px; display: grid; place-items: center; flex: 0 0 auto; border-radius: 7px; color: #176b46; background: #e4f2ea; }
.panel-title { min-width: 0; flex: 1; }
.panel-title span { display: block; color: #839088; font-size: 9px; line-height: 1.2; font-weight: 720; }
.panel-title h2 { margin: 3px 0 0; font-size: 14px; line-height: 1.25; font-weight: 680; }
.close-button { border-color: transparent; }
.new-view-form { min-height: 0; flex: 1; display: flex; flex-direction: column; }
.new-view-scroll { flex: 1; min-height: 0; overflow-y: auto; }
.form-stack { display: grid; gap: 14px; padding: 16px; }
.field-label { display: grid; gap: 7px; color: #55625a; font-size: 11px; font-weight: 630; }
.field-label input, .field-label textarea { width: 100%; min-height: 36px; box-sizing: border-box; border: 1px solid #d4dbd7; border-radius: 6px; background: #fff; padding: 8px 10px; color: #243129; font: inherit; font-size: 12px; resize: vertical; }
.field-label input:focus, .field-label textarea:focus { border-color: #4b9a72; outline: 2px solid #e2f0e8; }
.chart-type-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 7px; }
.chart-type-grid button { min-width: 0; min-height: 58px; display: flex; align-items: center; gap: 8px; border: 1px solid var(--line); border-radius: 6px; padding: 8px; background: #fff; color: #546159; text-align: left; cursor: pointer; }
.chart-type-grid button:hover { border-color: #9fc8b1; background: #f5faf7; }
.chart-type-grid button.active { border-color: #4b9a72; background: #edf7f1; color: #176b46; box-shadow: inset 0 0 0 1px #b7d8c4; }
.chart-type-grid button > svg { flex: 0 0 auto; }
.chart-type-grid button span { min-width: 0; display: grid; gap: 2px; }
.chart-type-grid button strong { font-size: 10px; white-space: nowrap; }
.chart-type-grid button small { color: #849087; font-size: 8px; line-height: 1.25; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sheet-source-picker { overflow: hidden; border: 1px solid #d9dfdb; border-radius: 6px; background: #fff; }
.sheet-source-list { max-height: 190px; overflow-y: auto; }
.sheet-source-option { min-height: 43px; display: flex; align-items: center; gap: 9px; padding: 7px 9px; border-bottom: 1px solid #edf0ee; color: #4d5b53; cursor: pointer; }
.sheet-source-option:last-child { border-bottom: 0; }
.sheet-source-option:hover, .sheet-source-option.selected { background: #f1f8f4; }
.sheet-source-option input { width: 15px; min-height: 15px; height: 15px; flex: 0 0 auto; padding: 0; accent-color: var(--green); }
.sheet-source-copy { min-width: 0; display: grid; gap: 2px; }
.sheet-source-copy strong { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #35433b; font-size: 10px; }
.sheet-source-copy small { color: #849087; font-size: 9px; font-weight: 450; }
.sheet-source-summary { min-height: 31px; display: flex; align-items: center; gap: 7px; padding: 6px 9px; border-top: 1px solid #dfe7e2; background: #f6f9f7; color: #557064; font-size: 9px; font-weight: 500; }
.reference-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(90px, 1fr)); gap: 7px; }
.reference-image, .image-upload { position: relative; aspect-ratio: 1.25; min-width: 0; border: 1px solid #d8deda; border-radius: 5px; overflow: hidden; background: #f6f8f7; }
.reference-image img { width: 100%; height: 100%; object-fit: cover; }
.reference-image button { position: absolute; top: 3px; right: 3px; display: grid; place-items: center; width: 20px; height: 20px; border: 0; border-radius: 4px; background: rgba(24,35,29,.82); color: #fff; cursor: pointer; }
.image-upload { display: grid; place-items: center; align-content: center; gap: 4px; color: #758179; font-size: 9px; cursor: pointer; }
.image-upload:hover { border-color: #8dbda2; color: #176b46; background: #f1f8f4; }
.image-upload input { display: none; }
.field-error { color: #a63333; font-size: 9px; font-weight: 500; }
.inline-note { display: flex; align-items: flex-start; gap: 8px; padding: 9px 10px; border-radius: 5px; background: #edf6f1; color: #32654d; font-size: 10px; line-height: 1.5; }
.inline-note svg { flex: 0 0 auto; }
.new-view-footer { flex: 0 0 auto; display: flex; justify-content: flex-end; gap: 8px; padding: 12px 16px; border-top: 1px solid var(--line); background: #fafbfa; }
.button { min-height: 36px; display: inline-flex; align-items: center; justify-content: center; gap: 7px; border-radius: 6px; border: 1px solid transparent; padding: 0 13px; font-size: 12px; font-weight: 630; cursor: pointer; white-space: nowrap; }
.button:disabled { opacity: .48; cursor: not-allowed; }
.button.primary { color: #fff; background: var(--green); border-color: var(--green); }
.button.secondary { color: #435148; background: #fff; border-color: #cfd7d2; }
.spin { animation: spin .85s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
@container (min-width: 560px) { .chart-type-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); } }
@container (min-width: 700px) { .chart-type-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); } }
</style>
