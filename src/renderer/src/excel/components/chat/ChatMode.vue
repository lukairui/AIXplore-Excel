<script setup name="AiExcelChatMode" lang="ts">
// @ts-nocheck —— 遗留文件：ChatAnalysisResult 递归类型在 vue-tsc computed 推断中触发 TS2589，待上游类型重构
import { ref, computed, onBeforeUnmount, onMounted, watch, toRaw } from 'vue'
import { ElMessageBox } from 'element-plus'
import {
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  CaretRight,
  Check,
  ChatDotRound,
  Close,
  Coin,
  DataLine,
  Delete,
  Download,
  Files,
  Grid,
  Loading,
  MagicStick,
  Plus,
  Promotion,
  Refresh,
  Setting,
  Upload,
  View,
  Warning,
} from '@element-plus/icons-vue'
import {
  CHAT_ANALYSIS_STEPS,
  chatResultToMarkdown,
  getChatSourceSheets,
  importChatSource,
  runChatAnalysis,
  type ChatAnalysisPhase,
  type ChatAnalysisResult,
  type ChatSourceFile,
  type ChatSourceSheet,
} from '../../lib/chat-analysis'
import { calculateRuntimeRows } from '../../lib/core'
import { projectToExcel } from '../../lib/excel'
import { listExcelProviders } from '@/api/excelTool'
import AnalysisChart from './AnalysisChart.vue'
import type { Project } from '../../lib/types'

type ChatMessage = {
  id: string
  role: 'user' | 'assistant'
  content: string
  attachment?: ChatArtifact
}

type ChatArtifact = {
  id: string
  name: string
  project: Project
}

type ChatReportTab = {
  id: string
  label: string
  result: ChatAnalysisResult
}

const props = defineProps({
  modelConfigured: { type: Boolean, default: false },
  modelName: { type: String, default: '' },
  provider: { type: String, default: '' },
})

const emit = defineEmits(['back', 'open-settings', 'change-provider'])

// === 模型 provider 列表（用于顶栏快速切换） ===
const providers = ref<Array<{ provider: string; providerName: string; modelName: string }>>([])
async function loadProviders() {
  try {
    const res = await listExcelProviders()
    const raw = ((res as any).data || []).map((p: any) => ({
      provider: p.provider,
      providerName: p.providerName,
      modelName: p.modelName,
    }))
    providers.value = raw
    if ((!props.provider || !providers.value.some((p) => p.provider === props.provider)) && providers.value.length) {
      emit('change-provider', providers.value[0].provider)
    }
  } catch {
    // 静默失败
  }
}

function onQuickChangeProvider(provider: string) {
  emit('change-provider', provider)
}

const DEFAULT_QUESTION = ''
const QUESTION_PLACEHOLDER =
  '可以针对上传的 Excel 提出任何需求，比如数据分析、修改列、新增列、修改行、绘制柱状图、饼图、折线图等......'

function sourceSize(size: number) {
  if (size < 1024) return `${size} B`
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`
  return `${(size / 1024 / 1024).toFixed(1)} MB`
}

function saveBlob(name: string, blob: Blob) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = name
  anchor.click()
  URL.revokeObjectURL(url)
}

function previewValue(value: unknown) {
  if (value == null || value === '') return ''
  if (typeof value === 'number') return new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 }).format(value)
  if (typeof value === 'boolean') return value ? '是' : '否'
  return String(value)
}

function clone<T>(value: T): T {
  if (typeof (globalThis as any).structuredClone === 'function') {
    return (globalThis as any).structuredClone(value)
  }
  return JSON.parse(JSON.stringify(value)) as T
}

// The chat-analysis lib exposes a BrowserModelConfig-typed surface but only the
// `provider` field is read at runtime (it is passed straight through to
// requestModelText, which itself accepts a bare provider string). We bridge the
// single-provider model layer to that surface here.
function providerConfig() {
  return {
    provider: props.provider,
    baseUrl: '',
    model: '',
    apiKey: '',
    apiMode: 'chat-completions' as const,
  }
}

// === State ===
const fileInputRef = ref<HTMLInputElement | null>(null)
const previewScrollRef = ref<HTMLDivElement | null>(null)

const sources = ref<ChatSourceFile[]>([])
const selectedSheetKey = ref('')
const expandedFiles = ref<Set<string>>(new Set())
const dragging = ref(false)
const phase = ref<ChatAnalysisPhase>('idle')
const question = ref(DEFAULT_QUESTION)
const messages = ref<ChatMessage[]>([])
const result = ref<ChatAnalysisResult | null>(null)
const reportTabs = ref<ChatReportTab[]>([])
const activeReportId = ref('')
const previewArtifact = ref<ChatArtifact | null>(null)
const activeResultTab = ref('overview')
const viewMode = ref<'preview' | 'result' | 'artifact'>('preview')
const error = ref('')
const conversationWidth = ref(340)
const resizingConversation = ref(false)
const horizontalPosition = ref(0)

let analysisController: AbortController | null = null
let reportSequence = 0
const conversationResizeOrigin = { startX: 0, startWidth: 340 }

// === Computed ===
const sourceSheets = computed(() => (sources.value as any[]).flatMap((source) => getChatSourceSheets(source)))
const selectedSheet = computed(
  () => sourceSheets.value.find((item) => item.key === selectedSheetKey.value) || sourceSheets.value[0],
)
const artifactSheet = computed(() => {
  const project = previewArtifact.value?.project
  if (!project) return undefined
  return project.sheets?.find((sheet) => sheet.id === project.activeSheetId) || project.sheets?.[0]
})
const artifactSourceSheet = computed<ChatSourceSheet | undefined>(() => {
  const artifact = previewArtifact.value
  const sheet = artifactSheet.value
  if (!artifact || !sheet) return undefined
  return {
    key: `${artifact.id}:${sheet.id}`,
    sourceId: artifact.id,
    sourceName: artifact.name,
    sheet,
  }
})
const busy = computed(() => phase.value === 'importing' || CHAT_ANALYSIS_STEPS.some((step) => step.phase === phase.value))
const displayedResult = computed(
  (): ChatAnalysisResult | null => (reportTabs.value.find((tab) => tab.id === activeReportId.value)?.result as ChatAnalysisResult) || null,
)
const hasReport = computed(() => Boolean(displayedResult.value))
const currentStep = computed(() => CHAT_ANALYSIS_STEPS.findIndex((step) => step.phase === phase.value))
const latestUserMessageId = computed(() => {
  for (let i = messages.value.length - 1; i >= 0; i -= 1) {
    if (messages.value[i].role === 'user') return messages.value[i].id
  }
  return undefined
})
const activeModule = computed(() => displayedResult.value?.modules.find((module) => module.id === activeResultTab.value))

const conversationMaxWidth = computed(() => {
  const sourceWidth = window.innerWidth <= 1100 ? 230 : 270
  return Math.max(280, Math.min(720, window.innerWidth - sourceWidth - 328))
})

// === Methods ===
async function addFiles(files: File[]) {
  if (!files.length || busy.value) return
  phase.value = 'importing'
  error.value = ''
  const imported: ChatSourceFile[] = []
  const errors: string[] = []
  const outcomes = await Promise.allSettled(files.map((file) => importChatSource(file, providerConfig())))
  outcomes.forEach((outcome, index) => {
    if (outcome.status === 'fulfilled') imported.push(outcome.value)
    else errors.push(outcome.reason instanceof Error ? outcome.reason.message : `${files[index].name} 导入失败`)
  })
  if (imported.length) {
    sources.value = [...sources.value, ...imported]
    expandedFiles.value = new Set([...expandedFiles.value, ...imported.map((source) => source.id)])
    const firstSheet = getChatSourceSheets(imported[0])[0]
    if (!selectedSheetKey.value && firstSheet) selectedSheetKey.value = firstSheet.key
    result.value = null
    previewArtifact.value = null
    activeResultTab.value = 'overview'
    viewMode.value = 'preview'
  }
  if (errors.length) error.value = errors.join('；')
  phase.value = 'idle'
}

function removeSource(sourceId: string) {
  const next = sources.value.filter((source) => source.id !== sourceId)
  sources.value = next
  const nextSheets = next.flatMap(getChatSourceSheets)
  if (!nextSheets.some((item) => item.key === selectedSheetKey.value)) {
    selectedSheetKey.value = nextSheets[0]?.key || ''
  }
  result.value = null
  previewArtifact.value = null
  viewMode.value = 'preview'
  activeResultTab.value = 'overview'
}

async function runAnalysis(nextQuestion: string = question.value) {
  const trimmed = nextQuestion.trim()
  if (!sources.value.length) {
    error.value = '请先上传至少一个数据文件'
    return
  }
  if (!props.modelConfigured) {
    emit('open-settings')
    return
  }
  if (!trimmed || busy.value) return
  analysisController?.abort()
  const controller = new AbortController()
  analysisController = controller
  error.value = ''
  messages.value = [
    ...messages.value,
    { id: `user-${Date.now()}`, role: 'user', content: trimmed },
  ]
  question.value = ''
  try {
    const analysisSources: ChatSourceFile[] = result.value?.processedProject
      ? [
          {
            id: `processed-${result.value.id}`,
            name: '上轮处理结果.xlsx',
            size: 0,
            project: result.value.processedProject,
            addedAt: Date.now(),
          },
        ]
      : sources.value
    const next = await runChatAnalysis({
      sources: analysisSources,
      question: trimmed,
      browserConfig: providerConfig(),
      previous: result.value,
      onPhase: (p) => (phase.value = p),
      signal: controller.signal,
    })
    result.value = next
    if (next.reportGenerated !== false) {
      reportSequence += 1
      const tab = { id: next.id, label: `分析结果${reportSequence}`, result: next }
      reportTabs.value = [...reportTabs.value, tab]
      activeReportId.value = next.id
    }
    const version = messages.value.filter((message) => message.attachment).length + 1
    const attachment: ChatArtifact | undefined =
      next.processedProject && next.operationSummaries?.length
        ? {
            id: `artifact-${Date.now()}`,
            name: `AI Excel 处理结果_v${version}.xlsx`,
            project: next.processedProject,
          }
        : undefined
    messages.value = [
      ...messages.value,
      { id: `assistant-${Date.now()}`, role: 'assistant', content: next.executiveSummary, attachment },
    ]
    activeResultTab.value = 'overview'
    viewMode.value = next.reportGenerated === false ? 'preview' : 'result'
  } catch (analysisError) {
    if (controller.signal.aborted) {
      phase.value = 'idle'
    } else {
      phase.value = 'error'
      error.value = analysisError instanceof Error ? analysisError.message : '分析失败'
    }
  } finally {
    if (analysisController === controller) analysisController = null
  }
}

function stopAnalysis() {
  analysisController?.abort()
  analysisController = null
  phase.value = 'idle'
}

function resetAnalysis() {
  const proceed = () => {
    analysisController?.abort()
    result.value = null
    reportTabs.value = []
    activeReportId.value = ''
    reportSequence = 0
    previewArtifact.value = null
    messages.value = []
    question.value = DEFAULT_QUESTION
    error.value = ''
    phase.value = 'idle'
    activeResultTab.value = 'overview'
    viewMode.value = 'preview'
  }
  if (!reportTabs.value.length) {
    proceed()
    return
  }
  ElMessageBox.confirm(
    `新分析将清空并删除当前 ${reportTabs.value.length} 份分析报告，是否继续？`,
    '新分析',
    { type: 'warning', confirmButtonText: '继续', cancelButtonText: '取消' },
  )
    .then(() => proceed())
    .catch(() => {})
}

function exportReport() {
  const current = displayedResult.value
  if (!current) return
  const blob = new Blob([chatResultToMarkdown(current)], { type: 'text/markdown;charset=utf-8' })
  saveBlob(`${current.title.replace(/[\\/:*?"<>|]/g, '_')}.md`, blob)
}

async function exportProcessedWorkbook(sourceProject: Project, fileName: string) {
  error.value = ''
  try {
    const project = clone(toRaw(sourceProject))
    const blob = await projectToExcel(project, calculateRuntimeRows(project.rows, project.columns))
    saveBlob(fileName, blob)
  } catch (exportError) {
    error.value = exportError instanceof Error ? exportError.message : '处理结果导出失败'
  }
}

function closeReport(tab: ChatReportTab) {
  const proceed = () => {
    const nextTabs = reportTabs.value.filter((item) => item.id !== tab.id)
    const previousIndex = reportTabs.value.findIndex((item) => item.id === tab.id)
    reportTabs.value = nextTabs
    if (activeReportId.value === tab.id) {
      const fallback = nextTabs[Math.min(previousIndex, nextTabs.length - 1)]
      activeReportId.value = fallback?.id || ''
      activeResultTab.value = 'overview'
      if (!fallback) viewMode.value = 'preview'
    }
  }
  ElMessageBox.confirm(`关闭后将清空并删除“${tab.label}”，此操作无法撤销。确定关闭吗？`, `关闭${tab.label}`, {
    type: 'warning',
    confirmButtonText: '关闭',
    cancelButtonText: '取消',
  })
    .then(() => proceed())
    .catch(() => {})
}

function backToModes() {
  const proceed = () => emit('back')
  if (!reportTabs.value.length) {
    proceed()
    return
  }
  ElMessageBox.confirm(
    `返回模式选择将清空并删除当前 ${reportTabs.value.length} 份分析报告，是否继续？`,
    '返回模式选择',
    { type: 'warning', confirmButtonText: '继续', cancelButtonText: '取消' },
  )
    .then(() => proceed())
    .catch(() => {})
}

function toggleExpand(sourceId: string) {
  const next = new Set(expandedFiles.value)
  if (next.has(sourceId)) next.delete(sourceId)
  else next.add(sourceId)
  expandedFiles.value = next
}

function selectSheet(item: ChatSourceSheet) {
  selectedSheetKey.value = item.key
  viewMode.value = 'preview'
}

function onFileInputChange(event: Event) {
  const input = event.target as HTMLInputElement
  addFiles(Array.from(input.files || []))
  input.value = ''
}

function onPreviewScroll(event: Event) {
  const target = event.currentTarget as HTMLDivElement
  const maximum = target.scrollWidth - target.clientWidth
  horizontalPosition.value = maximum > 0 ? Math.round((target.scrollLeft / maximum) * 1000) : 0
}

function onHorizontalRange(event: Event) {
  const preview = previewScrollRef.value
  if (!preview) return
  const position = Number((event.target as HTMLInputElement).value)
  preview.scrollLeft = ((preview.scrollWidth - preview.clientWidth) * position) / 1000
  horizontalPosition.value = position
}

function onComposerKeyDown(event: KeyboardEvent) {
  if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
    runAnalysis()
  }
}

function onDragEnter(event: DragEvent) {
  event.preventDefault()
  dragging.value = true
}
function onDragOver(event: DragEvent) {
  event.preventDefault()
  dragging.value = true
}
function onDragLeave(event: DragEvent) {
  if (event.currentTarget === event.target) dragging.value = false
}
function onDrop(event: DragEvent) {
  event.preventDefault()
  dragging.value = false
  addFiles(Array.from(event.dataTransfer?.files || []))
}

// === Conversation resizer ===
function onConversationPointerDown(event: PointerEvent) {
  if (event.button !== 0 || window.innerWidth <= 820) return
  event.preventDefault()
  const panel = (event.currentTarget as HTMLElement).nextElementSibling as HTMLElement | null
  conversationResizeOrigin.startX = event.clientX
  conversationResizeOrigin.startWidth = panel?.getBoundingClientRect().width || conversationWidth.value
  resizingConversation.value = true
}

function onConversationPointerMove(event: PointerEvent) {
  const sourceWidth = window.innerWidth <= 1100 ? 230 : 270
  const maxWidth = Math.max(280, Math.min(720, window.innerWidth - sourceWidth - 328))
  const nextWidth = conversationResizeOrigin.startWidth + conversationResizeOrigin.startX - event.clientX
  conversationWidth.value = Math.min(maxWidth, Math.max(280, nextWidth))
}
function onConversationPointerUp() {
  resizingConversation.value = false
}
function onConversationKeyDown(event: KeyboardEvent) {
  const sourceWidth = window.innerWidth <= 1100 ? 230 : 270
  const maxWidth = Math.max(280, Math.min(720, window.innerWidth - sourceWidth - 328))
  const step = event.shiftKey ? 40 : 16
  if (event.key === 'ArrowLeft') conversationWidth.value = Math.min(maxWidth, conversationWidth.value + step)
  else if (event.key === 'ArrowRight') conversationWidth.value = Math.max(280, conversationWidth.value - step)
  else if (event.key === 'Home') conversationWidth.value = 280
  else if (event.key === 'End') conversationWidth.value = maxWidth
  else return
  event.preventDefault()
}

watch(resizingConversation, (active) => {
  if (active) {
    document.body.classList.add('resizing-chat-conversation')
    window.addEventListener('pointermove', onConversationPointerMove)
    window.addEventListener('pointerup', onConversationPointerUp, { once: true })
  } else {
    document.body.classList.remove('resizing-chat-conversation')
    window.removeEventListener('pointermove', onConversationPointerMove)
    window.removeEventListener('pointerup', onConversationPointerUp)
  }
})

const warnBeforeClose = (event: BeforeUnloadEvent) => {
  event.preventDefault()
  event.returnValue = ''
}
watch(
  () => reportTabs.value.length,
  (count) => {
    if (count) window.addEventListener('beforeunload', warnBeforeClose)
    else window.removeEventListener('beforeunload', warnBeforeClose)
  },
)

onMounted(() => {
  loadProviders()
})

onBeforeUnmount(() => {
  analysisController?.abort()
  window.removeEventListener('pointermove', onConversationPointerMove)
  window.removeEventListener('pointerup', onConversationPointerUp)
  window.removeEventListener('beforeunload', warnBeforeClose)
  document.body.classList.remove('resizing-chat-conversation')
})
</script>

<template>
  <div class="chat-shell">
    <header class="chat-topbar">
      <button class="icon-button" aria-label="返回模式选择" title="返回模式选择" @click="backToModes">
        <el-icon><ArrowLeft /></el-icon>
      </button>
      <div class="brand">
        <div class="brand-mark"><el-icon><Grid /></el-icon></div>
        <span>AI Excel</span>
      </div>
      <div class="topbar-divider" />
      <span class="chat-mode-label">
        <el-icon><ChatDotRound /></el-icon>Chat 模式
      </span>
      <div class="topbar-spacer" />
      <el-select
        v-if="providers.length"
        :model-value="props.provider"
        size="small"
        class="model-quick-select"
        placeholder="切换模型"
        @update:model-value="onQuickChangeProvider"
      >
        <template #prefix><span class="model-dot online" /></template>
        <el-option
          v-for="item in providers"
          :key="item.provider"
          :value="item.provider"
          :label="item.providerName || item.provider"
        >
          <span>{{ item.providerName || item.provider }}</span>
        </el-option>
      </el-select>
      <button v-if="hasReport" class="button secondary compact" @click="exportReport">
        <el-icon><Download /></el-icon><span class="button-label">导出报告</span>
      </button>
      <button class="button secondary compact" @click="resetAnalysis">
        <el-icon><Refresh /></el-icon><span class="button-label">新分析</span>
      </button>
      <button class="icon-button" aria-label="模型设置" title="模型设置与测试连接" @click="emit('open-settings')">
        <el-icon><Setting /></el-icon>
      </button>
    </header>

    <main class="chat-workspace" :style="{ '--chat-conversation-width': conversationWidth + 'px' }">
      <aside class="chat-sources">
        <div class="chat-panel-heading">
          <div>
            <span class="eyebrow">数据源</span>
            <h2>分析文件</h2>
          </div>
          <button class="icon-button subtle" title="添加文件" aria-label="添加文件" @click="fileInputRef?.click()">
            <el-icon><Plus /></el-icon>
          </button>
        </div>
        <div class="chat-upload-zone" :class="{ dragging }" @dragenter="onDragEnter" @dragover="onDragOver" @dragleave="onDragLeave" @drop="onDrop">
          <el-icon v-if="phase === 'importing'" class="spin"><Loading /></el-icon>
          <el-icon v-else><Upload /></el-icon>
          <strong>{{ phase === 'importing' ? '正在读取文件' : '添加数据文件' }}</strong>
          <span>.xlsx · .csv · .pdf · 图片，可多选</span>
          <button class="button secondary compact" :disabled="busy" @click="fileInputRef?.click()">选择文件</button>
          <input
            ref="fileInputRef"
            hidden
            type="file"
            multiple
            accept=".xlsx,.csv,.pdf,.png,.jpg,.jpeg,.webp,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv,application/pdf,image/*"
            @change="onFileInputChange"
          />
        </div>
        <div class="chat-source-list">
          <div v-for="source in sources" :key="source.id" class="chat-source-file">
            <div class="chat-source-file-row">
              <button
                class="chat-source-expand"
                :aria-label="`${expandedFiles.has(source.id) ? '收起' : '展开'}${source.name}`"
                @click="toggleExpand(source.id)"
              >
                <el-icon><ArrowDown v-if="expandedFiles.has(source.id)" /><CaretRight v-else /></el-icon>
              </button>
              <el-icon class="source-file-icon"><Grid /></el-icon>
              <span>
                <strong :title="source.name">{{ source.name }}</strong>
                <small>
                  {{ getChatSourceSheets(source).length }} 张表 ·
                  {{ getChatSourceSheets(source).reduce((total, item) => total + item.sheet.rows.length, 0).toLocaleString('zh-CN') }} 行 ·
                  {{ sourceSize(source.size) }}
                </small>
              </span>
              <button
                class="icon-button subtle danger-hover"
                :aria-label="`移除 ${source.name}`"
                title="移除文件"
                @click="removeSource(source.id)"
              >
                <el-icon><Delete /></el-icon>
              </button>
            </div>
            <div v-if="expandedFiles.has(source.id)" class="chat-source-sheets">
              <button
                v-for="item in getChatSourceSheets(source)"
                :key="item.key"
                :class="{ active: selectedSheet?.key === item.key }"
                @click="selectSheet(item)"
              >
                <el-icon><Coin /></el-icon>
                <span>{{ item.sheet.name }}</span>
                <small>{{ item.sheet.rows.length }} × {{ item.sheet.columns.length }}</small>
              </button>
            </div>
          </div>
        </div>
      </aside>

      <section class="chat-center">
        <nav class="chat-center-tabs">
          <button :class="{ active: viewMode === 'preview' }" @click="viewMode = 'preview'">
            <el-icon><Grid /></el-icon>文件预览
          </button>
          <span
            v-for="tab in reportTabs"
            :key="tab.id"
            class="chat-center-report-tab"
            :class="{ active: viewMode === 'result' && activeReportId === tab.id }"
          >
            <button
              class="chat-center-report-open"
              @click="activeReportId = tab.id; activeResultTab = 'overview'; viewMode = 'result'"
            >
              <el-icon><MagicStick /></el-icon>{{ tab.label }}
            </button>
            <button
              class="chat-center-report-close"
              :aria-label="`关闭${tab.label}`"
              :title="`关闭${tab.label}`"
              @click="closeReport(tab)"
            >
              <el-icon><Close /></el-icon>
            </button>
          </span>
          <button v-if="previewArtifact" :class="{ active: viewMode === 'artifact' }" @click="viewMode = 'artifact'">
            <el-icon><View /></el-icon>处理文件
          </button>
        </nav>

        <!-- Artifact preview -->
        <template v-if="viewMode === 'artifact'">
          <section v-if="!artifactSourceSheet" class="chat-empty-stage">
            <el-icon><Grid /></el-icon>
            <h2>等待数据文件</h2>
          </section>
          <section v-else class="chat-preview-stage">
            <header class="chat-preview-header">
              <div>
                <span class="eyebrow">文件预览</span>
                <h2>{{ artifactSourceSheet.sourceName }} / {{ artifactSourceSheet.sheet.name }}</h2>
              </div>
              <span>{{ artifactSourceSheet.sheet.rows.length }} 行 × {{ artifactSourceSheet.sheet.columns.length }} 列</span>
            </header>
            <div ref="previewScrollRef" class="chat-preview-scroll" @scroll="onPreviewScroll">
              <table class="chat-preview-table" :style="{ minWidth: Math.max(1800, artifactSourceSheet.sheet.columns.slice(0, 24).length * 160 + 40) + 'px' }">
                <thead>
                  <tr>
                    <th>#</th>
                    <th v-for="column in artifactSourceSheet.sheet.columns.slice(0, 24)" :key="column.id">{{ column.name }}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, rowIndex) in artifactSourceSheet.sheet.rows.slice(0, 80)" :key="row.id">
                    <td>{{ rowIndex + 1 }}</td>
                    <td
                      v-for="column in artifactSourceSheet.sheet.columns.slice(0, 24)"
                      :key="column.id"
                      :title="previewValue(row.values[column.id])"
                    >{{ previewValue(row.values[column.id]) }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div class="chat-preview-horizontal-scroll">
              <input type="range" min="0" max="1000" :value="horizontalPosition" aria-label="表格横向滚动条" @input="onHorizontalRange" />
            </div>
            <footer
              v-if="artifactSourceSheet.sheet.columns.length > artifactSourceSheet.sheet.columns.slice(0, 24).length || artifactSourceSheet.sheet.rows.length > artifactSourceSheet.sheet.rows.slice(0, 80).length"
              class="chat-preview-limit"
            >预览前 {{ artifactSourceSheet.sheet.rows.slice(0, 80).length }} 行、{{ artifactSourceSheet.sheet.columns.slice(0, 24).length }} 列，分析使用全部有效数据</footer>
          </section>
        </template>

        <!-- Selected sheet preview -->
        <template v-else-if="viewMode === 'preview'">
          <section v-if="!selectedSheet" class="chat-empty-stage">
            <el-icon><Grid /></el-icon>
            <h2>等待数据文件</h2>
          </section>
          <section v-else class="chat-preview-stage">
            <header class="chat-preview-header">
              <div>
                <span class="eyebrow">文件预览</span>
                <h2>{{ selectedSheet.sourceName }} / {{ selectedSheet.sheet.name }}</h2>
              </div>
              <span>{{ selectedSheet.sheet.rows.length }} 行 × {{ selectedSheet.sheet.columns.length }} 列</span>
            </header>
            <div ref="previewScrollRef" class="chat-preview-scroll" @scroll="onPreviewScroll">
              <table
                class="chat-preview-table"
                :style="{ minWidth: Math.max(1800, selectedSheet.sheet.columns.slice(0, 24).length * 160 + 40) + 'px' }"
              >
                <thead>
                  <tr>
                    <th>#</th>
                    <th v-for="column in selectedSheet.sheet.columns.slice(0, 24)" :key="column.id">{{ column.name }}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, rowIndex) in selectedSheet.sheet.rows.slice(0, 80)" :key="row.id">
                    <td>{{ rowIndex + 1 }}</td>
                    <td
                      v-for="column in selectedSheet.sheet.columns.slice(0, 24)"
                      :key="column.id"
                      :title="previewValue(row.values[column.id])"
                    >{{ previewValue(row.values[column.id]) }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div class="chat-preview-horizontal-scroll">
              <input type="range" min="0" max="1000" :value="horizontalPosition" aria-label="表格横向滚动条" @input="onHorizontalRange" />
            </div>
            <footer
              v-if="selectedSheet.sheet.columns.length > selectedSheet.sheet.columns.slice(0, 24).length || selectedSheet.sheet.rows.length > selectedSheet.sheet.rows.slice(0, 80).length"
              class="chat-preview-limit"
            >预览前 {{ selectedSheet.sheet.rows.slice(0, 80).length }} 行、{{ selectedSheet.sheet.columns.slice(0, 24).length }} 列，分析使用全部有效数据</footer>
          </section>
        </template>

        <!-- Results stage -->
        <div v-else-if="displayedResult" class="chat-results-stage">
          <nav class="chat-result-tabs">
            <button :class="{ active: activeResultTab === 'overview' }" @click="activeResultTab = 'overview'">综合报告</button>
            <button
              v-for="(module, index) in displayedResult.modules"
              :key="module.id"
              :class="{ active: activeResultTab === module.id }"
              @click="activeResultTab = module.id"
            >{{ index + 1 }}. {{ module.title }}</button>
            <button :class="{ active: activeResultTab === 'relations' }" @click="activeResultTab = 'relations'">数据关系</button>
          </nav>

          <!-- Overview result -->
          <div v-if="activeResultTab === 'overview'" class="chat-result-scroll">
            <section class="chat-report-hero">
              <span class="eyebrow">综合分析</span>
              <h1>{{ displayedResult.title }}</h1>
              <p>{{ displayedResult.executiveSummary }}</p>
            </section>
            <section
              v-if="displayedResult.operationSummaries?.length"
              class="chat-report-section chat-operation-summary"
            >
              <div class="chat-section-title">
                <el-icon><Files /></el-icon><h2>数据处理结果</h2>
              </div>
              <ul>
                <li v-for="(item, index) in displayedResult.operationSummaries" :key="`${item}-${index}`">{{ item }}</li>
              </ul>
              <div
                v-if="displayedResult.processedProject?.sheets?.find((sheet) => sheet.id === displayedResult.processedProject?.activeSheetId) || displayedResult.processedProject?.sheets?.[0]"
                class="chat-module-table-wrap chat-processed-preview"
              >
                <table>
                  <thead>
                    <tr>
                      <th
                        v-for="column in (displayedResult.processedProject?.sheets?.find((sheet) => sheet.id === displayedResult.processedProject?.activeSheetId) || displayedResult.processedProject?.sheets?.[0])?.columns.slice(0, 8)"
                        :key="column.id"
                      >{{ column.name }}</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr
                      v-for="row in (displayedResult.processedProject?.sheets?.find((sheet) => sheet.id === displayedResult.processedProject?.activeSheetId) || displayedResult.processedProject?.sheets?.[0])?.rows.slice(0, 10)"
                      :key="row.id"
                    >
                      <td
                        v-for="column in (displayedResult.processedProject?.sheets?.find((sheet) => sheet.id === displayedResult.processedProject?.activeSheetId) || displayedResult.processedProject?.sheets?.[0])?.columns.slice(0, 8)"
                        :key="column.id"
                      >{{ previewValue(row.values[column.id]) }}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
            <section class="chat-report-metrics">
              <div><strong>{{ displayedResult.fileCount }}</strong><span>数据文件</span></div>
              <div><strong>{{ displayedResult.sheetCount }}</strong><span>工作表</span></div>
              <div><strong>{{ displayedResult.totalRows.toLocaleString('zh-CN') }}</strong><span>分析行</span></div>
              <div><strong>{{ displayedResult.modules.length }}</strong><span>分析模块</span></div>
            </section>
            <section class="chat-report-section">
              <div class="chat-section-title"><el-icon><MagicStick /></el-icon><h2>关键发现</h2></div>
              <ol class="chat-numbered-list">
                <li v-for="(item, index) in displayedResult.keyFindings" :key="`${item}-${index}`">{{ item }}</li>
              </ol>
            </section>
            <section class="chat-report-section">
              <div class="chat-section-title"><el-icon><DataLine /></el-icon><h2>分析模块</h2></div>
              <div class="chat-module-list">
                <button
                  v-for="(module, index) in displayedResult.modules"
                  :key="module.id"
                  @click="activeResultTab = module.id"
                >
                  <span>{{ String(index + 1).padStart(2, '0') }}</span>
                  <span>
                    <strong>{{ module.title }}</strong>
                    <small>{{ module.summary }}</small>
                  </span>
                  <el-icon><CaretRight /></el-icon>
                </button>
              </div>
            </section>
            <section class="chat-report-columns">
              <div class="chat-report-section">
                <div class="chat-section-title"><el-icon><Check /></el-icon><h2>行动建议</h2></div>
                <ul>
                  <li v-for="(item, index) in displayedResult.recommendations" :key="`${item}-${index}`">{{ item }}</li>
                </ul>
              </div>
              <div class="chat-report-section warning">
                <div class="chat-section-title"><el-icon><Warning /></el-icon><h2>风险与口径</h2></div>
                <template v-if="displayedResult.warnings.length">
                  <ul>
                    <li v-for="(item, index) in displayedResult.warnings" :key="`${item}-${index}`">{{ item }}</li>
                  </ul>
                </template>
                <p v-else>没有额外的数据质量预警。</p>
              </div>
            </section>
          </div>

          <!-- Relations result -->
          <div v-else-if="activeResultTab === 'relations'" class="chat-result-scroll">
            <section class="chat-module-heading">
              <span class="eyebrow">数据血缘</span>
              <h1>跨表关系</h1>
              <p>关系发现智能体识别高可信公共键，以及可纵向合并的同构文件。</p>
            </section>
            <section class="chat-relation-list">
              <template v-if="displayedResult.relationships.length">
                <div v-for="(relation, index) in displayedResult.relationships" :key="relation">
                  <span>{{ index + 1 }}</span>
                  <el-icon><Coin /></el-icon>
                  <strong>{{ relation }}</strong>
                </div>
              </template>
              <div v-else class="chat-no-relations">
                <el-icon><Warning /></el-icon>
                <span>未发现高可信公共键，分析结果按各工作表独立计算。</span>
              </div>
            </section>
          </div>

          <!-- Module result -->
          <div v-else-if="activeModule" class="chat-result-scroll">
            <section class="chat-module-heading">
              <span class="eyebrow">{{ activeModule.operation }} · {{ activeModule.sourceNames.join(' + ') }}</span>
              <h1>{{ activeModule.title }}</h1>
              <p>{{ activeModule.summary }}</p>
            </section>
            <section class="chat-module-metrics">
              <div v-for="metric in activeModule.metrics" :key="metric.label">
                <span>{{ metric.label }}</span>
                <strong>{{ metric.value }}</strong>
              </div>
            </section>
            <section
              class="chat-chart-panel"
              :style="{ height: (activeModule.chartType === 'bar' ? Math.max(300, Math.min(720, activeModule.data.length * 26 + 54)) : 340) + 'px' }"
            >
              <AnalysisChart :module="activeModule" />
            </section>
            <section v-if="activeModule.findings.length > 0" class="chat-report-section">
              <div class="chat-section-title"><el-icon><MagicStick /></el-icon><h2>模块发现</h2></div>
              <ul>
                <li v-for="(item, index) in activeModule.findings" :key="`${item}-${index}`">{{ item }}</li>
              </ul>
            </section>
            <section class="chat-data-table-section">
              <div class="chat-section-title"><el-icon><Coin /></el-icon><h2>计算结果</h2></div>
              <div class="chat-module-table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>分类</th>
                      <th>指标值</th>
                      <th v-if="activeModule.data.some((item) => item.secondary != null)">辅助指标</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-for="(item, index) in activeModule.data" :key="`${item.name}-${index}`">
                      <td>{{ item.name }}</td>
                      <td>{{ item.value.toLocaleString('zh-CN', { maximumFractionDigits: 2 }) }}</td>
                      <td v-if="activeModule.data.some((point) => point.secondary != null)">
                        {{ item.secondary?.toLocaleString('zh-CN', { maximumFractionDigits: 2 }) || '-' }}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
            <section v-if="activeModule.notes.length > 0" class="chat-method-notes">
              <el-icon><Coin /></el-icon>
              <div>
                <strong>计算口径</strong>
                <p v-for="note in activeModule.notes" :key="note">{{ note }}</p>
              </div>
            </section>
          </div>
        </div>
      </section>

      <div
        class="chat-conversation-resizer"
        role="separator"
        aria-label="调整业务聊天框宽度"
        aria-orientation="vertical"
        :aria-valuemin="280"
        :aria-valuemax="conversationMaxWidth"
        :aria-valuenow="conversationWidth"
        tabindex="0"
        title="拖动调整业务聊天框宽度，双击恢复默认宽度"
        @dblclick="conversationWidth = 340"
        @pointerdown="onConversationPointerDown"
        @keydown="onConversationKeyDown"
      />

      <aside class="chat-conversation">
        <div class="chat-panel-heading">
          <div>
            <span class="eyebrow">对话分析</span>
            <h2>业务问题</h2>
          </div>
          <el-icon><ChatDotRound /></el-icon>
        </div>
        <div class="chat-messages">
          <template v-if="messages.length">
            <template v-for="message in messages" :key="message.id">
              <div class="chat-message" :class="message.role">
                <span>
                  <el-icon v-if="message.role === 'assistant'"><MagicStick /></el-icon>
                  <template v-else>你</template>
                </span>
                <div class="chat-message-body">
                  <p>{{ message.content }}</p>
                  <div v-if="message.attachment" class="chat-file-card">
                    <span class="chat-file-icon"><el-icon><Grid /></el-icon></span>
                    <span class="chat-file-copy">
                      <strong>{{ message.attachment.name }}</strong>
                      <small>
                        {{ message.attachment.project.sheets?.length || 1 }} 张表 ·
                        {{ (message.attachment.project.sheets || []).reduce((total, sheet) => total + sheet.rows.length, 0).toLocaleString('zh-CN') }} 行
                      </small>
                    </span>
                    <button
                      class="icon-button"
                      :aria-label="`预览 ${message.attachment.name}`"
                      title="预览"
                      @click="previewArtifact = message.attachment; viewMode = 'artifact'"
                    >
                      <el-icon><View /></el-icon>
                    </button>
                    <button
                      class="icon-button"
                      :aria-label="`下载 ${message.attachment.name}`"
                      title="下载"
                      @click="exportProcessedWorkbook(message.attachment!.project, message.attachment!.name)"
                    >
                      <el-icon><Download /></el-icon>
                    </button>
                  </div>
                </div>
              </div>
              <section v-if="message.id === latestUserMessageId" class="chat-agent-flow">
                <div class="chat-panel-heading compact">
                  <div>
                    <span class="eyebrow">执行流程</span>
                    <h2>分析智能体</h2>
                  </div>
                  <el-icon><ChatDotRound /></el-icon>
                </div>
                <div class="chat-agent-steps">
                  <div
                    v-for="(step, index) in CHAT_ANALYSIS_STEPS.filter((s) => result?.reportGenerated !== false || s.phase !== 'synthesizing')"
                    :key="step.phase"
                    :class="{ complete: phase === 'done' || (currentStep >= 0 && index < currentStep), active: step.phase === phase }"
                  >
                    <span>
                      <el-icon v-if="phase === 'done' || (currentStep >= 0 && index < currentStep)"><Check /></el-icon>
                      <el-icon v-else-if="step.phase === phase" class="spin"><Loading /></el-icon>
                      <span v-else class="dot" />
                    </span>
                    <div>
                      <strong>{{ step.name }}</strong>
                      <small>{{ step.description }}</small>
                    </div>
                  </div>
                </div>
              </section>
            </template>
          </template>
          <div v-else class="chat-conversation-empty">
            <el-icon><ChatDotRound /></el-icon>
            <strong>全自动分析</strong>
            <span>当前会话将使用左侧全部数据源</span>
          </div>
          <div v-if="busy && phase !== 'importing'" class="chat-message assistant working">
            <span><el-icon class="spin"><Loading /></el-icon></span>
            <p>{{ CHAT_ANALYSIS_STEPS.find((step) => step.phase === phase)?.description || '正在分析' }}</p>
          </div>
        </div>
        <div v-if="displayedResult?.followUpQuestions.length" class="chat-followups">
          <button
            v-for="item in displayedResult.followUpQuestions.slice(0, 3)"
            :key="item"
            :disabled="busy"
            @click="question = item; runAnalysis(item)"
          >{{ item }}<el-icon><ArrowRight /></el-icon></button>
        </div>
        <div v-if="error" class="chat-error">
          <el-icon><Warning /></el-icon><span>{{ error }}</span>
        </div>
        <div class="chat-composer">
          <textarea
            v-model="question"
            :placeholder="QUESTION_PLACEHOLDER"
            rows="5"
            :disabled="busy"
            @keydown="onComposerKeyDown"
          />
          <div>
            <span>{{ sources.length }} 个文件 · {{ sourceSheets.length }} 张表</span>
            <button v-if="busy && phase !== 'importing'" class="button secondary compact" @click="stopAnalysis">停止</button>
            <button
              v-else
              class="button primary compact"
              :disabled="!sources.length || !question.trim()"
              v-ai-credit="'workflow:ai-excel-tool:generate'"
              @click="runAnalysis()"
            ><el-icon><Promotion /></el-icon>开始分析</button>
          </div>
        </div>
      </aside>
    </main>
  </div>
</template>

<style scoped>
/* 全局 styles.css（.excel-page 作用域）已定义所有 .chat-* 类名样式 */
/* 这里仅调整 element-plus 图标尺寸以匹配原 lucide 图标的视觉 */
.chat-shell :deep(.el-icon) {
  width: 1em;
  height: 1em;
  font-size: inherit;
}
.brand-mark :deep(.el-icon svg) {
  width: 19px;
  height: 19px;
}
.model-quick-select {
  width: 170px;
}
.model-quick-select :deep(.el-input__wrapper) {
  min-height: 32px;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 0 0 1px var(--line) inset;
}
.model-quick-select :deep(.el-input__wrapper:hover),
.model-quick-select :deep(.el-input__wrapper.is-focus) {
  box-shadow: 0 0 0 1px var(--line-strong) inset;
}
.model-quick-select :deep(.el-input__inner) {
  color: #46534c;
  font-size: 12px;
}
.quick-option {
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}
.chat-mode-label :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.icon-button :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.button :deep(.el-icon svg) {
  width: 15px;
  height: 15px;
}
.button.compact :deep(.el-icon svg) {
  width: 15px;
  height: 15px;
}
.chat-upload-zone :deep(.el-icon svg) {
  width: 21px;
  height: 21px;
}
.chat-upload-zone > :deep(.el-icon) {
  color: var(--green);
  font-size: 21px;
}
.chat-panel-heading :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.chat-panel-heading > :deep(.el-icon) {
  color: #5f7167;
}
.chat-source-file-row :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.chat-source-file-row > :deep(.source-file-icon) {
  width: 17px;
  min-width: 17px;
  flex: 0 0 17px;
  display: inline-flex;
  color: #33815d;
  font-size: 17px;
}
.chat-source-expand :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.chat-source-sheets :deep(.el-icon svg) {
  width: 13px;
  height: 13px;
}
.chat-source-sheets button > :deep(.el-icon) {
  width: 13px;
  min-width: 13px;
  height: 13px;
  display: inline-flex;
  overflow: visible;
  color: inherit;
  font-size: 13px;
}
.chat-center-tabs :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.chat-center-tabs button > :deep(.el-icon) {
  width: 14px;
  min-width: 14px;
  height: 14px;
  display: inline-flex;
  border-radius: 0;
  background: transparent;
  color: inherit;
  font-size: 14px;
}
.chat-center-report-open :deep(.el-icon svg),
.chat-center-report-close :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.chat-message :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.chat-file-card :deep(.el-icon svg) {
  width: 18px;
  height: 18px;
}
.chat-file-card .icon-button :deep(.el-icon svg) {
  width: 15px;
  height: 15px;
}
.chat-agent-flow :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.chat-agent-steps :deep(.el-icon svg) {
  width: 12px;
  height: 12px;
}
.chat-section-title :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.chat-composer textarea + div :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.chat-followups :deep(.el-icon svg) {
  width: 13px;
  height: 13px;
}
.chat-error :deep(.el-icon svg) {
  width: 15px;
  height: 15px;
}
.chat-conversation-empty > :deep(.el-icon) {
  color: #3977a8;
  font-size: 24px;
}
.chat-method-notes > :deep(.el-icon) {
  flex: 0 0 auto;
}
.spin {
  animation: chat-spin 1.1s linear infinite;
}
@keyframes chat-spin {
  to { transform: rotate(360deg); }
}
@media (max-width: 1100px) {
  .chat-topbar .button > :deep(.el-icon) {
    display: inline-flex;
    font-size: 15px;
  }
  .chat-topbar .button-label {
    display: none;
  }
}
@media (max-width: 820px) {
  .model-quick-select {
    width: 130px;
    margin-left: auto;
  }
}
@media (max-width: 480px) {
  .model-quick-select {
    display: none;
  }
}
</style>
