<template>
  <div class="app-shell">
    <header class="topbar">
      <button class="brand topbar-brand" aria-label="返回历史项目" title="返回历史项目" @click="exitToStart">
        <div class="brand-mark"><el-icon><Grid /></el-icon></div>
        <span>AI Excel</span>
      </button>
      <span class="topbar-divider" />
      <input
        class="project-name"
        :value="project.name"
        aria-label="项目名称"
        placeholder="项目名称"
        @change="onProjectNameChange(($event.target as HTMLInputElement).value)"
        @keydown.enter="($event.target as HTMLInputElement).blur()"
      />
      <span class="save-state" :class="saveStatus" :title="saveError">
        <span class="save-dot" />{{ saveStatusLabel }}
      </span>
      <span class="topbar-spacer" />
      <el-select
        v-if="providers.length"
        :model-value="props.selectedProvider"
        size="small"
        class="model-quick-select"
        placeholder="切换模型"
        @update:model-value="onProviderChange"
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
      <button v-else class="model-chip" @click="modal = 'settings'">
        <span class="model-dot demo" />
        <span>{{ modelName || '配置大模型' }}</span>
      </button>
      <div class="header-actions">
        <button class="icon-button" aria-label="返回项目首页" title="返回项目首页" @click="exitToStart"><el-icon><House /></el-icon></button>
        <button class="icon-button" aria-label="上传 Excel 新项目" title="上传 Excel 新项目" @click="triggerExcelUpload"><el-icon><Upload /></el-icon></button>
        <button class="icon-button" :disabled="exporting" aria-label="导出 Excel" title="导出 Excel" @click="handleExport"><el-icon><Download /></el-icon></button>
        <button class="icon-button" aria-label="重置项目" title="重置项目" @click="confirmResetProject"><el-icon><Refresh /></el-icon></button>
        <button class="icon-button" aria-label="模型设置" title="模型设置与测试连接" @click="modal = 'settings'"><el-icon><Setting /></el-icon></button>
      </div>
      <input
        ref="excelInputRef"
        class="hidden-file-input"
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        @change="onExcelInputChange"
      />
    </header>

    <nav class="tabbar">
      <div class="tabs-scroll">
        <button class="tab" :class="{ active: project.activeTabId === 'table' }" @click="setActiveTab('table')">
          <el-icon><Grid /></el-icon>数据表格
        </button>
        <button
          v-for="view in project.views"
          :key="view.id"
          class="tab"
          :class="{ active: project.activeTabId === view.id }"
          @click="setActiveTab(view.id)"
        >
          <el-icon><MagicStick /></el-icon>{{ view.name }}
          <span v-if="view.status === 'error'" class="tab-error" />
          <span class="tab-close" title="删除视图" @click.stop="confirmDeleteView(view.id)">×</span>
        </button>
        <button class="tab add-tab" @click="openNewView"><el-icon><Plus /></el-icon>新建视图</button>
      </div>
    </nav>

    <main class="workspace" :style="{ '--inspector-width': inspectorWidth + 'px' }">
      <section class="content-pane">
        <TableEditor
          v-if="project.activeTabId === 'table'"
          :columns="project.columns"
          :displayed-rows="displayedRows"
          :runtime-row-count="runtimeRows.length"
          :rule-active="Boolean(project.tableRule.generatedCode)"
          :rule-error="ruleError"
          :selection="selection"
          :sheet-name="sheets.find((sheet) => sheet.id === project.activeSheetId)?.name || '数据'"
          @add-column="modal = 'column'"
          @add-view="openNewView"
          @add-row="onAddRow"
          @delete-row="onDeleteRow"
          @select-rule="selection = { type: 'rule' }"
          @update-cell="onUpdateCell"
          @select-column="(id) => (selection = { type: 'column', id })"
        />
        <ViewStage
          v-else-if="activeView"
          :project="project"
          :view="activeView"
          :runtime-error="activeView.runtimeError || ''"
          @view-state-change="onViewStateChange"
          @runtime-error="onViewRuntimeError"
          @regenerate="onGenerateView(activeView!.id)"
        />
        <div v-else class="view-stage-placeholder">
          <span class="placeholder-icon"><ViewIcon /></span>
          <h3>视图不存在</h3>
          <p>该视图可能已被删除，请切换到其他标签。</p>
        </div>
      </section>

      <div
        class="inspector-resizer"
        role="separator"
        aria-orientation="vertical"
        tabindex="0"
        title="拖动调整右侧面板宽度，双击恢复默认"
        @dblclick="inspectorWidth = 360"
        @pointerdown="onResizerPointerDown"
        @keydown="onResizerKeyDown"
      />
      <aside class="inspector">
        <NewViewPanel
          v-if="modal === 'view'"
          :sheets="sheets"
          :active-sheet-id="project.activeSheetId"
          @close="closeNewView"
          @create="onCreateView"
        />
        <InspectorPanel
          v-else
          :selection="selection"
          :project="project"
          :generation-progress="generationProgress"
          :model-configured="modelConfigured"
          :model-name="modelName"
          @add-column="modal = 'column'"
          @add-view="openNewView"
          @select="(s) => (selection = s)"
          @generate-column="onGenerateColumn"
          @generate-rule="onGenerateRule"
          @generate-view="onGenerateView"
          @cancel-generation="cancelGeneration"
          @update-column-description="onUpdateColumnDescription"
          @update-rule-description="onUpdateRuleDescription"
          @update-view-description="onUpdateViewDescription"
          @clear-rule="onClearRule"
          @open-settings="modal = 'settings'"
          @delete-column="onDeleteColumn"
          @rename-column="onRenameColumn"
          @delete-view="confirmDeleteView"
          @rename-view="onRenameView"
        />
      </aside>
    </main>

    <footer class="statusbar">
      <div class="sheet-tabs">
        <el-icon><Coin /></el-icon>
        <button
          v-for="sheet in sheets"
          :key="sheet.id"
          :class="{ active: sheet.id === project.activeSheetId }"
          role="tab"
          :aria-selected="sheet.id === project.activeSheetId"
          @click="onSelectSheet(sheet.id)"
        >
          {{ sheet.name }}
        </button>
      </div>
      <span class="statusbar-count">{{ project.rows.length }} 行 × {{ project.columns.length }} 列</span>
      <span>本地项目</span>
      <span class="statusbar-spacer" />
      <span>schema v2</span>
    </footer>

    <!-- 弹窗 -->
    <AddColumnModal v-model:visible="addColumnVisible" @create="onCreateColumn" />
    <SettingsModal
      v-model:visible="settingsVisible"
      :providers="providers"
      :selected-provider="selectedProvider"
      @update:selected-provider="onProviderChange"
      @close="settingsVisible = false"
    />

    <div v-if="toast" class="toast">
      <el-icon><Check /></el-icon>
      <span>{{ toast }}</span>
    </div>
  </div>
</template>

<script setup name="AiExcelWorkbench" lang="ts">
import { ref, reactive, computed, watch, onMounted, onBeforeUnmount, toRaw } from 'vue'
import {
  Check, Coin, Download, Grid, House, MagicStick, Plus, Refresh, Setting, Upload,
  View as ViewIcon,
} from '@element-plus/icons-vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import TableEditor from './TableEditor.vue'
import InspectorPanel from './InspectorPanel.vue'
import AddColumnModal from './AddColumnModal.vue'
import SettingsModal from './SettingsModal.vue'
import ViewStage from './ViewStage.vue'
import NewViewPanel from './NewViewPanel.vue'
import {
  applyTableRule,
  calculateRuntimeRows,
  cleanModelCode,
  createDefaultProject,
  createGenerationState,
  createId,
  ensureProjectSheets,
  hydrateProjectActiveSheet,
  syncProjectActiveSheet,
} from '../lib/core'
import { saveProject } from '../lib/project-store'
import { projectToExcel } from '../lib/excel'
import {
  generateColumn,
  generateRule,
  generateView,
  readModelStatus,
} from '../lib/generation'
import {
  runSelfCorrectingGeneration,
  cleanGeneratedViewCode,
  humanizeGenerationError,
  SelfCorrectionExhaustedError,
} from '../lib/self-correction'
import { preflightGeneratedView } from '../lib/generated-view-host'
import { GENERATED_APP_RUNTIME_VERSION } from '../lib/generated-app-runtime'
import { resolveChartBindings } from '../lib/chart-templates'
import { projectForView, resolveSourceSheetIds, resolveViewData } from '../lib/data-sources'
import { listExcelProviders } from '@/api/excelTool'
import type {
  ChartType,
  ColumnDefinition,
  GenerationAttempt,
  GenerationState,
  Project,
  ReferenceImage,
  RuntimeRow,
  Selection,
  ViewDefinition,
} from '../lib/types'

const props = defineProps({
  initialProject: { type: Object as () => Project, required: true },
  selectedProvider: { type: String, default: '' },
})
const emit = defineEmits(['exit', 'provider-change', 'upload-excel'])

// === 项目状态（用 reactive 包装便于深层变更） ===
// makeReactive：规避 reactive 泛型对递归 JsonValue 的过深实例化（vue-tsc TS2589），运行时仍是深响应式对象
function makeReactive(source: Project): Project {
  return reactive(source as any) as Project
}
const project = makeReactive(cloneProject(props.initialProject))
hydrateProjectActiveSheet(project)

function cloneProject(src: Project): Project {
  return JSON.parse(JSON.stringify(src)) as Project
}

const selection = ref<Selection>({ type: 'project' })
const modal = ref<'column' | 'settings' | 'view' | null>(null)
const addColumnVisible = computed({
  get: () => modal.value === 'column',
  set: (v) => (modal.value = v ? 'column' : null),
})
const settingsVisible = computed({
  get: () => modal.value === 'settings',
  set: (v) => (modal.value = v ? 'settings' : null),
})

const inspectorWidth = ref(360)
const saveStatus = ref<'saving' | 'saved' | 'error'>('saved')
const saveError = ref('')
const toast = ref('')
const exporting = ref(false)

// === 模型状态 ===
const modelStatus = ref({ configured: false, provider: '未配置', model: '' })
const providers = ref<Array<{ provider: string; providerName: string; modelName: string }>>([])
const modelConfigured = computed(() => modelStatus.value.configured || providers.value.length > 0)
const modelName = computed(() => {
  if (props.selectedProvider) {
    const p = providers.value.find((x) => x.provider === props.selectedProvider)
    return p ? `${p.providerName || p.provider} · ${p.modelName || ''}` : props.selectedProvider
  }
  return modelStatus.value.model || ''
})

// === 派生数据 ===
const sheets = computed(() => ensureProjectSheets(project as Project))
const runtimeRows = computed(() => calculateRuntimeRows(project.rows, project.columns))
const displayedRowsInfo = computed(() => {
  try {
    return {
      rows: applyTableRule(runtimeRows.value, project.columns, project.tableRule.generatedCode),
      error: '',
    }
  } catch (error) {
    return { rows: runtimeRows.value, error: error instanceof Error ? error.message : String(error) }
  }
})
const displayedRows = computed(() => displayedRowsInfo.value.rows)
const ruleError = computed(() => displayedRowsInfo.value.error)
const activeView = computed(() => (project.views as ViewDefinition[]).find((v) => v.id === project.activeTabId))

// === 生成进度（key: column:<id> / view:<id> / rule） ===
const generationProgress = ref<Record<string, { attempt: number; maxAttempts: number }>>({})
const generationControllers = new Map<string, AbortController>()

const saveStatusLabel = computed(() => ({ saving: '保存中', saved: '已保存', error: '保存失败' } as Record<string, string>)[saveStatus.value] || '')

// === 保存策略：只在 AI 生成完成 / 关键结构变更时保存（避免每次按键都序列化大对象） ===
let saveTimer: number | null = null

function triggerSave(immediate = false) {
  if (saveTimer != null) {
    window.clearTimeout(saveTimer)
    saveTimer = null
  }
  const run = () => {
    saveStatus.value = 'saving'
    saveError.value = ''
    saveProject(toRaw(project))
      .then(() => (saveStatus.value = 'saved'))
      .catch((error) => {
        saveStatus.value = 'error'
        saveError.value = error instanceof Error ? error.message : String(error)
      })
  }
  if (immediate) {
    run()
  } else {
    // 防抖 800ms，避免短时间内多次触发（如批量生成）
    saveTimer = window.setTimeout(() => {
      saveTimer = null
      run()
    }, 800)
  }
}

onBeforeUnmount(() => {
  if (saveTimer != null) window.clearTimeout(saveTimer)
})

// === Toast ===
let toastTimer: number | null = null
watch(toast, (val) => {
  if (!val) return
  if (toastTimer != null) window.clearTimeout(toastTimer)
  toastTimer = window.setTimeout(() => (toast.value = ''), 2600)
})

// === 项目变更工具 ===
function mutate(recipe: (draft: Project) => void) {
  const draft = cloneProject(project as Project)
  recipe(draft)
  syncProjectActiveSheet(draft)
  draft.updatedAt = Date.now()
  // 用 Object.assign 覆盖 reactive 对象的顶层字段
  Object.keys(project).forEach((key) => {
    ;(project as any)[key] = (draft as any)[key]
  })
}

function onProjectNameChange(name: string) {
  mutate((draft) => (draft.name = name.trim() || '未命名项目'))
}

// === 行/列操作 ===
function onUpdateCell(rowId: string, columnId: string, value: unknown) {
  mutate((draft) => {
    const row = draft.rows.find((r) => r.id === rowId)
    if (row) row.values[columnId] = value
  })
}

function onAddRow() {
  mutate((draft) => {
    const newRow = { id: createId('row'), values: {} as Record<string, unknown> }
    draft.columns.forEach((col) => {
      newRow.values[col.id] = col.dataType === 'number' || col.dataType === 'boolean' ? null : ''
    })
    draft.rows.push(newRow)
  })
}

function onDeleteRow(rowId: string) {
  mutate((draft) => {
    draft.rows = draft.rows.filter((r) => r.id !== rowId)
  })
}

function onCreateColumn(payload: { name: string; kind: 'manual' | 'computed'; dataType: string }) {
  mutate((draft) => {
    const column: ColumnDefinition = {
      id: createId('col'),
      name: payload.name.trim(),
      kind: payload.kind,
      dataType: payload.dataType as any,
      generation: payload.kind === 'computed' ? createGenerationState(payload.name) : undefined,
    }
    draft.columns.push(column)
    if (column.kind === 'manual') {
      draft.rows.forEach((row) => {
        row.values[column.id] = column.dataType === 'number' || column.dataType === 'boolean' ? null : ''
      })
    }
    selection.value = { type: 'column', id: column.id }
  })
  modal.value = null
  ElMessage.success('列已创建')
}

function onDeleteColumn(columnId: string) {
  mutate((draft) => {
    draft.columns = draft.columns.filter((c) => c.id !== columnId)
    draft.rows.forEach((row) => delete row.values[columnId])
  })
  selection.value = { type: 'project' }
}

function onRenameColumn(payload: { columnId: string; name: string }) {
  mutate((draft) => {
    const col = draft.columns.find((c) => c.id === payload.columnId)
    if (col) col.name = payload.name.trim() || col.name
  })
}

function onRenameView(payload: { viewId: string; name: string }) {
  mutate((draft) => {
    const view = draft.views.find((item) => item.id === payload.viewId)
    if (view) view.name = payload.name.trim() || view.name
  })
}

function onSelectSheet(sheetId: string) {
  mutate((draft) => {
    const sheet = ensureProjectSheets(draft).find((s) => s.id === sheetId)
    if (sheet) {
      draft.activeSheetId = sheet.id
      draft.columns = sheet.columns
      draft.rows = sheet.rows
      draft.tableRule = sheet.tableRule
    }
  })
}

// === 视图 ===
function openNewView() {
  modal.value = 'view'
  selection.value = { type: 'project' }
  inspectorWidth.value = Math.min(720, Math.max(360, Math.floor(window.innerWidth * 0.48)))
}

function closeNewView() {
  modal.value = null
  selection.value = { type: 'project' }
}

function onCreateView(payload: {
  name: string
  description: string
  chartType: ChartType
  images: ReferenceImage[]
  sourceSheetIds: string[]
}) {
  const id = createId('view')
  mutate((draft) => {
    const sources = resolveSourceSheetIds(draft, payload.sourceSheetIds)
    const sourceProject = projectForView(draft, { sourceSheetIds: sources })
    const chartBindings = payload.chartType === 'custom' || sources.length > 1
      ? undefined
      : resolveChartBindings(sourceProject, payload.chartType, payload.description)
    const view: ViewDefinition = {
      id,
      name: payload.name,
      chartType: payload.chartType,
      sourceSheetIds: sources,
      chartBindings,
      chartBindingsLocked: false,
      ...createGenerationState(payload.description),
      referenceImages: payload.images,
      appState: {},
      runtimeVersion: GENERATED_APP_RUNTIME_VERSION,
      runtimeEpoch: 0,
      runtimeError: null,
    }
    draft.views.push(view)
    draft.activeTabId = id
  })
  selection.value = { type: 'view', id }
  modal.value = null
}

function confirmDeleteView(viewId: string) {
  ElMessageBox.confirm('删除该视图？', '删除视图', { type: 'warning' })
    .then(() => {
      mutate((draft) => {
        draft.views = draft.views.filter((v) => v.id !== viewId)
        if (draft.activeTabId === viewId) draft.activeTabId = 'table'
      })
      if (selection.value.type === 'view' && selection.value.id === viewId) {
        selection.value = { type: 'project' }
      }
    })
    .catch(() => {})
}

function setActiveTab(tabId: 'table' | string) {
  mutate((draft) => (draft.activeTabId = tabId as any))
  if (tabId === 'table') {
    selection.value = { type: 'project' }
  } else {
    selection.value = { type: 'view', id: tabId }
  }
}

// === 描述更新 ===
function onUpdateColumnDescription(payload: { columnId: string; description: string }) {
  mutate((draft) => {
    const col = draft.columns.find((c) => c.id === payload.columnId)
    if (col?.generation && col.generation.draftDescription !== payload.description) {
      col.generation.draftDescription = payload.description
      col.generation.sourceVersion += 1
    }
  })
}
function onUpdateRuleDescription(description: string) {
  mutate((draft) => {
    if (draft.tableRule.draftDescription !== description) {
      draft.tableRule.draftDescription = description
      draft.tableRule.sourceVersion += 1
    }
  })
}
function onUpdateViewDescription(payload: { viewId: string; description: string }) {
  mutate((draft) => {
    const view = draft.views.find((v) => v.id === payload.viewId)
    if (view && view.draftDescription !== payload.description) {
      view.draftDescription = payload.description
      view.sourceVersion += 1
    }
  })
}
function onClearRule() {
  mutate((draft) => {
    draft.tableRule.generatedCode = ''
    draft.tableRule.status = 'idle'
    draft.tableRule.error = null
  })
}

// === 生成（计算列、规则、视图）===
function requireModel(): boolean {
  if (modelConfigured.value) return true
  modal.value = 'settings'
  showToast('请先在右侧选择可用模型')
  return false
}

function showToast(msg: string) {
  toast.value = msg
}

function startGeneration(key: string): AbortController {
  generationControllers.get(key)?.abort()
  const controller = new AbortController()
  generationControllers.set(key, controller)
  return controller
}

function updateGenerationProgress(key: string, attempt: number, maxAttempts: number) {
  generationProgress.value = { ...generationProgress.value, [key]: { attempt, maxAttempts } }
}

function finishGeneration(key: string, controller: AbortController) {
  if (generationControllers.get(key) === controller) generationControllers.delete(key)
  const next = { ...generationProgress.value }
  delete next[key]
  generationProgress.value = next
}

function cancelGeneration(key: string) {
  generationControllers.get(key)?.abort()
}

function storedAttempt(failure: { attempt: number; rawError: string; feedback: string }): GenerationAttempt {
  return {
    attempt: failure.attempt,
    error: failure.feedback,
    rawError: failure.rawError,
    code: '',
    createdAt: Date.now(),
  }
}

function validateColumnCandidate(project: Project, columnId: string, code: string) {
  const candidate = cloneProject(project)
  const column = candidate.columns.find((item) => item.id === columnId)
  if (!column?.generation) throw new Error('计算列不存在')
  column.generation.generatedCode = code
  const rows = calculateRuntimeRows(candidate.rows, candidate.columns)
  const rowFailures = rows
    .map((row, index) => row.errors[columnId] ? `第 ${index + 1} 行：${row.errors[columnId]}` : '')
    .filter(Boolean)
  if (rowFailures.length) {
    throw new Error(`计算列在 ${rowFailures.length} 行运行失败：${rowFailures.slice(0, 3).join('；')}`)
  }
  const values = rows.map((row) => row.values[columnId]).filter((value) => value != null && value !== '')
  const invalid = values.find((value) => {
    if (column.dataType === 'number') return typeof value !== 'number' || !Number.isFinite(value)
    if (column.dataType === 'boolean') return typeof value !== 'boolean'
    if (column.dataType === 'text') return typeof value !== 'string'
    return !(value instanceof Date) && (typeof value !== 'string' || Number.isNaN(Date.parse(value)))
  })
  if (invalid !== undefined) {
    throw new Error(`计算列返回类型错误：目标类型是 ${column.dataType}，实际值示例为 ${JSON.stringify(invalid)}`)
  }
}

function updateGenerationState(
  target: GenerationState,
  recipe: (s: GenerationState) => void,
) {
  // 直接修改 reactive 内部字段
  recipe(target)
}

async function onGenerateColumn(columnId: string) {
  if (!requireModel()) return
  const column = project.columns.find((c) => c.id === columnId)
  if (!column?.generation) return
  const key = `column:${columnId}`
  const controller = startGeneration(key)
  const baseProject = cloneProject(project)
  const baseColumn = baseProject.columns.find((c) => c.id === columnId)!

  updateGenerationState(column.generation, (s) => {
    s.status = 'generating'
    s.error = null
    s.description = s.draftDescription
  })

  try {
    const result = await runSelfCorrectingGeneration({
      generate: async (ctx) => {
        updateGenerationProgress(key, ctx.attempt, ctx.maxAttempts)
        const repair = ctx.previousFailure
          ? {
              attempt: ctx.attempt,
              maxAttempts: ctx.maxAttempts,
              error: ctx.previousFailure.feedback,
              code: ctx.previousFailure.code,
              history: ctx.failures.map((f) => f.rawError),
            }
          : undefined
        return generateColumn(baseProject, baseColumn, props.selectedProvider, repair, controller.signal)
      },
      clean: cleanModelCode,
      validate: (code) => validateColumnCandidate(baseProject, columnId, code),
      humanize: (err, code) => humanizeGenerationError(err, code, baseProject.columns),
      signal: controller.signal,
      maxAttempts: 3,
      onFailure: (failure) => {
        updateGenerationState(column.generation!, (s) => {
          s.attempts = [...s.attempts, storedAttempt(failure)]
        })
      },
    })

    updateGenerationState(column.generation, (s) => {
      s.generatedCode = result.code
      s.lastAttemptCode = result.code
      s.status = 'ready'
      s.error = null
      s.generatedVersion = s.sourceVersion
      s.attempts = []
    })
    triggerSave()
    showToast(`计算列「${column.name}」已生成`)
  } catch (error) {
    const message = error instanceof SelfCorrectionExhaustedError
      ? error.message
      : error instanceof Error ? error.message : String(error)
    updateGenerationState(column.generation, (s) => {
      s.status = 'error'
      s.error = message
      s.lastAttemptCode = (error as any)?.lastCode || s.lastAttemptCode
    })
    if (!controller.signal.aborted) {
      ElMessage.error(message)
    }
  } finally {
    finishGeneration(key, controller)
  }
}

async function onGenerateRule() {
  if (!requireModel()) return
  const key = 'rule'
  const controller = startGeneration(key)
  const baseProject = cloneProject(project)

  updateGenerationState(project.tableRule, (s) => {
    s.status = 'generating'
    s.error = null
    s.description = s.draftDescription
  })

  try {
    const result = await runSelfCorrectingGeneration({
      generate: async (ctx) => {
        updateGenerationProgress(key, ctx.attempt, ctx.maxAttempts)
        const repair = ctx.previousFailure
          ? {
              attempt: ctx.attempt,
              maxAttempts: ctx.maxAttempts,
              error: ctx.previousFailure.feedback,
              code: ctx.previousFailure.code,
              history: ctx.failures.map((f) => f.rawError),
            }
          : undefined
        return generateRule(baseProject, props.selectedProvider, repair, controller.signal)
      },
      clean: cleanModelCode,
      validate: (code) => {
        applyTableRule(calculateRuntimeRows(baseProject.rows, baseProject.columns), baseProject.columns, code)
      },
      humanize: (err, code) => humanizeGenerationError(err, code, project.columns),
      signal: controller.signal,
      maxAttempts: 3,
    })

    updateGenerationState(project.tableRule, (s) => {
      s.generatedCode = result.code
      s.lastAttemptCode = result.code
      s.status = 'ready'
      s.error = null
      s.generatedVersion = s.sourceVersion
    })
    triggerSave()
    showToast('表格规则已生成')
  } catch (error) {
    if (controller.signal.aborted) {
      updateGenerationState(project.tableRule, (s) => {
        s.status = s.generatedCode ? 'ready' : 'idle'
        s.error = null
      })
      showToast('已停止生成')
      return
    }
    const message = error instanceof Error ? error.message : String(error)
    updateGenerationState(project.tableRule, (s) => {
      s.status = 'error'
      s.error = message
    })
    if (!controller.signal.aborted) ElMessage.error(message)
  } finally {
    finishGeneration(key, controller)
  }
}

async function onGenerateView(viewId: string) {
  const view = project.views.find((v) => v.id === viewId)
  if (!view) return
  const sourceSheetIds = resolveSourceSheetIds(project, view.sourceSheetIds)
  if ((view.chartType === 'custom' || sourceSheetIds.length > 1) && !requireModel()) return
  const key = `view:${viewId}`
  const controller = startGeneration(key)
  const baseProject = cloneProject(project)
  const baseView = baseProject.views.find((v) => v.id === viewId)!
  const startedDescription = baseView.draftDescription
  const startedSourceVersion = baseView.sourceVersion
  const sourceData = resolveViewData(baseProject, baseView.sourceSheetIds)

  updateGenerationState(view, (s) => {
    s.status = 'generating'
    s.error = null
    s.runtimeError = null
    s.attempts = []
  })

  try {
    const result = await runSelfCorrectingGeneration({
      generate: async (ctx) => {
        updateGenerationProgress(key, ctx.attempt, ctx.maxAttempts)
        const repair = ctx.previousFailure
          ? {
              attempt: ctx.attempt,
              maxAttempts: ctx.maxAttempts,
              error: ctx.previousFailure.feedback,
              code: ctx.previousFailure.code,
              history: ctx.failures.map((f) => f.rawError),
            }
          : false
        return generateView(baseProject, baseView, props.selectedProvider, repair || undefined, controller.signal)
      },
      clean: cleanGeneratedViewCode,
      validate: async (code) => {
        await preflightGeneratedView(
          code,
          sourceData.rows,
          sourceData.columns,
          baseView,
          sourceData.sheets,
          sourceData.relations,
        )
      },
      humanize: (err, code) => humanizeGenerationError(err, code, sourceData.columns),
      signal: controller.signal,
      maxAttempts: 3,
      onFailure: (failure) => {
        updateGenerationState(view, (s) => {
          s.attempts = [...s.attempts, storedAttempt(failure)]
        })
      },
    })

    updateGenerationState(view, (s) => {
      s.description = startedDescription
      s.generatedCode = result.code
      s.lastAttemptCode = result.code
      s.status = 'ready'
      s.error = null
      s.runtimeError = null
      s.generatedVersion = startedSourceVersion
    })
    triggerSave()
    showToast(`视图「${view.name}」已生成`)
  } catch (error) {
    if (controller.signal.aborted) {
      updateGenerationState(view, (s) => {
        s.status = s.generatedCode ? 'ready' : 'idle'
        s.error = null
      })
      view.runtimeError = null
      showToast('已停止生成')
      return
    }
    const message = error instanceof Error ? error.message : String(error)
    updateGenerationState(view, (s) => {
      s.status = 'error'
      s.error = message
    })
    view.runtimeError = message
    if (!controller.signal.aborted) ElMessage.error(message)
  } finally {
    finishGeneration(key, controller)
  }
}

// === ViewStage 事件 ===
function onViewStateChange(payload: { viewId: string; state: any; remount?: boolean }) {
  const current = project.views.find((view) => view.id === payload.viewId)
  if (!current) return
  if (!payload.remount && JSON.stringify(current.appState) === JSON.stringify(payload.state)) return
  mutate((draft) => {
    const view = draft.views.find((v) => v.id === payload.viewId)
    if (view) {
      view.appState = payload.state
      if (payload.remount) view.runtimeEpoch += 1
    }
  })
}

function onViewRuntimeError(payload: { viewId: string; message: string | null }) {
  const current = project.views.find((view) => view.id === payload.viewId)
  if (!current || current.runtimeError === payload.message) return
  mutate((draft) => {
    const view = draft.views.find((v) => v.id === payload.viewId)
    if (view && view.runtimeError !== payload.message) view.runtimeError = payload.message
  })
}

// === Inspector 宽度调整 ===
const resizeOrigin = { startX: 0, startWidth: 360 }
function onResizerPointerDown(event: PointerEvent) {
  if (event.button !== 0 || window.innerWidth <= 820) return
  event.preventDefault()
  resizeOrigin.startX = event.clientX
  resizeOrigin.startWidth = inspectorWidth.value
  window.addEventListener('pointermove', onResizerPointerMove)
  window.addEventListener('pointerup', onResizerPointerUp, { once: true })
  document.body.classList.add('resizing-inspector')
}
function onResizerPointerMove(event: PointerEvent) {
  const maxWidth = Math.max(320, Math.floor(window.innerWidth / 2))
  const next = resizeOrigin.startWidth + resizeOrigin.startX - event.clientX
  inspectorWidth.value = Math.min(maxWidth, Math.max(320, next))
}
function onResizerPointerUp() {
  window.removeEventListener('pointermove', onResizerPointerMove)
  document.body.classList.remove('resizing-inspector')
}
function onResizerKeyDown(event: KeyboardEvent) {
  const maxWidth = Math.max(320, Math.floor(window.innerWidth / 2))
  const step = event.shiftKey ? 40 : 16
  if (event.key === 'ArrowLeft') inspectorWidth.value = Math.min(maxWidth, inspectorWidth.value + step)
  else if (event.key === 'ArrowRight') inspectorWidth.value = Math.max(320, inspectorWidth.value - step)
  else if (event.key === 'Home') inspectorWidth.value = 320
  else if (event.key === 'End') inspectorWidth.value = maxWidth
  else return
  event.preventDefault()
}

// === Excel 上传/导出 ===
const excelInputRef = ref<HTMLInputElement | null>(null)
function triggerExcelUpload() {
  excelInputRef.value?.click()
}
function onExcelInputChange(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (file) emit('upload-excel', file)
}

async function handleExport() {
  if (exporting.value) return
  exporting.value = true
  try {
    const blob = await projectToExcel(toRaw(project), runtimeRows.value)
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = `${project.name || 'ai-excel'}.xlsx`
    anchor.click()
    URL.revokeObjectURL(url)
    ElMessage.success('已导出 Excel')
  } catch (error) {
    ElMessage.error('导出失败：' + (error instanceof Error ? error.message : String(error)))
  } finally {
    exporting.value = false
  }
}

// === Provider 变更 ===
function onProviderChange(provider: string) {
  emit('provider-change', provider)
  modal.value = null
  ElMessage.success('已选择模型')
}

function confirmResetProject() {
  ElMessageBox.confirm('重置当前项目内容？此操作无法撤销。', '重置项目', {
    type: 'warning',
    confirmButtonText: '重置',
    cancelButtonText: '取消',
  })
    .then(() => {
      mutate((draft) => {
        const replacement = createDefaultProject()
        Object.assign(draft, replacement, {
          id: draft.id,
          name: draft.name,
          createdAt: draft.createdAt,
        })
      })
      selection.value = { type: 'project' }
    })
    .catch(() => {})
}

// === 退出 ===
function exitToStart() {
  // 强制同步保存一次，再退出
  saveProject(toRaw(project))
    .catch(() => {})
    .finally(() => emit('exit'))
}

// === 初始化 ===
async function refreshModelStatus() {
  try {
    modelStatus.value = await readModelStatus()
    const res = await listExcelProviders()
    providers.value = (res.data || []).map((p: any) => ({
      provider: p.provider,
      providerName: p.providerName,
      modelName: p.modelName,
    }))
  } catch {
    // 静默失败，UI 会显示"未配置"
  }
}

onMounted(() => {
  refreshModelStatus()
})

onBeforeUnmount(() => {
  generationControllers.forEach((c) => c.abort())
  generationControllers.clear()
  window.removeEventListener('pointermove', onResizerPointerMove)
})

// 暴露给父组件的方法
defineExpose({
  refreshModelStatus,
  loadProject(newProject: Project) {
    const draft = cloneProject(newProject)
    hydrateProjectActiveSheet(draft)
    Object.keys(project).forEach((key) => {
      ;(project as any)[key] = (draft as any)[key]
    })
    selection.value = { type: 'project' }
  },
})
</script>

<style scoped>
.workbench-shell {
  --ink: #15221f;
  --muted: #6f7c78;
  --line: #dbe3df;
  --accent: #0f766e;
  height: 100%;
  min-height: 0;
  overflow: hidden;
  display: grid;
  grid-template-rows: 56px 42px minmax(0, 1fr) 26px;
  background: #f4f7f5;
  color: var(--ink);
  font-family: "PingFang SC", "Microsoft YaHei", sans-serif;
}

.workbench-topbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  padding: 0 18px;
  background: #fff;
  border-bottom: 1px solid var(--line);
}
.brand-block {
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
}
.brand-mark {
  width: 30px;
  height: 30px;
  display: grid;
  place-items: center;
  border-radius: 6px;
  color: #fff;
  background: var(--accent);
}
.brand-mark :deep(svg) {
  width: 16px;
  height: 16px;
}
.project-name-input {
  border: 0;
  background: transparent;
  font-size: 15px;
  font-weight: 600;
  color: var(--ink);
  min-width: 100px;
  max-width: 240px;
  padding: 4px 6px;
  border-radius: 4px;
}
.project-name-input:focus {
  outline: 2px solid rgba(15, 118, 110, 0.35);
  background: #fff;
}
.save-status {
  font-size: 11px;
  color: var(--muted);
  padding: 2px 6px;
}
.save-status.status-saving { color: #9a650e; }
.save-status.status-error { color: #b53636; }
.header-actions {
  display: flex;
  align-items: center;
  gap: 6px;
}
.model-chip {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 4px 10px;
  border: 1px solid var(--line);
  border-radius: 16px;
  font-size: 12px;
  color: var(--muted);
  cursor: pointer;
  transition: border-color 0.15s, color 0.15s;
}
.model-chip:hover {
  border-color: var(--accent);
  color: var(--ink);
}
.model-dot {
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: #93a09c;
}
.model-dot.online {
  background: #15803d;
}
.model-dot.demo {
  background: #d89015;
}
.model-quick-select {
  width: 160px;
}
.model-quick-select :deep(.el-input__wrapper) {
  border-radius: 16px;
  background: #f2f8f6;
  box-shadow: 0 0 0 1px var(--line) inset;
}
.model-quick-select :deep(.el-input__wrapper.is-focus),
.model-quick-select :deep(.el-input__wrapper:hover) {
  box-shadow: 0 0 0 1px var(--accent) inset;
}
.model-quick-select :deep(.el-input__inner) {
  font-size: 12px;
  color: var(--ink);
}
.quick-option {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  width: 100%;
}
.hidden-file-input {
  display: none;
}

.tab-bar {
  display: flex;
  align-items: center;
  gap: 2px;
  padding: 0 14px;
  background: #fff;
  border-bottom: 1px solid var(--line);
  overflow-x: auto;
}
.tab-button {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  height: 100%;
  padding: 0 14px;
  border: 0;
  background: transparent;
  color: var(--muted);
  font-size: 13px;
  border-bottom: 2px solid transparent;
  cursor: pointer;
  white-space: nowrap;
  transition: color 0.12s, border-color 0.12s;
}
.tab-button:hover {
  color: var(--ink);
}
.tab-button.active {
  color: var(--accent);
  border-bottom-color: var(--accent);
  font-weight: 600;
}
.tab-button :deep(svg) {
  width: 13px;
  height: 13px;
}
.tab-close {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 16px;
  height: 16px;
  border-radius: 50%;
  font-size: 14px;
  color: var(--muted);
  margin-left: 4px;
}
.tab-close:hover {
  background: #f0e5e5;
  color: #b53636;
}

.workbench-main {
  display: flex;
  min-height: 0;
  background: #fff;
}
.stage-area {
  flex: 1;
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
}
.view-stage-placeholder {
  flex: 1;
  display: grid;
  place-content: center;
  justify-items: center;
  padding: 40px;
  text-align: center;
  color: var(--muted);
}
.placeholder-icon {
  width: 56px;
  height: 56px;
  display: grid;
  place-items: center;
  border-radius: 8px;
  border: 1px solid var(--line);
  background: #f2f7f4;
  color: var(--accent);
  margin-bottom: 12px;
}
.view-stage-placeholder h3 {
  margin: 0 0 8px;
  font-size: 16px;
  color: var(--ink);
}
.view-stage-placeholder p {
  margin: 0 0 16px;
  font-size: 13px;
  max-width: 360px;
  line-height: 1.6;
}
.inspector-resizer {
  flex: 0 0 6px;
  cursor: col-resize;
  background: transparent;
  transition: background 0.15s;
}
.inspector-resizer:hover,
.inspector-resizer:focus-visible {
  background: rgba(15, 118, 110, 0.18);
  outline: none;
}
.inspector-aside {
  flex: 0 0 auto;
  min-height: 0;
  border-left: 1px solid var(--line);
  background: #fbfcfb;
  overflow: hidden;
}

.workbench-statusbar {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 0 14px;
  background: #f7f9f8;
  border-top: 1px solid var(--line);
  font-size: 11px;
  color: var(--muted);
}
.sheet-tabs {
  display: inline-flex;
  align-items: center;
  gap: 4px;
}
.sheet-tab-button {
  border: 0;
  background: transparent;
  font-size: 11px;
  color: var(--muted);
  padding: 2px 8px;
  border-radius: 4px;
  cursor: pointer;
}
.sheet-tab-button:hover {
  background: #eef2ef;
}
.sheet-tab-button.active {
  background: var(--accent);
  color: #fff;
}
.statusbar-count {
  font-weight: 500;
}
.statusbar-spacer {
  flex: 1;
}

.toast {
  position: fixed;
  bottom: 50px;
  left: 50%;
  transform: translateX(-50%);
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 10px 18px;
  background: #18231d;
  color: #fff;
  border-radius: 24px;
  font-size: 13px;
  z-index: 9999;
  box-shadow: 0 6px 16px rgba(0, 0, 0, 0.16);
}
.toast :deep(svg) {
  width: 14px;
  height: 14px;
  color: #6cd399;
}

:global(body.resizing-inspector) {
  cursor: col-resize;
  user-select: none;
}

/* Vue/Element Plus adapters for the migrated React layout. */
.app-shell {
  --accent: var(--green);
}
.app-shell .brand-mark {
  border-radius: 7px;
  background: #1d7b50;
}
.app-shell .brand-mark :deep(.el-icon svg),
.app-shell .icon-button :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.app-shell .header-actions {
  gap: 2px;
}
.app-shell .header-actions .icon-button:disabled {
  opacity: .48;
  cursor: wait;
}
.app-shell .model-quick-select {
  width: 170px;
}
.app-shell .model-quick-select :deep(.el-input__wrapper) {
  min-height: 32px;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 0 0 1px var(--line) inset;
}
.app-shell .model-quick-select :deep(.el-input__wrapper:hover),
.app-shell .model-quick-select :deep(.el-input__wrapper.is-focus) {
  box-shadow: 0 0 0 1px var(--line-strong) inset;
}
.app-shell .model-quick-select :deep(.el-input__inner) {
  color: #46534c;
  font-size: 12px;
}
.app-shell .model-chip {
  height: 32px;
  border-radius: 6px;
  padding: 0 9px;
  background: #fff;
}
.app-shell .inspector-resizer {
  min-width: 8px;
  background: #f3f5f4;
}
.app-shell .content-pane {
  position: relative;
}
.app-shell .sheet-tabs {
  min-width: 0;
  max-width: 52vw;
  height: 24px;
  display: flex;
  gap: 2px;
}
.app-shell .sheet-tabs > :deep(.el-icon) {
  flex: 0 0 auto;
  margin: 0 3px;
}
.app-shell .tab > :deep(.el-icon),
.app-shell .sheet-tabs > :deep(.el-icon) {
  width: 15px;
  height: 15px;
  display: inline-flex;
}
.app-shell .tab > :deep(.el-icon svg),
.app-shell .sheet-tabs > :deep(.el-icon svg) {
  width: 14px;
  height: 14px;
}
.app-shell .tab-close {
  width: 16px;
  height: 16px;
  display: grid;
  place-items: center;
  margin-left: 2px;
  border-radius: 50%;
  color: #89938d;
  opacity: 0;
}
.app-shell .tab:hover .tab-close,
.app-shell .tab:focus-visible .tab-close {
  opacity: 1;
}
.app-shell .tab-close:hover {
  color: var(--red);
  background: var(--red-soft);
}
.app-shell .toast {
  min-height: 38px;
  bottom: 38px;
  padding: 0 13px;
  border-radius: 6px;
  font-size: 11px;
}
@media (max-width: 820px) {
  .app-shell .model-quick-select {
    width: 130px;
    margin-left: auto;
  }
}
@media (max-width: 480px) {
  .app-shell .model-quick-select {
    display: none;
  }
}
</style>
