<template>
  <div class="inspector-panel">
    <!-- Sticky header -->
    <div class="inspector-header">
      <button
        v-if="selection.type !== 'project'"
        class="icon-button inspector-back"
        type="button"
        aria-label="返回项目概览"
        title="返回项目概览"
        @click="emit('select', { type: 'project' })"
      >
        <el-icon><ArrowLeft /></el-icon>
      </button>
      <span v-else class="object-icon"><el-icon><Coin /></el-icon></span>
      <div class="inspector-title">
        <template v-if="selection.type === 'project'">
          <span class="eyebrow">项目概览</span>
          <h2>{{ project.name || '未命名项目' }}</h2>
        </template>
        <template v-else-if="selection.type === 'column' && currentColumn">
          <span class="eyebrow">{{ currentColumn.kind === 'computed' ? '计算列' : '普通列' }}</span>
          <h2>{{ currentColumn.name }}</h2>
        </template>
        <template v-else-if="selection.type === 'rule'">
          <span class="eyebrow">数据表格</span>
          <h2>筛选与排序</h2>
        </template>
        <template v-else-if="selection.type === 'view' && currentView">
          <span class="eyebrow">AI 视图</span>
          <h2>{{ currentView.name }}</h2>
        </template>
      </div>
    </div>

    <div ref="bodyRef" class="inspector-scroll inspector-body">
      <!-- ============ Project overview ============ -->
      <template v-if="selection.type === 'project'">
        <section class="inspector-section overview-grid">
          <div><strong>{{ project.rows.length }}</strong><span>当前表行</span></div>
          <div><strong>{{ project.columns.length }}</strong><span>当前表列</span></div>
          <div><strong>{{ sheetCount }}</strong><span>工作表</span></div>
          <div><strong>{{ project.views.length }}</strong><span>数据应用</span></div>
        </section>

        <section class="inspector-section">
          <div class="section-heading">
            <div><span>从这里开始</span><small>选择一个对象继续</small></div>
          </div>
          <div class="quick-actions">
            <button @click="emit('add-column')"><el-icon><Plus /></el-icon><span><strong>新增数据列</strong><small>普通列或自然语言计算列</small></span></button>
            <button @click="emit('select', { type: 'rule' })"><el-icon><Filter /></el-icon><span><strong>筛选与排序</strong><small>用一句话定义展示规则</small></span></button>
            <button @click="emit('add-view')"><el-icon><MagicStick /></el-icon><span><strong>创建 AI 应用</strong><small>看板、甘特、日历、工具或游戏</small></span></button>
          </div>
        </section>

        <section class="inspector-section">
          <div class="helper-copy"><el-icon><Cloudy /></el-icon><p>{{ modelConfigured ? `当前生成请求由 ${modelName || '已选模型'} 处理。` : '尚未配置大模型。计算列、自然语言规则、AI 应用和跨表图表需要先配置模型。' }}</p></div>
        </section>
      </template>

      <!-- ============ Column ============ -->
      <template v-else-if="selection.type === 'column' && currentColumn">
        <section class="inspector-section">
          <label class="field">
            <span class="field-label">列名称</span>
            <el-input
              v-model="currentColumn.name"
              size="small"
              @change="onColumnNameChange"
              @keyup.enter="onColumnNameChange"
            />
          </label>
          <div class="field">
            <span class="field-label">{{ currentColumn.kind === 'computed' ? '返回类型' : '数据类型' }}</span>
            <el-tag size="small" effect="plain">{{ dataTypeLabel(currentColumn.dataType) }}</el-tag>
          </div>
          <div class="field">
            <span class="field-label">列类型</span>
            <el-tag size="small" :type="currentColumn.kind === 'computed' ? 'success' : 'info'" effect="light">
              {{ currentColumn.kind === 'computed' ? '计算列' : '普通列' }}
            </el-tag>
          </div>
          <el-button
            text
            type="danger"
            :icon="Delete"
            class="danger-link"
            @click="confirmDeleteColumn"
          >
            删除此列
          </el-button>
        </section>

        <section v-if="currentColumn.kind === 'computed' && currentColumn.generation" class="inspector-section">
          <h3 class="section-title">描述</h3>
          <el-input
            type="textarea"
            :rows="4"
            :model-value="currentColumn.generation.draftDescription"
            placeholder="例如：单价乘以数量；任一为空时返回空。"
            @input="(value) => emit('update-column-description', { columnId: currentColumn.id, description: value })"
          />

          <el-button
            class="full-btn primary-btn"
            :icon="Refresh"
            :loading="isColumnGenerating"
            :disabled="isColumnGenerating"
            v-ai-credit="'workflow:ai-excel-tool:generate'"
            @click="handleGenerateColumn"
          >
            <template v-if="isColumnGenerating">
              生成中（{{ columnProgress.attempt }}/{{ columnProgress.maxAttempts }}）…
            </template>
            <template v-else>
              {{ currentColumn.generation.generatedCode ? '重新生成' : '生成代码' }}
            </template>
          </el-button>

          <div class="status-card" :class="columnStateTone">
            <span class="status-dot" :class="columnStateTone" />
            <span>{{ generationLabel(currentColumn.generation) }}</span>
          </div>

          <div v-if="currentColumn.generation.error" class="error-box">
            {{ currentColumn.generation.error }}
          </div>

          <details v-if="currentColumn.generation.generatedCode" class="code-details">
            <summary>查看当前代码</summary>
            <pre>{{ currentColumn.generation.generatedCode }}</pre>
          </details>
        </section>

        <section v-else class="inspector-section">
          <div class="helper">
            普通列保存原始数据。在表格中点击单元格即可编辑，计算列会自动重新计算。
          </div>
        </section>
      </template>

      <!-- ============ Rule ============ -->
      <template v-else-if="selection.type === 'rule'">
        <section class="inspector-section">
          <h3 class="section-title">表格规则描述</h3>
          <el-input
            type="textarea"
            :rows="5"
            :model-value="project.tableRule.draftDescription"
            placeholder="如：仅显示进行中的任务，按优先级从高到低"
            @input="(value) => emit('update-rule-description', value)"
          />
          <div class="hint">
            规则只改变当前工作表的展示行与顺序，不会修改原始数据。
          </div>

          <el-button
            class="full-btn primary-btn"
            :icon="Refresh"
            :loading="isRuleGenerating"
            :disabled="isRuleGenerating"
            v-ai-credit="'workflow:ai-excel-tool:generate'"
            @click="handleGenerateRule"
          >
            <template v-if="isRuleGenerating">
              生成中（{{ ruleProgress.attempt }}/{{ ruleProgress.maxAttempts }}）…
            </template>
            <template v-else>
              生成规则
            </template>
          </el-button>

          <div class="status-card" :class="ruleStateTone">
            <span class="status-dot" :class="ruleStateTone" />
            <span>{{ generationLabel(project.tableRule) }}</span>
          </div>

          <div v-if="project.tableRule.error" class="error-box">
            {{ project.tableRule.error }}
          </div>

          <details v-if="project.tableRule.generatedCode" class="code-details">
            <summary>查看规则代码</summary>
            <pre>{{ project.tableRule.generatedCode }}</pre>
          </details>

          <el-button
            v-if="project.tableRule.generatedCode"
            text
            type="danger"
            :icon="Delete"
            class="danger-link"
            @click="confirmClearRule"
          >
            清除规则
          </el-button>
        </section>
      </template>

      <!-- ============ View ============ -->
      <template v-else-if="selection.type === 'view' && currentView">
        <section class="inspector-section">
          <label class="field">
            <span class="field-label">视图名称</span>
            <el-input
              v-model="currentView.name"
              size="small"
              @change="emit('rename-view', { viewId: currentView.id, name: currentView.name })"
            />
          </label>
          <div class="field">
            <span class="field-label">视图类型</span>
            <el-tag size="small" effect="plain">{{ currentView.chartType === 'custom' ? 'AI 应用' : currentView.chartType }}</el-tag>
          </div>
          <div class="field">
            <span class="field-label">数据来源</span>
            <div class="source-tags">
              <el-tag v-for="sheet in currentViewSheets" :key="sheet.id" size="small" type="success" effect="light">
                {{ sheet.name }} · {{ sheet.rows.length }} 行
              </el-tag>
            </div>
          </div>
          <div class="status-card" :class="viewStateTone">
            <span class="status-dot" :class="viewStateTone" />
            <span>{{ generationLabel(currentView) }}</span>
          </div>
        </section>

        <section class="inspector-section">
          <label class="field">
            <span class="field-label">提示词</span>
            <el-input
              type="textarea"
              :rows="7"
              :model-value="currentView.draftDescription"
              placeholder="例如：展示总金额、进行中任务数和任务金额分布，状态可点击切换。"
              @input="(value) => emit('update-view-description', { viewId: currentView.id, description: value })"
            />
          </label>
          <div v-if="currentView.generatedCode && currentView.sourceVersion !== currentView.generatedVersion" class="draft-warning">
            预览仍使用上一版描述。
          </div>
          <el-button
            v-if="isViewGenerating"
            class="full-btn"
            type="danger"
            plain
            @click="emit('cancel-generation', `view:${currentView.id}`)"
          >
            停止生成（{{ viewProgress.attempt }}/{{ viewProgress.maxAttempts }}）
          </el-button>
          <el-button
            v-else
            class="full-btn primary-btn"
            :icon="Refresh"
            :disabled="!currentView.draftDescription.trim()"
            v-ai-credit="'workflow:ai-excel-tool:generate'"
            @click="handleGenerateView"
          >
            {{ currentView.generatedCode ? '重新生成 AI 应用' : '生成 AI 应用' }}
          </el-button>

          <div v-if="currentView.error || currentView.runtimeError" class="error-box">
            {{ currentView.error || currentView.runtimeError }}
          </div>
        </section>

        <section v-if="currentView.attempts.length" class="inspector-section">
          <details class="code-details" open>
            <summary>自动纠错记录（{{ currentView.attempts.length }}）</summary>
            <div class="attempt-list">
              <article v-for="item in currentView.attempts" :key="`${currentView.id}-${item.attempt}-${item.createdAt}`" class="attempt-card">
                <header>
                  <strong>第 {{ item.attempt }} 轮</strong>
                  <time>{{ formatAttemptTime(item.createdAt) }}</time>
                </header>
                <p>{{ item.error }}</p>
                <details class="attempt-code-details">
                  <summary>查看该轮代码</summary>
                  <pre>{{ item.code }}</pre>
                </details>
              </article>
            </div>
          </details>
        </section>

        <section v-if="currentView.generatedCode" class="inspector-section">
          <details class="code-details">
            <summary>查看 JSX 源码</summary>
            <pre>{{ currentView.generatedCode }}</pre>
          </details>
        </section>

        <section class="inspector-section">
          <el-button text type="danger" :icon="Delete" class="danger-link" @click="emit('delete-view', currentView.id)">
            删除此视图
          </el-button>
        </section>
      </template>

      <!-- ============ Fallback ============ -->
      <template v-else>
        <section class="inspector-section">
          <div class="helper">未选择任何对象。</div>
        </section>
      </template>
    </div>
  </div>
</template>

<script setup name="AiExcelInspectorPanel" lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import {
  ArrowLeft,
  Cloudy,
  Coin,
  Delete,
  Filter,
  MagicStick,
  Plus,
  Refresh,
} from '@element-plus/icons-vue'
// 注：lib/core.ts 中未导出 generationLabel，这里使用本地实现
import type { Project, Selection } from '../lib/types'

const props = defineProps({
  selection: { type: Object as () => Selection, required: true },
  project: { type: Object as () => Project, required: true },
  generationProgress: { type: Object, default: () => ({}) },
  modelConfigured: { type: Boolean, default: false },
  modelName: { type: String, default: '' },
})

const emit = defineEmits([
  'add-column',
  'add-view',
  'select',
  'generate-column',
  'generate-rule',
  'generate-view',
  'cancel-generation',
  'update-column-description',
  'update-rule-description',
  'update-view-description',
  'open-settings',
  'delete-column',
  'rename-column',
  'clear-rule',
  'rename-view',
  'delete-view',
])
const bodyRef = ref<HTMLElement | null>(null)

/* -------------------------------------------------------------------------- */
/*  Helpers                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * lib/core.ts 中未导出 generationLabel，按规格在此本地实现。
 */
function generationLabel(state) {
  if (!state) return ''
  if (state.status === 'generating') return '正在生成…'
  if (state.status === 'error') return '生成失败'
  if (state.status === 'ready' && state.generatedCode) return '已生成'
  return '未生成'
}

function stateTone(state) {
  if (!state) return 'neutral'
  if (state.status === 'error') return 'danger'
  if (state.status === 'generating') return 'busy'
  if (state.status === 'ready' && state.generatedCode) return 'success'
  return 'neutral'
}

const DATA_TYPE_LABELS = {
  text: '文本',
  number: '数字',
  boolean: '布尔',
  date: '日期',
}
function dataTypeLabel(type: string) {
  return DATA_TYPE_LABELS[type as keyof typeof DATA_TYPE_LABELS] || type
}

function formatAttemptTime(value: string | number | Date) {
  if (!value) return ''
  try {
    return new Date(value).toLocaleString('zh-CN', { hour12: false })
  } catch {
    return ''
  }
}

watch(
  () => `${props.selection.type}:${'id' in props.selection ? props.selection.id || '' : ''}`,
  () => {
    nextTick(() => {
      bodyRef.value?.scrollTo({ top: 0 })
    })
  },
)

/* -------------------------------------------------------------------------- */
/*  Project overview                                                           */
/* -------------------------------------------------------------------------- */

const sheets = computed(() => props.project.sheets || [])
const sheetCount = computed(() => sheets.value.length)

/* -------------------------------------------------------------------------- */
/*  Column branch                                                              */
/* -------------------------------------------------------------------------- */

const currentColumn = computed(() => {
  const selection = props.selection
  if (selection.type !== 'column' || !selection.id) return null
  return props.project.columns.find(c => c.id === selection.id) || null
})

const columnProgress = computed(() => {
  if (!currentColumn.value) return { attempt: 0, maxAttempts: 0 }
  return props.generationProgress[`column:${currentColumn.value.id}`] || { attempt: 0, maxAttempts: 0 }
})

const isColumnGenerating = computed(
  () => Boolean(columnProgress.value && columnProgress.value.attempt > 0) ||
    currentColumn.value?.generation?.status === 'generating',
)

const columnStateTone = computed(() =>
  currentColumn.value?.generation ? stateTone(currentColumn.value.generation) : 'neutral',
)

function onColumnNameChange() {
  if (currentColumn.value) {
    emit('rename-column', { columnId: currentColumn.value.id, name: currentColumn.value.name })
  }
}

async function confirmDeleteColumn() {
  if (!currentColumn.value) return
  try {
    await ElMessageBox.confirm(
      '删除该列及其所有原始数据？此操作无法撤销。',
      '删除列',
      { type: 'warning', confirmButtonText: '删除', cancelButtonText: '取消' },
    )
    emit('delete-column', currentColumn.value.id)
  } catch {
    /* 用户取消 */
  }
}

function handleGenerateColumn() {
  if (!currentColumn.value) return
  if (!props.modelConfigured) {
    ElMessage.warning('请先配置模型')
    emit('open-settings')
    return
  }
  emit('generate-column', currentColumn.value.id)
}

/* -------------------------------------------------------------------------- */
/*  Rule branch                                                                */
/* -------------------------------------------------------------------------- */

const ruleProgress = computed(() =>
  props.generationProgress.rule || { attempt: 0, maxAttempts: 0 },
)

const isRuleGenerating = computed(
  () => Boolean(ruleProgress.value && ruleProgress.value.attempt > 0) ||
    props.project.tableRule.status === 'generating',
)

const ruleStateTone = computed(() => stateTone(props.project.tableRule))

function handleGenerateRule() {
  if (!props.modelConfigured) {
    ElMessage.warning('请先配置模型')
    emit('open-settings')
    return
  }
  emit('generate-rule')
}

async function confirmClearRule() {
  try {
    await ElMessageBox.confirm('确定清除当前的筛选与排序规则？', '清除规则', {
      type: 'warning',
      confirmButtonText: '清除',
      cancelButtonText: '取消',
    })
    emit('clear-rule')
  } catch {
    /* 用户取消 */
  }
}

/* -------------------------------------------------------------------------- */
/*  View branch                                                                */
/* -------------------------------------------------------------------------- */

const currentView = computed(() => {
  const selection = props.selection
  if (selection.type !== 'view' || !selection.id) return null
  return props.project.views.find(v => v.id === selection.id) || null
})

const viewProgress = computed(() => {
  if (!currentView.value) return { attempt: 0, maxAttempts: 0 }
  return props.generationProgress[`view:${currentView.value.id}`] || { attempt: 0, maxAttempts: 0 }
})

const isViewGenerating = computed(
  () => Boolean(viewProgress.value && viewProgress.value.attempt > 0) ||
    currentView.value?.status === 'generating',
)

const viewStateTone = computed(() => currentView.value ? stateTone(currentView.value) : 'neutral')
const currentViewSheets = computed(() => {
  const ids = currentView.value?.sourceSheetIds || []
  return sheets.value.filter(sheet => ids.includes(sheet.id))
})

function handleGenerateView() {
  if (!currentView.value) return
  const requiresModel = currentView.value.chartType === 'custom' || (currentView.value.sourceSheetIds?.length || 1) > 1
  if (requiresModel && !props.modelConfigured) {
    ElMessage.warning('请先配置模型')
    emit('open-settings')
    return
  }
  emit('generate-view', currentView.value.id)
}
</script>

<style scoped>
.inspector-panel {
  --ink: #15221f;
  --muted: #6f7c78;
  --line: #dbe3df;
  --accent: #0f766e;
  width: 100%;
  height: 100%;
  display: flex;
  flex-direction: column;
  min-height: 0;
  color: var(--ink);
  background: #fff;
  font-family: "PingFang SC", "Microsoft YaHei", sans-serif;
}

/* ------------------------------ Header ---------------------------------- */
.inspector-header {
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 14px 16px;
  background: #fff;
  border-bottom: 1px solid var(--line);
}
.back-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 28px;
  height: 28px;
  flex: 0 0 auto;
  padding: 0;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
  color: var(--ink);
  cursor: pointer;
  transition: background 0.12s, border-color 0.12s, color 0.12s;
}
.back-btn:hover {
  background: #f2f7f4;
  border-color: var(--accent);
  color: var(--accent);
}
.back-btn :deep(svg) {
  width: 16px;
  height: 16px;
}
.inspector-titles {
  display: flex;
  flex-direction: column;
  min-width: 0;
  gap: 2px;
}
.eyebrow {
  color: var(--muted);
  font-size: 11px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.04em;
}
.inspector-title {
  margin: 0;
  font-size: 15px;
  font-weight: 700;
  color: var(--ink);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/* ------------------------------- Body ----------------------------------- */
.inspector-body {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  padding: 0;
}
.inspector-section {
  padding: 16px;
}
.inspector-section + .inspector-section {
  border-top: 1px solid var(--line);
}
.section-title {
  margin: 0 0 10px;
  font-size: 12px;
  font-weight: 600;
  color: var(--muted);
  letter-spacing: 0.04em;
  text-transform: uppercase;
}

/* ------------------------------ Fields ---------------------------------- */
.field {
  display: block;
  margin-bottom: 12px;
}
.field-label {
  display: block;
  margin-bottom: 4px;
  color: var(--ink);
  font-size: 12.5px;
  font-weight: 600;
}
.field :deep(.el-input__wrapper),
.field :deep(.el-textarea__inner) {
  border-radius: 6px;
}
.field :deep(.el-input__wrapper.is-focus),
.field :deep(.el-input__wrapper:hover) {
  box-shadow: 0 0 0 1px var(--accent) inset;
}

/* --------------------------- Project overview --------------------------- */
.stats-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 8px;
  margin-top: 8px;
}
.stat {
  display: flex;
  flex-direction: column;
  gap: 2px;
  padding: 10px 12px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fafbfa;
}
.stat strong {
  font-size: 17px;
  font-weight: 700;
  color: var(--ink);
}
.stat span {
  font-size: 11.5px;
  color: var(--muted);
}

.sheet-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.sheet-row {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 2px;
  padding: 8px 10px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
  text-align: left;
  cursor: pointer;
  transition: background 0.12s, border-color 0.12s;
}
.sheet-row:hover {
  background: #f2f7f4;
  border-color: #bfe0d8;
}
.sheet-row.active {
  background: #eaf3ee;
  border-color: var(--accent);
}
.sheet-name {
  font-size: 13px;
  font-weight: 600;
  color: var(--ink);
}
.sheet-meta {
  font-size: 11px;
  color: var(--muted);
}

.rule-status {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-top: 8px;
  font-size: 12px;
  color: var(--muted);
}

/* ------------------------------- Buttons -------------------------------- */
.full-btn {
  width: 100%;
  border-radius: 6px;
  margin-top: 10px;
}
.full-btn:first-child {
  margin-top: 0;
}
.full-btn :deep(svg) {
  margin-right: 4px;
}
.primary-btn {
  --el-color-primary: var(--accent);
  --el-color-primary-light-3: #2f8a82;
  --el-color-primary-light-5: #5aa39c;
  --el-color-primary-light-7: #8fc0bb;
  --el-color-primary-light-8: #b3d4d0;
  --el-color-primary-light-9: #d6e9e6;
  --el-color-primary-dark-2: #0c655e;
}

.danger-link {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 0;
  color: #c0392b;
  --el-color-danger: #c0392b;
}
.danger-link :deep(svg) {
  width: 14px;
  height: 14px;
}

/* --------------------------- Status / error ----------------------------- */
.status-card {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-top: 10px;
  padding: 6px 10px;
  border-radius: 6px;
  font-size: 12px;
  background: #f4f7f5;
  color: var(--ink);
}
.status-card.success {
  background: #ecf6ef;
  color: #15803d;
}
.status-card.busy {
  background: #fff8ec;
  color: #8a6310;
}
.status-card.danger {
  background: #fdf2f2;
  color: #b91c1c;
}

.status-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  flex: 0 0 auto;
  background: #b6c2bd;
}
.status-dot.success { background: #15803d; }
.status-dot.busy { background: #d89015; }
.status-dot.danger { background: #dc2626; }
.status-dot.neutral { background: #b6c2bd; }

.error-box {
  margin-top: 8px;
  padding: 8px 10px;
  border-radius: 6px;
  background: #fdf2f2;
  border-left: 3px solid #dc2626;
  color: #7f1d1d;
  font-size: 12px;
  line-height: 1.6;
  white-space: pre-wrap;
  word-break: break-word;
}

/* ------------------------------ Code block ------------------------------ */
.code-details {
  margin-top: 10px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fafbfa;
  overflow: hidden;
}
.code-details > summary {
  padding: 8px 10px;
  font-size: 12px;
  font-weight: 600;
  color: var(--ink);
  cursor: pointer;
  user-select: none;
}
.code-details > summary:hover {
  background: #f2f7f4;
}
.code-details pre {
  margin: 0;
  padding: 10px 12px;
  max-height: 240px;
  overflow: auto;
  border-top: 1px solid var(--line);
  background: #0f1f1c;
  color: #d8f3e8;
  font-family: ui-monospace, Menlo, monospace;
  font-size: 11.5px;
  line-height: 1.55;
  white-space: pre;
}

.attempt-list {
  display: grid;
  gap: 10px;
  padding: 10px 12px 12px;
  border-top: 1px solid var(--line);
}

.attempt-card {
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
  overflow: hidden;
}

.attempt-card header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 8px 10px;
  border-bottom: 1px solid #eef2ef;
  background: #f8faf9;
}

.attempt-card strong {
  font-size: 12px;
  color: var(--ink);
}

.attempt-card time {
  color: var(--muted);
  font-size: 10px;
  white-space: nowrap;
}

.attempt-card p {
  margin: 0;
  padding: 10px;
  color: #4f5d56;
  font-size: 11px;
  line-height: 1.6;
  white-space: pre-wrap;
  word-break: break-word;
}

.attempt-code-details {
  border-top: 1px solid #eef2ef;
}

.attempt-code-details > summary {
  padding: 8px 10px;
  font-size: 11px;
  color: #516159;
  cursor: pointer;
  user-select: none;
}

.attempt-code-details > summary:hover {
  background: #f5f8f6;
}

.attempt-code-details pre {
  margin: 0;
  padding: 10px 12px;
  max-height: 220px;
  overflow: auto;
  border-top: 1px solid #eef2ef;
  background: #0f1f1c;
  color: #d8f3e8;
  font-family: ui-monospace, Menlo, monospace;
  font-size: 11px;
  line-height: 1.55;
}

/* ------------------------------- Misc ----------------------------------- */
.helper {
  padding: 10px 12px;
  border-radius: 6px;
  background: #f4f7f5;
  color: var(--muted);
  font-size: 12.5px;
  line-height: 1.6;
}
.hint {
  margin-top: 6px;
  font-size: 11.5px;
  color: var(--muted);
  line-height: 1.6;
}

.source-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}
.draft-warning {
  margin: 8px 0;
  padding: 7px 8px;
  border-left: 2px solid #d39a37;
  background: #fff9e9;
  color: #805c1e;
  font-size: 11px;
}

/* Match the original workbench inspector while retaining Vue form controls. */
.inspector-panel .inspector-header {
  position: static;
  min-height: 74px;
  gap: 11px;
  padding: 13px 16px;
}
.inspector-panel .inspector-title {
  min-width: 0;
  flex: 1;
  display: block;
  margin: 0;
  color: inherit;
}
.inspector-panel .inspector-title h2 {
  margin: 3px 0 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 14px;
  line-height: 1.25;
  font-weight: 680;
}
.inspector-panel .eyebrow {
  color: #839088;
  font-size: 9px;
  line-height: 1.2;
  font-weight: 720;
  letter-spacing: 0;
}
.inspector-panel .object-icon :deep(.el-icon svg) {
  width: 18px;
  height: 18px;
}
.inspector-panel .inspector-back :deep(.el-icon svg) {
  width: 16px;
  height: 16px;
}
.inspector-panel .inspector-section {
  padding: 16px;
  border-bottom: 1px solid #e9edea;
}
.inspector-panel .inspector-section + .inspector-section {
  border-top: 0;
}
.inspector-panel .overview-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 1px;
  padding: 0;
  background: var(--line);
}
.inspector-panel .overview-grid > div {
  display: grid;
  gap: 3px;
  padding: 15px 16px;
  background: #fff;
}
.inspector-panel .overview-grid strong {
  color: var(--ink);
  font-size: 20px;
  font-variant-numeric: tabular-nums;
}
.inspector-panel .overview-grid span {
  color: #7d8981;
  font-size: 10px;
}
.inspector-panel .quick-actions button > :deep(.el-icon) {
  flex: 0 0 auto;
  color: #237b52;
}
.inspector-panel .quick-actions :deep(.el-icon svg) {
  width: 16px;
  height: 16px;
}
.inspector-panel .helper-copy > :deep(.el-icon) {
  flex: 0 0 auto;
  margin-top: 1px;
}
</style>
