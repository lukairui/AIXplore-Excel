<template>
  <div class="table-stage">
    <div class="table-context-bar">
      <div>
        <strong>{{ displayedRows.length }}</strong>
        <span> / {{ runtimeRowCount }} 行</span>
        <span class="sheet-current-name">· {{ sheetName }}</span>
        <span v-if="ruleActive" class="rule-applied"><el-icon><Filter /></el-icon>规则已应用</span>
      </div>
      <div class="table-context-actions">
        <button class="button compact secondary" @click="emit('add-column')"><el-icon><Plus /></el-icon>新增数据列</button>
        <button class="button compact secondary" @click="emit('add-view')"><el-icon><MagicStick /></el-icon>创建 AI 应用</button>
        <button class="button compact" :class="selection?.type === 'rule' ? 'selected' : 'secondary'" @click="emit('select-rule')"><el-icon><Filter /></el-icon>筛选与排序</button>
        <button class="button compact secondary" @click="emit('add-row')"><el-icon><Plus /></el-icon>新增行</button>
      </div>
    </div>

    <div v-if="ruleError" class="inline-error" :title="ruleError">
      <el-icon><Warning /></el-icon>
      <span>规则运行失败，已显示全部数据：{{ ruleError }}</span>
    </div>

    <div v-if="!columns.length" class="table-empty">
      <p>当前工作表没有列。点击“新增数据列”开始。</p>
    </div>

    <div v-else ref="scrollRef" class="sheet-scroll" @scroll="onScroll">
      <table class="data-sheet">
        <thead>
          <tr>
            <th class="row-index-cell">#</th>
            <th
              v-for="column in columns"
              :key="column.id"
              :class="{ 'selected-column': selection?.type === 'column' && selection.id === column.id }"
              title="点击查看列设置"
              @click="emit('select-column', column.id)"
            >
              <span class="column-type-icon">
                <el-icon v-if="column.kind === 'computed'"><MagicStick /></el-icon>
                <template v-else>{{ column.dataType === 'number' ? '#' : column.dataType === 'date' ? 'D' : column.dataType === 'boolean' ? 'B' : 'T' }}</template>
              </span>
              <span>{{ column.name }}</span>
              <span v-if="column.kind === 'computed'" class="computed-tag">AI</span>
            </th>
            <th class="row-action-head">
              <button class="icon-button subtle" aria-label="新增列" title="新增列" @click="emit('add-column')"><el-icon><Plus /></el-icon></button>
            </th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, rowIndex) in displayedRows" :key="row.id">
            <td class="row-index-cell">{{ rowIndex + 1 }}</td>
            <td
              v-for="column in columns"
              :key="column.id"
              :class="{
                'selected-column': selection?.type === 'column' && selection.id === column.id,
                'computed-cell': column.kind === 'computed',
              }"
            >
              <template v-if="column.kind === 'computed'">
                <span :class="row.errors[column.id] ? 'cell-error' : 'computed-value'" :title="row.errors[column.id] || '计算结果'">
                  {{ row.errors[column.id] ? '#ERROR' : formatCellValue(row.values[column.id], column.dataType) || '—' }}
                </span>
              </template>
              <template v-else-if="column.dataType === 'boolean'">
                <label class="cell-checkbox">
                  <input type="checkbox" :checked="Boolean(row.values[column.id])" @change="onCheckboxChange(row.id, column.id, $event)" />
                </label>
              </template>
              <template v-else>
                <input
                  class="cell-input"
                  :type="column.dataType === 'number' ? 'number' : column.dataType === 'date' ? 'date' : 'text'"
                  :value="row.values[column.id] == null ? '' : String(row.values[column.id])"
                  @input="onInputChange(row.id, column.id, column.dataType, $event)"
                />
              </template>
            </td>
            <td class="row-actions">
              <button class="icon-button subtle danger-hover" :aria-label="`删除第 ${rowIndex + 1} 行`" title="删除行" @click="emit('delete-row', row.id)"><el-icon><Delete /></el-icon></button>
            </td>
          </tr>
          <tr class="add-row-line">
            <td class="row-index-cell"><el-icon><Plus /></el-icon></td>
            <td :colspan="columns.length + 1"><button @click="emit('add-row')">新增一行</button></td>
          </tr>
        </tbody>
      </table>
    </div>
    <div v-if="columns.length" class="agent-horizontal-scroll">
      <input type="range" min="0" max="1000" :value="horizontalPosition" aria-label="表格横向滚动条" @input="onHorizontalRange" />
    </div>
  </div>
</template>

<script setup name="AiExcelTableEditor" lang="ts">
import { ref } from 'vue'
import { Delete, Filter, MagicStick, Plus, Warning } from '@element-plus/icons-vue'
import { formatCellValue } from '../lib/core'
import type { ColumnDefinition, RuntimeRow, Selection } from '../lib/types'

const props = defineProps({
  columns: { type: Array as () => ColumnDefinition[], required: true },
  displayedRows: { type: Array as () => RuntimeRow[], required: true },
  runtimeRowCount: { type: Number, default: 0 },
  ruleActive: { type: Boolean, default: false },
  ruleError: { type: String, default: '' },
  selection: { type: Object as () => Selection | null, default: null },
  sheetName: { type: String, default: '数据' },
})

const emit = defineEmits([
  'add-column',
  'add-view',
  'add-row',
  'delete-row',
  'select-rule',
  'update-cell',
  'select-column',
])

const scrollRef = ref<HTMLElement | null>(null)
const horizontalPosition = ref(0)

function coerceInput(value: string, type: string) {
  if (value === '') return type === 'text' ? '' : null
  if (type === 'number') return value === '' ? null : Number(value)
  return value
}

function updateCell(rowId: string, columnId: string, value: unknown) {
  emit('update-cell', rowId, columnId, value)
}

function onInputChange(rowId: string, columnId: string, type: string, event: Event) {
  updateCell(rowId, columnId, coerceInput((event.target as HTMLInputElement).value, type))
}

function onCheckboxChange(rowId: string, columnId: string, event: Event) {
  updateCell(rowId, columnId, (event.target as HTMLInputElement).checked)
}

function onScroll(event: Event) {
  const target = event.currentTarget as HTMLElement
  const maximum = target.scrollWidth - target.clientWidth
  horizontalPosition.value = maximum > 0 ? Math.round((target.scrollLeft / maximum) * 1000) : 0
}

function onHorizontalRange(event: Event) {
  const scroll = scrollRef.value
  if (!scroll) return
  const position = Number((event.target as HTMLInputElement).value)
  scroll.scrollLeft = ((scroll.scrollWidth - scroll.clientWidth) * position) / 1000
  horizontalPosition.value = position
}
</script>

<style scoped>
.table-editor {
  --ink: #15221f;
  --muted: #6f7c78;
  --line: #dbe3df;
  --accent: #0f766e;
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
  background: #fff;
}
.rule-error {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 14px;
  background: #fff5f0;
  color: #b53636;
  font-size: 12px;
  border-bottom: 1px solid #f5d3c4;
}
.table-empty {
  flex: 1;
  display: grid;
  place-content: center;
  color: var(--muted);
  padding: 40px;
  text-align: center;
}
.table-scroll {
  flex: 1;
  min-height: 0;
  overflow: auto;
  position: relative;
}
.data-table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  table-layout: auto;
  font-size: 13px;
}
.data-table th,
.data-table td {
  border-right: 1px solid var(--line);
  border-bottom: 1px solid var(--line);
  padding: 0;
  vertical-align: middle;
}
.data-table thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  background: #f7f9f8;
  font-weight: 600;
  text-align: left;
  height: 38px;
}
.row-index-col {
  width: 44px;
  min-width: 44px;
  text-align: center;
  color: var(--muted);
  font-size: 11px;
  background: #fafbfa;
}
.column-header {
  cursor: pointer;
  min-width: 120px;
  transition: background 0.12s;
}
.column-header:hover {
  background: #eaf3ee;
}
.column-header.selected {
  background: var(--accent);
  color: #fff;
  box-shadow: inset 0 2px 0 0 #0a4a44;
}
.column-header.selected .column-meta {
  color: rgba(255, 255, 255, 0.78);
}
.column-header-inner {
  display: flex;
  flex-direction: column;
  padding: 4px 10px;
  gap: 2px;
}
.column-name {
  font-weight: 600;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 200px;
}
.column-meta {
  font-size: 10px;
  color: var(--muted);
  font-weight: 400;
}
.column-header.computed .column-name::before {
  content: "ƒ ";
  color: var(--accent);
}
.column-header.selected.computed .column-name::before {
  color: #fff;
}
.add-col-cell {
  width: 100px;
  min-width: 100px;
  text-align: center;
  background: #fafbfa;
}
.data-cell {
  height: 34px;
  padding: 0 !important;
}
.data-cell.is-computed {
  background: #fafdfa;
}
.data-cell.has-error {
  background: #fff5f0;
}
.cell-static {
  display: block;
  padding: 0 10px;
  height: 34px;
  line-height: 34px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.cell-error-tag {
  position: absolute;
}
.cell-input,
.data-cell :deep(.el-input__inner),
.data-cell :deep(.el-input-number) {
  width: 100%;
  height: 34px;
}
.cell-input {
  border: 0;
  background: transparent;
  padding: 0 10px;
  height: 34px;
  line-height: 34px;
  font: inherit;
  color: var(--ink);
}
.cell-input:focus {
  outline: 2px solid rgba(15, 118, 110, 0.45);
  background: #fff;
}
.data-cell :deep(.el-input-number) {
  width: 100%;
}
.data-cell :deep(.el-input-number .el-input__inner) {
  text-align: left;
}
.row-actions-cell {
  width: 44px;
  min-width: 44px;
  text-align: center;
  background: #fafbfa;
  opacity: 0.4;
  transition: opacity 0.12s;
}
tr:hover .row-actions-cell {
  opacity: 1;
}
.table-empty-row {
  text-align: center;
  padding: 32px;
  color: var(--muted);
}
.table-footer {
  display: flex;
  justify-content: space-between;
  align-items: center;
  height: 36px;
  padding: 0 14px;
  border-top: 1px solid var(--line);
  background: #fafbfa;
  font-size: 12px;
}
.footer-info {
  color: var(--muted);
}

/* Vue icon wrappers keep the migrated table geometry intact. */
.table-stage .cell-input {
  width: 100%;
  height: 37px;
  border: 1px solid transparent;
  border-radius: 0;
  padding: 0 10px;
  background: transparent;
  color: #26332c;
  font-size: 12px;
  line-height: normal;
}
.table-stage .cell-input:hover {
  border-color: #ced6d1;
}
.table-stage .cell-input:focus {
  position: relative;
  z-index: 1;
  border-color: #2b8a5f;
  outline: 1px solid #2b8a5f;
}
.table-stage .column-type-icon :deep(.el-icon svg) {
  width: 13px;
  height: 13px;
}
.table-stage .button :deep(.el-icon svg),
.table-stage .icon-button :deep(.el-icon svg),
.table-stage .inline-error :deep(.el-icon svg) {
  width: 15px;
  height: 15px;
}
.table-stage .button > :deep(.el-icon),
.table-stage .icon-button > :deep(.el-icon) {
  font-size: 15px;
}
.table-stage .rule-applied :deep(.el-icon svg),
.table-stage .add-row-line :deep(.el-icon svg) {
  width: 12px;
  height: 12px;
}
</style>
