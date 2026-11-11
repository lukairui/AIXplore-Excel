<template>
  <div ref="containerRef" class="generated-view-host"></div>
</template>

<script setup name="AiExcelGeneratedViewHost" lang="ts">
import { ref, watch, onBeforeUnmount, shallowRef, toRaw } from 'vue'
import {
  createGeneratedViewHost,
  type GeneratedViewHostBindings,
} from '../lib/generated-view-host'
import { resolveViewData } from '../lib/data-sources'
import { calculateRuntimeRows } from '../lib/core'
import type { Project, ViewDefinition, JsonValue } from '../lib/types'

const props = defineProps({
  project: { type: Object as () => Project, required: true },
  view: { type: Object as () => ViewDefinition, required: true },
})

const emit = defineEmits<{
  (e: 'view-state-change', payload: { viewId: string; state: JsonValue; remount?: boolean }): void
  (e: 'runtime-error', payload: { viewId: string; message: string | null }): void
}>()

const containerRef = ref<HTMLElement | null>(null)
const hostRef = shallowRef<ReturnType<typeof createGeneratedViewHost> | null>(null)

function buildBindings(): GeneratedViewHostBindings {
  // 注意：不要 toRaw(props.project)，否则子组件内部的 computed 无法响应数据变化。
  // 直接保留 reactive 引用，actions 修改 props.project 的字段时会自动触发子组件更新。
  const project = props.project
  const view = props.view
  const resolved = resolveViewData(project, view.sourceSheetIds)
  const actions = {
    updateCell: (rowId: string, columnId: string, value: unknown) => {
      const row = project.rows.find((r) => r.id === rowId)
      if (row) row.values[columnId] = value
    },
    updateRow: (rowId: string, values: Record<string, unknown>) => {
      const row = project.rows.find((r) => r.id === rowId)
      if (row) Object.assign(row.values, values)
    },
    batchUpdateRows: (updates: Array<{ rowId: string; values: Record<string, unknown> }>) => {
      updates.forEach((u) => {
        const row = project.rows.find((r) => r.id === u.rowId)
        if (row) Object.assign(row.values, u.values)
      })
    },
    addRow: (values?: Record<string, unknown>) => {
      const id = `row_${Math.random().toString(36).slice(2, 9)}`
      const empty: Record<string, unknown> = {}
      project.columns.forEach((c) => {
        empty[c.id] = c.dataType === 'number' || c.dataType === 'boolean' ? null : ''
      })
      project.rows.push({ id, values: { ...empty, ...(values || {}) } })
      return id
    },
    deleteRow: (rowId: string) => {
      const idx = project.rows.findIndex((r) => r.id === rowId)
      if (idx >= 0) project.rows.splice(idx, 1)
    },
    deleteRows: (rowIds: string[]) => {
      const set = new Set(rowIds)
      const removed = project.rows.filter((r) => !set.has(r.id))
      project.rows.splice(0, project.rows.length, ...removed)
    },
    reorderRows: (rowIds: string[]) => {
      const orderMap = new Map(rowIds.map((id, idx) => [id, idx]))
      const listed = rowIds
        .map((id) => project.rows.find((r) => r.id === id))
        .filter((r): r is NonNullable<typeof r> => Boolean(r))
      const rest = project.rows.filter((r) => !orderMap.has(r.id))
      const next = [...listed, ...rest]
      project.rows.splice(0, project.rows.length, ...next)
    },
  }

  return {
    view,
    rows: resolved.rows,
    allRows: resolved.allRows,
    columns: resolved.columns,
    sourceSheets: resolved.sheets,
    relations: resolved.relations,
    ...actions,
    onViewStateChange: (viewId, state, remount) => {
      emit('view-state-change', { viewId, state, remount })
    },
    onRuntimeError: (message) => {
      emit('runtime-error', { viewId: view.id, message })
    },
  }
}

function mountHost() {
  if (!containerRef.value) return
  if (hostRef.value) {
    hostRef.value.dispose()
    hostRef.value = null
  }
  containerRef.value.innerHTML = ''
  try {
    const host = createGeneratedViewHost(containerRef.value, buildBindings())
    host.mount()
    hostRef.value = host
  } catch (error) {
    // 同步挂载错误上抛
    emit('runtime-error', {
      viewId: props.view.id,
      message: error instanceof Error ? error.message : String(error),
    })
  }
}

// 当 generatedCode 或 runtimeEpoch 变化时重新挂载（对齐 React 原版：只有代码变了才重建）
watch(
  () => [props.view.generatedCode, props.view.runtimeEpoch, props.view.id],
  () => {
    mountHost()
  },
  { immediate: false }
)

// 注意：数据行/列/appState 变化时【不 remount】。
// 生成的组件通过 props 拿到的是 reactive 对象引用，组件内部的 computed/watch
// 会自动响应数据变化（如同 React 原版的 reconciliation）。
// 之前的实现每次 rows.length 变化都 createApp+mount，对含 echarts 的大屏会极慢。

import { onMounted } from 'vue'
onMounted(() => {
  mountHost()
})

onBeforeUnmount(() => {
  if (hostRef.value) {
    hostRef.value.dispose()
    hostRef.value = null
  }
})

// 暴露给父组件的方法
defineExpose({
  remount: mountHost,
})
</script>

<style scoped>
.generated-view-host {
  width: 100%;
  height: 100%;
  min-height: 0;
  overflow: auto;
  background: #fff;
}
.generated-view-host :deep(.empty-view-stage),
.generated-view-host :deep(.view-compile-empty),
.generated-view-host :deep(.view-runtime-error) {
  /* 这些 class 来自全局 styles.css */
}
</style>
