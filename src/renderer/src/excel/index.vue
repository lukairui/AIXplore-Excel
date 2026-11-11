<template>
  <div class="excel-page">
    <!-- 工作台模式 -->
    <Workbench
      v-if="mode === 'workbench' && currentProject"
      :initial-project="currentProject"
      :selected-provider="selectedProvider"
      @exit="exitWorkbench"
      @provider-change="selectedProvider = $event"
      @upload-excel="onUploadNewExcel"
    />

    <!-- Chat 模式 -->
    <ChatMode
      v-else-if="mode === 'chat'"
      :model-configured="modelConfigured"
      :model-name="modelName"
      :provider="selectedProvider"
      @back="mode = 'select'"
      @open-settings="settingsVisible = true"
      @change-provider="selectedProvider = $event"
    />

    <!-- 手动模式入口：项目列表 + 上传 -->
    <StartScreen
      v-else-if="mode === 'agent'"
      :projects="projects"
      :loading="loadingProjects"
      :busy="uploading"
      :model-name="modelName"
      @upload="onUpload"
      @open="onOpenProject"
      @remove="onRemoveProject"
      @settings="settingsVisible = true"
      @back-to-modes="mode = 'select'"
    />

    <!-- 模式选择页（默认入口） -->
    <ModeSelection
      v-else
      :model-name="modelName"
      @select-agent="enterAgent"
      @select-chat="mode = 'chat'"
      @open-settings="settingsVisible = true"
    />

    <!-- 模型设置弹窗 -->
    <SettingsModal
      v-model:visible="settingsVisible"
      :providers="providers"
      :selected-provider="selectedProvider"
      @update:selected-provider="selectedProvider = $event"
      @close="settingsVisible = false"
      @saved="refreshModel"
    />

    <div v-if="toast" class="toast">
      <el-icon><Check /></el-icon>
      <span>{{ toast }}</span>
    </div>
  </div>
</template>

<script setup name="AiExcelTool" lang="ts">
import { ref, computed, onMounted, watch } from 'vue'
import { Check } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import StartScreen from './components/StartScreen.vue'
import Workbench from './components/Workbench.vue'
import ChatMode from './components/chat/ChatMode.vue'
import ModeSelection from './components/ModeSelection.vue'
import SettingsModal from './components/SettingsModal.vue'
// 引入 ai_excel 原全局样式（chat-*、mode-select-*、start-*、modal-* 等类名）
import './lib/styles.css'
import { listExcelProviders, getExcelStatus } from '@/api/excelTool'
import {
  initializeProjectStore,
  listProjectSummaries,
  loadProjectById,
  saveProject,
  deleteStoredProject,
  type ProjectSummary,
} from './lib/project-store'
import { projectFromExcel } from './lib/excel'
import { projectFromCsv } from './lib/core'
import type { Project } from './lib/types'

type Mode = 'select' | 'agent' | 'workbench' | 'chat'

// 默认进入模式选择页
const mode = ref<Mode>('select')
const currentProject = ref<Project | null>(null)
const projects = ref<ProjectSummary[]>([])
const loadingProjects = ref(true)
const uploading = ref(false)
const settingsVisible = ref(false)
const toast = ref('')

const providers = ref<Array<{ provider: string; providerName: string; modelName: string }>>([])
const selectedProvider = ref<string>(localStorage.getItem('ai-excel.selected-provider') || '')
const modelStatus = ref({ configured: false, model: '' })

const modelConfigured = computed(() => modelStatus.value.configured || providers.value.length > 0)
const modelName = computed(() => {
  if (selectedProvider.value) {
    const p = providers.value.find((x) => x.provider === selectedProvider.value)
    return p ? `${p.providerName || p.provider} · ${p.modelName || ''}` : selectedProvider.value
  }
  return modelStatus.value.model || ''
})

watch(selectedProvider, (val) => {
  localStorage.setItem('ai-excel.selected-provider', val)
})

let toastTimer: number | null = null
function showToast(msg: string) {
  toast.value = msg
  if (toastTimer != null) window.clearTimeout(toastTimer)
  toastTimer = window.setTimeout(() => (toast.value = ''), 2600)
}

async function refreshProjects() {
  loadingProjects.value = true
  try {
    projects.value = await listProjectSummaries()
  } catch (error) {
    ElMessage.error('读取项目列表失败：' + (error instanceof Error ? error.message : String(error)))
  } finally {
    loadingProjects.value = false
  }
}

async function refreshModel() {
  try {
    const statusRes = await getExcelStatus()
    const data = statusRes.data || {}
    modelStatus.value = {
      configured: Boolean(data.configured),
      model: data.configured ? (data.providers?.[0]?.modelName || '') : '',
    }
    const providerRes = await listExcelProviders()
    const rawProviders = ((providerRes.data as any[]) || []).map((p) => ({
      provider: p.provider,
      providerName: p.providerName,
      modelName: p.modelName,
    }))
    providers.value = rawProviders
    const next = providers.value.some((item) => item.provider === selectedProvider.value)
      ? selectedProvider.value
      : providers.value[0]?.provider || ''
    if (next !== selectedProvider.value) {
      selectedProvider.value = next
    }
  } catch {
    // 静默失败
  }
}

async function onUpload(file: File) {
  if (uploading.value) return
  uploading.value = true
  try {
    const project = await parseFileToProject(file)
    project.name = file.name.replace(/\.(xlsx|csv)$/i, '') || '新项目'
    await saveProject(project)
    await refreshProjects()
    currentProject.value = project
    mode.value = 'workbench'
    ElMessage.success('已创建项目并进入工作台')
  } catch (error) {
    ElMessage.error('打开文件失败：' + (error instanceof Error ? error.message : String(error)))
  } finally {
    uploading.value = false
  }
}

async function onUploadNewExcel(file: File) {
  // 工作台中"上传新 Excel" → 创建新项目并切换
  await onUpload(file)
}

async function parseFileToProject(file: File): Promise<Project> {
  const lower = file.name.toLowerCase()
  if (lower.endsWith('.csv')) {
    const text = await file.text()
    return projectFromCsv(text, file.name)
  }
  if (lower.endsWith('.xlsx')) {
    const buffer = await file.arrayBuffer()
    return await projectFromExcel(buffer, file.name)
  }
  throw new Error('仅支持 .xlsx 和 .csv 文件')
}

async function onOpenProject(projectId: string) {
  try {
    const project = await loadProjectById(projectId)
    if (!project) {
      ElMessage.error('项目不存在或已被删除')
      await refreshProjects()
      return
    }
    currentProject.value = project
    mode.value = 'workbench'
  } catch (error) {
    ElMessage.error('打开项目失败：' + (error instanceof Error ? error.message : String(error)))
  }
}

async function onRemoveProject(projectId: string) {
  try {
    await deleteStoredProject(projectId)
    await refreshProjects()
    showToast('项目已删除')
  } catch (error) {
    ElMessage.error('删除失败：' + (error instanceof Error ? error.message : String(error)))
  }
}

function enterAgent() {
  // 进入手动模式：刷新项目列表后展示 StartScreen
  mode.value = 'agent'
  refreshProjects()
}

function exitWorkbench() {
  // 从工作台返回到手动模式的项目列表
  mode.value = 'agent'
  currentProject.value = null
  refreshProjects()
}

onMounted(async () => {
  await Promise.all([
    initializeProjectStore().then(refreshProjects).catch(refreshProjects),
    refreshModel(),
  ])
})
</script>

<style scoped>
.excel-page {
  height: 100%;
  min-height: 0;
  overflow: hidden;
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

@media (max-width: 820px) {
  .excel-page {
    overflow: auto;
  }
}
</style>
