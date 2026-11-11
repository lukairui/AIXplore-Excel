<template>
  <div class="start-shell">
    <header class="start-topbar">
      <button class="icon-button" aria-label="返回模式选择" title="返回模式选择" @click="emit('back-to-modes')">
        <el-icon><ArrowLeft /></el-icon>
      </button>
      <div class="brand">
        <div class="brand-mark"><el-icon><Grid /></el-icon></div>
        <span>AI Excel</span>
      </div>
      <div class="topbar-spacer" />
      <button class="model-chip" @click="emit('settings')">
        <span class="model-dot" :class="modelName ? 'online' : 'demo'" />
        <span>{{ modelName || '配置大模型' }}</span>
        <el-icon class="chip-caret"><ArrowDown /></el-icon>
      </button>
      <button class="icon-button" aria-label="模型设置" title="模型设置" @click="emit('settings')">
        <el-icon><Setting /></el-icon>
      </button>
    </header>

    <main class="start-main">
      <section class="start-content">
        <!-- 左侧：上传区 -->
        <div class="start-upload-pane">
          <div class="start-title">
            <span class="eyebrow">本地工作区</span>
            <h1>打开 Excel 开始</h1>
          </div>
          <div
            class="upload-zone"
            :class="{ dragging }"
            @dragenter.prevent="dragging = true"
            @dragover.prevent="dragging = true"
            @dragleave.prevent="dragging = $event.currentTarget === $event.target ? false : dragging"
            @drop.prevent="onDrop"
          >
            <div class="upload-icon">
              <el-icon v-if="busy" class="spin"><Loading /></el-icon>
              <el-icon v-else><UploadFilled /></el-icon>
            </div>
            <strong>{{ busy ? '正在打开工作簿' : '上传 Excel 工作簿' }}</strong>
            <span>支持 .xlsx 文件，也可以直接拖到这里</span>
            <button class="button primary" :disabled="busy" @click="fileInputRef?.click()">
              <el-icon><Upload /></el-icon>选择 Excel 文件
            </button>
            <input
              ref="fileInputRef"
              hidden
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              @change="onInputChange"
            />
          </div>
        </div>

        <!-- 右侧：历史项目 -->
        <section class="history-pane" aria-label="历史项目">
          <div class="history-header">
            <div>
              <el-icon class="history-icon"><Clock /></el-icon>
              <span>
                <strong>历史项目</strong>
                <small>{{ projects.length }} 个本地项目</small>
              </span>
            </div>
            <input
              v-if="projects.length > 3"
              v-model="query"
              placeholder="搜索项目"
              aria-label="搜索历史项目"
            />
          </div>
          <div class="history-list">
            <div v-if="loading" class="history-empty">
              <el-icon class="spin"><Loading /></el-icon>
              <span>正在读取项目</span>
            </div>
            <template v-else-if="visibleProjects.length">
              <div v-for="item in visibleProjects" :key="item.id" class="history-row">
                <button class="history-open" :disabled="busy" @click="emit('open', item.id)">
                  <el-icon><UploadFilled /></el-icon>
                  <span>
                    <strong>{{ item.name }}</strong>
                    <small>
                      {{ item.rowCount }} 行 · {{ item.columnCount }} 列<template v-if="item.sheetCount && item.sheetCount > 1"> · {{ item.sheetCount }} 个工作表</template> · {{ item.viewCount }} 个应用
                    </small>
                  </span>
                  <time>{{ formatTime(item.updatedAt) }}</time>
                </button>
                <button
                  class="icon-button subtle danger-hover"
                  :disabled="busy"
                  :aria-label="`删除项目 ${item.name}`"
                  title="删除项目"
                  @click="confirmRemove(item)"
                >
                  <el-icon><Delete /></el-icon>
                </button>
              </div>
            </template>
            <div v-else class="history-empty">
              <el-icon class="history-icon"><Clock /></el-icon>
              <span>{{ query ? '没有匹配的项目' : '上传 Excel 后，项目会出现在这里' }}</span>
            </div>
          </div>
        </section>
      </section>
    </main>

    <footer class="start-footer">
      <span>项目保存在 IndexedDB</span>
      <span>Excel .xlsx</span>
    </footer>
  </div>
</template>

<script setup name="AiExcelStartScreen" lang="ts">
import { ref, computed } from 'vue'
import { ElMessageBox } from 'element-plus'
import { ArrowDown, ArrowLeft, Clock, Delete, Grid, Loading, Setting, Upload, UploadFilled } from '@element-plus/icons-vue'
import type { ProjectSummary } from '../lib/project-store'

const props = defineProps({
  projects: { type: Array as () => ProjectSummary[], default: () => [] },
  loading: { type: Boolean, default: false },
  busy: { type: Boolean, default: false },
  modelName: { type: String, default: '' },
})

const emit = defineEmits(['upload', 'open', 'remove', 'settings', 'back-to-modes'])

const fileInputRef = ref<HTMLInputElement | null>(null)
const dragging = ref(false)
const query = ref('')

const visibleProjects = computed(() =>
  props.projects.filter(p =>
    p.name.toLocaleLowerCase().includes(query.value.trim().toLocaleLowerCase())
  )
)

const timeFormatter = new Intl.DateTimeFormat('zh-CN', {
  month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit'
})
function formatTime(ts: number) {
  return timeFormatter.format(ts)
}

function selectFile(files: FileList | null) {
  const file = files?.[0]
  if (file && !props.busy) emit('upload', file)
}

function onInputChange(event: Event) {
  selectFile((event.target as HTMLInputElement).files)
  ;(event.target as HTMLInputElement).value = ''
}

function onDrop(event: DragEvent) {
  dragging.value = false
  if (event.dataTransfer?.files) selectFile(event.dataTransfer.files)
}

async function confirmRemove(item: ProjectSummary) {
  try {
    await ElMessageBox.confirm(
      `删除项目"${item.name}"？所有本地数据将丢失。`,
      '删除项目',
      { type: 'warning' }
    )
    emit('remove', item.id)
  } catch {
    /* 取消 */
  }
}
</script>

<style scoped>
/* 全局 styles.css 已定义 .start-shell / .upload-zone / .history-pane 等所有样式 */
/* 这里只调整 element-plus 图标在原生按钮中的尺寸 */
.start-shell :deep(.el-icon) {
  width: 1em;
  height: 1em;
  font-size: inherit;
}
.brand-mark :deep(.el-icon svg) {
  width: 19px;
  height: 19px;
}
.upload-icon :deep(.el-icon svg) {
  width: 28px;
  height: 28px;
}
.button :deep(.el-icon svg) {
  width: 16px;
  height: 16px;
}
.icon-button :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.icon-button.subtle :deep(.el-icon svg) {
  width: 15px;
  height: 15px;
}
.chip-caret :deep(.el-icon svg) {
  width: 13px;
  height: 13px;
}
.history-icon :deep(.el-icon svg) {
  width: 18px;
  height: 18px;
}
.history-header .history-icon {
  color: var(--green);
}
.history-open :deep(.el-icon svg) {
  width: 18px;
  height: 18px;
}
.history-open > :deep(.el-icon) {
  flex: 0 0 auto;
  color: #4a765e;
}
.history-empty :deep(.el-icon svg) {
  width: 18px;
  height: 18px;
}

.spin {
  animation: spin 1.1s linear infinite;
}
@keyframes spin {
  to { transform: rotate(360deg); }
}
</style>
