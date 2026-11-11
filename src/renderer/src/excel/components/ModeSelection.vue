<template>
  <div class="mode-select-shell">
    <header class="start-topbar">
      <div class="brand">
        <div class="brand-mark"><el-icon><Grid /></el-icon></div>
        <span>AI Excel</span>
      </div>
      <div class="topbar-spacer" />
      <button class="model-chip" @click="emit('open-settings')">
        <span class="model-dot" :class="modelName ? 'online' : 'demo'" />
        <span>{{ modelName || '配置大模型' }}</span>
        <el-icon class="chip-caret"><ArrowDown /></el-icon>
      </button>
      <button class="icon-button" aria-label="模型设置" title="模型设置" @click="emit('open-settings')">
        <el-icon><Setting /></el-icon>
      </button>
    </header>

    <main class="mode-select-main">
      <section class="mode-select-content">
        <div class="mode-select-heading">
          <span class="eyebrow">工作方式</span>
          <h1>选择使用模式</h1>
        </div>
        <div class="mode-options">
          <button class="mode-option agent" @click="emit('select-agent')">
            <span class="mode-option-icon"><el-icon><Grid /></el-icon></span>
            <span class="mode-option-copy">
              <strong>手动模式</strong>
              <small>进入表格工作台，编辑数据并创建 AI 应用</small>
            </span>
            <el-icon class="mode-option-arrow"><ArrowRight /></el-icon>
          </button>
          <button class="mode-option chat" @click="emit('select-chat')">
            <span class="mode-option-icon"><el-icon><ChatDotRound /></el-icon></span>
            <span class="mode-option-copy">
              <strong>聊天模式</strong>
              <small>上传多个文件，自动完成跨文件业务分析</small>
            </span>
            <el-icon class="mode-option-arrow"><ArrowRight /></el-icon>
          </button>
        </div>
      </section>
    </main>

    <footer class="start-footer">
      <span>AI Excel 本地数据工作区</span>
      <span>Agent · Chat</span>
    </footer>
  </div>
</template>

<script setup name="AiExcelModeSelection" lang="ts">
import { ArrowDown, ArrowRight, ChatDotRound, Grid, Setting } from '@element-plus/icons-vue'

defineProps({
  modelName: { type: String, default: '' },
})

const emit = defineEmits(['select-agent', 'select-chat', 'open-settings'])
</script>

<style scoped>
/* 所有样式来自全局 styles.css（.excel-page 作用域下），这里只补 element-plus 图标尺寸对齐 */
.mode-select-shell :deep(.el-icon) {
  width: 1em;
  height: 1em;
  font-size: inherit;
}
.brand-mark :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.mode-option-icon :deep(.el-icon svg) {
  width: 22px;
  height: 22px;
}
.mode-option-arrow :deep(.el-icon svg) {
  width: 18px;
  height: 18px;
}
.mode-option-arrow {
  color: #89958e;
}
.icon-button :deep(.el-icon svg) {
  width: 17px;
  height: 17px;
}
.chip-caret :deep(.el-icon svg) {
  width: 13px;
  height: 13px;
}
</style>
