<template>
  <div class="view-stage">
    <div v-if="runtimeError" class="runtime-error-banner">
      <el-icon><Warning /></el-icon>
      <span>{{ runtimeError }}</span>
      <el-button text size="small" :icon="Refresh" @click="$emit('regenerate')">重新生成</el-button>
    </div>
    <div
      v-if="view.status === 'generating'"
      class="view-stage-loading"
      role="status"
      aria-live="polite"
      aria-label="正在生成 AI 应用"
    >
      <div class="loading-scene" aria-hidden="true">
        <span class="loading-sparkle sparkle-one">✦</span>
        <span class="loading-sparkle sparkle-two">✦</span>
        <span class="loading-sparkle sparkle-three">✦</span>
        <div class="sheet-mascot">
          <span class="sheet-fold" />
          <div class="sheet-grid">
            <span v-for="cell in 12" :key="cell" :class="{ filled: [2, 5, 7, 11].includes(cell) }" />
          </div>
          <div class="sheet-face"><i /><i /><b /></div>
        </div>
        <span class="loading-badge badge-chart">▥</span>
        <span class="loading-badge badge-check">✓</span>
      </div>

      <div class="loading-copy">
        <span class="loading-kicker">AI Excel 正在工作</span>
        <h3>正在把想法整理成应用</h3>
        <div class="loading-message">
          <Transition name="message-swap" mode="out-in">
            <p :key="loadingMessageIndex">{{ loadingMessages[loadingMessageIndex] }}</p>
          </Transition>
        </div>
        <div class="loading-progress" aria-hidden="true"><span /></div>
        <div class="loading-dots" aria-hidden="true">
          <span
            v-for="(_, index) in loadingMessages"
            :key="index"
            :class="{ active: index === loadingMessageIndex }"
          />
        </div>
      </div>
    </div>
    <div v-else-if="!view.generatedCode" class="view-stage-empty">
      <span class="empty-mark">✦</span>
      <h3>{{ view.name || '未命名视图' }}</h3>
      <p>在右侧"描述"中写下需求，点击"生成"创建一个只属于这份数据的视图。</p>
    </div>
    <GeneratedViewHost
      v-else
      :project="project"
      :view="view"
      @view-state-change="onViewStateChange"
      @runtime-error="onRuntimeError"
    />
  </div>
</template>

<script setup name="AiExcelViewStage" lang="ts">
import { onBeforeUnmount, ref, watch } from 'vue'
import { Refresh, Warning } from '@element-plus/icons-vue'
import GeneratedViewHost from './GeneratedViewHost.vue'
import type { Project, ViewDefinition, JsonValue } from '../lib/types'

const props = defineProps({
  project: { type: Object as () => Project, required: true },
  view: { type: Object as () => ViewDefinition, required: true },
  // 父级传入的"当前运行时错误"（由父组件统一管理）
  runtimeError: { type: String, default: '' },
})

const emit = defineEmits<{
  (e: 'view-state-change', payload: { viewId: string; state: JsonValue; remount?: boolean }): void
  (e: 'runtime-error', payload: { viewId: string; message: string | null }): void
  (e: 'regenerate'): void
}>()

const loadingMessages = [
  '正在读懂表格里的每一列…',
  '正在把你的想法拆成小步骤…',
  '正在整理指标、关系和图表…',
  '正在检查交互有没有小遗漏…',
  '正在给结果做最后一次梳理…',
]
const loadingMessageIndex = ref(0)
let loadingTimer: number | null = null

function stopLoadingMessages() {
  if (loadingTimer != null) window.clearInterval(loadingTimer)
  loadingTimer = null
}

watch(
  () => props.view.status,
  (status) => {
    stopLoadingMessages()
    loadingMessageIndex.value = 0
    if (status === 'generating') {
      loadingTimer = window.setInterval(() => {
        loadingMessageIndex.value = (loadingMessageIndex.value + 1) % loadingMessages.length
      }, 2200)
    }
  },
  { immediate: true }
)

onBeforeUnmount(stopLoadingMessages)

// 不用 :key 强制重建。GeneratedViewHost 内部的 watch(generatedCode, runtimeEpoch, view.id)
// 已经会在代码变化时重新 mountHost。再加 :key 会触发双重重建（vnode 销毁 + watch 触发）。
function onViewStateChange(payload: { viewId: string; state: JsonValue; remount?: boolean }) {
  emit('view-state-change', payload)
}
function onRuntimeError(payload: { viewId: string; message: string | null }) {
  emit('runtime-error', payload)
}
</script>

<style scoped>
.view-stage {
  --ink: #15221f;
  --muted: #6f7c78;
  --line: #dbe3df;
  --accent: #0f766e;
  display: flex;
  flex-direction: column;
  width: 100%;
  height: 100%;
  min-height: 0;
  background: #fff;
}
.runtime-error-banner {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 14px;
  background: #fff5f0;
  color: #b53636;
  font-size: 12px;
  border-bottom: 1px solid #f5d3c4;
}
.runtime-error-banner span {
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.view-stage-empty {
  flex: 1;
  display: grid;
  place-content: center;
  justify-items: center;
  padding: 40px;
  text-align: center;
  color: var(--muted);
}
.empty-mark {
  font-size: 40px;
  color: var(--accent);
  margin-bottom: 12px;
}
.view-stage-empty h3 {
  margin: 0 0 8px;
  font-size: 16px;
  color: var(--ink);
}
.view-stage-empty p {
  margin: 0;
  font-size: 13px;
  max-width: 360px;
  line-height: 1.6;
}
.view-stage-loading {
  flex: 1;
  min-height: 0;
  display: grid;
  place-content: center;
  justify-items: center;
  gap: 24px;
  padding: 40px 24px;
  overflow: hidden;
  color: var(--ink);
  text-align: center;
  background-color: #fbfcfb;
  background-image:
    linear-gradient(rgba(29, 117, 94, .045) 1px, transparent 1px),
    linear-gradient(90deg, rgba(29, 117, 94, .045) 1px, transparent 1px);
  background-size: 32px 32px;
}
.loading-scene {
  position: relative;
  width: 190px;
  height: 150px;
}
.sheet-mascot {
  position: absolute;
  top: 18px;
  left: 50%;
  width: 118px;
  height: 120px;
  padding: 20px 15px 14px;
  border: 2px solid #176f5b;
  border-radius: 8px;
  background: #fff;
  box-shadow: 8px 10px 0 #d9eee6;
  transform: translateX(-50%);
  animation: sheet-float 2.6s ease-in-out infinite;
}
.sheet-fold {
  position: absolute;
  top: -2px;
  right: -2px;
  width: 25px;
  height: 25px;
  border-left: 2px solid #176f5b;
  border-bottom: 2px solid #176f5b;
  border-radius: 0 6px 0 5px;
  background: #f6c85f;
}
.sheet-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 4px;
}
.sheet-grid span {
  height: 12px;
  border: 1px solid #cddbd6;
  border-radius: 2px;
  background: #f7faf9;
}
.sheet-grid span.filled {
  border-color: #7fcab4;
  background: #aee0d0;
  animation: cell-pulse 1.8s ease-in-out infinite alternate;
}
.sheet-face {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 17px;
  margin-top: 15px;
}
.sheet-face i {
  width: 5px;
  height: 7px;
  border-radius: 50%;
  background: #23423a;
}
.sheet-face b {
  position: absolute;
  bottom: 13px;
  left: 50%;
  width: 15px;
  height: 7px;
  border-bottom: 2px solid #e16f61;
  border-radius: 0 0 12px 12px;
  transform: translateX(-50%);
}
.loading-sparkle {
  position: absolute;
  color: #e5a52f;
  line-height: 1;
  animation: sparkle-pop 1.8s ease-in-out infinite;
}
.sparkle-one {
  top: 8px;
  left: 18px;
  font-size: 22px;
}
.sparkle-two {
  top: 48px;
  right: 7px;
  color: #e16f61;
  font-size: 15px;
  animation-delay: .55s;
}
.sparkle-three {
  bottom: 8px;
  left: 21px;
  color: #4f7fc1;
  font-size: 13px;
  animation-delay: 1.1s;
}
.loading-badge {
  position: absolute;
  display: grid;
  place-items: center;
  width: 32px;
  height: 32px;
  border: 1px solid #cedad6;
  border-radius: 50%;
  background: #fff;
  box-shadow: 0 5px 14px rgba(35, 66, 58, .12);
  font-size: 16px;
  font-weight: 700;
}
.badge-chart {
  right: 12px;
  bottom: 17px;
  color: #4f7fc1;
  animation: badge-bob 2.2s ease-in-out infinite .3s;
}
.badge-check {
  top: 58px;
  left: 2px;
  color: #16856c;
  animation: badge-bob 2.2s ease-in-out infinite .9s;
}
.loading-copy {
  width: min(420px, calc(100vw - 48px));
}
.loading-kicker {
  color: #16856c;
  font-size: 11px;
  font-weight: 700;
  letter-spacing: 0;
}
.loading-copy h3 {
  margin: 8px 0 7px;
  color: #17332c;
  font-size: 20px;
  font-weight: 720;
  letter-spacing: 0;
}
.loading-message {
  min-height: 24px;
}
.loading-message p {
  margin: 0;
  color: #667770;
  font-size: 13px;
  line-height: 24px;
}
.loading-progress {
  width: min(260px, 74%);
  height: 4px;
  margin: 15px auto 13px;
  overflow: hidden;
  border-radius: 2px;
  background: #e5ece9;
}
.loading-progress span {
  display: block;
  width: 42%;
  height: 100%;
  border-radius: inherit;
  background: #16856c;
  animation: progress-travel 1.7s ease-in-out infinite;
}
.loading-dots {
  display: flex;
  justify-content: center;
  gap: 6px;
  height: 7px;
}
.loading-dots span {
  width: 5px;
  height: 5px;
  border-radius: 50%;
  background: #c9d5d1;
  transition: width .25s ease, background-color .25s ease;
}
.loading-dots span.active {
  width: 17px;
  border-radius: 3px;
  background: #e5a52f;
}
.message-swap-enter-active,
.message-swap-leave-active {
  transition: opacity .25s ease, transform .25s ease;
}
.message-swap-enter-from {
  opacity: 0;
  transform: translateY(5px);
}
.message-swap-leave-to {
  opacity: 0;
  transform: translateY(-5px);
}
@keyframes sheet-float {
  0%, 100% { transform: translate(-50%, 0) rotate(-1deg); }
  50% { transform: translate(-50%, -7px) rotate(1deg); }
}
@keyframes cell-pulse {
  from { opacity: .55; }
  to { opacity: 1; }
}
@keyframes sparkle-pop {
  0%, 100% { opacity: .35; transform: scale(.72) rotate(0); }
  50% { opacity: 1; transform: scale(1.08) rotate(12deg); }
}
@keyframes badge-bob {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-5px); }
}
@keyframes progress-travel {
  0% { transform: translateX(-110%); }
  100% { transform: translateX(340%); }
}
@media (max-height: 620px) {
  .view-stage-loading {
    grid-template-columns: 170px minmax(240px, 380px);
    gap: 18px;
  }
  .loading-scene {
    transform: scale(.82);
  }
}
@media (max-width: 640px) {
  .view-stage-loading {
    grid-template-columns: 1fr;
  }
}
@media (prefers-reduced-motion: reduce) {
  .sheet-mascot,
  .sheet-grid span.filled,
  .loading-sparkle,
  .loading-badge,
  .loading-progress span {
    animation: none;
  }
}
</style>
