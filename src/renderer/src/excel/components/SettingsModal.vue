<template>
  <el-dialog
    v-model="dialogVisible"
    title="模型设置"
    width="640px"
    :close-on-click-modal="false"
    append-to-body
    class="settings-dialog"
    @open="loadLocalConfig"
    @close="handleClose"
  >
    <!-- 已配置模型列表 -->
    <div v-if="draftProviders.length > 0" class="provider-list">
      <div
        v-for="(item, index) in draftProviders"
        :key="index"
        class="provider-card"
        :class="{ active: item.provider === draftProvider && item.complete }"
        @click="selectProvider(item)"
      >
        <div class="provider-info">
          <div class="provider-name">
            {{ item.providerName || item.provider }}
            <el-tag v-if="!item.complete" type="warning" size="small" effect="light">未配置完整</el-tag>
          </div>
          <div class="provider-meta">{{ item.modelName || '未设置模型' }} · {{ item.baseUrl || '未设置 Base URL' }}</div>
        </div>
        <div class="provider-actions" @click.stop>
          <el-button link size="small" @click="editProvider(index)">编辑</el-button>
          <el-button link size="small" type="danger" @click="removeProvider(index)">删除</el-button>
        </div>
      </div>
    </div>
    <el-alert
      v-else
      class="empty-alert"
      type="info"
      :closable="false"
      show-icon
      title="还没有模型配置，添加一个 OpenAI 兼容模型即可开始使用"
    />

    <el-button class="add-btn" :icon="Plus" plain size="small" @click="addProvider">+ 添加模型</el-button>

    <!-- 编辑表单 -->
    <el-form v-if="editingActive" label-position="top" class="settings-form">
      <el-divider content-position="left">
        {{ editingIndex == null ? '新增模型' : `编辑：${editing.providerName || editing.provider}` }}
      </el-divider>
      <el-form-item label="名称（用于区分，如：通义千问）" required>
        <el-input v-model="editing.providerName" placeholder="例如：通义千问 / DeepSeek / OpenAI" maxlength="30" />
      </el-form-item>
      <el-form-item label="Base URL（OpenAI 兼容地址）" required>
        <el-input v-model="editing.baseUrl" placeholder="例如：https://dashscope.aliyuncs.com/compatible-mode/v1" />
      </el-form-item>
      <el-form-item label="模型名" required>
        <el-input v-model="editing.modelName" placeholder="例如：qwen-max / deepseek-chat / gpt-4o" />
      </el-form-item>
      <el-form-item label="API Key（多个用逗号或换行分隔，调用时自动轮换）" required>
        <el-input v-model="editing.apiKeys" type="textarea" :rows="2" placeholder="sk-..." />
      </el-form-item>
      <div class="form-actions">
        <el-button size="small" @click="cancelEdit">取消</el-button>
        <el-button size="small" type="primary" @click="applyEdit">保存模型配置</el-button>
      </div>
    </el-form>

    <!-- 默认模型选择 -->
    <el-form v-if="completeProviders.length > 0 && !editingActive" label-position="top" class="settings-form">
      <el-divider content-position="left">默认模型</el-divider>
      <el-form-item>
        <div class="provider-row">
          <el-select v-model="draftProvider" placeholder="请选择默认模型" filterable class="provider-select">
            <el-option
              v-for="item in completeProviders"
              :key="item.provider"
              :label="`${item.providerName} (${item.modelName})`"
              :value="item.provider"
            />
          </el-select>
          <el-button
            :loading="testState === 'testing'"
            :disabled="!draftProvider || testState === 'testing'"
            class="test-btn"
            @click="handleTest"
          >
            {{ testState === 'testing' ? '测试中…' : '测试连接' }}
          </el-button>
        </div>
      </el-form-item>

      <div v-if="testState === 'testing'" class="test-result testing">
        <span class="test-dot" />
        <span class="test-text">正在测试 {{ currentDraftLabel }} 的连通性…</span>
      </div>
      <div v-else-if="testState === 'success'" class="test-result success">
        <div class="test-head">
          <el-tag type="success" effect="light" size="small">连接成功</el-tag>
          <span class="test-model">{{ currentDraftLabel }}</span>
        </div>
        <p v-if="testMessage" class="test-message">{{ testMessage }}</p>
      </div>
      <div v-else-if="testState === 'error'" class="test-result error">
        <div class="test-head">
          <el-tag type="danger" effect="light" size="small">连接失败</el-tag>
          <span class="test-model">{{ currentDraftLabel }}</span>
        </div>
        <p v-if="testError" class="test-message">{{ testError }}</p>
      </div>
    </el-form>

    <template #footer>
      <span class="dialog-footer">
        <el-button @click="handleClose">取消</el-button>
        <el-button type="primary" :disabled="editingActive" @click="handleSave">保存设置</el-button>
      </span>
    </template>
  </el-dialog>
</template>

<script setup name="AiExcelSettingsModal" lang="ts">
import { ref, computed, watch } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Plus } from '@element-plus/icons-vue'
import { testExcelConnection, getLocalConfig, saveLocalConfig, type LocalProviderConfig } from '@/api/excelTool'

interface DraftProvider extends LocalProviderConfig {
  complete: boolean
}

const props = defineProps({
  visible: {
    type: Boolean,
    default: false
  },
  providers: {
    type: Array,
    default: () => []
  },
  selectedProvider: {
    type: String,
    default: ''
  }
})

const emit = defineEmits(['update:visible', 'update:selectedProvider', 'close', 'saved'])

const dialogVisible = computed({
  get: () => props.visible,
  set: value => emit('update:visible', value)
})

const draftProvider = ref(props.selectedProvider)
const draftProviders = ref<DraftProvider[]>([])
const editingActive = ref(false)
const editing = ref<LocalProviderConfig>(blankProvider())
const editingIndex = ref<number | null>(null)

const completeProviders = computed(() => draftProviders.value.filter(item => item.complete))

const testState = ref('idle')
const testMessage = ref('')
const testError = ref('')

const currentDraftLabel = computed(() => {
  const matched = draftProviders.value.find(item => item.provider === draftProvider.value)
  if (!matched) return draftProvider.value || ''
  return `${matched.providerName} (${matched.modelName})`
})

watch(
  () => props.selectedProvider,
  value => {
    draftProvider.value = value
  }
)

watch(
  () => props.visible,
  open => {
    if (open) {
      resetTestState()
    }
  }
)

function blankProvider(): LocalProviderConfig {
  return { provider: '', providerName: '', baseUrl: '', modelName: '', apiKeys: '' }
}

function normalizeDraft(list: LocalProviderConfig[]): DraftProvider[] {
  return list.map(item => ({
    ...item,
    complete: Boolean(item.baseUrl?.trim() && item.modelName?.trim() && item.apiKeys?.trim())
  }))
}

async function loadLocalConfig() {
  try {
    const config = await getLocalConfig()
    draftProviders.value = normalizeDraft(config?.providers || [])
    draftProvider.value = props.selectedProvider || completeProviders.value[0]?.provider || ''
    if (draftProviders.value.length === 0) {
      editing.value = blankProvider()
      editingIndex.value = null
      editingActive.value = true
    }
  } catch (error) {
    ElMessage.error('读取本地配置失败：' + (error instanceof Error ? error.message : String(error)))
  }
}

function addProvider() {
  editing.value = blankProvider()
  editingIndex.value = null
  editingActive.value = true
}

function editProvider(index: number) {
  editing.value = { ...draftProviders.value[index] }
  editingIndex.value = index
  editingActive.value = true
}

function cancelEdit() {
  editingActive.value = false
  editingIndex.value = null
}

function applyEdit() {
  const draft = editing.value
  if (!draft.providerName.trim() || !draft.baseUrl.trim() || !draft.modelName.trim() || !draft.apiKeys.trim()) {
    ElMessage.warning('请完整填写名称、Base URL、模型名和 API Key')
    return
  }
  const provider = draft.provider || slugify(draft.providerName)
  const next: LocalProviderConfig = {
    provider,
    providerName: draft.providerName.trim(),
    baseUrl: draft.baseUrl.trim(),
    modelName: draft.modelName.trim(),
    apiKeys: draft.apiKeys.trim()
  }
  if (editingIndex.value == null) {
    // 新增：slug 冲突时追加序号
    let unique = next.provider
    let suffix = 2
    while (draftProviders.value.some(item => item.provider === unique)) {
      unique = `${next.provider}-${suffix++}`
    }
    next.provider = unique
    draftProviders.value.push({ ...next, complete: true })
  } else {
    draftProviders.value[editingIndex.value] = { ...next, complete: true }
  }
  draftProviders.value = normalizeDraft(draftProviders.value)
  editingActive.value = false
  editingIndex.value = null
}

async function removeProvider(index: number) {
  const target = draftProviders.value[index]
  try {
    await ElMessageBox.confirm(`确定删除模型「${target.providerName || target.provider}」？`, '删除确认', {
      type: 'warning',
      confirmButtonText: '删除',
      cancelButtonText: '取消'
    })
  } catch {
    return
  }
  draftProviders.value.splice(index, 1)
  if (draftProvider.value === target.provider) {
    draftProvider.value = completeProviders.value[0]?.provider || ''
  }
}

function selectProvider(item: DraftProvider) {
  if (item.complete) draftProvider.value = item.provider
}

function slugify(input: string): string {
  const base = String(input || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '-')
    .replace(/^-+|-+$/g, '')
  return base || `provider-${Date.now().toString(36)}`
}

function resetTestState() {
  testState.value = 'idle'
  testMessage.value = ''
  testError.value = ''
}

function truncate(text: unknown, max = 600) {
  if (!text) return ''
  const value = String(text)
  return value.length > max ? `${value.slice(0, max)}…` : value
}

async function handleTest() {
  if (!draftProvider.value || testState.value === 'testing') return
  testState.value = 'testing'
  testMessage.value = ''
  testError.value = ''
  try {
    // 测试前先把当前草稿持久化，保证测试的是最新配置
    await saveLocalConfig({ providers: draftProviders.value })
    const response = await testExcelConnection({ provider: draftProvider.value })
    const data = response?.data || response || {}
    if (data && data.ok === false) {
      testState.value = 'error'
      testError.value = truncate(data.message || data.error || '连接失败')
      return
    }
    testState.value = 'success'
    testMessage.value = truncate(data.message || '连接成功')
  } catch (error: any) {
    testState.value = 'error'
    testError.value = truncate(error?.message || '连接失败，请检查模型配置')
  }
}

async function handleSave() {
  try {
    await saveLocalConfig({ providers: draftProviders.value })
  } catch (error) {
    ElMessage.error('保存失败：' + (error instanceof Error ? error.message : String(error)))
    return
  }
  if (draftProvider.value) {
    emit('update:selectedProvider', draftProvider.value)
  }
  ElMessage.success('模型设置已保存')
  emit('saved')
  emit('update:visible', false)
  emit('close')
}

function handleClose() {
  emit('update:visible', false)
  emit('close')
}
</script>

<style scoped>
.settings-dialog {
  --ink: #15221f;
  --muted: #6f7c78;
  --line: #dbe3df;
  --accent: #0f766e;
  font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;
}

.settings-dialog :deep(.el-dialog__header) {
  margin-right: 0;
  padding: 18px 22px 14px;
  border-bottom: 1px solid var(--line);
}
.settings-dialog :deep(.el-dialog__title) {
  color: var(--ink);
  font-size: 17px;
  font-weight: 700;
}
.settings-dialog :deep(.el-dialog__body) {
  padding: 18px 22px;
  max-height: 62vh;
  overflow-y: auto;
}
.settings-dialog :deep(.el-dialog__footer) {
  padding: 12px 22px 18px;
  border-top: 1px solid var(--line);
}

.empty-alert {
  border: 1px solid #bcd9e8;
  border-radius: 8px;
  background: #eef7fc;
}

.provider-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.provider-card {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 10px 12px;
  border: 1px solid var(--line);
  border-radius: 8px;
  cursor: pointer;
  transition: border-color 0.15s ease, background 0.15s ease;
}
.provider-card:hover {
  border-color: var(--accent);
}
.provider-card.active {
  border-color: var(--accent);
  background: #f2f8f6;
  box-shadow: 0 0 0 1px var(--accent) inset;
}

.provider-info {
  min-width: 0;
  flex: 1;
}
.provider-name {
  display: flex;
  align-items: center;
  gap: 6px;
  color: var(--ink);
  font-weight: 600;
  font-size: 13.5px;
}
.provider-meta {
  margin-top: 2px;
  color: var(--muted);
  font-size: 12px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.provider-actions {
  flex: 0 0 auto;
  display: flex;
  gap: 2px;
}

.add-btn {
  margin-top: 10px;
  width: 100%;
  border-style: dashed;
  color: var(--accent);
  border-color: #bfe0d8;
  background: #f8fbfa;
}

.settings-form :deep(.el-divider--text-left) {
  margin: 18px 0 14px;
  color: var(--muted);
  font-size: 12.5px;
  font-weight: 600;
}
.settings-form :deep(.el-form-item) {
  margin-bottom: 14px;
}
.settings-form :deep(.el-form-item__label) {
  padding-bottom: 6px;
  color: var(--ink);
  font-size: 13px;
  font-weight: 600;
  line-height: 1.5;
}

.form-actions {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}

.provider-row {
  display: flex;
  align-items: center;
  gap: 10px;
  width: 100%;
}
.provider-select {
  flex: 1;
  min-width: 0;
}
.provider-select :deep(.el-input__wrapper) {
  border-radius: 6px;
}
.provider-select :deep(.el-input__wrapper.is-focus) {
  box-shadow: 0 0 0 1px var(--accent) inset;
}

.test-btn {
  flex: 0 0 auto;
  height: 32px;
  padding: 0 16px;
  border-radius: 6px;
  color: var(--accent);
  border-color: #bfe0d8;
  background: #f2f8f6;
}
.test-btn:not(.is-disabled):hover {
  color: #0c655e;
  border-color: var(--accent);
  background: #e6f3f0;
}
.test-btn.is-disabled {
  color: var(--muted);
  background: #f4f6f5;
}

.test-result {
  margin-top: 12px;
  padding: 10px 12px;
  border-radius: 6px;
  font-size: 12.5px;
  line-height: 1.6;
}
.test-result.testing {
  display: flex;
  align-items: center;
  gap: 8px;
  border-left: 3px solid #d89015;
  background: #fff8ec;
  color: #8a6310;
}
.test-result.testing .test-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: #d89015;
  animation: pulse 1s ease-in-out infinite;
}
@keyframes pulse {
  0%,
  100% {
    opacity: 1;
  }
  50% {
    opacity: 0.35;
  }
}
.test-result.success {
  border-left: 3px solid #15803d;
  background: #f2faf4;
  color: #14532d;
}
.test-result.error {
  border-left: 3px solid #dc2626;
  background: #fdf2f2;
  color: #7f1d1d;
}
.test-head {
  display: flex;
  align-items: center;
  gap: 8px;
}
.test-model {
  color: var(--muted);
  font-size: 12px;
}
.test-message {
  margin: 6px 0 0;
  white-space: pre-wrap;
  word-break: break-word;
}

.dialog-footer {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}
.dialog-footer .el-button {
  border-radius: 6px;
  min-width: 72px;
}
.dialog-footer .el-button--primary {
  --el-color-primary: var(--accent);
  --el-color-primary-light-3: #2f8a82;
  --el-color-primary-light-5: #5aa39c;
  --el-color-primary-light-7: #8fc0bb;
  --el-color-primary-light-8: #b3d4d0;
  --el-color-primary-light-9: #d6e9e6;
  --el-color-primary-dark-2: #0c655e;
}
</style>
