<template>
  <el-dialog
    v-model="dialogVisible"
    title="新增列"
    width="440px"
    :close-on-click-modal="false"
    append-to-body
    class="add-column-dialog"
    @close="handleClose"
  >
    <el-form label-position="top" class="add-column-form" @submit.prevent="handleCreate">
      <el-form-item label="列名称" required>
        <el-input
          v-model="name"
          placeholder="列名，如：销售额"
          maxlength="60"
          clearable
        />
      </el-form-item>

      <el-form-item label="列类型">
        <el-radio-group v-model="kind" class="kind-group">
          <el-radio value="manual">普通列</el-radio>
          <el-radio value="computed">计算列</el-radio>
        </el-radio-group>
        <div v-if="kind === 'computed'" class="kind-hint">
          <el-icon><MagicStick /></el-icon>
          <span>由 AI 根据描述生成可执行代码</span>
        </div>
      </el-form-item>

      <el-form-item :label="kind === 'computed' ? '返回类型' : '数据类型'">
        <el-select v-model="dataType" class="type-select">
          <el-option label="文本" value="text" />
          <el-option label="数字" value="number" />
          <el-option label="布尔" value="boolean" />
          <el-option label="日期" value="date" />
        </el-select>
      </el-form-item>
    </el-form>

    <template #footer>
      <span class="dialog-footer">
        <el-button @click="handleClose">取消</el-button>
        <el-button
          class="confirm-button"
          type="primary"
          :disabled="!name.trim()"
          :icon="Plus"
          @click="handleCreate"
        >
          确定
        </el-button>
      </span>
    </template>
  </el-dialog>
</template>

<script setup name="AiExcelAddColumnModal" lang="ts">
import { ref, computed, watch } from 'vue'
import { Plus, MagicStick } from '@element-plus/icons-vue'

const props = defineProps({
  visible: {
    type: Boolean,
    default: false,
  },
})

const emit = defineEmits(['update:visible', 'create', 'close'])

const name = ref('')
const kind = ref<'manual' | 'computed'>('manual')
const dataType = ref<'text' | 'number' | 'boolean' | 'date'>('text')

const dialogVisible = computed({
  get: () => props.visible,
  set: (value: boolean) => emit('update:visible', value),
})

// 每次打开重置表单
watch(
  () => props.visible,
  open => {
    if (open) {
      name.value = ''
      kind.value = 'manual'
      dataType.value = 'text'
    }
  },
)

function handleCreate() {
  const trimmed = name.value.trim()
  if (!trimmed) return
  emit('create', {
    name: trimmed,
    kind: kind.value,
    dataType: dataType.value,
  })
  emit('update:visible', false)
  emit('close')
}

function handleClose() {
  emit('update:visible', false)
  emit('close')
}
</script>

<style scoped>
:global(.add-column-dialog) {
  --ink: #15221f;
  --muted: #6f7c78;
  --line: #dbe3df;
  --accent: #0f766e;
  font-family: "PingFang SC", "Microsoft YaHei", sans-serif;
}

.add-column-dialog :deep(.el-dialog__header) {
  margin-right: 0;
  padding: 18px 22px 14px;
  border-bottom: 1px solid var(--line);
}
.add-column-dialog :deep(.el-dialog__title) {
  color: var(--ink);
  font-size: 17px;
  font-weight: 700;
}
.add-column-dialog :deep(.el-dialog__body) {
  padding: 20px 22px 4px;
}
.add-column-dialog :deep(.el-dialog__footer) {
  padding: 12px 22px 18px;
  border-top: 1px solid var(--line);
}

.add-column-form :deep(.el-form-item) {
  margin-bottom: 18px;
}
.add-column-form :deep(.el-form-item__label) {
  padding-bottom: 6px;
  color: var(--ink);
  font-size: 13px;
  font-weight: 600;
  line-height: 1.5;
}
.add-column-form :deep(.el-form-item.is-required .el-form-item__label)::before {
  color: #dc2626;
}

.add-column-form :deep(.el-input__wrapper),
.add-column-form :deep(.el-select .el-input__wrapper) {
  border-radius: 6px;
}
.add-column-form :deep(.el-input__wrapper.is-focus),
.add-column-form :deep(.el-input__wrapper:hover) {
  box-shadow: 0 0 0 1px var(--accent) inset;
}

.kind-group {
  display: flex;
  gap: 18px;
}
.kind-group :deep(.el-radio__label) {
  color: var(--ink);
  font-size: 13px;
}
.kind-group :deep(.el-radio__input.is-checked .el-radio__inner) {
  background: var(--accent);
  border-color: var(--accent);
}
.kind-group :deep(.el-radio__input.is-checked + .el-radio__label) {
  color: var(--accent);
}

.kind-hint {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-top: 8px;
  color: var(--accent);
  font-size: 12px;
  line-height: 1.6;
}
.kind-hint :deep(svg) {
  width: 14px;
  height: 14px;
}

.type-select {
  width: 100%;
}

.dialog-footer {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}
.dialog-footer .el-button {
  border-radius: 6px;
  min-width: 80px;
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
:global(.add-column-dialog .confirm-button) {
  border-color: #0f766e;
  color: #fff;
  background: #0f766e;
}
:global(.add-column-dialog .confirm-button:hover:not(.is-disabled)) {
  border-color: #0c655e;
  background: #0c655e;
}
</style>
