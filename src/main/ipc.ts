/**
 * IPC 通道注册：渲染进程 ↔ 主进程。
 *
 * 普通请求走 invoke/handle；流式请求由主进程主动推送
 * 'aixplore:stream-event' 事件（带 streamId 区分），中断通过 AbortController。
 */

import { ipcMain, type WebContents } from 'electron'
import { loadConfig, saveConfig, toProviderSummary, type AppConfig } from './config'
import * as excelChat from './llm/excel-chat'

const STREAM_EVENT_CHANNEL = 'aixplore:stream-event'

interface ChatEnvelope {
  payload: excelChat.ExcelChatRequest
  requestId: string
}

interface StreamEnvelope {
  payload: excelChat.ExcelChatRequest
  streamId: string
}

const abortControllers = new Map<string, AbortController>()

function registerAbort(id: string): AbortController {
  const existing = abortControllers.get(id)
  if (existing) return existing
  const controller = new AbortController()
  abortControllers.set(id, controller)
  return controller
}

function releaseAbort(id: string): void {
  abortControllers.delete(id)
}

function pushStreamEvent(
  sender: WebContents,
  streamId: string,
  eventName: string,
  payload: Record<string, unknown>
): void {
  try {
    if (!sender.isDestroyed()) {
      sender.send(STREAM_EVENT_CHANNEL, { streamId, eventName, payload })
    }
  } catch {
    /* 窗口可能已关闭 */
  }
}

export function registerIpcHandlers(): void {
  // ---- 模型配置状态 ----
  ipcMain.handle('excel:status', () => ({ code: 200, data: excelChat.status() }))

  ipcMain.handle('excel:providers', () => ({
    code: 200,
    data: excelChat.selectAvailableProviders().map(toProviderSummary)
  }))

  // ---- 本地配置 CRUD（设置页） ----
  ipcMain.handle('config:get', () => loadConfig())

  ipcMain.handle('config:save', (_event, config: AppConfig) => saveConfig(config))

  // ---- 非流式生成 ----
  ipcMain.handle('excel:chat', async (event, envelope: ChatEnvelope) => {
    const { payload, requestId } = envelope || {}
    const controller = registerAbort(requestId)
    try {
      const data = await excelChat.generate(payload, controller.signal)
      return { code: 200, data }
    } finally {
      releaseAbort(requestId)
    }
  })

  // ---- 流式生成（事件推送） ----
  ipcMain.handle('excel:chat-stream', async (event, envelope: StreamEnvelope) => {
    const { payload, streamId } = envelope || {}
    const controller = registerAbort(streamId)
    const sender = event.sender
    // 后台执行，事件通过 pushStreamEvent 推送；ack 立即返回
    void excelChat.streamGenerate(
      payload,
      {
        onStatus: (p) => pushStreamEvent(sender, streamId, 'status', p),
        onDelta: (p) => pushStreamEvent(sender, streamId, 'delta', p),
        onDone: (p) => pushStreamEvent(sender, streamId, 'done', p),
        onFail: (p) => pushStreamEvent(sender, streamId, 'fail', p)
      },
      controller.signal
    )
    return { code: 200, data: { ok: true } }
  })

  // ---- 中断 ----
  ipcMain.handle('excel:abort', (_event, envelope: { id: string }) => {
    const controller = abortControllers.get(envelope?.id || '')
    if (controller) controller.abort()
    return { code: 200 }
  })

  // ---- 连通性测试 ----
  ipcMain.handle('excel:test', async (_event, payload: excelChat.ExcelChatRequest) => ({
    code: 200,
    data: await excelChat.testConnection(payload)
  }))
}
