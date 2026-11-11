/**
 * AI Excel 工具 API（IPC 版）。
 *
 * 与原 RuoYi 平台版 src/api/ai/excelTool.js 保持完全一致的函数签名与
 * 返回结构（{ code, data }，消费方一律 response.data 取值），
 * 内部改走 Electron 主进程 LLM 代理，API Key 不出本机。
 */

interface AixploreBridge {
  invoke: (channel: string, payload?: unknown) => Promise<any>
  onStreamEvent: (
    streamId: string,
    callback: (event: { eventName: string; payload: any }) => void
  ) => () => void
  abort: (id: string) => Promise<unknown>
}

declare const window: Window & { aixplore?: AixploreBridge }

const bridge: AixploreBridge =
  window.aixplore ||
  (({
    invoke: async () => {
      throw new Error('Electron bridge 不可用：请在桌面应用中运行')
    },
    onStreamEvent: () => () => {},
    abort: async () => {}
  } as AixploreBridge))

function randomId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) return crypto.randomUUID()
  return `id-${Date.now()}-${Math.random().toString(36).slice(2)}`
}

/** 模型配置状态（是否已配置 + 可用 provider 列表） */
export async function getExcelStatus() {
  return await bridge.invoke('excel:status')
}

/** 可用模型 provider 列表（不含 API Key） */
export async function listExcelProviders() {
  return await bridge.invoke('excel:providers')
}

/**
 * 执行一次生成请求。
 * @param payload { kind, systemPrompt, userPrompt, images?, files?, provider?, temperature?, maxOutputTokens?, disableThinking? }
 * @returns {Promise<{code: string, provider: string, providerName: string, model: string}>}
 */
export async function generateExcel(payload: any, signal?: AbortSignal) {
  const requestId = randomId()
  if (signal) {
    const onAbort = () => {
      void bridge.abort(requestId)
    }
    signal.addEventListener('abort', onAbort, { once: true })
  }
  try {
    return await bridge.invoke('excel:chat', { payload, requestId })
  } catch (error: any) {
    if (signal?.aborted) throw new Error('已停止生成')
    throw new Error(error?.message || String(error))
  }
}

export interface ExcelStreamHandlers {
  signal?: AbortSignal
  onStatus?: (payload: any) => void
  onDelta?: (text: string, payload: any) => void
  onDone?: (payload: any) => void
}

/**
 * 流式执行一次生成请求（事件协议与原 SSE 版一致：status / delta / done / fail）。
 * @returns 完整结果 { code, provider, providerName, model }
 */
export async function generateExcelStream(payload: any, handlers: ExcelStreamHandlers = {}) {
  const streamId = randomId()
  let fullText = ''
  let donePayload: any = null
  let failed: Error | null = null

  const unsubscribe = bridge.onStreamEvent(streamId, ({ eventName, payload }) => {
    if (eventName === 'delta') {
      const text = String(payload?.text ?? '')
      if (text) fullText += text
      handlers.onDelta?.(text, payload)
      return
    }
    if (eventName === 'status') {
      handlers.onStatus?.(payload)
      return
    }
    if (eventName === 'done') {
      donePayload = payload
      handlers.onDone?.(payload)
      return
    }
    if (eventName === 'fail') {
      failed = new Error(payload?.message || '模型生成失败')
    }
  })

  const finish = () => {
    unsubscribe()
  }

  if (handlers.signal) {
    handlers.signal.addEventListener(
      'abort',
      () => {
        void bridge.abort(streamId)
      },
      { once: true }
    )
  }

  try {
    await bridge.invoke('excel:chat-stream', { payload, streamId })
    // 事件是异步推送的：等待 done / fail 到达
    await new Promise<void>((resolve, reject) => {
      const startedAt = Date.now()
      const poll = () => {
        if (failed) {
          reject(failed)
          return
        }
        if (donePayload || handlers.signal?.aborted) {
          resolve()
          return
        }
        // 20 分钟兜底超时（与后端 STREAM_TIMEOUT 一致量级）
        if (Date.now() - startedAt > 20 * 60 * 1000) {
          reject(new Error('模型生成超时'))
          return
        }
        setTimeout(poll, 60)
      }
      poll()
    })
  } finally {
    finish()
  }

  return {
    ...(donePayload || {}),
    code: fullText || donePayload?.code || ''
  }
}

/**
 * 测试当前 provider 连通性。
 * @returns {Promise<{code: 200, data: {ok: boolean, message?: string, provider?: string, model?: string}}>}
 */
export async function testExcelConnection(payload: any) {
  return await bridge.invoke('excel:test', payload)
}

// ---- 本地配置（设置页使用） ----

export interface LocalProviderConfig {
  provider: string
  providerName: string
  baseUrl: string
  modelName: string
  apiKeys: string
}

export async function getLocalConfig(): Promise<{ providers: LocalProviderConfig[] }> {
  return await bridge.invoke('config:get')
}

export async function saveLocalConfig(config: { providers: LocalProviderConfig[] }) {
  return await bridge.invoke('config:save', config)
}
