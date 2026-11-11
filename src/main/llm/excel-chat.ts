/**
 * AI Excel 对话代理服务（移植自 AiExcelChatServiceImpl.java）。
 *
 * 不持久化任何 Excel 数据；仅承担模型 API Key 注入与 SSE 流式转发。
 * 与原实现保持一致的行为：
 *  - 3 次重试（408/425/429/5xx 与网络异常），400ms × attempt 退避
 *  - 多 API Key 随机轮换
 *  - Qwen 系模型自动禁用思考
 *  - 事件协议：status / delta / done / fail
 */

import { loadConfig, toProviderSummary, type ProviderConfig } from '../config'
import {
  buildRequestBody,
  extractResponseText,
  extractSseContent,
  extractStreamContent,
  resolveModelUrl
} from './chat-support'

const MODEL_TIMEOUT_MS = 120_000
const DEFAULT_MAX_OUTPUT_TOKENS = 4096
const MAX_MODEL_ATTEMPTS = 3
const USER_AGENT = 'Mozilla/5.0 AIXplore-Excel/1.0'

export class ServiceError extends Error {}

export interface ExcelChatFile {
  name?: string
  dataUrl?: string
}

export interface ExcelChatRequest {
  kind?: string
  systemPrompt?: string
  userPrompt?: string
  images?: string[]
  files?: ExcelChatFile[]
  provider?: string
  temperature?: number
  maxOutputTokens?: number
  disableThinking?: boolean
}

export interface StreamCallbacks {
  onStatus?: (payload: Record<string, unknown>) => void
  onDelta?: (payload: { text: string }) => void
  onDone?: (payload: Record<string, unknown>) => void
  onFail?: (payload: Record<string, unknown>) => void
}

export function selectAvailableProviders(): ProviderConfig[] {
  return loadConfig().providers.filter(
    (p) => p.baseUrl.trim() !== '' && p.modelName.trim() !== '' && p.apiKeys.trim() !== ''
  )
}

export function status(): { configured: boolean; providers: ReturnType<typeof toProviderSummary>[] } {
  const providers = selectAvailableProviders()
  return {
    configured: providers.length > 0,
    providers: providers.map(toProviderSummary)
  }
}

function resolveModelConfig(provider?: string): ProviderConfig {
  const providers = selectAvailableProviders()
  if (providers.length === 0) {
    throw new ServiceError('尚未配置可用模型，请在「模型设置」中添加')
  }
  const selected = String(provider || '').trim()
  const config = selected !== '' ? providers.find((p) => p.provider === selected) : providers[0]
  if (!config) {
    throw new ServiceError(`模型 ${selected} 不存在或未完整配置`)
  }
  if (config.baseUrl.trim() === '' || config.modelName.trim() === '') {
    throw new ServiceError('AI Excel 只能使用已完整配置的文本模型')
  }
  if (config.apiKeys.trim() === '') {
    throw new ServiceError(`模型未配置 API Key：${config.providerName}`)
  }
  return config
}

function buildMessages(request: ExcelChatRequest): any[] {
  const messages: any[] = [{ role: 'system', content: request.systemPrompt || '' }]
  const images = request.images || []
  const files = request.files || []
  const hasImages = images.length > 0
  const hasFiles = files.some((f) => f && (f.dataUrl || '').trim() !== '')
  if (hasImages || hasFiles) {
    const parts: any[] = []
    for (const image of images) {
      if (!image) continue
      parts.push({ type: 'image_url', image_url: { url: String(image) } })
    }
    for (const file of files) {
      if (!file || !file.dataUrl || file.dataUrl.trim() === '') continue
      parts.push({
        type: 'file',
        file: { filename: file.name || 'upload.bin', file_data: file.dataUrl }
      })
    }
    parts.push({ type: 'text', text: request.userPrompt || '' })
    messages.push({ role: 'user', content: parts })
  } else {
    messages.push({ role: 'user', content: request.userPrompt || '' })
  }
  return messages
}

function chooseApiKey(config: ProviderConfig): string {
  const keys = String(config.apiKeys || '')
    .split(/[,\\n\\r]+/)
    .map((key) => key.trim())
    .filter((key) => key !== '')
  if (keys.length === 0) {
    throw new ServiceError(`模型未配置 API Key：${config.providerName}`)
  }
  return keys[Math.floor(Math.random() * keys.length)]
}

function isQwenLike(config: ProviderConfig): boolean {
  const text = [
    config.provider,
    config.providerName,
    config.modelName,
    config.baseUrl
  ]
    .map((part) => String(part || '').toLowerCase())
    .join(' ')
  return text.includes('qwen') || text.includes('dashscope.aliyuncs.com') || text.includes('阿里百炼')
}

function sleepForRetry(attempt: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 400 * attempt))
}

function modelErrorMessage(statusCode: number, body: string): string {
  let raw = String(body || '')
    .trim()
    .replace(/\s+/g, ' ')
  if (raw.length > 600) raw = raw.slice(0, 600)
  const detail = raw
  try {
    const json = JSON.parse(raw)
    const code = String(json.code || '')
    const message = String(
      json.error?.message || json.message || json.msg || json.detail || ''
    )
    if (code === 'INSUFFICIENT_BALANCE') {
      return '中转站账户余额不足（INSUFFICIENT_BALANCE）'
    }
    if (message !== '' && code !== '') return `${message}（${code}）`
    if (message !== '') return message
  } catch {
    /* 非 JSON 错误体 */
  }
  return `模型请求失败（HTTP ${statusCode}）` + (detail === '' ? '' : `：${detail}`)
}

function isRetryableStatus(status: number): boolean {
  return status === 408 || status === 425 || status === 429 || status >= 500
}

/** 非流式调用（含重试） */
export async function generate(
  request: ExcelChatRequest,
  signal?: AbortSignal
): Promise<{ code: string; provider: string; providerName: string; model: string }> {
  if (!request || !request.systemPrompt?.trim() || !request.userPrompt?.trim()) {
    throw new ServiceError('缺少生成提示词')
  }
  const config = resolveModelConfig(request.provider)
  const messages = buildMessages(request)
  const temperature = request.temperature == null ? 0.2 : request.temperature
  const disableThinking = request.disableThinking === true || isQwenLike(config)
  const maxOutputTokens = request.maxOutputTokens ?? DEFAULT_MAX_OUTPUT_TOKENS
  const body = buildRequestBody(config, messages, {
    temperature,
    stream: false,
    disableThinking,
    maxOutputTokens
  })
  const code = await callModel(config, body, signal)
  return {
    code,
    provider: config.provider,
    providerName: config.providerName,
    model: config.modelName
  }
}

async function callModel(
  config: ProviderConfig,
  body: Record<string, any>,
  signal?: AbortSignal
): Promise<string> {
  const apiKey = chooseApiKey(config)
  const url = resolveModelUrl(config)
  let lastError = '未知错误'
  for (let attempt = 1; attempt <= MAX_MODEL_ATTEMPTS; attempt++) {
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          Accept: 'application/json',
          'User-Agent': USER_AGENT,
          Authorization: `Bearer ${apiKey}`
        },
        body: JSON.stringify(body),
        signal
      })
      const responseText = await response.text()
      if (!response.ok) {
        const message = modelErrorMessage(response.status, responseText)
        if (isRetryableStatus(response.status) && attempt < MAX_MODEL_ATTEMPTS) {
          lastError = message
          await sleepForRetry(attempt)
          continue
        }
        throw new ServiceError(message)
      }
      const text = extractResponseText(responseText)
      if (text.trim() === '') {
        throw new ServiceError('模型没有返回可解析内容')
      }
      return text
    } catch (error) {
      if (error instanceof ServiceError) throw error
      if (signal?.aborted) throw new ServiceError('已停止生成')
      lastError = error instanceof Error && error.message ? error.message : String(error)
      console.warn(`AI Excel 模型调用异常，第 ${attempt}/${MAX_MODEL_ATTEMPTS} 次：`, lastError)
      if (attempt < MAX_MODEL_ATTEMPTS) {
        await sleepForRetry(attempt)
        continue
      }
    }
  }
  throw new ServiceError(`模型调用异常（已重试 ${MAX_MODEL_ATTEMPTS} 次）：${lastError}`)
}

/** 流式调用（含重试），通过 callbacks 推送 status/delta/done/fail 事件 */
export async function streamGenerate(
  request: ExcelChatRequest,
  callbacks: StreamCallbacks,
  signal?: AbortSignal
): Promise<void> {
  if (!request || !request.systemPrompt?.trim() || !request.userPrompt?.trim()) {
    throw new ServiceError('缺少生成提示词')
  }
  const config = resolveModelConfig(request.provider)
  const messages = buildMessages(request)
  const temperature = request.temperature == null ? 0.2 : request.temperature
  const disableThinking = request.disableThinking === true || isQwenLike(config)
  const maxOutputTokens = request.maxOutputTokens ?? DEFAULT_MAX_OUTPUT_TOKENS
  const body = buildRequestBody(config, messages, {
    temperature,
    stream: true,
    disableThinking,
    maxOutputTokens
  })
  try {
    callbacks.onStatus?.({
      message: '正在调用模型生成代码...',
      provider: config.provider,
      providerName: config.providerName,
      model: config.modelName
    })
    const responseText = await callModelStreaming(config, body, (delta) => {
      callbacks.onDelta?.({ text: delta })
    }, signal)
    callbacks.onDone?.({
      code: responseText,
      provider: config.provider,
      providerName: config.providerName,
      model: config.modelName
    })
  } catch (error) {
    if (signal?.aborted) {
      callbacks.onFail?.({ message: '已停止生成' })
      return
    }
    const message = error instanceof Error && error.message ? error.message : String(error)
    callbacks.onFail?.({ message: message || 'AI Excel 生成失败' })
  }
}

async function callModelStreaming(
  config: ProviderConfig,
  body: Record<string, any>,
  deltaHandler: (delta: string) => void,
  signal?: AbortSignal
): Promise<string> {
  const apiKey = chooseApiKey(config)
  const url = resolveModelUrl(config)
  let lastError = '未知错误'
  for (let attempt = 1; attempt <= MAX_MODEL_ATTEMPTS; attempt++) {
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          Accept: 'text/event-stream',
          'Cache-Control': 'no-cache',
          'Accept-Encoding': 'identity',
          'User-Agent': USER_AGENT,
          Authorization: `Bearer ${apiKey}`
        },
        body: JSON.stringify(body),
        signal
      })
      if (!response.ok) {
        const responseText = await response.text().catch(() => '')
        const message = modelErrorMessage(response.status, responseText)
        if (isRetryableStatus(response.status) && attempt < MAX_MODEL_ATTEMPTS) {
          lastError = message
          await sleepForRetry(attempt)
          continue
        }
        throw new ServiceError(message)
      }
      return await readStreamBody(response, deltaHandler)
    } catch (error) {
      if (error instanceof ServiceError) throw error
      if (signal?.aborted) throw new ServiceError('已停止生成')
      lastError = error instanceof Error && error.message ? error.message : String(error)
      console.warn(`AI Excel 流式模型调用异常，第 ${attempt}/${MAX_MODEL_ATTEMPTS} 次：`, lastError)
      if (attempt < MAX_MODEL_ATTEMPTS) {
        await sleepForRetry(attempt)
        continue
      }
    }
  }
  throw new ServiceError(`模型调用异常（已重试 ${MAX_MODEL_ATTEMPTS} 次）：${lastError}`)
}

/** 逐行解析 SSE 流：chat completions chunk 与 responses API 事件双兼容 */
async function readStreamBody(
  response: Response,
  deltaHandler: (delta: string) => void
): Promise<string> {
  if (!response.body) {
    throw new ServiceError('模型没有返回流式内容')
  }
  const reader = response.body.getReader()
  const decoder = new TextDecoder('utf-8')
  let full = ''
  let raw = ''
  let eventData = ''
  let eventName = ''

  const appendStreamEvent = () => {
    if (eventData === '') return
    let json: Record<string, any> | null = null
    try {
      json = JSON.parse(eventData)
    } catch {
      eventData = ''
      return
    }
    const delta = extractStreamContent(eventName, json)
    if (delta != null && delta.trim() !== '') {
      full += delta
      deltaHandler(delta)
    }
    eventData = ''
  }

  const processLine = (line: string): 'continue' | 'break' => {
    raw += line + '\n'
    if (line === '') {
      const completed = eventName === 'response.completed' || eventName === 'response.done'
      appendStreamEvent()
      eventName = ''
      if (completed) return 'break'
      return 'continue'
    }
    if (line.startsWith('event:')) {
      eventName = line.slice(6).trim()
      return 'continue'
    }
    if (!line.startsWith('data:')) return 'continue'
    const value = line.slice(5).trim()
    if (value === '[DONE]') return 'break'
    if (eventData !== '') eventData += '\n'
    eventData += value
    return 'continue'
  }

  let buffer = ''
  outer: while (true) {
    const { value, done } = await reader.read()
    if (done) break
    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split(/\r?\n/)
    buffer = lines.pop() || ''
    for (const line of lines) {
      if (processLine(line) === 'break') break outer
    }
  }
  buffer += decoder.decode()
  for (const line of buffer.split(/\r?\n/)) {
    if (processLine(line) === 'break') break
  }

  if (full !== '') return full
  const fallbackSse = extractSseContent(raw)
  if (fallbackSse.trim() !== '') return fallbackSse
  const fallbackJson = extractResponseText(raw)
  if (fallbackJson.trim() !== '') return fallbackJson
  throw new ServiceError('模型没有返回可解析内容')
}

/** 连通性测试 */
export async function testConnection(
  request: ExcelChatRequest
): Promise<{ ok: boolean; message: string; provider?: string; providerName?: string; model?: string }> {
  let config: ProviderConfig
  try {
    config = resolveModelConfig(request?.provider)
  } catch (error) {
    return { ok: false, message: error instanceof Error ? error.message : String(error) }
  }
  const body = buildRequestBody(
    config,
    [{ role: 'user', content: 'hi' }],
    { temperature: 0.2, stream: false, disableThinking: isQwenLike(config), maxOutputTokens: 64 }
  )
  try {
    const text = await callModel(config, body)
    return {
      ok: true,
      message: text || '',
      provider: config.provider,
      providerName: config.providerName,
      model: config.modelName
    }
  } catch (error) {
    return { ok: false, message: error instanceof Error ? error.message : String(error) }
  }
}
