/**
 * OpenAI 兼容模型请求支撑（移植自 ChatModelRequestSupport.java）。
 *
 * 负责：请求 URL 解析（/v1/chat/completions vs /v1/responses 双协议）、
 * 请求体组装（Qwen 思考禁用 / Gemini 双 token 字段 / 智谱模型名归一化 / Kimi 温度）、
 * 以及响应文本与 SSE 增量的多格式抽取。
 */

import type { ProviderConfig } from '../config'

type Json = Record<string, any>

export interface ModelMessages {
  messages: any[]
}

function defaultString(value: unknown): string {
  return value == null ? '' : String(value)
}

function firstNonBlank(...values: (string | undefined | null)[]): string {
  for (const value of values) {
    if (value != null && String(value).trim() !== '') return String(value)
  }
  return ''
}

export function shouldUseResponsesApi(config: ProviderConfig | null | undefined): boolean {
  const modelName = defaultString(config?.modelName).trim().toLowerCase()
  return modelName.startsWith('gpt-')
}

function normalizeBaseUrl(config: ProviderConfig | null | undefined): string {
  const normalized = defaultString(config?.baseUrl).replace(/\/+$/, '')
  if (normalized.startsWith('https://www.uuapi.net')) {
    return 'https://uuapi.net' + normalized.slice('https://www.uuapi.net'.length)
  }
  if (normalized.startsWith('http://www.uuapi.net')) {
    return 'https://uuapi.net' + normalized.slice('http://www.uuapi.net'.length)
  }
  return normalized
}

function normalizeModelName(config: ProviderConfig | null | undefined): string {
  const modelName = defaultString(config?.modelName).trim()
  const provider = defaultString(config?.provider).toLowerCase()
  if (provider === 'zhipu') {
    return modelName.replace(/_/g, '-').toLowerCase()
  }
  return modelName
}

function normalizeTemperature(config: ProviderConfig | null | undefined, defaultTemperature: number): number {
  const provider = defaultString(config?.provider).toLowerCase()
  const modelName = defaultString(config?.modelName).toLowerCase()
  if (provider === 'kimi' || modelName.startsWith('kimi-')) {
    return 1
  }
  return defaultTemperature
}

function isGeminiLike(config: ProviderConfig | null | undefined): boolean {
  const provider = defaultString(config?.provider).toLowerCase()
  const modelName = defaultString(config?.modelName).toLowerCase()
  return provider === 'gemini' || modelName.startsWith('gemini')
}

export function resolveChatUrl(config: ProviderConfig): string {
  const normalized = normalizeBaseUrl(config)
  if (normalized.endsWith('/responses')) {
    return normalized.replace(/\/responses$/, '/chat/completions')
  }
  if (normalized.endsWith('/chat/completions')) {
    return normalized
  }
  if (normalized.endsWith('/v1')) {
    return normalized + '/chat/completions'
  }
  return normalized + '/v1/chat/completions'
}

export function resolveResponsesUrl(config: ProviderConfig): string {
  const normalized = normalizeBaseUrl(config)
  if (normalized.endsWith('/responses')) {
    return normalized
  }
  if (normalized.endsWith('/chat/completions')) {
    return normalized.replace(/\/chat\/completions$/, '/responses')
  }
  if (normalized.endsWith('/v1')) {
    return normalized + '/responses'
  }
  return normalized + '/v1/responses'
}

export function resolveModelUrl(config: ProviderConfig): string {
  return shouldUseResponsesApi(config) ? resolveResponsesUrl(config) : resolveChatUrl(config)
}

export interface BuildRequestBodyOptions {
  temperature: number
  stream: boolean
  disableThinking?: boolean
  maxOutputTokens?: number | null
  jsonObjectMode?: boolean
}

export function buildRequestBody(
  config: ProviderConfig,
  messages: any[],
  options: BuildRequestBodyOptions
): Json {
  return shouldUseResponsesApi(config)
    ? buildResponsesBody(config, messages, options)
    : buildChatCompletionsBody(config, messages, options)
}

function buildChatCompletionsBody(
  config: ProviderConfig,
  messages: any[],
  options: BuildRequestBodyOptions
): Json {
  const body: Json = {
    model: normalizeModelName(config),
    temperature: normalizeTemperature(config, options.temperature),
    stream: options.stream,
    messages
  }
  if (options.disableThinking) {
    body.enable_thinking = false
    body.thinking = { type: 'disabled' }
  }
  if (options.maxOutputTokens != null && options.maxOutputTokens > 0) {
    body.max_tokens = options.maxOutputTokens
    if (isGeminiLike(config)) {
      body.max_output_tokens = options.maxOutputTokens
    }
  }
  if (options.jsonObjectMode) {
    body.response_format = { type: 'json_object' }
  }
  return body
}

function buildResponsesBody(
  config: ProviderConfig,
  messages: any[],
  options: BuildRequestBodyOptions
): Json {
  const body: Json = {
    model: normalizeModelName(config),
    temperature: normalizeTemperature(config, options.temperature),
    stream: options.stream
  }
  if (options.maxOutputTokens != null && options.maxOutputTokens > 0) {
    body.max_output_tokens = options.maxOutputTokens
  }
  if (options.jsonObjectMode) {
    body.text = { format: { type: 'json_object' } }
  }
  if (hasMediaContent(messages)) {
    body.input = toResponsesInputMessages(messages)
  } else {
    body.input = buildResponsesTextInput(messages)
  }
  return body
}

function hasMediaContent(messages: any[]): boolean {
  if (!Array.isArray(messages)) return false
  for (const message of messages) {
    if (!message) continue
    const content = message.content
    if (!Array.isArray(content)) continue
    for (const part of content) {
      const type = defaultString(part?.type).toLowerCase()
      if (type === 'image_url' || type === 'input_image' || type === 'input_file') {
        return true
      }
    }
  }
  return false
}

function toResponsesInputMessages(messages: any[]): any[] {
  const input: any[] = []
  for (const message of messages || []) {
    if (!message) continue
    input.push({
      type: 'message',
      role: normalizeRole(message.role),
      content: toResponsesInputContent(message.content)
    })
  }
  return input
}

function toResponsesInputContent(content: unknown): unknown {
  if (!Array.isArray(content)) {
    return content == null ? '' : String(content)
  }
  const parts: any[] = []
  for (const part of content) {
    if (!part) continue
    const type = defaultString(part.type).toLowerCase()
    if (type === 'text' || type === 'input_text') {
      parts.push({
        type: 'input_text',
        text: firstNonBlank(part.text, part.value) || ''
      })
    } else if (type === 'image_url' || type === 'input_image') {
      let imageUrl = ''
      const raw = part.image_url
      if (raw && typeof raw === 'object') {
        imageUrl = defaultString((raw as Json).url)
      } else if (raw != null) {
        imageUrl = String(raw)
      }
      if (!imageUrl) imageUrl = defaultString(part.image_url)
      if (imageUrl.trim() !== '') {
        parts.push({ type: 'input_image', image_url: imageUrl })
      }
    }
  }
  return parts.length > 0 ? parts : extractContentText(content)
}

function buildResponsesTextInput(messages: any[]): string {
  const sections: string[] = []
  for (const message of messages || []) {
    if (!message) continue
    const content = extractContentText(message.content)
    if (content.trim() === '') continue
    sections.push(`${responseRoleLabel(message.role)}：\n${content}`)
  }
  return sections.join('\n\n')
}

function normalizeRole(role: unknown): string {
  const normalized = firstNonBlank(role != null ? String(role) : '', 'user').trim().toLowerCase()
  if (normalized === 'system' || normalized === 'developer' || normalized === 'assistant') {
    return normalized
  }
  return 'user'
}

function responseRoleLabel(role: unknown): string {
  const normalized = normalizeRole(role)
  if (normalized === 'system' || normalized === 'developer') return '系统'
  if (normalized === 'assistant') return '助手'
  return '用户'
}

/** 多格式抽取响应正文文本（output_text / choices / content / answer / reply ...） */
export function extractResponseText(responseBody: string): string {
  let json: Json | null = null
  try {
    json = JSON.parse(defaultString(responseBody)) as Json
  } catch {
    return ''
  }
  if (!json) return ''
  return firstNonBlank(
    extractContentText(json.output_text),
    extractResponseOutputContent(json.output),
    extractChoicesText(json.choices),
    extractContentText(json.content),
    extractContentText(json.text),
    extractContentText(json.answer),
    extractContentText(json.reply),
    extractContentText(json.message),
    extractContentText(json.output),
    extractContentText(json.data),
    extractContentText(json.result),
    extractContentText(json.response)
  )
}

/** 从整段 SSE 文本兜底抽取内容（逐 data: 行解析增量并拼接） */
export function extractSseContent(responseBody: string): string {
  const body = defaultString(responseBody).trim()
  if (body === '') return ''
  let full = ''
  let eventData = ''
  let eventName = ''
  const appendEventText = () => {
    if (eventData === '') return
    const data = eventData.trim()
    eventData = ''
    if (data === '' || data === '[DONE]' || !data.startsWith('{')) return
    try {
      const text = extractStreamContent(eventName, JSON.parse(data) as Json)
      if (text.trim() !== '') full += text
    } catch {
      /* 忽略无法解析的事件 */
    }
  }
  for (const line of body.split(/\r?\n/)) {
    const trimmed = line.trim()
    if (trimmed === '') {
      appendEventText()
      eventName = ''
      continue
    }
    if (trimmed.startsWith('event:')) {
      eventName = trimmed.slice(6).trim()
      continue
    }
    if (!trimmed.startsWith('data:')) continue
    if (eventData !== '') eventData += '\n'
    eventData += trimmed.slice(5).trim()
  }
  appendEventText()
  return full !== '' ? full : extractResponseText(responseBody)
}

/** 抽取单个流式事件中的增量文本（chat completions chunk 与 responses API 事件双兼容） */
export function extractStreamContent(eventName: string, json: Json | null | undefined): string {
  if (!json) return ''
  const event = defaultString(eventName)
  const type = defaultString(json.type)
  if (
    event === 'response.output_text.delta' ||
    event === 'response.refusal.delta' ||
    type === 'response.output_text.delta' ||
    type === 'response.refusal.delta'
  ) {
    return defaultString(json.delta)
  }
  const choices = json.choices
  if (!Array.isArray(choices) || choices.length === 0) return ''
  const choice = choices[0] as Json
  return firstNonBlank(
    extractChatMessageText(choice.delta),
    extractChatMessageText(choice.message),
    extractContentText(choice.text),
    extractContentText(choice.content),
    extractContentText(choice.output_text)
  )
}

/** 抽取单个流式事件中的思考文本（reasoning_content 等） */
export function extractStreamReasoning(json: Json | null | undefined): string {
  if (!json) return ''
  const choices = json.choices
  if (!Array.isArray(choices) || choices.length === 0) return ''
  const delta = (choices[0] as Json)?.delta as Json | undefined
  if (!delta) return ''
  return firstNonBlank(
    delta.reasoning_content,
    delta.reasoning,
    delta.reasoningContent,
    delta.thoughts,
    delta.thinking
  )
}

/** 递归抽取任意 JSON 值中的文本 */
export function extractContentText(value: unknown): string {
  if (value == null) return ''
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  if (Array.isArray(value)) {
    let builder = ''
    for (const item of value) {
      const text = extractContentText(item)
      if (text.trim() === '') continue
      if (builder !== '') builder += '\n'
      builder += text
    }
    return builder
  }
  if (typeof value === 'object') {
    const object = value as Json
    return firstNonBlank(
      extractChoicesText(object.choices),
      extractResponseOutputContent(object.output),
      extractContentText(object.text),
      extractContentText(object.output_text),
      extractContentText(object.content),
      extractContentText(object.value),
      extractContentText(object.answer),
      extractContentText(object.reply),
      extractContentText(object.refusal),
      extractContentText(object.message),
      extractContentText(object.delta),
      extractContentText(object.data),
      extractContentText(object.result),
      extractContentText(object.response)
    )
  }
  return String(value)
}

function extractChoicesText(choices: unknown): string {
  if (!Array.isArray(choices) || choices.length === 0) return ''
  const choice = choices[0] as Json
  return firstNonBlank(
    extractChatMessageText(choice.message),
    extractChatMessageText(choice.delta),
    extractContentText(choice.text),
    extractContentText(choice.content),
    extractContentText(choice.output_text)
  )
}

function extractChatMessageText(source: Json | null | undefined): string {
  if (!source) return ''
  return firstNonBlank(
    extractContentText(source.content),
    extractContentText(source.output_text),
    extractContentText(source.text),
    extractContentText(source.answer),
    extractContentText(source.refusal)
  )
}

function extractResponseOutputContent(output: unknown): string {
  if (!Array.isArray(output) || output.length === 0) return ''
  let builder = ''
  for (const item of output) {
    const text = extractContentText((item as Json)?.content)
    if (text.trim() === '') continue
    if (builder !== '') builder += '\n'
    builder += text
  }
  return builder
}
