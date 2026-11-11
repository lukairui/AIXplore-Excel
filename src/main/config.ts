/**
 * 模型配置（本地持久化，等价于原平台的 ai_model_config 表）。
 *
 * 原平台由管理员在数据库集中维护模型 Base URL / 模型名 / API Key；
 * 桌面版改为每个用户在本机设置页维护，存储于 userData/config.json。
 */

export interface ProviderConfig {
  /** 唯一 slug（等价原 provider 字段），用于生成请求时选择模型 */
  provider: string
  /** 显示名 */
  providerName: string
  /** OpenAI 兼容 Base URL，如 https://api.openai.com 或 https://xxx/v1 */
  baseUrl: string
  /** 模型名 */
  modelName: string
  /** API Key，多个以逗号或换行分隔（调用时随机轮换） */
  apiKeys: string
}

export interface AppConfig {
  version: number
  providers: ProviderConfig[]
}

import { app } from 'electron'
import fs from 'node:fs'
import path from 'node:path'

const CONFIG_VERSION = 1

function configPath(): string {
  return path.join(app.getPath('userData'), 'config.json')
}

export function loadConfig(): AppConfig {
  try {
    const raw = fs.readFileSync(configPath(), 'utf-8')
    const parsed = JSON.parse(raw) as Partial<AppConfig>
    const providers = Array.isArray(parsed.providers) ? parsed.providers : []
    return {
      version: CONFIG_VERSION,
      providers: providers
        .filter((p) => p && typeof p.provider === 'string' && p.provider)
        .map((p) => ({
          provider: String(p.provider),
          providerName: String(p.providerName || p.provider),
          baseUrl: String(p.baseUrl || ''),
          modelName: String(p.modelName || ''),
          apiKeys: String(p.apiKeys || '')
        }))
    }
  } catch {
    return { version: CONFIG_VERSION, providers: [] }
  }
}

export function saveConfig(config: AppConfig): AppConfig {
  const normalized: AppConfig = {
    version: CONFIG_VERSION,
    providers: (config.providers || [])
      .filter((p) => p && (p.provider || p.providerName))
      .map((p, index) => ({
        provider: slugify(p.provider || p.providerName || `provider-${index + 1}`),
        providerName: String(p.providerName || p.provider || `模型 ${index + 1}`),
        baseUrl: String(p.baseUrl || '').trim(),
        modelName: String(p.modelName || '').trim(),
        apiKeys: String(p.apiKeys || '').trim()
      }))
  }
  // slug 去重（追加序号）
  const seen = new Map<string, number>()
  normalized.providers.forEach((p) => {
    const count = seen.get(p.provider) || 0
    seen.set(p.provider, count + 1)
    if (count > 0) p.provider = `${p.provider}-${count + 1}`
  })
  fs.mkdirSync(path.dirname(configPath()), { recursive: true })
  fs.writeFileSync(configPath(), JSON.stringify(normalized, null, 2), 'utf-8')
  return normalized
}

/** 管理端可见的精简视图（不含 API Key） */
export function toProviderSummary(p: ProviderConfig): {
  provider: string
  providerName: string
  modelName: string
} {
  return { provider: p.provider, providerName: p.providerName, modelName: p.modelName }
}

/** 从显示名生成 slug */
export function slugify(input: string): string {
  const base = String(input || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '-')
    .replace(/^-+|-+$/g, '')
  return base || `provider-${Date.now().toString(36)}`
}
