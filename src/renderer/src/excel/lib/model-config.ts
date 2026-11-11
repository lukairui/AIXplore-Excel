import type { BrowserModelConfig } from "./types";

export type ModelProviderPreset = {
  id: string;
  name: string;
  baseUrl: string;
  hint: string;
};

export const MODEL_PROVIDERS: ModelProviderPreset[] = [
  { id: "openai", name: "OpenAI", baseUrl: "https://api.openai.com/v1", hint: "支持 Chat Completions 与 Responses API" },
  { id: "kimi", name: "Kimi / Moonshot", baseUrl: "https://api.moonshot.cn/v1", hint: "填写控制台中的模型 ID" },
  { id: "deepseek", name: "DeepSeek", baseUrl: "https://api.deepseek.com/v1", hint: "OpenAI-compatible 接口" },
  { id: "qwen", name: "通义千问 / DashScope", baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1", hint: "OpenAI-compatible 模式" },
  { id: "zhipu", name: "智谱 GLM", baseUrl: "https://open.bigmodel.cn/api/paas/v4", hint: "OpenAI-compatible 接口" },
  { id: "siliconflow", name: "SiliconFlow", baseUrl: "https://api.siliconflow.cn/v1", hint: "可使用平台提供的多种开源模型" },
  { id: "openrouter", name: "OpenRouter", baseUrl: "https://openrouter.ai/api/v1", hint: "多模型聚合平台" },
  { id: "relay", name: "中转站 / 自定义", baseUrl: "", hint: "填写中转站提供的 Base URL 与模型 ID" },
];

export const DEFAULT_MODEL_CONFIG: BrowserModelConfig = {
  provider: "openai",
  baseUrl: "https://api.openai.com/v1",
  model: "",
  apiKey: "",
  apiMode: "chat-completions",
};

export function migrateModelConfig(value: Partial<BrowserModelConfig> & { enabled?: boolean }) {
  return {
    ...DEFAULT_MODEL_CONFIG,
    ...value,
    provider: value.provider || "relay",
    apiMode: value.apiMode || "chat-completions",
  } satisfies BrowserModelConfig;
}

export function isBrowserModelConfigured(config: BrowserModelConfig) {
  return Boolean(config.baseUrl.trim() && config.model.trim() && config.apiKey.trim());
}

export function providerName(providerId: string) {
  return MODEL_PROVIDERS.find((provider) => provider.id === providerId)?.name || "自定义模型";
}
