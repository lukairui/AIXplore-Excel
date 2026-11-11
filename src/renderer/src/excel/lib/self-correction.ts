import type { ColumnDefinition } from "./types";

export const MAX_GENERATION_ATTEMPTS = 3;

export type SelfCorrectionFailure = {
  attempt: number;
  rawError: string;
  feedback: string;
  code: string;
};

export type SelfCorrectionContext = {
  attempt: number;
  maxAttempts: number;
  previousFailure: SelfCorrectionFailure | null;
  failures: SelfCorrectionFailure[];
};

type SelfCorrectionOptions = {
  generate: (context: SelfCorrectionContext) => Promise<string>;
  clean: (output: string) => string;
  validate: (code: string) => void | Promise<void>;
  humanize: (error: unknown, code: string) => string;
  maxAttempts?: number;
  signal?: AbortSignal;
  onAttempt?: (context: SelfCorrectionContext) => void;
  onFailure?: (failure: SelfCorrectionFailure) => void;
};

export class SelfCorrectionExhaustedError extends Error {
  failures: SelfCorrectionFailure[];
  lastCode: string;

  constructor(failures: SelfCorrectionFailure[]) {
    const last = failures.at(-1);
    super(`${last?.feedback || "生成结果未通过校验"}\n已完成 ${failures.length} 轮自动纠错，仍未通过确定性校验。`);
    this.name = "SelfCorrectionExhaustedError";
    this.failures = failures;
    this.lastCode = last?.code || "";
  }
}

function throwIfAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  const error = new Error("已停止自动纠错");
  error.name = "AbortError";
  throw error;
}

export async function runSelfCorrectingGeneration(options: SelfCorrectionOptions) {
  const maxAttempts = Math.max(1, Math.floor(options.maxAttempts || MAX_GENERATION_ATTEMPTS));
  const failures: SelfCorrectionFailure[] = [];

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    throwIfAborted(options.signal);
    const context: SelfCorrectionContext = {
      attempt,
      maxAttempts,
      previousFailure: failures.at(-1) || null,
      failures: [...failures],
    };
    options.onAttempt?.(context);
    const output = await options.generate(context);
    throwIfAborted(options.signal);
    const code = options.clean(output);
    try {
      await options.validate(code);
      return { code, attempts: attempt, failures };
    } catch (error) {
      const failure: SelfCorrectionFailure = {
        attempt,
        rawError: error instanceof Error ? error.message : String(error),
        feedback: options.humanize(error, code),
        code,
      };
      failures.push(failure);
      options.onFailure?.(failure);
    }
  }

  throw new SelfCorrectionExhaustedError(failures);
}

function findClosingDelimiter(source: string, openingIndex: number, opening: string, closing: string) {
  let depth = 0;
  let quote: "'" | '"' | "`" | null = null;
  let escaped = false;
  let lineComment = false;
  let blockComment = false;

  for (let index = openingIndex; index < source.length; index += 1) {
    const char = source[index];
    const next = source[index + 1];
    if (lineComment) {
      if (char === "\n") lineComment = false;
      continue;
    }
    if (blockComment) {
      if (char === "*" && next === "/") {
        blockComment = false;
        index += 1;
      }
      continue;
    }
    if (quote) {
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (char === quote) {
        quote = null;
      }
      continue;
    }
    if (char === "/" && next === "/") {
      lineComment = true;
      index += 1;
      continue;
    }
    if (char === "/" && next === "*") {
      blockComment = true;
      index += 1;
      continue;
    }
    if (char === "'" || char === '"' || char === "`") {
      quote = char;
      continue;
    }
    if (char === opening) depth += 1;
    if (char === closing && --depth === 0) return index;
  }
  return -1;
}

export function cleanGeneratedViewCode(output: string) {
  const cleaned = output
    .trim()
    .replace(/^```(?:javascript|js|jsx|tsx|vue)?\s*/i, "")
    .replace(/\s*```[\s\S]*$/i, "")
    .trim();
  const match = /function\s+GeneratedView\s*\(/.exec(cleaned);
  if (!match) return cleaned;
  const openingParen = cleaned.indexOf("(", match.index);
  const closingParen = findClosingDelimiter(cleaned, openingParen, "(", ")");
  const openingBrace = cleaned.indexOf("{", closingParen + 1);
  if (openingBrace < 0) return cleaned.slice(match.index).trim();
  const closingBrace = findClosingDelimiter(cleaned, openingBrace, "{", "}");
  return closingBrace < 0
    ? cleaned.slice(match.index).trim()
    : cleaned.slice(match.index, closingBrace + 1).trim();
}

function codeSnippet(sourceCode: string, message: string) {
  const lineMatch = message.match(/\((\d+):(\d+)\)|(?:line|第)\s*(\d+)\s*(?:行)?/i);
  const lineNumber = Number(lineMatch?.[1] || lineMatch?.[3]);
  if (!Number.isFinite(lineNumber) || lineNumber < 1) return "";
  const lines = sourceCode.split(/\r?\n/);
  const start = Math.max(0, lineNumber - 3);
  const end = Math.min(lines.length, lineNumber + 2);
  return lines.slice(start, end).map((line, index) => `${start + index + 1}: ${line}`).join("\n");
}

const HELPERS_SIGNATURES = `helpers 可用签名：
- groupBy(rows, field)、countBy(rows, field)、unique(rows, field)
- sum(rows, field)、avg(rows, field)、min(rows, field)、max(rows, field)
- formatNumber(value, options)、formatDate(value, options)、daysBetween(start, end)
- resolveColumn(columns, idOrName)、getValue(row, columns, idOrName)`;

/**
 * 生成错误的可读化（Vue 版）。
 *
 * 相对 ai_excel React 版的变化：
 * - 移除 React.createElement/useState 相关提示；
 * - 增加 Vue h()/ref()/onMounted 等常见错误的修复建议；
 * - 提醒"使用 JSX"或"返回渲染函数"。
 */
export function humanizeGenerationError(
  error: unknown,
  sourceCode = "",
  columns: ColumnDefinition[] = [],
) {
  const raw = (error instanceof Error ? error.message : String(error)).trim() || "未知错误";
  const guidance: string[] = [];

  const missingMatch = raw.match(/未正确绑定描述中的字段：([^\n]+)/);
  if (missingMatch) {
    const names = missingMatch[1].split(/[、,，]/).map((name) => name.trim()).filter(Boolean);
    const mappings = names.map((name) => {
      const column = columns.find((item) => item.name === name);
      return column ? `${column.name} 必须直接引用列 ID ${column.id}` : `${name} 必须直接绑定当前 Schema`;
    });
    guidance.push(mappings.join("；"));
  }

  // Vue 相关错误
  if (/\bh\b is not defined|Vue is not defined|Vue\.\w+ is not a function|libs\.Vue is undefined/i.test(raw)) {
    guidance.push("Vue 已作为组件闭包的顶层变量注入。直接使用 Vue.h、Vue.ref、Vue.reactive、Vue.computed、Vue.onMounted 等；也可从 libs.Vue 解构 const { h, ref, onMounted } = libs.Vue。不要从 window、globalThis 或 props 获取 Vue。");
  }
  if (/(?:window|document|navigator|echarts) is not defined|Cannot read .* of (?:null|undefined)/i.test(raw)) {
    guidance.push("首屏同步渲染不能访问 window、document、navigator 或 echarts.init。把浏览器 API 放入 onMounted 或事件回调中，并为首屏提供不依赖浏览器对象的初值。");
  }
  if (/helpers\.[A-Za-z_$][\w$]* is not a function|helpers.*(?:undefined|not defined)/i.test(raw)) {
    guidance.push(HELPERS_SIGNATURES);
  }
  if (/缺少 function GeneratedView|未找到 GeneratedView 组件|渲染函数|render function|is not a function/i.test(raw)) {
    guidance.push("入口必须是 function GeneratedView(props)，函数体返回：一个渲染函数 () => <div>...</div>，或一个 defineComponent 配置对象。只输出一个完整的 GeneratedView，不要 Markdown、说明、文件路径或满意度询问。");
  }
  const snippet = codeSnippet(sourceCode, raw);
  if (snippet) guidance.push(`出错位置附近：\n${snippet}`);
  if (!guidance.length) {
    guidance.push("根据错误逐项修正上一版代码，保留原需求；不要绕过字段绑定、类型检查或运行时校验。");
  }

  return `上一版未通过确定性校验：${raw}\n修复要求：\n${guidance.map((item, index) => `${index + 1}. ${item}`).join("\n")}`;
}
