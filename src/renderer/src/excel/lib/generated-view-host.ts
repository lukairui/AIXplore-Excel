import * as Vue from "vue";
import { createApp, h, defineComponent, getCurrentInstance } from "vue";
import * as Babel from "@babel/standalone";
import * as echarts from "echarts";
import * as d3 from "d3";
import * as Motion from "motion-v";
import * as Lucide from "lucide-vue-next";
import * as dateFns from "date-fns";
import type { ColumnDefinition, JsonValue, RuntimeRow, ViewDefinition } from "./types";
import type { ViewSourceSheet } from "./data-sources";
import {
  createGeneratedDataRows,
  GENERATED_APP_RUNTIME_VERSION,
  generatedValue,
  generatedViewHelpers,
  normalizeViewState,
  resolveGeneratedColumn,
  validateGeneratedViewBindings,
} from "./generated-app-runtime";

/**
 * 生成式视图宿主（Vue 版）。
 *
 * 对应 ai_excel 的 src/generated-view-host.tsx（React 版）。
 *
 * 设计要点：
 * 1. 模型输出的代码字符串为 function GeneratedView(props)，内部使用
 *    Vue 渲染函数 + JSX（或 defineComponent）。
 * 2. 宿主用 @babel/standalone 将 JSX 编译为 Vue.h() 调用。
 * 3. 通过 new Function 注入 Vue 与 libs（echarts/d3/Motion/Lucide/dateFns）。
 * 4. 编译产物可能返回：
 *    - 一个渲染函数 () => VNode；
 *    - 一个 defineComponent 配置对象；
 *    - 一个已编译的 Vue 组件（如通过 defineComponent 包装）。
 *    宿主统一适配为可被 createApp/h(Component) 挂载的形式。
 */

const generatedViewLibs = {
  Vue,
  echarts,
  d3,
  Motion,
  Lucide,
  dateFns,
};

const MAX_GENERATED_VIEW_SOURCE_LENGTH = 120_000;
const blockedGeneratedViewPatterns: Array<[RegExp, string]> = [
  [/\b(?:while|for)\s*(?:await\s*)?\(|\bdo\s*\{/, "禁止使用可能阻塞页面的显式循环，请改用数据数组的 map/filter/reduce"],
  [/\b(?:setTimeout|setInterval|requestAnimationFrame|requestIdleCallback|queueMicrotask|nextTick)\s*\(/, "禁止使用可能持续占用主线程的调度任务"],
  [/\b(?:onUpdated|watchEffect|watchPostEffect|watchSyncEffect)\s*\(/, "禁止使用可能反复触发渲染的生命周期或监听器"],
];

function assertGeneratedViewSafety(sourceCode: string) {
  if (sourceCode.length > MAX_GENERATED_VIEW_SOURCE_LENGTH) {
    throw new Error("AI 应用代码过大，请缩小界面复杂度后重新生成");
  }
  const blocked = blockedGeneratedViewPatterns.find(([pattern]) => pattern.test(sourceCode));
  if (blocked) throw new Error(blocked[1]);
}

type RowUpdate = { rowId: string; values: Record<string, unknown> };

export type GeneratedViewHostBindings = {
  view: ViewDefinition;
  rows: RuntimeRow[];
  allRows?: RuntimeRow[];
  columns: ColumnDefinition[];
  sourceSheets?: ViewSourceSheet[];
  relations?: any[];
  updateCell: (rowId: string, columnId: string, value: unknown) => void;
  updateRow: (rowId: string, values: Record<string, unknown>) => void;
  batchUpdateRows: (updates: RowUpdate[]) => void;
  addRow: (values?: Record<string, unknown>) => string;
  deleteRow: (rowId: string) => void;
  deleteRows: (rowIds: string[]) => void;
  reorderRows: (rowIds: string[]) => void;
  onViewStateChange: (viewId: string, state: JsonValue, remount?: boolean) => void;
  onRuntimeError: (message: string | null) => void;
};

function createGeneratedSheet(sheet: ViewSourceSheet) {
  const rows = createGeneratedDataRows(sheet.rows, sheet.columns);
  const allRows = createGeneratedDataRows(sheet.allRows, sheet.columns);
  return {
    id: sheet.id,
    name: sheet.name,
    rows,
    allRows,
    columns: sheet.columns,
    byId: Object.fromEntries(rows.map((row) => [row.id, row])),
    resolveColumn: (idOrName: string) => resolveGeneratedColumn(sheet.columns, idOrName),
    value: (row: (typeof rows)[number], idOrName: string) => generatedValue(row, sheet.columns, idOrName),
  };
}

type CompiledEntry =
  | { kind: "factory"; factory: (props: any) => any }
  | { kind: "empty"; reason: string };

function componentFromFactory(factory: (props: any) => any, name: string) {
  return defineComponent({
    name,
    inheritAttrs: false,
    setup(_, context) {
      const result = factory(context.attrs);
      if (typeof result === "function") return result;
      if (result && typeof result === "object" && "__v_isVNode" in result) return () => result;
      if (result && typeof result === "object" && (typeof result.render === "function" || typeof result.setup === "function")) {
        return () => h(result);
      }
      throw new Error("GeneratedView 未返回渲染函数、VNode 或 Vue 组件");
    },
  });
}

function adaptFactoryResult(result: any): CompiledEntry {
  if (result == null) {
    return { kind: "empty", reason: "GeneratedView 未返回任何内容" };
  }
  if (typeof result === "function") {
    return { kind: "factory", factory: result };
  }
  return { kind: "empty", reason: "未找到 GeneratedView 函数" };
}

export function compileGeneratedView(sourceCode: string): CompiledEntry {
  assertGeneratedViewSafety(sourceCode);
  if (!/function\s+GeneratedView\s*\(/.test(sourceCode)) {
    throw new Error("模型输出缺少 function GeneratedView");
  }
  if (/(?:^|\n)\s*(?:import|export)\b|\bimport\s*\(|\brequire\s*\(/.test(sourceCode)) {
    throw new Error("AI 应用不能导入外部模块");
  }
  if (/\bclass(?:Name)?\s*=|(?:^|[{,])\s*class(?:Name)?\s*:/.test(sourceCode)) {
    throw new Error("运行时不提供 Tailwind 或外部 class 样式；请删除 class/className，并将全部排版写入 JSX style={{ ... }} 内联样式");
  }
  const transformed = Babel.transform(sourceCode, {
    presets: [["react", { runtime: "classic", pragma: "Vue.h", pragmaFrag: "Vue.Fragment" }]],
  }).code;
  if (!transformed) throw new Error("视图代码编译结果为空");
  const factory = new Function(
    "Vue",
    "libs",
    `"use strict";\nconst { h, defineComponent, ref, reactive, computed, watch, onMounted, onBeforeUnmount, toRefs, nextTick, shallowRef, shallowReactive } = Vue;\n${transformed}\nreturn typeof GeneratedView === "function" ? GeneratedView : (typeof generatedView !== "undefined" ? generatedView : null);`,
  ) as (vueApi: typeof Vue, libs: typeof generatedViewLibs) => any;
  const generated = factory(Vue, generatedViewLibs);
  if (typeof generated !== "function") {
    throw new Error("未找到 GeneratedView 组件");
  }
  return adaptFactoryResult(generated);
}

/**
 * 预检：在正式挂载前发现编译错误和字段绑定缺失。
 *
 * 注意：不再使用 renderToString 做 SSR 渲染。原因是 Vue SSR 对包含
 * onMounted 里 echarts.init 的组件会反复触发副作用或退化卡死，
 * 表现为"模型都返回了，preflight 还在卡"。
 *
 * 改为只做：
 * 1. validateGeneratedViewBindings（字段绑定静态校验，原版也做）
 * 2. compileGeneratedView（Babel 编译 + factory 注入 + 入口检查）
 *
 * 真正的运行时错误由 GeneratedViewHost 的错误边界捕获并显示。
 * 这是性能关键路径：3 轮自纠错会调用 preflight 3 次，SSR 会成倍放大开销。
 */
export function preflightGeneratedView(
  sourceCode: string,
  rows: RuntimeRow[],
  columns: ColumnDefinition[],
  view: ViewDefinition,
  _sourceSheets: ViewSourceSheet[] = [],
  _relations: any[] = [],
) {
  // 1. 字段绑定静态校验（保留原版行为）
  validateGeneratedViewBindings(sourceCode, columns, view.draftDescription || view.description);
  // 2. Babel 编译 + factory 执行 + 入口检查（捕获语法错误、缺 GeneratedView 等）
  const entry = compileGeneratedView(sourceCode);
  if (entry.kind === "empty") {
    throw new Error(entry.reason || "AI 应用编译失败");
  }
  return Promise.resolve("");
}

export type GeneratedViewHost = ReturnType<typeof createGeneratedViewHost>;

/**
 * 在指定 DOM 容器中挂载生成式视图，返回一个可销毁的句柄。
 * 错误通过 onRuntimeError 回调上抛，避免单条视图崩溃影响整个工作台。
 */
export function createGeneratedViewHost(container: HTMLElement, bindings: GeneratedViewHostBindings) {
  const {
    view,
    rows,
    allRows = rows,
    columns,
    sourceSheets = [],
    relations = [],
    updateCell,
    updateRow,
    batchUpdateRows,
    addRow,
    deleteRow,
    deleteRows,
    reorderRows,
    onViewStateChange,
    onRuntimeError,
  } = bindings;

  let app: Vue.App<Element> | null = null;
  let disposed = false;
  let lastCompiledCode: string | null = null;
  let lastCompiledEntry: CompiledEntry | null = null;

  function ensureCompiled() {
    if (!view.generatedCode) {
      lastCompiledCode = null;
      lastCompiledEntry = null;
      return null;
    }
    if (lastCompiledCode === view.generatedCode) return lastCompiledEntry;
    lastCompiledCode = view.generatedCode;
    try {
      validateGeneratedViewBindings(view.generatedCode, columns, view.draftDescription || view.description);
      lastCompiledEntry = compileGeneratedView(view.generatedCode);
    } catch (error) {
      lastCompiledEntry = { kind: "empty", reason: error instanceof Error ? error.message : String(error) };
      onRuntimeError(lastCompiledEntry.reason);
    }
    return lastCompiledEntry;
  }

  function mount() {
    if (disposed) return;
    const entry = ensureCompiled();

    // 没有代码 → 渲染占位
    if (!view.generatedCode) {
      app = createApp({
        render: () =>
          h("div", { class: "empty-view-stage" }, [
            h("div", { class: "empty-view-mark" }, "✦"),
            h("h2", null, "从一句描述开始"),
            h("p", null, "在右侧写下你想要的界面、布局和交互，再生成一个只属于这份数据的视图。"),
          ]),
      });
      app.mount(container);
      return;
    }

    if (!entry || entry.kind === "empty") {
      const reason = (entry as { reason?: string } | undefined | null)?.reason || "当前代码无法编译，请在右侧查看错误并重新生成。";
      onRuntimeError(reason);
      app = createApp({
        render: () =>
          h("div", { class: "view-compile-empty" }, [
            h("div", { class: "view-compile-empty__title" }, "视图无法编译"),
            h("div", { class: "view-compile-empty__desc" }, reason),
          ]),
      });
      app.mount(container);
      return;
    }

    onRuntimeError(null);

    const flatRows = createGeneratedDataRows(rows, columns);
    const flatAllRows = createGeneratedDataRows(allRows, columns);
    const sheets = sourceSheets.map(createGeneratedSheet);
    const data = {
      rows: flatRows,
      allRows: flatAllRows,
      columns,
      byId: Object.fromEntries(flatRows.map((row) => [row.id, row])),
      resolveColumn: (idOrName: string) => resolveGeneratedColumn(columns, idOrName),
      value: (row: (typeof flatRows)[number], idOrName: string) => generatedValue(row, columns, idOrName),
      sheets,
      sheet: (idOrName: string) => sheets.find((sheet) => sheet.id === idOrName || sheet.name === idOrName),
      relations,
    };
    const actions = {
      updateCell,
      updateRow,
      batchUpdateRows,
      addRow,
      deleteRow,
      deleteRows,
      reorderRows,
    };

    let stateRef = normalizeViewState(view.appState ?? {});
    let serializedState = JSON.stringify(stateRef);
    const setViewState = (next: JsonValue | ((current: JsonValue) => JsonValue)) => {
      try {
        const resolved = typeof next === "function" ? next(stateRef) : next;
        const normalized = normalizeViewState(resolved);
        const nextSerializedState = JSON.stringify(normalized);
        if (nextSerializedState === serializedState) return;
        stateRef = normalized;
        serializedState = nextSerializedState;
        onViewStateChange(view.id, normalized);
      } catch (error) {
        onRuntimeError(error instanceof Error ? error.message : String(error));
      }
    };
    const resetViewState = () => {
      stateRef = {};
      serializedState = "{}";
      onViewStateChange(view.id, {}, true);
    };
    const runtime = {
      version: GENERATED_APP_RUNTIME_VERSION,
      viewId: view.id,
      assets: view.referenceImages,
      referenceImages: view.referenceImages,
      state: view.appState ?? {},
      setState: setViewState,
      resetState: resetViewState,
    };

    const props = {
      rows: flatRows,
      allRows: flatAllRows,
      columns,
      sheets,
      assets: view.referenceImages,
      viewState: view.appState ?? {},
      setViewState,
      ...actions,
      data,
      actions,
      runtime,
      helpers: generatedViewHelpers,
      libs: generatedViewLibs,
    };

    // 错误边界：捕获生成组件渲染/提交阶段的异常
    const ErrorBoundary = defineComponent({
      name: "GeneratedViewErrorBoundary",
      setup(_, { slots, expose }) {
        const instance = getCurrentInstance();
        const error = Vue.ref<Error | null>(null);
        const onErr = (err: unknown) => {
          error.value = err instanceof Error ? err : new Error(String(err));
          onRuntimeError(error.value.message);
        };
        // Vue 3 的 errorCaptured 钩子：捕获子组件树抛出的任何错误
        const stop = Vue.onErrorCaptured((err) => {
          onErr(err);
          return false; // 阻止继续向上冒泡
        });
        expose({ reset: () => { error.value = null; }, stop });
        return () => {
          if (error.value) {
            return h("div", { class: "view-runtime-error" }, [
              h("div", { class: "error-icon" }, "!"),
              h("div", null, [
                h("strong", null, "视图运行失败"),
                h("p", null, error.value.message),
              ]),
            ]);
          }
          return slots.default?.();
        };
      },
    });

    const targetComponent = componentFromFactory(entry.factory, "GeneratedViewRenderer");

    app = createApp({
      name: "GeneratedViewRoot",
      render: () =>
        h(ErrorBoundary, null, {
          default: () => h(targetComponent, props),
        }),
    });
    app.config.errorHandler = (err) => {
      onErrSilent(err, onRuntimeError);
    };
    app.mount(container);
  }

  function remount() {
    if (disposed) return;
    if (app) {
      try { app.unmount(); } catch { /* noop */ }
      app = null;
    }
    // 清空容器，避免 unmount 残留
    container.innerHTML = "";
    mount();
  }

  function dispose() {
    disposed = true;
    if (app) {
      try { app.unmount(); } catch { /* noop */ }
      app = null;
    }
    container.innerHTML = "";
  }

  return { mount, remount, dispose };
}

function onErrSilent(err: unknown, onRuntimeError: (message: string | null) => void) {
  const message = err instanceof Error ? err.message : String(err);
  // 避免无限循环上报
  try { onRuntimeError(message); } catch { /* noop */ }
}
