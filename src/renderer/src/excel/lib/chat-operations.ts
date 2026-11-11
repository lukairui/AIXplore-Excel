import {
  cleanModelCode,
  calculateRuntimeRows,
  compileColumnCode,
  createGenerationState,
  createId,
  ensureProjectSheets,
  helpers,
} from "./core";
import type { ChatSourceFile } from "./chat-analysis";
import type { ColumnDataType, ColumnDefinition, Project, SheetDefinition, TableRow } from "./types";

type ConditionOperator = "eq" | "neq" | "contains" | "not-contains" | "gt" | "gte" | "lt" | "lte" | "between" | "empty" | "not-empty" | "in" | "regex";
type Aggregation = "count" | "distinct" | "sum" | "average" | "min" | "max";
type TableProgramColumn = { id: string; name: string; dataType: ColumnDataType };
type ReadonlyProgramTable = {
  id: string;
  name: string;
  columns: ReadonlyArray<TableProgramColumn>;
  rows: ReadonlyArray<Readonly<Record<string, unknown>>>;
};

export type ChatCondition = {
  columnId: string;
  operator: ConditionOperator;
  value?: unknown;
  secondValue?: unknown;
};

type OperationBase = { id: string; title: string; sheetId: string };

export type ChatDataOperation =
  | OperationBase & { type: "filter"; conditions: ChatCondition[]; match: "all" | "any"; keep: boolean }
  | OperationBase & { type: "sort"; columnId: string; direction: "asc" | "desc"; limit?: number }
  | OperationBase & { type: "select-columns"; columnIds: string[] }
  | OperationBase & { type: "delete-column"; columnId: string }
  | OperationBase & { type: "rename-column"; columnId: string; name: string }
  | OperationBase & { type: "transform-column"; columnId: string; code: string; dataType?: ColumnDataType }
  | OperationBase & { type: "computed-column"; columnId: string; name: string; code: string; dataType: ColumnDataType }
  | OperationBase & { type: "join"; relatedSheetId: string; leftKeyColumnId: string; rightKeyColumnId: string; includeColumnIds: string[]; joinType: "left" | "inner" }
  | OperationBase & { type: "append"; relatedSheetId: string }
  | OperationBase & { type: "split"; columnId: string; keepOriginal: boolean }
  | OperationBase & { type: "table-program"; code: string; outputColumns: TableProgramColumn[]; relatedSheetIds: string[] }
  | OperationBase & { type: "aggregate"; groupByColumnIds: string[]; metrics: Array<{ columnId?: string; aggregation: Aggregation; name: string }> };

export type ChatOperationResult = {
  project: Project;
  summaries: string[];
  errors: string[];
  completedOperationIds: string[];
  failedOperationIds: string[];
};

const OPERATORS = new Set<ConditionOperator>(["eq", "neq", "contains", "not-contains", "gt", "gte", "lt", "lte", "between", "empty", "not-empty", "in", "regex"]);
const AGGREGATIONS = new Set<Aggregation>(["count", "distinct", "sum", "average", "min", "max"]);
const DATA_TYPES = new Set<ColumnDataType>(["text", "number", "boolean", "date"]);
const OPERATION_TYPES = new Set(["filter", "sort", "select-columns", "delete-column", "rename-column", "transform-column", "computed-column", "join", "append", "split", "table-program", "aggregate"]);
const UNSAFE_CODE = /\b(?:window|document|globalThis|self|fetch|XMLHttpRequest|WebSocket|Worker|SharedWorker|navigator|location|localStorage|sessionStorage|indexedDB|eval|Function|require|process|constructor|__proto__|prototype|setTimeout|setInterval|importScripts)\b|\bimport\s*\(|\b(?:while|for)\s*\(/;
const SAFE_HELPERS = Object.freeze({ ...helpers });
const COLUMN_ID = /^[a-zA-Z0-9_-]{1,60}$/;
// ponytail: browser-side execution cap; move table programs to a worker if larger sheets become common.
const MAX_PROGRAM_ROWS = 250_000;
const MAX_PROGRAM_COLUMNS = 1_000;

function stringValue(value: unknown, fallback = "") {
  return typeof value === "string" && value.trim() ? value.trim() : fallback;
}

function stringArray(value: unknown) {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string" && Boolean(item.trim())).map((item) => item.trim()) : [];
}

function safeCode(value: unknown) {
  const code = cleanModelCode(stringValue(value));
  if (!code || code.length > 4_000 || UNSAFE_CODE.test(code)) return "";
  try {
    compileColumnCode(code);
    return code;
  } catch {
    return "";
  }
}

function compileTableProgram(code: string) {
  return new Function("table", "helpers", `"use strict";\n${code}`) as (
    table: ReadonlyProgramTable & { relatedSheets: ReadonlyArray<ReadonlyProgramTable> },
    runtimeHelpers: typeof helpers,
  ) => unknown;
}

function safeTableCode(value: unknown) {
  const code = cleanModelCode(stringValue(value));
  if (!code || code.length > 12_000 || UNSAFE_CODE.test(code)) return "";
  try {
    compileTableProgram(code);
    return code;
  } catch {
    return "";
  }
}

function declaredProgramColumns(value: unknown) {
  if (!Array.isArray(value)) return [];
  const used = new Set<string>();
  return value.flatMap((item): TableProgramColumn[] => {
    if (!item || typeof item !== "object") return [];
    const raw = item as Record<string, unknown>;
    const id = stringValue(raw.id);
    const name = stringValue(raw.name).slice(0, 100);
    const dataType = stringValue(raw.dataType, "text") as ColumnDataType;
    if (!COLUMN_ID.test(id) || used.has(id) || !name || !DATA_TYPES.has(dataType)) return [];
    used.add(id);
    return [{ id, name, dataType }];
  }).slice(0, MAX_PROGRAM_COLUMNS);
}

function sourceSheets(sources: ChatSourceFile[]) {
  return sources.flatMap((source) => ensureProjectSheets(source.project));
}

export function normalizeChatOperations(value: unknown, sources: ChatSourceFile[]): ChatDataOperation[] {
  if (!Array.isArray(value)) return [];
  const sheets = sourceSheets(sources);
  const bySheet = new Map(sheets.map((sheet) => [sheet.id, sheet]));
  const columnsBySheet = new Map(sheets.map((sheet) => [sheet.id, new Set(sheet.columns.map((column) => column.id))]));
  const hasColumn = (sheetId: string, columnId: string) => Boolean(columnsBySheet.get(sheetId)?.has(columnId));
  return value.flatMap((item, index): ChatDataOperation[] => {
    if (!item || typeof item !== "object") return [];
    const raw = item as Record<string, unknown>;
    const type = stringValue(raw.type);
    const sheetId = stringValue(raw.sheetId);
    if (!OPERATION_TYPES.has(type) || !bySheet.has(sheetId)) return [];
    const base = { id: stringValue(raw.id, `operation-${index + 1}`).slice(0, 60), title: stringValue(raw.title, `数据处理 ${index + 1}`).slice(0, 100), sheetId };
    if (type === "filter") {
      const conditions = Array.isArray(raw.conditions) ? raw.conditions.flatMap((condition): ChatCondition[] => {
        if (!condition || typeof condition !== "object") return [];
        const candidate = condition as Record<string, unknown>;
        const columnId = stringValue(candidate.columnId);
        const operator = stringValue(candidate.operator) as ConditionOperator;
        return hasColumn(sheetId, columnId) && OPERATORS.has(operator) ? [{ columnId, operator, value: candidate.value, secondValue: candidate.secondValue }] : [];
      }) : [];
      return conditions.length ? [{ ...base, type, conditions, match: raw.match === "any" ? "any" : "all", keep: raw.keep !== false }] : [];
    }
    if (type === "sort") {
      const columnId = stringValue(raw.columnId);
      const limit = Number(raw.limit);
      return hasColumn(sheetId, columnId) ? [{ ...base, type, columnId, direction: raw.direction === "desc" ? "desc" : "asc", limit: Number.isInteger(limit) && limit > 0 ? Math.min(10_000, limit) : undefined }] : [];
    }
    if (type === "select-columns") {
      const columnIds = stringArray(raw.columnIds).filter((id) => hasColumn(sheetId, id));
      if (!columnIds.length) return [];
      columnsBySheet.set(sheetId, new Set(columnIds));
      return [{ ...base, type, columnIds }];
    }
    if (type === "delete-column") {
      const columnId = stringValue(raw.columnId);
      const columns = columnsBySheet.get(sheetId);
      if (!columns?.has(columnId) || columns.size <= 1) return [];
      columns.delete(columnId);
      return [{ ...base, type, columnId }];
    }
    if (type === "rename-column") {
      const columnId = stringValue(raw.columnId);
      const name = stringValue(raw.name).slice(0, 100);
      return hasColumn(sheetId, columnId) && name ? [{ ...base, type, columnId, name }] : [];
    }
    if (type === "transform-column") {
      const columnId = stringValue(raw.columnId);
      const code = safeCode(raw.code);
      const dataType = stringValue(raw.dataType) as ColumnDataType;
      return hasColumn(sheetId, columnId) && code ? [{ ...base, type, columnId, code, dataType: DATA_TYPES.has(dataType) ? dataType : undefined }] : [];
    }
    if (type === "computed-column") {
      const columnId = stringValue(raw.columnId, createId("col")).replace(/[^a-zA-Z0-9_-]/g, "-").slice(0, 60);
      const name = stringValue(raw.name).slice(0, 100);
      const code = safeCode(raw.code);
      const dataType = stringValue(raw.dataType, "text") as ColumnDataType;
      if (!name || !code || !DATA_TYPES.has(dataType) || hasColumn(sheetId, columnId)) return [];
      columnsBySheet.get(sheetId)?.add(columnId);
      return [{ ...base, type, columnId, name, code, dataType }];
    }
    if (type === "join") {
      const relatedSheetId = stringValue(raw.relatedSheetId);
      const leftKeyColumnId = stringValue(raw.leftKeyColumnId);
      const rightKeyColumnId = stringValue(raw.rightKeyColumnId);
      const includeColumnIds = stringArray(raw.includeColumnIds).filter((id) => hasColumn(relatedSheetId, id) && id !== rightKeyColumnId);
      return bySheet.has(relatedSheetId) && hasColumn(sheetId, leftKeyColumnId) && hasColumn(relatedSheetId, rightKeyColumnId) && includeColumnIds.length
        ? [{ ...base, type, relatedSheetId, leftKeyColumnId, rightKeyColumnId, includeColumnIds, joinType: raw.joinType === "inner" ? "inner" : "left" }]
        : [];
    }
    if (type === "append") {
      const relatedSheetId = stringValue(raw.relatedSheetId);
      return bySheet.has(relatedSheetId) && relatedSheetId !== sheetId ? [{ ...base, type, relatedSheetId }] : [];
    }
    if (type === "split") {
      const columnId = stringValue(raw.columnId);
      return hasColumn(sheetId, columnId) ? [{ ...base, type, columnId, keepOriginal: raw.keepOriginal !== false }] : [];
    }
    if (type === "table-program") {
      const code = safeTableCode(raw.code);
      const outputColumns = declaredProgramColumns(raw.outputColumns);
      const relatedSheetIds = stringArray(raw.relatedSheetIds).filter((id) => id !== sheetId && bySheet.has(id));
      if (Array.isArray(raw.outputColumns) && raw.outputColumns.length && outputColumns.length !== raw.outputColumns.length) return [];
      if (!code) return [];
      if (outputColumns.length) columnsBySheet.set(sheetId, new Set(outputColumns.map((column) => column.id)));
      return [{ ...base, type, code, outputColumns, relatedSheetIds }];
    }
    const groupByColumnIds = stringArray(raw.groupByColumnIds).filter((id) => hasColumn(sheetId, id));
    const metrics = Array.isArray(raw.metrics) ? raw.metrics.flatMap((metric): Array<{ columnId?: string; aggregation: Aggregation; name: string }> => {
      if (!metric || typeof metric !== "object") return [];
      const candidate = metric as Record<string, unknown>;
      const aggregation = stringValue(candidate.aggregation) as Aggregation;
      const columnId = stringValue(candidate.columnId);
      const name = stringValue(candidate.name).slice(0, 100);
      if (!AGGREGATIONS.has(aggregation) || !name || (aggregation !== "count" && !hasColumn(sheetId, columnId))) return [];
      return [{ columnId: columnId || undefined, aggregation, name }];
    }) : [];
    return metrics.length ? [{ ...base, type: "aggregate", groupByColumnIds, metrics }] : [];
  }).slice(0, 20);
}

function normalizedCellValue(value: unknown): unknown {
  if (value == null || typeof value === "string" || typeof value === "number" || typeof value === "boolean" || value instanceof Date) return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function normalizeTableProgramResult(output: unknown, sheet: SheetDefinition, declaredColumns: TableProgramColumn[]) {
  const result = Array.isArray(output) ? { rows: output } : output;
  if (!result || typeof result !== "object" || !Array.isArray((result as Record<string, unknown>).rows)) {
    throw new Error("通用表格程序必须返回 rows 数组");
  }
  const rawResult = result as Record<string, unknown>;
  const rawColumns = Array.isArray(rawResult.columns) ? rawResult.columns : declaredColumns.length ? declaredColumns : sheet.columns;
  const existing = new Map(sheet.columns.map((column) => [column.id, column]));
  const used = new Set<string>();
  const columns = rawColumns.map((item, index): ColumnDefinition => {
    if (!item || typeof item !== "object") throw new Error(`第 ${index + 1} 个字段定义无效`);
    const raw = item as Record<string, unknown>;
    const id = stringValue(raw.id);
    const name = stringValue(raw.name).slice(0, 100);
    const dataType = stringValue(raw.dataType, existing.get(id)?.dataType || "text") as ColumnDataType;
    if (!COLUMN_ID.test(id) || used.has(id) || !name || !DATA_TYPES.has(dataType)) throw new Error(`第 ${index + 1} 个字段定义无效`);
    used.add(id);
    return existing.has(id) ? { ...existing.get(id)!, name, dataType } : { id, name, dataType, kind: "manual" };
  });
  if (!columns.length || columns.length > MAX_PROGRAM_COLUMNS) throw new Error("通用表格程序返回的字段数量无效");
  declaredColumns.forEach((column) => {
    if (!used.has(column.id)) throw new Error(`通用表格程序缺少声明字段 ${column.name}`);
  });
  const rawRows = rawResult.rows as unknown[];
  if (rawRows.length > MAX_PROGRAM_ROWS) throw new Error(`通用表格程序返回行数超过 ${MAX_PROGRAM_ROWS} 行限制`);
  const rows = rawRows.map((item, index): TableRow => {
    if (!item || typeof item !== "object" || Array.isArray(item)) throw new Error(`第 ${index + 1} 行不是有效对象`);
    const raw = item as Record<string, unknown>;
    const values = raw.values && typeof raw.values === "object" && !Array.isArray(raw.values) ? raw.values as Record<string, unknown> : raw;
    return {
      id: typeof raw.id === "string" && raw.id ? raw.id : createId("row"),
      values: Object.fromEntries(columns.map((column) => [column.id, normalizedCellValue(values[column.id]) ?? null])),
    };
  });
  return { columns, rows };
}

function normalized(value: unknown) {
  return String(value ?? "").trim().toLocaleLowerCase();
}

function numberValue(value: unknown) {
  const number = typeof value === "number" ? value : Number(String(value ?? "").replaceAll(",", ""));
  return Number.isFinite(number) ? number : null;
}

function matchesCondition(row: TableRow, condition: ChatCondition) {
  const value = row.values[condition.columnId];
  const left = normalized(value);
  const right = normalized(condition.value);
  if (condition.operator === "empty") return value == null || left === "";
  if (condition.operator === "not-empty") return value != null && left !== "";
  if (condition.operator === "eq") return left === right;
  if (condition.operator === "neq") return left !== right;
  if (condition.operator === "contains") return left.includes(right);
  if (condition.operator === "not-contains") return !left.includes(right);
  if (condition.operator === "in") return Array.isArray(condition.value) && condition.value.some((item) => normalized(item) === left);
  if (condition.operator === "regex") {
    try { return new RegExp(String(condition.value ?? ""), "i").test(String(value ?? "")); } catch { return false; }
  }
  const numeric = numberValue(value);
  const target = numberValue(condition.value);
  if (numeric == null || target == null) return false;
  if (condition.operator === "gt") return numeric > target;
  if (condition.operator === "gte") return numeric >= target;
  if (condition.operator === "lt") return numeric < target;
  if (condition.operator === "lte") return numeric <= target;
  const second = numberValue(condition.secondValue);
  return second != null && numeric >= Math.min(target, second) && numeric <= Math.max(target, second);
}

function aggregate(values: unknown[], aggregation: Aggregation, rowCount: number) {
  if (aggregation === "count") return rowCount;
  if (aggregation === "distinct") return new Set(values.map(normalized).filter(Boolean)).size;
  const numbers = values.map(numberValue).filter((value): value is number => value != null);
  if (!numbers.length) return null;
  if (aggregation === "sum") return helpers.sum(numbers);
  if (aggregation === "average") return helpers.avg(numbers);
  if (aggregation === "min") return Math.min(...numbers);
  return Math.max(...numbers);
}

function uniqueColumnName(columns: ColumnDefinition[], requested: string) {
  const names = new Set(columns.map((column) => column.name));
  if (!names.has(requested)) return requested;
  let index = 2;
  while (names.has(`${requested} (${index})`)) index += 1;
  return `${requested} (${index})`;
}

function resultProject(sources: ChatSourceFile[]) {
  const sheets = sources.flatMap((source) => ensureProjectSheets(JSON.parse(JSON.stringify(source.project)) as Project).map((sheet) => ({ ...sheet, name: sources.length > 1 ? `${source.name} - ${sheet.name}` : sheet.name })));
  const first = sheets[0];
  const now = Date.now();
  return {
    id: createId("project"),
    name: "AI Excel 处理结果",
    columns: first.columns,
    rows: first.rows,
    views: [],
    tableRule: first.tableRule,
    sheets,
    activeSheetId: first.id,
    activeTabId: "table",
    createdAt: now,
    updatedAt: now,
  } satisfies Project;
}

export function executeChatOperations(sources: ChatSourceFile[], operations: ChatDataOperation[]): ChatOperationResult {
  const project = resultProject(sources);
  const summaries: string[] = [];
  const errors: string[] = [];
  const completedOperationIds: string[] = [];
  const failedOperationIds: string[] = [];
  let activeSheetId = project.activeSheetId;
  const sheetById = () => new Map((project.sheets || []).map((sheet) => [sheet.id, sheet]));
  operations.forEach((operation) => {
    try {
      const sheets = sheetById();
      const sheet = sheets.get(operation.sheetId);
      if (!sheet) throw new Error("工作表不存在");
      activeSheetId = sheet.id;
      const before = sheet.rows.length;
      if (operation.type === "filter") {
        sheet.rows = sheet.rows.filter((row) => {
          const matched = operation.match === "any" ? operation.conditions.some((condition) => matchesCondition(row, condition)) : operation.conditions.every((condition) => matchesCondition(row, condition));
          return operation.keep ? matched : !matched;
        });
        summaries.push(`${operation.title}：${before.toLocaleString("zh-CN")} 行处理为 ${sheet.rows.length.toLocaleString("zh-CN")} 行`);
      } else if (operation.type === "sort") {
        sheet.rows.sort((left, right) => {
          const leftNumber = numberValue(left.values[operation.columnId]);
          const rightNumber = numberValue(right.values[operation.columnId]);
          const compared = leftNumber != null && rightNumber != null ? leftNumber - rightNumber : String(left.values[operation.columnId] ?? "").localeCompare(String(right.values[operation.columnId] ?? ""), "zh-CN");
          return operation.direction === "desc" ? -compared : compared;
        });
        if (operation.limit) sheet.rows = sheet.rows.slice(0, operation.limit);
        summaries.push(`${operation.title}：已按指定字段${operation.direction === "desc" ? "降序" : "升序"}排列${operation.limit ? `并保留前 ${operation.limit} 行` : ` ${sheet.rows.length.toLocaleString("zh-CN")} 行`}`);
      } else if (operation.type === "select-columns") {
        const selected = new Set(operation.columnIds);
        sheet.columns = sheet.columns.filter((column) => selected.has(column.id));
        sheet.rows = sheet.rows.map((row) => ({ ...row, values: Object.fromEntries(sheet.columns.map((column) => [column.id, row.values[column.id]])) }));
        summaries.push(`${operation.title}：保留 ${sheet.columns.length} 个字段`);
      } else if (operation.type === "delete-column") {
        const column = sheet.columns.find((item) => item.id === operation.columnId)!;
        sheet.columns = sheet.columns.filter((item) => item.id !== operation.columnId);
        sheet.rows.forEach((row) => { delete row.values[operation.columnId]; });
        summaries.push(`${operation.title}：已删除 ${column.name}`);
      } else if (operation.type === "rename-column") {
        const column = sheet.columns.find((item) => item.id === operation.columnId)!;
        const previous = column.name;
        column.name = uniqueColumnName(sheet.columns.filter((item) => item.id !== column.id), operation.name);
        summaries.push(`${operation.title}：${previous} 改为 ${column.name}`);
      } else if (operation.type === "transform-column") {
        const compute = compileColumnCode(operation.code);
        const runtimeRows = calculateRuntimeRows(sheet.rows, sheet.columns);
        const allRows = Object.freeze(runtimeRows.map((row) => Object.freeze({ ...row.values })));
        const values = runtimeRows.map((row) => compute(Object.freeze({ ...row.values }), allRows as unknown as Record<string, unknown>[], SAFE_HELPERS));
        sheet.rows.forEach((row, index) => { row.values[operation.columnId] = values[index]; });
        const column = sheet.columns.find((item) => item.id === operation.columnId)!;
        column.kind = "manual";
        column.generation = undefined;
        if (operation.dataType) column.dataType = operation.dataType;
        summaries.push(`${operation.title}：已更新 ${sheet.rows.length.toLocaleString("zh-CN")} 行 ${column.name}`);
      } else if (operation.type === "computed-column") {
        const columnId = sheet.columns.some((column) => column.id === operation.columnId) ? createId("col") : operation.columnId;
        const column: ColumnDefinition = { id: columnId, name: uniqueColumnName(sheet.columns, operation.name), kind: "manual", dataType: operation.dataType };
        const compute = compileColumnCode(operation.code);
        const runtimeRows = calculateRuntimeRows(sheet.rows, sheet.columns);
        const allRows = Object.freeze(runtimeRows.map((row) => Object.freeze({ ...row.values })));
        const values = runtimeRows.map((row) => compute(Object.freeze({ ...row.values }), allRows as unknown as Record<string, unknown>[], SAFE_HELPERS));
        sheet.columns.push(column);
        sheet.rows.forEach((row, index) => { row.values[column.id] = values[index]; });
        summaries.push(`${operation.title}：新增 ${column.name}，计算 ${sheet.rows.length.toLocaleString("zh-CN")} 行`);
      } else if (operation.type === "join") {
        const right = sheets.get(operation.relatedSheetId)!;
        const rightByKey = new Map<string, TableRow>();
        right.rows.forEach((row) => { const key = normalized(row.values[operation.rightKeyColumnId]); if (key && !rightByKey.has(key)) rightByKey.set(key, row); });
        const mappings = operation.includeColumnIds.map((rightId) => {
          const source = right.columns.find((column) => column.id === rightId)!;
          const target = { ...source, id: createId("col"), name: uniqueColumnName(sheet.columns, source.name), kind: "manual" as const, generation: undefined };
          sheet.columns.push(target);
          return { source, target };
        });
        sheet.rows = sheet.rows.flatMap((row) => {
          const match = rightByKey.get(normalized(row.values[operation.leftKeyColumnId]));
          if (!match && operation.joinType === "inner") return [];
          return [{ ...row, values: { ...row.values, ...Object.fromEntries(mappings.map(({ source, target }) => [target.id, match?.values[source.id] ?? null])) } }];
        });
        summaries.push(`${operation.title}：补充 ${mappings.length} 个字段，匹配 ${sheet.rows.filter((row) => mappings.some(({ target }) => row.values[target.id] != null)).length.toLocaleString("zh-CN")} 行`);
      } else if (operation.type === "append") {
        const right = sheets.get(operation.relatedSheetId)!;
        const rightByName = new Map(right.columns.map((column) => [column.name.trim().toLocaleLowerCase(), column]));
        const mappings = sheet.columns.map((column) => [column.id, rightByName.get(column.name.trim().toLocaleLowerCase())?.id] as const);
        const appended = right.rows.map((row) => ({ id: createId("row"), values: Object.fromEntries(mappings.map(([leftId, rightId]) => [leftId, rightId ? row.values[rightId] : null])) }));
        sheet.rows.push(...appended);
        summaries.push(`${operation.title}：纵向合并 ${appended.length.toLocaleString("zh-CN")} 行，共 ${sheet.rows.length.toLocaleString("zh-CN")} 行`);
      } else if (operation.type === "split") {
        const column = sheet.columns.find((item) => item.id === operation.columnId)!;
        const groups = new Map<string, TableRow[]>();
        sheet.rows.forEach((row) => { const name = String(row.values[operation.columnId] ?? "未填写").trim() || "未填写"; const rows = groups.get(name) || []; rows.push(structuredClone(row)); groups.set(name, rows); });
        const created = [...groups.entries()].slice(0, 100).map(([name, rows]) => ({ ...structuredClone(sheet), id: createId("sheet"), name: `${column.name}-${name}`.slice(0, 31), rows }));
        if (!created.length) throw new Error("没有可拆分的数据行");
        project.sheets = [...(project.sheets || []).filter((item) => operation.keepOriginal || item.id !== sheet.id), ...created];
        activeSheetId = created[0]?.id || activeSheetId;
        summaries.push(`${operation.title}：按 ${column.name} 拆分为 ${created.length} 张工作表`);
      } else if (operation.type === "table-program") {
        const table = Object.freeze({
          id: sheet.id,
          name: sheet.name,
          columns: Object.freeze(sheet.columns.map((column) => Object.freeze({ id: column.id, name: column.name, dataType: column.dataType }))),
          rows: Object.freeze(sheet.rows.map((row) => Object.freeze(structuredClone(row.values)))),
          relatedSheets: Object.freeze(operation.relatedSheetIds.flatMap((sheetId): ReadonlyProgramTable[] => {
            const related = sheets.get(sheetId);
            return related ? [Object.freeze({
              id: related.id,
              name: related.name,
              columns: Object.freeze(related.columns.map((column) => Object.freeze({ id: column.id, name: column.name, dataType: column.dataType }))),
              rows: Object.freeze(related.rows.map((row) => Object.freeze(structuredClone(row.values)))),
            })] : [];
          })),
        });
        const output = compileTableProgram(operation.code)(table, SAFE_HELPERS);
        if (output && typeof (output as PromiseLike<unknown>).then === "function") throw new Error("通用表格程序不能异步执行");
        const normalized = normalizeTableProgramResult(output, sheet, operation.outputColumns);
        sheet.columns = normalized.columns;
        sheet.rows = normalized.rows;
        summaries.push(`${operation.title}：通用表格程序将 ${before.toLocaleString("zh-CN")} 行处理为 ${sheet.rows.length.toLocaleString("zh-CN")} 行，保留 ${sheet.columns.length} 个字段`);
      } else if (operation.type === "aggregate") {
        const groups = new Map<string, TableRow[]>();
        sheet.rows.forEach((row) => { const key = JSON.stringify(operation.groupByColumnIds.map((id) => row.values[id] ?? null)); const rows = groups.get(key) || []; rows.push(row); groups.set(key, rows); });
        if (!operation.groupByColumnIds.length) groups.set("[]", sheet.rows);
        const groupColumns = operation.groupByColumnIds.map((id) => sheet.columns.find((column) => column.id === id)!).filter(Boolean).map((column) => ({ ...column, kind: "manual" as const, generation: undefined }));
        const metricColumns = operation.metrics.map((metric) => ({ id: createId("col"), name: metric.name, kind: "manual" as const, dataType: "number" as const }));
        const rows = [...groups.entries()].map(([key, members]) => {
          const groupValues = JSON.parse(key) as unknown[];
          return {
            id: createId("row"),
            values: {
              ...Object.fromEntries(groupColumns.map((column, index) => [column.id, groupValues[index]])),
              ...Object.fromEntries(operation.metrics.map((metric, index) => [metricColumns[index].id, aggregate(members.map((row) => metric.columnId ? row.values[metric.columnId] : null), metric.aggregation, members.length)])),
            },
          };
        });
        const resultSheet: SheetDefinition = { id: createId("sheet"), name: operation.title.slice(0, 31), columns: [...groupColumns, ...metricColumns], rows, tableRule: { id: "table-rule", ...createGenerationState() } };
        project.sheets = [...(project.sheets || []), resultSheet];
        activeSheetId = resultSheet.id;
        summaries.push(`${operation.title}：生成 ${rows.length.toLocaleString("zh-CN")} 行汇总结果`);
      }
      completedOperationIds.push(operation.id);
    } catch (error) {
      failedOperationIds.push(operation.id);
      errors.push(`${operation.title}：${error instanceof Error ? error.message : String(error)}`);
    }
  });
  const active = (project.sheets || []).find((sheet) => sheet.id === activeSheetId) || (project.sheets || [])[0];
  project.activeSheetId = active.id;
  project.columns = active.columns;
  project.rows = active.rows;
  project.tableRule = active.tableRule;
  project.updatedAt = Date.now();
  return { project, summaries, errors, completedOperationIds, failedOperationIds };
}
