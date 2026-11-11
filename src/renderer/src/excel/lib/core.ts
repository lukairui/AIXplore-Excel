import type {
  ColumnDataType,
  ColumnDefinition,
  GenerationState,
  PersistedPayload,
  Project,
  RuntimeRow,
  SheetDefinition,
  TableRow,
} from "./types";
import { GENERATED_APP_RUNTIME_VERSION, normalizeViewState } from "./generated-app-runtime";

export const PROJECT_STORAGE_KEY = "nl-sheet.project.v1";
export const MODEL_STORAGE_KEY = "nl-sheet.model-settings.v1";

export const createId = (prefix: string) =>
  `${prefix}_${Math.random().toString(36).slice(2, 9)}`;

export function ensureProjectSheets(project: Project): SheetDefinition[] {
  if (!project.sheets?.length) {
    const id = project.activeSheetId || createId("sheet");
    project.sheets = [{
      id,
      name: "数据",
      columns: project.columns,
      rows: project.rows,
      tableRule: project.tableRule,
    }];
    project.activeSheetId = id;
  }
  if (!project.activeSheetId || !project.sheets.some((sheet) => sheet.id === project.activeSheetId)) {
    project.activeSheetId = project.sheets[0].id;
  }
  return project.sheets;
}

export function hydrateProjectActiveSheet(project: Project) {
  const sheets = ensureProjectSheets(project);
  const active = sheets.find((sheet) => sheet.id === project.activeSheetId) || sheets[0];
  project.activeSheetId = active.id;
  project.columns = active.columns;
  project.rows = active.rows;
  project.tableRule = active.tableRule;
  return active;
}

export function syncProjectActiveSheet(project: Project) {
  const sheets = ensureProjectSheets(project);
  const active = sheets.find((sheet) => sheet.id === project.activeSheetId) || sheets[0];
  project.activeSheetId = active.id;
  active.columns = project.columns;
  active.rows = project.rows;
  active.tableRule = project.tableRule;
  return active;
}

export function activateProjectSheet(project: Project, sheetId: string) {
  syncProjectActiveSheet(project);
  const target = ensureProjectSheets(project).find((sheet) => sheet.id === sheetId);
  if (!target) return false;
  project.activeSheetId = target.id;
  project.columns = target.columns;
  project.rows = target.rows;
  project.tableRule = target.tableRule;
  return true;
}

export function createGenerationState(
  description = "",
  generatedCode = "",
): GenerationState {
  return {
    description,
    draftDescription: description,
    generatedCode,
    lastAttemptCode: generatedCode,
    attempts: [],
    status: generatedCode ? "ready" : "idle",
    error: null,
    sourceVersion: description ? 1 : 0,
    generatedVersion: generatedCode ? 1 : 0,
  };
}

export function createDefaultProject(): Project {
  const now = Date.now();
  const columns: ColumnDefinition[] = [
    { id: "col_task", name: "任务", kind: "manual", dataType: "text" },
    { id: "col_owner", name: "负责人", kind: "manual", dataType: "text" },
    { id: "col_status", name: "状态", kind: "manual", dataType: "text" },
    { id: "col_price", name: "单价", kind: "manual", dataType: "number" },
    { id: "col_quantity", name: "数量", kind: "manual", dataType: "number" },
    { id: "col_due", name: "截止日期", kind: "manual", dataType: "date" },
    {
      id: "col_amount",
      name: "金额",
      kind: "computed",
      dataType: "number",
      generation: createGenerationState(
        "单价乘以数量；任一为空时返回空。",
        'const price = row["col_price"];\nconst quantity = row["col_quantity"];\nif (price == null || quantity == null) return null;\nreturn helpers.round(Number(price) * Number(quantity), 2);',
      ),
    },
  ];

  const source: Array<[string, string, string, number, number, string]> = [
    ["门店陈列更新", "林晓", "进行中", 128, 24, "2026-07-24"],
    ["夏季物料补货", "周然", "待处理", 86.5, 42, "2026-07-27"],
    ["重点客户回访", "陈墨", "已完成", 320, 12, "2026-07-18"],
    ["渠道数据复核", "林晓", "进行中", 56, 65, "2026-07-22"],
    ["新品试销追踪", "周然", "进行中", 218, 18, "2026-07-30"],
    ["月度报告归档", "陈墨", "待处理", 45, 30, "2026-08-02"],
  ];
  const rows = source.map(([task, owner, status, price, quantity, due]) => ({
    id: createId("row"),
    values: {
      col_task: task,
      col_owner: owner,
      col_status: status,
      col_price: price,
      col_quantity: quantity,
      col_due: due,
    },
  }));

  const tableRule = { id: "table-rule" as const, ...createGenerationState() };
  const sheet: SheetDefinition = { id: createId("sheet"), name: "数据", columns, rows, tableRule };
  return {
    id: createId("project"),
    name: "零售运营工作台",
    columns,
    rows,
    views: [],
    tableRule,
    sheets: [sheet],
    activeSheetId: sheet.id,
    activeTabId: "table",
    createdAt: now,
    updatedAt: now,
  };
}

export const helpers = {
  round(value: number, digits = 2) {
    const base = 10 ** digits;
    return Math.round(value * base) / base;
  },
  sum(values: unknown[]) {
    return values.reduce<number>((total, value) => total + Number(value || 0), 0);
  },
  avg(values: unknown[]) {
    if (!values.length) return null;
    return values.reduce<number>((sum, value) => sum + Number(value || 0), 0) / values.length;
  },
  formatDate(value: unknown) {
    if (!value) return "";
    return new Date(String(value)).toLocaleDateString("zh-CN");
  },
};

export function cleanModelCode(output: string) {
  return output
    .trim()
    .replace(/^```(?:javascript|js|jsx|tsx)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();
}

export function compileColumnCode(code: string) {
  return new Function("row", "rows", "helpers", `"use strict";\n${code}`) as (
    row: Record<string, unknown>,
    rows: Record<string, unknown>[],
    runtimeHelpers: typeof helpers,
  ) => unknown;
}

export function calculateRuntimeRows(
  sourceRows: TableRow[],
  columns: ColumnDefinition[],
): RuntimeRow[] {
  const compiled = new Map<string, ReturnType<typeof compileColumnCode>>();
  for (const column of columns) {
    if (column.kind === "computed" && column.generation?.generatedCode) {
      try {
        compiled.set(column.id, compileColumnCode(column.generation.generatedCode));
      } catch {
        // A failed persisted build remains empty until the user regenerates it.
      }
    }
  }

  const runtimeRows: RuntimeRow[] = sourceRows.map((row) => ({
    id: row.id,
    values: { ...row.values },
    errors: {},
  }));

  for (const runtimeRow of runtimeRows) {
    for (const column of columns) {
      if (column.kind !== "computed") continue;
      const compute = compiled.get(column.id);
      if (!compute) {
        runtimeRow.values[column.id] = null;
        continue;
      }
      try {
        runtimeRow.values[column.id] = compute(
          runtimeRow.values,
          runtimeRows.map((item) => item.values),
          helpers,
        );
      } catch (error) {
        runtimeRow.values[column.id] = null;
        runtimeRow.errors[column.id] = error instanceof Error ? error.message : String(error);
      }
    }
  }
  return runtimeRows;
}

const SUMMARY_LABEL = /^(?:总计|合计|小计|汇总|总数|总额|总和|subtotal|grand\s*total|total)$/i;
const EMPTY_PLACEHOLDER = new Set(["", "-", "—", "--", "n/a", "na", "null", "undefined"]);

function normalizedCell(value: unknown) {
  return String(value ?? "").trim().replace(/[：:]+$/, "").toLocaleLowerCase();
}

/**
 * Summary rows are useful in the raw sheet but would double-count charts and dashboards.
 * Keep them available to the table while excluding only unambiguous totals/placeholders
 * from analysis views.
 */
export function isAnalysisSummaryRow(row: RuntimeRow, columns: ColumnDefinition[]) {
  const descriptiveColumns = columns.filter((column) => column.dataType !== "number");
  if (!descriptiveColumns.length) return false;
  const descriptiveValues = descriptiveColumns.map((column) => normalizedCell(row.values[column.id]));
  if (descriptiveValues.some((value) => SUMMARY_LABEL.test(value))) return true;
  const allPlaceholders = descriptiveValues.every((value) => EMPTY_PLACEHOLDER.has(value));
  const hasNumericValue = columns
    .filter((column) => column.dataType === "number")
    .some((column) => {
      const value = row.values[column.id];
      return value != null && value !== "" && Number.isFinite(Number(value)) && Number(value) !== 0;
    });
  return allPlaceholders && !hasNumericValue;
}

export function filterAnalysisRows(rows: RuntimeRow[], columns: ColumnDefinition[]) {
  return rows.filter((row) => !isAnalysisSummaryRow(row, columns));
}

export function applyTableRule(
  rows: RuntimeRow[],
  columns: ColumnDefinition[],
  code?: string,
) {
  if (!code?.trim()) return rows;
  const flatRows = rows.map((row) => ({ id: row.id, ...row.values }));
  const fn = new Function(
    "rows",
    "columns",
    "helpers",
    `"use strict";\n${code}`,
  ) as (
    rows: Record<string, unknown>[],
    columns: ColumnDefinition[],
    runtimeHelpers: typeof helpers,
  ) => unknown;
  const result = fn([...flatRows], columns, helpers);
  if (!Array.isArray(result)) throw new Error("表格规则必须返回行数组");
  const allowed = new Set(rows.map((row) => row.id));
  if (result.some((row) => !row || typeof row !== "object" || !allowed.has(String(row.id)))) {
    throw new Error("表格规则返回了未知行");
  }
  const byId = new Map(rows.map((row) => [row.id, row]));
  return result.map((row) => byId.get(String(row.id))!);
}

export function recoverInterruptedProject(project: Project): Project {
  const interruptedMessage = "上次生成因页面刷新而中断，请重新生成";
  const recover = <T extends GenerationState>(state: T): T => {
    const recovered = {
      ...state,
      attempts: Array.isArray(state.attempts) ? state.attempts : [],
    };
    return state.status === "generating"
      ? { ...recovered, status: "error", error: interruptedMessage }
      : recovered;
  };
  const sheets = ensureProjectSheets(project).map((sheet) => ({
    ...sheet,
    columns: sheet.columns.map((column) =>
      column.generation ? { ...column, generation: recover(column.generation) } : column,
    ),
    tableRule: recover(sheet.tableRule),
  }));
  const recovered: Project = {
    ...project,
    sheets,
    views: project.views.map((view) => ({
      ...recover(view),
      chartType: view.chartType || "custom",
      sourceSheetIds: view.sourceSheetIds?.length ? view.sourceSheetIds : [project.activeSheetId || sheets[0].id],
      appState: normalizeViewState(view.appState),
      runtimeVersion: GENERATED_APP_RUNTIME_VERSION,
      runtimeEpoch: Number.isFinite(view.runtimeEpoch) ? view.runtimeEpoch : 0,
    })),
    tableRule: recover(project.tableRule),
  };
  hydrateProjectActiveSheet(recovered);
  return recovered;
}

export function serializeProject(project: Project): string {
  syncProjectActiveSheet(project);
  const payload: PersistedPayload = { schemaVersion: 2, project };
  return JSON.stringify(payload, null, 2);
}

export function parseProjectJson(raw: string): Project {
  const payload = JSON.parse(raw) as Partial<PersistedPayload>;
  if (
    (payload.schemaVersion !== 1 && payload.schemaVersion !== 2) ||
    !payload.project ||
    !Array.isArray(payload.project.columns) ||
    !Array.isArray(payload.project.rows) ||
    !Array.isArray(payload.project.views)
  ) {
    throw new Error("不是有效的 AI Excel 项目文件");
  }
  return recoverInterruptedProject(payload.project);
}

function parseCsvRows(input: string): string[][] {
  const text = input.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === '"' && field === "") {
      quoted = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n" || char === "\r") {
      if (char === "\r" && text[index + 1] === "\n") index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((item) => item.some((value) => value !== ""));
}

function inferType(values: string[]): ColumnDataType {
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean);
  if (!nonEmpty.length) return "text";
  if (nonEmpty.every((value) => /^(true|false|是|否)$/i.test(value))) return "boolean";
  if (nonEmpty.every((value) => /^[-+]?\d+(?:\.\d+)?$/.test(value))) return "number";
  if (
    nonEmpty.every(
      (value) => /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(value) && !Number.isNaN(Date.parse(value)),
    )
  ) return "date";
  return "text";
}

function castCsvValue(value: string, type: ColumnDataType): unknown {
  if (value === "") return null;
  if (type === "number") return Number(value);
  if (type === "boolean") return /^(true|是)$/i.test(value);
  if (type === "date") return value.replaceAll("/", "-");
  return value;
}

export function projectFromCsv(raw: string, fileName: string): Project {
  const parsed = parseCsvRows(raw);
  if (parsed.length < 2) throw new Error("CSV 至少需要表头和一行数据");
  const width = Math.max(...parsed.map((row) => row.length));
  const used = new Map<string, number>();
  const names = Array.from({ length: width }, (_, index) => {
    const base = parsed[0][index]?.trim() || `未命名列 ${index + 1}`;
    const count = (used.get(base) || 0) + 1;
    used.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
  const dataRows = parsed.slice(1);
  const columns: ColumnDefinition[] = names.map((name, index) => ({
    id: createId("col"),
    name,
    kind: "manual",
    dataType: inferType(dataRows.map((row) => row[index] || "")),
  }));
  const rows: TableRow[] = dataRows.map((source) => ({
    id: createId("row"),
    values: Object.fromEntries(
      columns.map((column, index) => [
        column.id,
        castCsvValue(source[index] || "", column.dataType),
      ]),
    ),
  }));
  const now = Date.now();
  const tableRule = { id: "table-rule" as const, ...createGenerationState() };
  const sheet: SheetDefinition = { id: createId("sheet"), name: "数据", columns, rows, tableRule };
  return {
    id: createId("project"),
    name: fileName.replace(/\.csv$/i, "") || "CSV 项目",
    columns,
    rows,
    views: [],
    tableRule,
    sheets: [sheet],
    activeSheetId: sheet.id,
    activeTabId: "table",
    createdAt: now,
    updatedAt: now,
  };
}

export function formatCellValue(value: unknown, type: ColumnDataType) {
  if (value == null || value === "") return "";
  if (type === "number" && typeof value === "number") {
    return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2 }).format(value);
  }
  if (type === "boolean") return value ? "是" : "否";
  return String(value);
}
