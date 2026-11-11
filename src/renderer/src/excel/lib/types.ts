export type ColumnDataType = "text" | "number" | "boolean" | "date";
export type GenerationStatus = "idle" | "generating" | "ready" | "error";
export type GenerationAttempt = {
  attempt: number;
  error: string;
  rawError: string;
  code: string;
  createdAt: number;
};
export type JsonPrimitive = string | number | boolean | null;
export type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };
export type ChartType =
  | "custom"
  | "pie"
  | "bar"
  | "line"
  | "area"
  | "scatter"
  | "bubble"
  | "histogram"
  | "radar"
  | "waterfall"
  | "funnel"
  | "treemap";

export type ChartBindings = {
  dimensionColumnId?: string;
  metricColumnId?: string;
  secondaryMetricColumnId?: string;
  sizeColumnId?: string;
};

export type GenerationState = {
  description: string;
  draftDescription: string;
  generatedCode: string;
  lastAttemptCode: string;
  attempts: GenerationAttempt[];
  status: GenerationStatus;
  error: string | null;
  sourceVersion: number;
  generatedVersion: number;
  /** 视图运行时报错（仅 ViewDefinition 使用，置于基类便于统一重置） */
  runtimeError?: string | null;
};

export type ColumnDefinition = {
  id: string;
  name: string;
  kind: "manual" | "computed";
  dataType: ColumnDataType;
  generation?: GenerationState;
  sourceSheetId?: string;
  sourceSheetName?: string;
};

export type TableRow = {
  id: string;
  values: Record<string, unknown>;
};

export type ReferenceImage = {
  id: string;
  name: string;
  dataUrl: string;
};

export type ViewDefinition = GenerationState & {
  id: string;
  name: string;
  chartType: ChartType;
  sourceSheetIds?: string[];
  chartBindings?: ChartBindings;
  chartBindingsLocked?: boolean;
  referenceImages: ReferenceImage[];
  appState: JsonValue;
  runtimeVersion: number;
  runtimeEpoch: number;
  runtimeError: string | null;
};

export type TableRule = GenerationState & { id: "table-rule" };

export type SheetDefinition = {
  id: string;
  name: string;
  columns: ColumnDefinition[];
  rows: TableRow[];
  tableRule: TableRule;
};

export type SheetRelation = {
  leftSheetId: string;
  leftSheetName: string;
  leftColumnId: string;
  leftColumnName: string;
  rightSheetId: string;
  rightSheetName: string;
  rightColumnId: string;
  rightColumnName: string;
  matchRate: number;
  cardinality: "one-to-one" | "many-to-one" | "one-to-many" | "many-to-many";
  joined: boolean;
};

export type Project = {
  id: string;
  name: string;
  columns: ColumnDefinition[];
  rows: TableRow[];
  views: ViewDefinition[];
  tableRule: TableRule;
  sheets?: SheetDefinition[];
  activeSheetId?: string;
  activeTabId: "table" | string;
  createdAt: number;
  updatedAt: number;
};

export type RuntimeRow = {
  id: string;
  values: Record<string, unknown>;
  errors: Record<string, string>;
};

export type PersistedPayload = {
  schemaVersion: 1 | 2;
  project: Project;
};

export type BrowserModelConfig = {
  provider: string;
  baseUrl: string;
  model: string;
  apiKey: string;
  apiMode: "chat-completions" | "responses";
};

export type ModelStatus = {
  configured: boolean;
  provider: string;
  model: string;
};

export type Selection =
  | { type: "project" }
  | { type: "column"; id: string }
  | { type: "view"; id: string }
  | { type: "rule" };
