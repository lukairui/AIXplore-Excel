import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  calculateRuntimeRows,
  createGenerationState,
  createId,
  parseProjectJson,
  serializeProject,
} from "./core";
import { resolveChartBindings } from "./chart-templates";
import { projectForView, resolveViewData } from "./data-sources";
import type { ChartType, ColumnDataType, ColumnDefinition, Project, RuntimeRow, SheetDefinition, ViewDefinition } from "./types";

const META_SHEET = "_NL_SHEET_META";
const META_CHUNK_SIZE = 30_000;

function cellValue(value: ExcelJS.CellValue): unknown {
  if (value == null) return null;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value !== "object") return value;
  if ("result" in value) return cellValue(value.result as ExcelJS.CellValue);
  if ("richText" in value) return value.richText.map((part) => part.text).join("");
  if ("text" in value) return value.text;
  return String(value);
}

function inferExcelType(values: unknown[]): ColumnDataType {
  const present = values.filter((value) => value != null && value !== "");
  if (!present.length) return "text";
  if (present.every((value) => typeof value === "number")) return "number";
  if (present.every((value) => typeof value === "boolean")) return "boolean";
  if (present.every((value) => typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value))) return "date";
  return "text";
}

function uniqueNames(rawNames: string[]) {
  const used = new Map<string, number>();
  return rawNames.map((raw, index) => {
    const base = raw.trim() || `未命名列 ${index + 1}`;
    const count = (used.get(base) || 0) + 1;
    used.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
}

const HEADER_HINT = /(?:^|\b)(id|name|date|time|amount|price|count|status|type|code|number)(?:$|\b)|名称|姓名|日期|时间|金额|数量|单价|状态|类型|编号|编码|地区|城市|部门|职位|学历|年度|月份|备注/i;

function headerRowIndex(worksheet: ExcelJS.Worksheet, width: number) {
  const lastCandidate = Math.min(10, Math.max(1, worksheet.rowCount - 1));
  let best = { index: 1, score: Number.NEGATIVE_INFINITY };
  for (let index = 1; index <= lastCandidate; index += 1) {
    const values = Array.from({ length: width }, (_, columnIndex) => cellValue(worksheet.getRow(index).getCell(columnIndex + 1).value));
    const present = values.filter((value) => value != null && String(value).trim() !== "");
    if (present.length < Math.min(2, width)) continue;
    const textValues = present.filter((value): value is string => typeof value === "string" && Boolean(value.trim()));
    const uniqueRatio = new Set(present.map((value) => String(value).trim().toLocaleLowerCase())).size / present.length;
    const density = present.length / Math.max(1, width);
    const textRatio = textValues.length / present.length;
    const hintRatio = textValues.filter((value) => HEADER_HINT.test(value)).length / Math.max(1, textValues.length);
    const yearHeaderRatio = present.filter((value) => {
      const number = Number(value);
      return Number.isInteger(number) && number >= 1900 && number <= 2100;
    }).length / present.length;
    const nextValues = Array.from({ length: width }, (_, columnIndex) => cellValue(worksheet.getRow(index + 1).getCell(columnIndex + 1).value));
    const typeContrast = values.filter((value, columnIndex) =>
      typeof value === "string"
      && value.trim()
      && nextValues[columnIndex] != null
      && typeof nextValues[columnIndex] !== "string",
    ).length / Math.max(1, present.length);
    const score = density * 4 + textRatio * 1.5 + uniqueRatio + hintRatio * 2 + yearHeaderRatio * 1.5 + typeContrast - (index - 1) * 0.08;
    if (score > best.score) best = { index, score };
  }
  return best.index;
}

function sheetFromWorksheet(worksheet: ExcelJS.Worksheet): SheetDefinition | null {
  if (worksheet.rowCount < 2) return null;
  const width = worksheet.columnCount;
  const headerIndex = headerRowIndex(worksheet, width);
  const names = uniqueNames(Array.from({ length: width }, (_, index) => String(cellValue(worksheet.getRow(headerIndex).getCell(index + 1).value) ?? "")));
  const rawRows = Array.from({ length: Math.max(0, worksheet.rowCount - headerIndex) }, (_, rowIndex) =>
    Array.from({ length: width }, (_, columnIndex) => cellValue(worksheet.getRow(rowIndex + headerIndex + 1).getCell(columnIndex + 1).value)),
  ).filter((row) => row.some((value) => value != null && value !== ""));
  if (!rawRows.length) return null;
  const columns: ColumnDefinition[] = names.map((name, index) => ({
    id: createId("col"),
    name,
    kind: "manual",
    dataType: inferExcelType(rawRows.map((row) => row[index])),
  }));
  const rows = rawRows.map((source) => ({
    id: createId("row"),
    values: Object.fromEntries(columns.map((column, index) => [column.id, source[index] ?? null])),
  }));
  return {
    id: createId("sheet"),
    name: worksheet.name,
    columns,
    rows,
    tableRule: { id: "table-rule", ...createGenerationState() },
  };
}

export async function projectFromExcel(buffer: ArrayBuffer, fileName: string): Promise<Project> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const metadata = workbook.getWorksheet(META_SHEET);
  if (metadata) {
    const chunks: string[] = [];
    metadata.eachRow((row) => {
      const value = row.getCell(1).value;
      if (typeof value === "string") chunks.push(value);
    });
    if (chunks.length) {
      try {
        return parseProjectJson(chunks.join(""));
      } catch {
        // Fall back to importing the visible data sheet.
      }
    }
  }

  const sheets = workbook.worksheets
    .filter((sheet) => sheet.name !== META_SHEET && sheet.state === "visible")
    .map(sheetFromWorksheet)
    .filter((sheet): sheet is SheetDefinition => Boolean(sheet));
  if (!sheets.length) throw new Error("Excel 至少需要一个表头行和一行数据");
  const first = sheets[0];
  const now = Date.now();
  return {
    id: createId("project"),
    name: fileName.replace(/\.xlsx$/i, "") || first.name || "Excel 项目",
    columns: first.columns,
    rows: first.rows,
    views: [],
    tableRule: first.tableRule,
    sheets,
    activeSheetId: first.id,
    activeTabId: "table",
    createdAt: now,
    updatedAt: now,
  };
}

function excelValue(value: unknown, type: ColumnDataType) {
  if (value == null || value === "") return null;
  if (type === "date") {
    const date = new Date(`${String(value).slice(0, 10)}T00:00:00Z`);
    return Number.isNaN(date.getTime()) ? String(value) : date;
  }
  if (type === "number") return Number(value);
  if (type === "boolean") return Boolean(value);
  return String(value);
}

type ExportChartPoint = {
  name: string;
  primary: number;
  secondary: number;
  size: number;
};

type ExportChartDataset = {
  points: ExportChartPoint[];
  histogram: Array<{ name: string; count: number }>;
  waterfall: Array<ExportChartPoint & { base: number; delta: number }>;
  dimensionName: string;
  metricName: string;
  secondaryName: string;
  sizeName: string;
};

const EXPORT_CHART_WIDTH = 760;
const EXPORT_CHART_HEIGHT = 380;
const EXPORT_COLORS = ["#19734a", "#d3922e", "#3977a8", "#b85151", "#7a64a3", "#4f8f8a", "#87944c", "#c26f3a"];

function exportNumber(value: unknown, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function escapeXml(value: unknown) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function exportLabel(value: string, maxLength = 12) {
  return value.length > maxLength ? `${value.slice(0, maxLength - 1)}…` : value;
}

function exportChartDataset(project: Project, view: ViewDefinition, runtimeRows: RuntimeRow[]): ExportChartDataset {
  const bindings = resolveChartBindings(
    project,
    view.chartType,
    view.draftDescription,
    view.chartBindingsLocked ? view.chartBindings : {},
  );
  const dimension = project.columns.find((column) => column.id === bindings.dimensionColumnId);
  const metric = project.columns.find((column) => column.id === bindings.metricColumnId);
  const secondary = project.columns.find((column) => column.id === bindings.secondaryMetricColumnId);
  const size = project.columns.find((column) => column.id === bindings.sizeColumnId);
  const rawPoints = runtimeRows.map((row, index) => ({
    name: String(row.values[dimension?.id || ""] ?? `记录 ${index + 1}`),
    primary: exportNumber(row.values[metric?.id || ""]),
    secondary: exportNumber(row.values[secondary?.id || ""]),
    size: Math.max(1, Math.abs(exportNumber(row.values[size?.id || ""], 1))),
  }));
  const grouped = new Map<string, ExportChartPoint>();
  rawPoints.forEach((point) => {
    const current = grouped.get(point.name) || { name: point.name, primary: 0, secondary: 0, size: 0 };
    current.primary += point.primary;
    current.secondary += point.secondary;
    current.size += point.size;
    grouped.set(point.name, current);
  });
  const points = ["scatter", "bubble", "histogram", "waterfall"].includes(view.chartType) ? rawPoints : [...grouped.values()];
  const values = points.map((point) => point.primary);
  const minValue = values.length ? Math.min(...values) : 0;
  const maxValue = values.length ? Math.max(...values) : 0;
  const bucketCount = Math.max(1, Math.min(6, Math.ceil(Math.sqrt(values.length || 1))));
  const bucketSize = Math.max(1, (maxValue - minValue) / bucketCount);
  const histogram = Array.from({ length: bucketCount }, (_, index) => ({
    name: `${Math.round(minValue + index * bucketSize)}–${Math.round(minValue + (index + 1) * bucketSize)}`,
    count: 0,
  }));
  values.forEach((value) => {
    const index = Math.min(bucketCount - 1, Math.floor((value - minValue) / bucketSize));
    histogram[index].count += 1;
  });
  let runningTotal = 0;
  const waterfall = points.map((point) => {
    const base = runningTotal;
    runningTotal += point.primary;
    return { ...point, base, delta: point.primary };
  });
  return {
    points,
    histogram,
    waterfall,
    dimensionName: dimension?.name || "分类",
    metricName: metric?.name || "数值",
    secondaryName: secondary?.name || "指标二",
    sizeName: size?.name || "大小",
  };
}

function chartSvg(type: Exclude<ChartType, "custom">, title: string, dataset: ExportChartDataset) {
  const width = EXPORT_CHART_WIDTH;
  const height = EXPORT_CHART_HEIGHT;
  const left = 66;
  const top = 48;
  const right = width - 24;
  const bottom = height - 58;
  const plotWidth = right - left;
  const plotHeight = bottom - top;
  const points = type === "histogram"
    ? dataset.histogram.map((item) => ({ name: item.name, primary: item.count, secondary: 0, size: 1 }))
    : dataset.points;
  const values = points.map((point) => point.primary);
  const scaleValues = type === "waterfall"
    ? dataset.waterfall.flatMap((point) => [point.base, point.base + point.delta])
    : values;
  const maxValue = Math.max(1, ...scaleValues);
  const maxMagnitude = Math.max(1, ...values.map((value) => Math.abs(value)));
  const minValue = Math.min(0, ...scaleValues);
  const valueRange = Math.max(1, maxValue - minValue);
  const xFor = (index: number) => left + (points.length <= 1 ? plotWidth / 2 : (index / (points.length - 1)) * plotWidth);
  const yFor = (value: number) => bottom - ((value - minValue) / valueRange) * plotHeight;
  const zeroY = yFor(0);
  const text = (value: unknown, x: number, y: number, size = 12, anchor = "middle", fill = "#405148") => `<text x="${x}" y="${y}" text-anchor="${anchor}" font-family="Arial, Microsoft YaHei, sans-serif" font-size="${size}" fill="${fill}">${escapeXml(value)}</text>`;
  const grid = [0, 0.25, 0.5, 0.75, 1].map((step) => {
    const y = top + plotHeight * step;
    const value = maxValue - valueRange * step;
    return `<line x1="${left}" y1="${y}" x2="${right}" y2="${y}" stroke="#e2e8e4" stroke-dasharray="4 4" />${text(Math.round(value).toLocaleString("zh-CN"), left - 8, y + 4, 10, "end", "#76837b")}`;
  }).join("");
  const axis = `<line x1="${left}" y1="${top}" x2="${left}" y2="${bottom}" stroke="#aebbb3" /><line x1="${left}" y1="${bottom}" x2="${right}" y2="${bottom}" stroke="#aebbb3" />`;
  const labels = points.map((point, index) => text(exportLabel(point.name), xFor(index), bottom + 23, 10)).join("");
  let marks = "";

  if (["bar", "histogram"].includes(type)) {
    const slot = plotWidth / Math.max(1, points.length);
    const barWidth = Math.max(12, Math.min(58, slot * 0.68));
    marks = points.map((point, index) => {
      const x = left + slot * index + (slot - barWidth) / 2;
      const y = yFor(point.primary);
      const rectY = point.primary >= 0 ? y : zeroY;
      const rectHeight = Math.max(1, Math.abs(zeroY - y));
      return `<rect x="${x}" y="${rectY}" width="${barWidth}" height="${rectHeight}" rx="3" fill="${EXPORT_COLORS[index % EXPORT_COLORS.length]}" />`;
    }).join("");
  } else if (["line", "area"].includes(type)) {
    const path = points.map((point, index) => `${xFor(index)},${yFor(point.primary)}`).join(" ");
    const area = type === "area" ? `<polygon points="${left},${zeroY} ${path} ${right},${zeroY}" fill="#bfe0cd" opacity="0.75" />` : "";
    marks = `${area}<polyline points="${path}" fill="none" stroke="#19734a" stroke-width="3" stroke-linejoin="round" />${points.map((point, index) => `<circle cx="${xFor(index)}" cy="${yFor(point.primary)}" r="4" fill="#19734a" />`).join("")}`;
  } else if (["scatter", "bubble"].includes(type)) {
    const xValues = points.map((point) => point.secondary);
    const xMin = Math.min(0, ...xValues);
    const xRange = Math.max(1, Math.max(...xValues) - xMin);
    marks = points.map((point, index) => {
      const x = left + ((point.secondary - xMin) / xRange) * plotWidth;
      const radius = type === "bubble" ? Math.min(28, 7 + Math.sqrt(point.size) * 2) : 6;
      return `<circle cx="${x}" cy="${yFor(point.primary)}" r="${radius}" fill="${EXPORT_COLORS[index % EXPORT_COLORS.length]}" fill-opacity="0.72" stroke="#ffffff" />`;
    }).join("");
  } else if (type === "pie") {
    const total = points.reduce((sum, point) => sum + Math.max(0, point.primary), 0) || 1;
    let angle = -Math.PI / 2;
    const cx = 270;
    const cy = 195;
    const radius = 112;
    marks = points.map((point, index) => {
      const slice = (Math.max(0, point.primary) / total) * Math.PI * 2;
      const end = angle + slice;
      const large = slice > Math.PI ? 1 : 0;
      const startX = cx + Math.cos(angle) * radius;
      const startY = cy + Math.sin(angle) * radius;
      const endX = cx + Math.cos(end) * radius;
      const endY = cy + Math.sin(end) * radius;
      const path = `<path d="M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${large} 1 ${endX} ${endY} Z" fill="${EXPORT_COLORS[index % EXPORT_COLORS.length]}" stroke="#ffffff" stroke-width="2" />`;
      const legendColumn = Math.floor(index / 8);
      const legendRow = index % 8;
      const legendX = 440 + legendColumn * 150;
      const legend = `<rect x="${legendX}" y="${90 + legendRow * 22}" width="10" height="10" rx="2" fill="${EXPORT_COLORS[index % EXPORT_COLORS.length]}" />${text(exportLabel(point.name, 15), legendX + 17, 100 + legendRow * 22, 10, "start")}`;
      angle = end;
      return path + legend;
    }).join("");
  } else if (type === "radar") {
    const cx = width / 2;
    const cy = 198;
    const radius = 126;
    const count = Math.max(3, points.length);
    const coordinate = (index: number, value: number) => {
      const angle = -Math.PI / 2 + (index / count) * Math.PI * 2;
      const r = radius * Math.max(0, value) / maxMagnitude;
      return [cx + Math.cos(angle) * r, cy + Math.sin(angle) * r];
    };
    const gridPolygons = [0.25, 0.5, 0.75, 1].map((scale) => `<polygon points="${Array.from({ length: count }, (_, index) => { const [x, y] = coordinate(index, maxMagnitude * scale); return `${x},${y}`; }).join(" ")}" fill="none" stroke="#dce5df" />`).join("");
    const axisLines = Array.from({ length: count }, (_, index) => { const [x, y] = coordinate(index, maxMagnitude); return `<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#dce5df" />${text(exportLabel(points[index]?.name || `维度 ${index + 1}`), x, y + (y < cy ? -7 : 16), 10)}`; }).join("");
    const dataPolygon = `<polygon points="${points.map((point, index) => { const [x, y] = coordinate(index, point.primary); return `${x},${y}`; }).join(" ")}" fill="#52a579" fill-opacity="0.42" stroke="#19734a" stroke-width="3" />`;
    marks = gridPolygons + axisLines + dataPolygon;
  } else if (type === "waterfall") {
    const slot = plotWidth / Math.max(1, points.length);
    const barWidth = Math.max(12, Math.min(58, slot * 0.68));
    marks = dataset.waterfall.map((point, index) => {
      const x = left + slot * index + (slot - barWidth) / 2;
      const from = yFor(point.base);
      const to = yFor(point.base + point.delta);
      const y = Math.min(from, to);
      const height = Math.max(1, Math.abs(to - from));
      return `<rect x="${x}" y="${y}" width="${barWidth}" height="${height}" rx="3" fill="${point.delta >= 0 ? "#2c8b5f" : "#b85151"}" />`;
    }).join("");
  } else if (type === "funnel") {
    const funnel = [...points].sort((a, b) => b.primary - a.primary);
    const max = Math.max(1, ...funnel.map((point) => Math.max(0, point.primary)));
    marks = funnel.map((point, index) => {
      const y = 68 + index * Math.min(48, 250 / Math.max(1, funnel.length));
      const itemWidth = Math.max(50, (Math.max(0, point.primary) / max) * 480);
      const x = (width - itemWidth) / 2;
      return `<path d="M ${x} ${y} H ${x + itemWidth} L ${x + itemWidth * 0.84} ${y + 34} H ${x + itemWidth * 0.16} Z" fill="${EXPORT_COLORS[index % EXPORT_COLORS.length]}" opacity="0.86" />${text(`${exportLabel(point.name, 18)}  ${point.primary.toLocaleString("zh-CN")}`, width / 2, y + 23, 11, "middle", "#ffffff")}`;
    }).join("");
  } else if (type === "treemap") {
    const sorted = [...points].sort((a, b) => b.primary - a.primary);
    const positiveTotal = sorted.reduce((sum, point) => sum + Math.max(0, point.primary), 0);
    const total = positiveTotal || 1;
    let cursorX = left;
    marks = sorted.map((point, index) => {
      const itemWidth = positiveTotal ? (Math.max(0, point.primary) / total) * plotWidth : plotWidth / Math.max(1, sorted.length);
      const actualWidth = Math.max(1, Math.min(itemWidth, right - cursorX));
      const rect = `<rect x="${cursorX}" y="120" width="${actualWidth}" height="150" fill="${EXPORT_COLORS[index % EXPORT_COLORS.length]}" stroke="#ffffff" stroke-width="2" />${text(exportLabel(point.name, 14), cursorX + actualWidth / 2, 190, 11, "middle", "#ffffff")}`;
      cursorX += actualWidth;
      return rect;
    }).join("");
  }

  const showAxes = !["pie", "radar", "funnel", "treemap"].includes(type);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#ffffff" /><text x="24" y="28" font-family="Arial, Microsoft YaHei, sans-serif" font-size="16" font-weight="700" fill="#18231d">${escapeXml(title)}</text>${showAxes ? grid + axis + labels : ""}${marks}${showAxes ? text(dataset.dimensionName, (left + right) / 2, height - 13, 10) + text(dataset.metricName, 14, top - 10, 10, "start") : ""}</svg>`;
}

export function createExcelChartSvg(project: Project, view: ViewDefinition, runtimeRows: RuntimeRow[]) {
  if (view.chartType === "custom") return "";
  return chartSvg(view.chartType, view.name || "图表", exportChartDataset(project, view, runtimeRows));
}

async function svgToPng(svg: string) {
  if (
    typeof document === "undefined"
    || typeof Image === "undefined"
    || (typeof navigator !== "undefined" && navigator.userAgent.toLowerCase().includes("jsdom"))
  ) return null;
  try {
    const canvas = document.createElement("canvas");
    canvas.width = EXPORT_CHART_WIDTH;
    canvas.height = EXPORT_CHART_HEIGHT;
    const context = canvas.getContext("2d");
    if (!context) return null;
    const image = new Image();
    const dataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
    return await new Promise<string | null>((resolve) => {
      image.onload = () => {
        context.fillStyle = "#ffffff";
        context.fillRect(0, 0, canvas.width, canvas.height);
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        try {
          resolve(canvas.toDataURL("image/png"));
        } catch {
          resolve(null);
        }
      };
      image.onerror = () => resolve(null);
      image.src = dataUrl;
    });
  } catch {
    return null;
  }
}

function styleExportChartSheet(sheet: ExcelJS.Worksheet) {
  sheet.getColumn(1).width = 14;
  for (let index = 2; index <= 9; index += 1) sheet.getColumn(index).width = 12;
  sheet.getColumn(11).width = 22;
  sheet.getColumn(12).width = 14;
  sheet.getColumn(13).width = 14;
  sheet.getColumn(14).width = 14;
  sheet.getColumn(15).width = 14;
}

type NativeChartPoint = ExportChartPoint;

type NativeChartSpec = {
  sheetName: string;
  title: string;
  chartType: Exclude<ChartType, "custom">;
  sourceStartRow: number;
  sourceEndRow: number;
  categoryColumn: string;
  primaryColumn: string;
  secondaryColumn: string;
  sizeColumn: string;
  points: NativeChartPoint[];
  anchor: { fromRow: number; fromCol: number; toRow: number; toCol: number };
  waterfall?: boolean;
  primaryName?: string;
  secondaryName?: string;
  primaryColor?: string;
  secondaryColor?: string;
  palette?: string[];
  doughnut?: boolean;
  comboLine?: boolean;
  primaryNumberFormat?: string;
  secondaryNumberFormat?: string;
  showLegend?: boolean;
  plotVisibleOnly?: boolean;
};

const DASHBOARD_COLORS = ["3B82F6", "10B981", "F59E0B", "8B5CF6", "EC4899", "06B6D4", "6366F1", "64748B"];

const excelLetters = (column: number) => {
  let value = column;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
};

function xmlText(value: unknown) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function chartTextXml(value: string) {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="zh-CN" sz="1200" b="1"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>${xmlText(value)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

function seriesTextXml(value: string) {
  return `<c:tx><c:v>${xmlText(value)}</c:v></c:tx>`;
}

function stringCache(points: NativeChartPoint[]) {
  return `<c:strCache><c:ptCount val="${points.length}"/>${points.map((point, index) => `<c:pt idx="${index}"><c:v>${xmlText(point.name)}</c:v></c:pt>`).join("")}</c:strCache>`;
}

function numberCache(points: NativeChartPoint[], field: keyof NativeChartPoint, formatCode = "#,##0.00") {
  return `<c:numCache><c:formatCode>${xmlText(formatCode)}</c:formatCode><c:ptCount val="${points.length}"/>${points.map((point, index) => `<c:pt idx="${index}"><c:v>${Number(point[field] || 0)}</c:v></c:pt>`).join("")}</c:numCache>`;
}

function strRef(sheetName: string, column: string, startRow: number, endRow: number, points: NativeChartPoint[]) {
  return `<c:strRef><c:f>'${sheetName.replaceAll("'", "''")}'!$${column}$${startRow}:$${column}$${endRow}</c:f>${stringCache(points)}</c:strRef>`;
}

function numRef(sheetName: string, column: string, startRow: number, endRow: number, points: NativeChartPoint[], field: keyof NativeChartPoint, formatCode = "#,##0.00") {
  return `<c:numRef><c:f>'${sheetName.replaceAll("'", "''")}'!$${column}$${startRow}:$${column}$${endRow}</c:f>${numberCache(points, field, formatCode)}</c:numRef>`;
}

function chartSeriesXml(spec: NativeChartSpec, index: number, field: keyof NativeChartPoint = "primary", name = "", color = "19734A", options = "") {
  const formatCode = field === "secondary" ? spec.secondaryNumberFormat : spec.primaryNumberFormat;
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${seriesTextXml(name || spec.title)}<c:spPr>${options || `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>`}</c:spPr><c:cat>${strRef(spec.sheetName, spec.categoryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points)}</c:cat><c:val>${numRef(spec.sheetName, field === "secondary" ? spec.secondaryColumn : spec.primaryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points, field, formatCode)}</c:val></c:ser>`;
}

function lineSeriesXml(spec: NativeChartSpec, index: number) {
  const color = spec.secondaryColor || "10B981";
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${seriesTextXml(spec.secondaryName || "指标二")}<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr></c:marker><c:cat>${strRef(spec.sheetName, spec.categoryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points)}</c:cat><c:val>${numRef(spec.sheetName, spec.secondaryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points, "secondary", spec.secondaryNumberFormat || "0.0%")}</c:val><c:smooth val="1"/></c:ser>`;
}

function pieSeriesXml(spec: NativeChartSpec) {
  const palette = spec.palette?.length ? spec.palette : DASHBOARD_COLORS;
  const pointStyles = spec.points.map((_, index) => {
    const color = palette[index % palette.length];
    return `<c:dPt><c:idx val="${index}"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`;
  }).join("");
  return `<c:ser><c:idx val="0"/><c:order val="0"/>${seriesTextXml(spec.primaryName || spec.title)}${pointStyles}<c:cat>${strRef(spec.sheetName, spec.categoryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points)}</c:cat><c:val>${numRef(spec.sheetName, spec.primaryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points, "primary", spec.primaryNumberFormat || "#,##0")}</c:val></c:ser>`;
}

function chartAxes() {
  return `<c:axId val="48650112"/><c:axId val="48672768"/>`;
}

function axisTextXml() {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="850"><a:solidFill><a:srgbClr val="64748B"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></c:txPr>`;
}

function gridlineXml() {
  return `<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E7EB"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>`;
}

function axisXml(primaryNumberFormat = "#,##0.00") {
  return `<c:catAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="low"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${gridlineXml()}<c:numFmt formatCode="${xmlText(primaryNumberFormat)}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
}

function scatterAxisXml() {
  return `<c:valAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx><c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${gridlineXml()}<c:numFmt formatCode="#,##0.00" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
}

function comboAxisXml(spec: NativeChartSpec) {
  const primaryFormat = xmlText(spec.primaryNumberFormat || "#,##0");
  const secondaryFormat = xmlText(spec.secondaryNumberFormat || "0.0%");
  return `<c:catAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="low"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${gridlineXml()}<c:numFmt formatCode="${primaryFormat}" sourceLinked="0"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx><c:valAx><c:axId val="48672769"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="${secondaryFormat}" sourceLinked="0"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisTextXml()}<c:crossAx val="48650112"/><c:crosses val="max"/><c:crossBetween val="midCat"/></c:valAx>`;
}

function nativeChartXml(spec: NativeChartSpec) {
  const type = spec.chartType;
  const fallbackBar = type === "histogram" || type === "funnel" || type === "treemap" || type === "waterfall";
  const chartType = fallbackBar ? "bar" : type;
  let plot = "";
  let axes = "";
  if (spec.comboLine) {
    const primaryColor = spec.primaryColor || "3B82F6";
    plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${chartSeriesXml(spec, 0, "primary", spec.primaryName || spec.title, primaryColor)}<c:gapWidth val="70"/><c:overlap val="0"/><c:axId val="48650112"/><c:axId val="48672768"/></c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lineSeriesXml(spec, 1)}<c:marker val="1"/><c:axId val="48650112"/><c:axId val="48672769"/></c:lineChart>`;
    axes = comboAxisXml(spec);
  } else if (chartType === "pie") {
    const pieTag = spec.doughnut ? "doughnutChart" : "pieChart";
    plot = `<c:${pieTag}><c:varyColors val="1"/>${pieSeriesXml(spec)}<c:firstSliceAng val="270"/>${spec.doughnut ? `<c:holeSize val="58"/>` : ""}</c:${pieTag}>`;
  } else if (chartType === "scatter" || chartType === "bubble") {
    const scatterSeries = `<c:ser><c:idx val="0"/><c:order val="0"/>${seriesTextXml(spec.title)}<c:spPr><a:solidFill><a:srgbClr val="${spec.primaryColor || "3B82F6"}"/></a:solidFill></c:spPr><c:xVal>${numRef(spec.sheetName, spec.secondaryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points, "secondary")}</c:xVal><c:yVal>${numRef(spec.sheetName, spec.primaryColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points, "primary")}</c:yVal>${chartType === "bubble" ? `<c:bubbleSize>${numRef(spec.sheetName, spec.sizeColumn, spec.sourceStartRow, spec.sourceEndRow, spec.points, "size")}</c:bubbleSize>` : ""}</c:ser>`;
    plot = chartType === "bubble"
      ? `<c:bubbleChart><c:varyColors val="0"/>${scatterSeries}<c:bubble3D val="0"/><c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/>${chartAxes()}</c:bubbleChart>`
      : `<c:scatterChart><c:scatterStyle val="marker"/><c:varyColors val="0"/>${scatterSeries}${chartAxes()}</c:scatterChart>`;
    axes = scatterAxisXml();
  } else if (chartType === "line") {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${chartSeriesXml(spec, 0, "primary", spec.primaryName || spec.title, spec.primaryColor || "3B82F6")}${chartAxes()}</c:lineChart>`;
    axes = axisXml(spec.primaryNumberFormat);
  } else if (chartType === "area") {
    plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${chartSeriesXml(spec, 0, "primary", spec.primaryName || spec.title, spec.primaryColor || "3B82F6")}${chartAxes()}</c:areaChart>`;
    axes = axisXml(spec.primaryNumberFormat);
  } else if (chartType === "radar") {
    plot = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${chartSeriesXml(spec, 0, "primary", spec.primaryName || spec.title, spec.primaryColor || "3B82F6")}${chartAxes()}</c:radarChart>`;
    axes = axisXml(spec.primaryNumberFormat);
  } else {
    const barDir = type === "funnel" ? "bar" : "col";
    const series = type === "waterfall"
      ? `${chartSeriesXml(spec, 0, "secondary", "基准", "FFFFFF", `<a:noFill/><a:ln><a:noFill/></a:ln>`)}${chartSeriesXml(spec, 1, "primary", "变化值", "19734A")}`
      : chartSeriesXml(spec, 0, "primary", spec.primaryName || spec.title, spec.primaryColor || "3B82F6");
    plot = `<c:barChart><c:barDir val="${barDir}"/><c:grouping val="${type === "waterfall" ? "stacked" : "clustered"}"/><c:varyColors val="0"/>${series}<c:gapWidth val="${barDir === "bar" ? "55" : "85"}"/>${chartAxes()}</c:barChart>`;
    axes = axisXml(spec.primaryNumberFormat);
  }
  const showLegend = spec.showLegend ?? (chartType === "pie" || Boolean(spec.comboLine));
  const legendPosition = chartType === "pie" ? "r" : "b";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="zh-CN"/><c:roundedCorners val="0"/><c:style val="2"/><c:chart><c:title>${chartTextXml(spec.title)}<c:layout/><c:overlay val="0"/></c:title><c:plotArea><c:layout/>${plot}${axes}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${showLegend ? `<c:legend><c:legendPos val="${legendPosition}"/><c:layout/><c:overlay val="0"/>${axisTextXml()}</c:legend>` : ""}<c:plotVisOnly val="${spec.plotVisibleOnly === false ? "0" : "1"}"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:pPr><a:endParaRPr lang="zh-CN"/></a:p></c:txPr></c:chartSpace>`;
}

function drawingXml(specs: NativeChartSpec[]) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${specs.map((spec, index) => `<xdr:twoCellAnchor editAs="twoCell"><xdr:from><xdr:col>${spec.anchor.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${spec.anchor.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${spec.anchor.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${spec.anchor.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 2}" name="${xmlText(spec.title)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId${index + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join("")}</xdr:wsDr>`;
}

function drawingRelsXml(chartIndices: number[]) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${chartIndices.map((chartIndex, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIndex}.xml"/>`).join("")}</Relationships>`;
}

function uniqueSheetName(workbook: ExcelJS.Workbook, prefix: string) {
  const base = prefix.replace(/[\\/?*\[\]:]/g, " ").trim().slice(0, 28) || "大屏";
  let candidate = base;
  let index = 2;
  while (workbook.getWorksheet(candidate)) candidate = `${base.slice(0, 28 - String(index).length - 1)}-${index++}`;
  return candidate;
}

function addChartSource(
  sheet: ExcelJS.Worksheet,
  view: ViewDefinition,
  project: Project,
  runtimeRows: RuntimeRow[],
  startRow: number,
  sourceColumn: number,
  anchor: NativeChartSpec["anchor"],
): NativeChartSpec {
  if (view.chartType === "custom") throw new Error("AI 应用不能作为单个预制图表导出");
  const dataset = exportChartDataset(project, view, runtimeRows);
  const chartTitle = view.name || "图表";
  const sourcePoints = view.chartType === "histogram"
    ? dataset.histogram.map((item) => ({ name: item.name, primary: item.count, secondary: 0, size: 1 }))
    : view.chartType === "waterfall"
      ? dataset.waterfall.map((item) => ({ name: item.name, primary: item.delta, secondary: item.base, size: item.size }))
      : view.chartType === "funnel"
        ? [...dataset.points].sort((a, b) => b.primary - a.primary)
        : dataset.points;
  const points = sourcePoints.length ? sourcePoints : [{ name: "暂无数据", primary: 0, secondary: 0, size: 1 }];
  const headers = view.chartType === "waterfall"
    ? ["分类", "变化值", "基准", dataset.sizeName]
    : ["分类", dataset.metricName, dataset.secondaryName, dataset.sizeName];
  headers.forEach((value, index) => {
    const cell = sheet.getCell(startRow, sourceColumn + index);
    cell.value = value;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF19734A" } };
  });
  points.forEach((point, index) => {
    const row = startRow + index + 1;
    [point.name, point.primary, point.secondary, point.size].forEach((value, columnIndex) => {
      const cell = sheet.getCell(row, sourceColumn + columnIndex);
      cell.value = value;
      cell.alignment = { vertical: "middle" };
      if (columnIndex > 0) cell.numFmt = "#,##0.00";
      cell.border = { bottom: { style: "thin", color: { argb: "FFE2E8E4" } } };
    });
  });
  return {
    sheetName: sheet.name,
    title: chartTitle,
    chartType: view.chartType,
    sourceStartRow: startRow + 1,
    sourceEndRow: startRow + points.length,
    categoryColumn: excelLetters(sourceColumn),
    primaryColumn: excelLetters(sourceColumn + 1),
    secondaryColumn: excelLetters(sourceColumn + 2),
    sizeColumn: excelLetters(sourceColumn + 3),
    points,
    anchor,
    waterfall: view.chartType === "waterfall",
  };
}

type DashboardColumns = {
  dimension?: ColumnDefinition;
  primary?: ColumnDefinition;
  secondary?: ColumnDefinition;
  secondaryMode?: "margin" | "rate" | "value";
  quantity?: ColumnDefinition;
};

type DashboardFilterColumns = {
  category?: ColumnDefinition;
  date?: ColumnDefinition;
};

type DashboardGroup = {
  key: string;
  name: string;
  primary: number;
  secondary: number;
  quantity: number;
  count: number;
};

function dashboardColumns(project: Project, view: ViewDefinition): DashboardColumns {
  const context = `${view.name} ${view.draftDescription}`.toLocaleLowerCase();
  const numeric = project.columns.filter((column) => column.dataType === "number");
  const descriptive = project.columns.filter((column) => column.dataType !== "number" && column.dataType !== "date");
  const score = (column: ColumnDefinition, patterns: RegExp[]) => {
    const name = column.name.toLocaleLowerCase();
    const semanticScore = patterns.reduce((total, pattern, index) => total + (pattern.test(name) ? 60 - index * 5 : 0), 0);
    return semanticScore + (semanticScore > 0 && context.includes(name) ? 100 : 0);
  };
  const choose = (columns: ColumnDefinition[], patterns: RegExp[], exclude: Set<string> = new Set()) => columns
    .filter((column) => !exclude.has(column.id))
    .map((column) => ({ column, score: score(column, patterns) }))
    .sort((left, right) => right.score - left.score)
    .find((item) => item.score > 0)?.column;

  const mentionedDescriptive = descriptive.filter((column) => context.includes(column.name.toLocaleLowerCase()) && !/(?:编号|代码|序号|^id$|id$|code$)/i.test(column.name));
  const dimension = mentionedDescriptive[0]
    || choose(descriptive, [/(?:地区|区域|省份|城市|国家|region|area)/i, /(?:分类|类别|品类|渠道|部门|门店|状态|负责人)/i])
    || descriptive.find((column) => !/(?:编号|代码|序号|^id$|id$|code$)/i.test(column.name))
    || descriptive[0];
  const quantity = choose(numeric, [/(?:数量|件数|销量|订单数|笔数|count|qty|quantity)/i]);
  const mentionedNumeric = numeric.filter((column) => context.includes(column.name.toLocaleLowerCase()));
  const primary = mentionedNumeric.find((column) => column.id !== quantity?.id)
    || choose(numeric, [/(?:销售额|营业额|营收|收入|成交额|交易额|revenue|sales|amount)/i, /(?:金额|总额|价格|单价|value)/i], new Set(quantity ? [quantity.id] : []))
    || numeric.find((column) => column.id !== quantity?.id)
    || numeric[0];
  const excludedSecondary = new Set([primary?.id, quantity?.id].filter(Boolean) as string[]);
  const marginColumn = choose(numeric, [/(?:利润|毛利|净利|收益|profit|margin)/i], excludedSecondary);
  const secondary = marginColumn || mentionedNumeric.find((column) => !excludedSecondary.has(column.id));
  const secondaryMode = secondary
    ? marginColumn
      ? "margin"
      : /(?:率|比例|占比|百分比|rate|ratio|percent|pct)/i.test(secondary.name)
        ? "rate"
        : "value"
    : undefined;
  return { dimension, primary, secondary, secondaryMode, quantity };
}

function dashboardFilterColumns(project: Project, view: ViewDefinition, runtimeRows: RuntimeRow[], dimension?: ColumnDefinition): DashboardFilterColumns {
  const context = `${view.name} ${view.draftDescription}`.toLocaleLowerCase();
  const categoryCandidates = project.columns
    .filter((column) => column.id !== dimension?.id && column.dataType !== "number" && column.dataType !== "date" && !/(?:编号|代码|序号|^id$|id$|code$)/i.test(column.name))
    .map((column) => {
      const distinct = new Set(runtimeRows.map((row) => String(row.values[column.id] ?? "").trim()).filter(Boolean)).size;
      const semantic = /(?:分类|类别|品类|渠道|部门|状态|负责人|门店|类型|组别|category|status|channel|owner|type|group)/i.test(column.name);
      const score = (context.includes(column.name.toLocaleLowerCase()) ? 100 : 0)
        + (semantic ? 50 : 0)
        + (distinct >= 2 && distinct <= 30 ? 25 : 0)
        + (column.id !== dimension?.id ? 10 : 0)
        - (distinct > 100 ? 80 : 0);
      return { column, score, distinct };
    })
    .filter((item) => item.distinct > 0)
    .sort((left, right) => right.score - left.score);
  const dateCandidates = project.columns.filter((column) => column.dataType === "date");
  const date = dateCandidates.find((column) => context.includes(column.name.toLocaleLowerCase()))
    || dateCandidates.find((column) => /(?:日期|时间|date|time)/i.test(column.name))
    || dateCandidates[0];
  return { category: categoryCandidates[0]?.column, date };
}

function isMoneyColumn(column?: ColumnDefinition) {
  return Boolean(column && /(?:销售额|营业额|营收|收入|成交额|交易额|金额|总额|利润|毛利|净利|收益|价格|单价|revenue|sales|amount|profit)/i.test(column.name));
}

function dashboardNumberFormat(column?: ColumnDefinition) {
  return isMoneyColumn(column) ? "¥#,##0" : "#,##0";
}

function stylePanelRange(sheet: ExcelJS.Worksheet, fromRow: number, fromCol: number, toRow: number, toCol: number, fill: string) {
  const borderColor = { argb: "FFE2E8F0" };
  for (let row = fromRow; row <= toRow; row += 1) {
    for (let column = fromCol; column <= toCol; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${fill}` } };
      cell.border = {
        ...(row === fromRow ? { top: { style: "thin", color: borderColor } as ExcelJS.Border } : {}),
        ...(row === toRow ? { bottom: { style: "thin", color: borderColor } as ExcelJS.Border } : {}),
        ...(column === fromCol ? { left: { style: "thin", color: borderColor } as ExcelJS.Border } : {}),
        ...(column === toCol ? { right: { style: "thin", color: borderColor } as ExcelJS.Border } : {}),
      };
    }
  }
}

function addDashboardSheet(project: Project, view: ViewDefinition, runtimeRows: RuntimeRow[], workbook: ExcelJS.Workbook, skipCardLayout = false): NativeChartSpec[] {
  if (!skipCardLayout && /卡牌|卡片|card/i.test(`${view.name} ${view.draftDescription}`)) {
    return addCardDashboardSheet(project, view, runtimeRows, workbook);
  }
  const sheetName = uniqueSheetName(workbook, `大屏-${view.name}`);
  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 4, showGridLines: false, zoomScale: 80 }],
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
      margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 },
    },
  });
  sheet.properties.defaultRowHeight = 18;
  for (let index = 1; index <= 16; index += 1) sheet.getColumn(index).width = 10.5;
  for (let index = 17; index <= Math.max(19, project.columns.length); index += 1) sheet.getColumn(index).width = 12;
  for (let index = 20; index <= 25; index += 1) {
    sheet.getColumn(index).width = 14;
    sheet.getColumn(index).hidden = true;
  }

  const { dimension, primary, secondary, secondaryMode, quantity } = dashboardColumns(project, view);
  const { category: categoryFilter, date: dateFilter } = dashboardFilterColumns(project, view, runtimeRows, dimension);
  const detailsHeaderRow = 54;
  const detailsStartRow = detailsHeaderRow + 1;
  const detailsEndRow = Math.max(detailsStartRow, detailsStartRow + runtimeRows.length - 1);
  const dashboardEndRow = Math.max(66, detailsEndRow);
  for (let row = 1; row <= dashboardEndRow; row += 1) {
    for (let column = 1; column <= 16; column += 1) {
      sheet.getCell(row, column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F5F7" } };
    }
  }

  stylePanelRange(sheet, 1, 1, 5, 16, "FFFFFF");
  sheet.mergeCells(1, 1, 2, 16);
  const titleCell = sheet.getCell(1, 1);
  titleCell.value = view.name || "AI 数据分析大屏";
  titleCell.font = { name: "Microsoft YaHei", bold: true, size: 20, color: { argb: "FF111827" } };
  titleCell.alignment = { vertical: "middle", horizontal: "left" };
  sheet.mergeCells(3, 1, 3, 16);
  const descriptionCell = sheet.getCell(3, 1);
  descriptionCell.value = view.draftDescription || "基于当前表格数据生成的结构化分析大屏";
  descriptionCell.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF64748B" } };
  descriptionCell.alignment = { vertical: "middle", wrapText: true };
  if (categoryFilter) {
    sheet.mergeCells(4, 1, 4, 2);
    const label = sheet.getCell(4, 1);
    label.value = `${categoryFilter.name}筛选`;
    label.font = { name: "Microsoft YaHei", bold: true, size: 9, color: { argb: "FF475569" } };
    label.alignment = { vertical: "middle", horizontal: "right" };
    sheet.mergeCells(4, 3, 4, 6);
    const selector = sheet.getCell(4, 3);
    selector.value = "全部";
    selector.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF1F2937" } };
    selector.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    selector.border = { top: { style: "thin", color: { argb: "FFCBD5E1" } }, bottom: { style: "thin", color: { argb: "FFCBD5E1" } }, left: { style: "thin", color: { argb: "FFCBD5E1" } }, right: { style: "thin", color: { argb: "FFCBD5E1" } } };
    selector.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  }
  if (dateFilter) {
    sheet.mergeCells(4, 7, 4, 8);
    const label = sheet.getCell(4, 7);
    label.value = `${dateFilter.name}区间`;
    label.font = { name: "Microsoft YaHei", bold: true, size: 9, color: { argb: "FF475569" } };
    label.alignment = { vertical: "middle", horizontal: "right" };
    sheet.mergeCells(4, 9, 4, 12);
    sheet.mergeCells(4, 13, 4, 16);
    [sheet.getCell(4, 9), sheet.getCell(4, 13)].forEach((cell, index) => {
      cell.value = null;
      cell.numFmt = "yyyy-mm-dd";
      cell.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF1F2937" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      cell.border = { top: { style: "thin", color: { argb: "FFCBD5E1" } }, bottom: { style: "thin", color: { argb: "FFCBD5E1" } }, left: { style: "thin", color: { argb: "FFCBD5E1" } }, right: { style: "thin", color: { argb: "FFCBD5E1" } } };
      cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      cell.note = index === 0 ? "开始日期；清空表示不限开始时间。" : "结束日期；清空表示不限结束时间。";
    });
  }
  sheet.mergeCells(5, 1, 5, 16);
  const metaCell = sheet.getCell(5, 1);
  metaCell.value = `${runtimeRows.length} 条分析记录 · ${project.columns.length} 个字段 · 选择“全部”并清空日期即可重置 · 图表为可编辑的原生 Excel 对象`;
  metaCell.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF94A3B8" } };
  sheet.getRow(1).height = 28;
  sheet.getRow(2).height = 16;
  sheet.getRow(3).height = 24;
  sheet.getRow(4).height = 28;
  sheet.getRow(5).height = 20;
  sheet.getRow(6).height = 9;

  const columnRange = (column?: ColumnDefinition) => {
    if (!column) return null;
    const letter = excelLetters(project.columns.findIndex((item) => item.id === column.id) + 1);
    return `$${letter}$${detailsStartRow}:$${letter}$${detailsEndRow}`;
  };
  const sumValue = (column?: ColumnDefinition) => runtimeRows.reduce((total, row) => total + exportNumber(row.values[column?.id || ""]), 0);
  const primaryTotal = sumValue(primary);
  const secondaryTotal = sumValue(secondary);
  const quantityTotal = quantity ? sumValue(quantity) : runtimeRows.length;
  const primaryRange = columnRange(primary);
  const secondaryRange = columnRange(secondary);
  const quantityRange = columnRange(quantity);
  const firstDataRange = `$A$${detailsStartRow}:$A$${detailsEndRow}`;
  const categoryFilterRange = columnRange(categoryFilter);
  const dateFilterRange = columnRange(dateFilter);
  const filterConditions = [`--(ROW(${firstDataRange})>=${detailsStartRow})`];
  if (categoryFilterRange) filterConditions.push(`--((($C$4="")+($C$4="全部")+(${categoryFilterRange}=$C$4))>0)`);
  if (dateFilterRange) {
    filterConditions.push(`--((($I$4="")+(${dateFilterRange}>=$I$4))>0)`);
    filterConditions.push(`--((($M$4="")+(${dateFilterRange}<=$M$4))>0)`);
  }
  const filterMultiplier = filterConditions.join("*");
  const filteredSumFormula = (range: string | null) => range ? `SUMPRODUCT(${filterMultiplier},${range})` : "0";
  const filteredCountFormula = () => `SUMPRODUCT(${filterMultiplier})`;

  if (categoryFilter) {
    const options = [...new Set(runtimeRows.map((row) => String(row.values[categoryFilter.id] ?? "").trim()).filter(Boolean))]
      .sort((left, right) => left.localeCompare(right, "zh-CN"))
      .slice(0, 1000);
    sheet.getCell("X3").value = `${categoryFilter.name}筛选项`;
    ["全部", ...options].forEach((value, index) => {
      sheet.getCell(4 + index, 24).value = value;
    });
    sheet.getCell("C4").dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`'${sheetName.replaceAll("'", "''")}'!$X$4:$X$${4 + options.length}`],
      showErrorMessage: true,
      errorTitle: "无效筛选值",
      error: "请从下拉列表中选择筛选值。",
    };
    sheet.getCell("C4").note = `选择${categoryFilter.name}后，KPI、图表和排行榜会一起更新。`;
  }
  if (dateFilter) {
    const dateValues = runtimeRows
      .map((row) => new Date(`${String(row.values[dateFilter.id] ?? "").slice(0, 10)}T00:00:00Z`))
      .filter((date) => !Number.isNaN(date.getTime()));
    if (dateValues.length) {
      const earliest = new Date(Math.min(...dateValues.map((date) => date.getTime())));
      const latest = new Date(Math.max(...dateValues.map((date) => date.getTime())));
      [sheet.getCell("I4"), sheet.getCell("M4")].forEach((cell) => {
        cell.dataValidation = {
          type: "date",
          operator: "between",
          allowBlank: true,
          formulae: [earliest, latest],
          showErrorMessage: true,
          errorTitle: "日期超出数据范围",
          error: `请输入 ${earliest.toISOString().slice(0, 10)} 至 ${latest.toISOString().slice(0, 10)} 之间的日期。`,
        };
      });
    }
  }
  const cards: Array<{ label: string; formula: string; result: number; format: string; fill: string; color: string }> = secondary
    ? [
      { label: `当前筛选${primary?.name || "指标"}`, formula: filteredSumFormula(primaryRange), result: primaryTotal, format: dashboardNumberFormat(primary), fill: "EFF6FF", color: "2563EB" },
      { label: secondaryMode === "rate" ? `当前筛选平均${secondary.name}` : `当前筛选${secondary.name}`, formula: secondaryMode === "rate" ? `IFERROR(${filteredSumFormula(secondaryRange)}/${filteredCountFormula()},0)` : filteredSumFormula(secondaryRange), result: secondaryMode === "rate" ? (runtimeRows.length ? secondaryTotal / runtimeRows.length : 0) : secondaryTotal, format: secondaryMode === "rate" ? "0.0%" : dashboardNumberFormat(secondary), fill: "ECFDF5", color: "059669" },
      { label: quantity ? `当前筛选${quantity.name}` : "当前筛选记录数", formula: quantityRange ? filteredSumFormula(quantityRange) : filteredCountFormula(), result: quantityTotal, format: "#,##0", fill: "FFF7ED", color: "D97706" },
      { label: secondaryMode === "margin" ? `${secondary.name}率` : `平均${primary?.name || "指标"}`, formula: secondaryMode === "margin" ? "IFERROR($E$9/$A$9,0)" : `IFERROR($A$9/${filteredCountFormula()},0)`, result: secondaryMode === "margin" ? (primaryTotal ? secondaryTotal / primaryTotal : 0) : (runtimeRows.length ? primaryTotal / runtimeRows.length : 0), format: secondaryMode === "margin" ? "0.0%" : dashboardNumberFormat(primary), fill: "F5F3FF", color: "7C3AED" },
    ]
    : [
      { label: `当前筛选${primary?.name || "指标"}`, formula: filteredSumFormula(primaryRange), result: primaryTotal, format: dashboardNumberFormat(primary), fill: "EFF6FF", color: "2563EB" },
      { label: quantity ? `当前筛选${quantity.name}` : "当前筛选记录数", formula: quantityRange ? filteredSumFormula(quantityRange) : filteredCountFormula(), result: quantityTotal, format: "#,##0", fill: "ECFDF5", color: "059669" },
      { label: "当前筛选记录数", formula: filteredCountFormula(), result: runtimeRows.length, format: "#,##0", fill: "FFF7ED", color: "D97706" },
      { label: `平均${primary?.name || "指标"}`, formula: "IFERROR($A$9/$I$9,0)", result: runtimeRows.length ? primaryTotal / runtimeRows.length : 0, format: dashboardNumberFormat(primary), fill: "F5F3FF", color: "7C3AED" },
    ];
  cards.forEach((card, index) => {
    const startColumn = 1 + index * 4;
    const endColumn = startColumn + 3;
    stylePanelRange(sheet, 7, startColumn, 11, endColumn, card.fill);
    sheet.mergeCells(7, startColumn, 8, endColumn);
    sheet.mergeCells(9, startColumn, 11, endColumn);
    const labelCell = sheet.getCell(7, startColumn);
    labelCell.value = card.label;
    labelCell.font = { name: "Microsoft YaHei", bold: true, size: 10, color: { argb: "FF64748B" } };
    labelCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    const valueCell = sheet.getCell(9, startColumn);
    valueCell.value = { formula: card.formula, result: card.result };
    valueCell.numFmt = card.format;
    valueCell.font = { name: "Microsoft YaHei", bold: true, size: 20, color: { argb: `FF${card.color}` } };
    valueCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  });
  for (let row = 7; row <= 11; row += 1) sheet.getRow(row).height = 18;
  sheet.getRow(12).height = 9;

  stylePanelRange(sheet, 13, 1, 31, 8, "FFFFFF");
  stylePanelRange(sheet, 13, 9, 31, 16, "FFFFFF");
  stylePanelRange(sheet, 33, 1, 50, 8, "FFFFFF");
  stylePanelRange(sheet, 33, 9, 50, 16, "FFFFFF");

  const grouped = new Map<string, DashboardGroup>();
  runtimeRows.forEach((row, index) => {
    const rawDimension = dimension ? row.values[dimension.id] : `记录 ${index + 1}`;
    const key = String(rawDimension ?? "").trim();
    const name = key || "未分组";
    const current = grouped.get(key) || { key, name, primary: 0, secondary: 0, quantity: 0, count: 0 };
    current.primary += exportNumber(row.values[primary?.id || ""]);
    current.secondary += exportNumber(row.values[secondary?.id || ""]);
    current.quantity += quantity ? exportNumber(row.values[quantity.id]) : 1;
    current.count += 1;
    grouped.set(key, current);
  });
  const allGroups = [...grouped.values()].sort((left, right) => right.primary - left.primary);
  const groups = (allGroups.length ? allGroups : [{ key: "", name: "暂无数据", primary: 0, secondary: 0, quantity: 0, count: 0 }]).slice(0, 8);
  const secondaryDisplayValue = (group: DashboardGroup) => {
    if (secondaryMode === "margin") return group.primary ? group.secondary / group.primary : 0;
    if (secondaryMode === "rate") return group.count ? group.secondary / group.count : 0;
    return group.secondary;
  };
  const points: NativeChartPoint[] = groups.map((group) => ({
    name: group.name,
    primary: group.primary,
    secondary: secondaryDisplayValue(group),
    size: group.count,
  }));

  const helperHeaderRow = 3;
  const helperStartRow = helperHeaderRow + 1;
  const dimensionRange = columnRange(dimension);
  const helperSecondaryName = secondary
    ? secondaryMode === "margin" ? `${secondary.name}率` : secondary.name
    : "指标二";
  const helperHeaders = [dimension?.name || "分类", primary?.name || "指标", helperSecondaryName, "记录数"];
  helperHeaders.forEach((value, index) => {
    const cell = sheet.getCell(helperHeaderRow, 20 + index);
    cell.value = value;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
  });
  sheet.getCell("Y3").value = secondary ? `${secondary.name}筛选合计` : quantity ? `${quantity.name}筛选合计` : "筛选合计";
  groups.forEach((group, index) => {
    const row = helperStartRow + index;
    const categoryCell = sheet.getCell(row, 20);
    categoryCell.value = group.name;
    const criterion = group.key ? `$T$${row}` : '""';
    const groupMultiplier = dimensionRange ? `${filterMultiplier}*--(${dimensionRange}=${criterion})` : filterMultiplier;
    const groupSumFormula = (range: string | null) => range ? `SUMPRODUCT(${groupMultiplier},${range})` : "0";
    const primaryFormula = groupSumFormula(primaryRange);
    const groupCountFormula = `SUMPRODUCT(${groupMultiplier})`;
    const secondaryTotalFormula = groupSumFormula(secondaryRange);
    const secondaryFormula = secondaryMode === "margin"
      ? `IFERROR(${secondaryTotalFormula}/${primaryFormula},0)`
      : secondaryMode === "rate"
        ? `IFERROR(${secondaryTotalFormula}/${groupCountFormula},0)`
        : secondaryTotalFormula;
    [
      { column: 21, formula: primaryFormula, result: group.primary, format: dashboardNumberFormat(primary) },
      { column: 22, formula: secondaryFormula, result: secondaryDisplayValue(group), format: secondaryMode === "margin" || secondaryMode === "rate" ? "0.0%" : dashboardNumberFormat(secondary) },
      { column: 23, formula: groupCountFormula, result: group.count, format: "#,##0" },
      { column: 25, formula: secondary ? secondaryTotalFormula : groupSumFormula(quantityRange), result: secondary ? group.secondary : group.quantity, format: secondary ? dashboardNumberFormat(secondary) : "#,##0" },
    ].forEach((item) => {
      const cell = sheet.getCell(row, item.column);
      cell.value = { formula: item.formula, result: item.result };
      cell.numFmt = item.format;
    });
  });

  sheet.mergeCells(33, 9, 33, 16);
  const rankingTitle = sheet.getCell(33, 9);
  rankingTitle.value = `${dimension?.name || "分类"}绩效排行榜`;
  rankingTitle.font = { name: "Microsoft YaHei", bold: true, size: 12, color: { argb: "FF1F2937" } };
  rankingTitle.alignment = { vertical: "middle", indent: 1 };
  sheet.mergeCells(34, 9, 34, 16);
  const rankingMeta = sheet.getCell(34, 9);
  rankingMeta.value = `按${primary?.name || "指标"}从高到低，最多展示 8 项`;
  rankingMeta.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF94A3B8" } };
  rankingMeta.alignment = { indent: 1 };
  const rankingHeaderRow = 35;
  [
    { from: 9, to: 10, label: dimension?.name || "分类" },
    { from: 11, to: 12, label: primary?.name || "指标" },
    { from: 13, to: 14, label: secondary?.name || quantity?.name || "数量" },
    { from: 15, to: 15, label: secondary ? helperSecondaryName : "占比" },
    { from: 16, to: 16, label: "记录数" },
  ].forEach((header) => {
    if (header.from !== header.to) sheet.mergeCells(rankingHeaderRow, header.from, rankingHeaderRow, header.to);
    const cell = sheet.getCell(rankingHeaderRow, header.from);
    cell.value = header.label;
    cell.font = { name: "Microsoft YaHei", bold: true, size: 9, color: { argb: "FF475569" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    cell.alignment = { vertical: "middle", horizontal: header.from === 9 ? "left" : "right" };
  });
  groups.forEach((group, index) => {
    const row = rankingHeaderRow + index + 1;
    sheet.mergeCells(row, 9, row, 10);
    sheet.mergeCells(row, 11, row, 12);
    sheet.mergeCells(row, 13, row, 14);
    const helperRow = helperStartRow + index;
    const values = [
      { column: 9, formula: `$T$${helperRow}`, result: group.name, format: "@" },
      { column: 11, formula: `$U$${helperRow}`, result: group.primary, format: dashboardNumberFormat(primary) },
      { column: 13, formula: `$Y$${helperRow}`, result: secondary ? group.secondary : group.quantity, format: secondary ? dashboardNumberFormat(secondary) : "#,##0" },
      { column: 15, formula: secondary ? `$V$${helperRow}` : primaryTotal ? `IFERROR($U$${helperRow}/$A$9,0)` : "0", result: secondary ? secondaryDisplayValue(group) : (primaryTotal ? group.primary / primaryTotal : 0), format: secondaryMode === "value" ? dashboardNumberFormat(secondary) : "0.0%" },
      { column: 16, formula: `$W$${helperRow}`, result: group.count, format: "#,##0" },
    ];
    values.forEach((item) => {
      const cell = sheet.getCell(row, item.column);
      cell.value = { formula: item.formula, result: item.result };
      cell.numFmt = item.format;
      cell.font = { name: "Microsoft YaHei", size: 9, color: { argb: item.column === 13 && secondary && Number(item.result) < 0 ? "FFDC2626" : "FF334155" } };
      cell.alignment = { vertical: "middle", horizontal: item.column === 9 ? "left" : "right" };
    });
    for (let column = 9; column <= 16; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: index === 0 ? "FFEFF6FF" : "FFFFFFFF" } };
      cell.border = { bottom: { style: "thin", color: { argb: "FFE2E8F0" } } };
    }
    sheet.getRow(row).height = 21;
  });

  sheet.mergeCells(52, 1, 52, 16);
  const detailTitle = sheet.getCell(52, 1);
  detailTitle.value = "交易明细";
  detailTitle.font = { name: "Microsoft YaHei", bold: true, size: 12, color: { argb: "FF1F2937" } };
  detailTitle.alignment = { vertical: "middle" };
  sheet.mergeCells(53, 1, 53, 16);
  const detailMeta = sheet.getCell(53, 1);
  detailMeta.value = `共 ${runtimeRows.length} 条分析记录；总计、小计和空白占位行不会进入图表聚合。`;
  detailMeta.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF94A3B8" } };
  stylePanelRange(sheet, 52, 1, dashboardEndRow, Math.min(16, Math.max(1, project.columns.length)), "FFFFFF");
  project.columns.forEach((column, index) => {
    const cell = sheet.getCell(detailsHeaderRow, index + 1);
    cell.value = column.name;
    cell.font = { name: "Microsoft YaHei", bold: true, size: 9, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
    cell.alignment = { vertical: "middle", horizontal: column.dataType === "number" ? "right" : "left" };
  });
  sheet.getRow(detailsHeaderRow).height = 24;
  runtimeRows.forEach((row, rowIndex) => {
    const excelRow = detailsStartRow + rowIndex;
    project.columns.forEach((column, columnIndex) => {
      const cell = sheet.getCell(excelRow, columnIndex + 1);
      cell.value = excelValue(row.values[column.id], column.dataType);
      if (column.dataType === "number") cell.numFmt = dashboardNumberFormat(column);
      if (column.dataType === "date") cell.numFmt = "yyyy-mm-dd";
      cell.font = { name: "Microsoft YaHei", size: 9, color: { argb: column.dataType === "number" && /利润|收益/i.test(column.name) && exportNumber(row.values[column.id]) < 0 ? "FFDC2626" : "FF334155" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowIndex % 2 ? "FFF8FAFC" : "FFFFFFFF" } };
      cell.alignment = { vertical: "middle", horizontal: column.dataType === "number" ? "right" : "left" };
      cell.border = { bottom: { style: "thin", color: { argb: "FFE2E8F0" } } };
    });
    sheet.getRow(excelRow).height = 21;
  });
  if (project.columns.length) {
    sheet.autoFilter = { from: { row: detailsHeaderRow, column: 1 }, to: { row: detailsHeaderRow, column: project.columns.length } };
  }
  sheet.pageSetup.printArea = `A1:P${dashboardEndRow}`;

  const sourceEndRow = helperStartRow + points.length - 1;
  const commonSpec = {
    sheetName,
    sourceStartRow: helperStartRow,
    sourceEndRow,
    categoryColumn: "T",
    primaryColumn: "U",
    secondaryColumn: "V",
    sizeColumn: "W",
    points,
    primaryName: primary?.name || "指标",
    primaryColor: "3B82F6",
    primaryNumberFormat: dashboardNumberFormat(primary),
    plotVisibleOnly: false,
  };
  const comparisonTitle = secondary
    ? `${dimension?.name || "分类"}${primary?.name || "指标"}与${helperSecondaryName}`
    : `${dimension?.name || "分类"}${primary?.name || "指标"}对比`;
  const comparisonSpec: NativeChartSpec = {
    ...commonSpec,
    title: comparisonTitle,
    chartType: "bar",
    anchor: { fromRow: 12, fromCol: 0, toRow: 31, toCol: 8 },
    comboLine: Boolean(secondary),
    secondaryName: secondary ? helperSecondaryName : undefined,
    secondaryColor: "10B981",
    secondaryNumberFormat: secondaryMode === "margin" || secondaryMode === "rate" ? "0.0%" : dashboardNumberFormat(secondary),
    showLegend: Boolean(secondary),
  };
  const shareSpec: NativeChartSpec = {
    ...commonSpec,
    title: `${dimension?.name || "分类"}${primary?.name || "指标"}占比`,
    chartType: "pie",
    anchor: { fromRow: 12, fromCol: 8, toRow: 31, toCol: 16 },
    doughnut: true,
    palette: DASHBOARD_COLORS,
    showLegend: true,
  };
  const rankingSpec: NativeChartSpec = {
    ...commonSpec,
    title: `${dimension?.name || "分类"}${primary?.name || "指标"}排名`,
    chartType: "funnel",
    anchor: { fromRow: 32, fromCol: 0, toRow: 50, toCol: 8 },
    primaryColor: "F59E0B",
    showLegend: false,
  };
  return [comparisonSpec, shareSpec, rankingSpec];
}

function addCardDashboardSheet(project: Project, view: ViewDefinition, runtimeRows: RuntimeRow[], workbook: ExcelJS.Workbook): NativeChartSpec[] {
  const { dimension, primary, secondary, secondaryMode, quantity } = dashboardColumns(project, view);
  if (!dimension) return addDashboardSheet(project, view, runtimeRows, workbook, true);
  const sheetName = uniqueSheetName(workbook, `大屏-${view.name}`);
  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 6, showGridLines: false, zoomScale: 85 }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9, margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 } },
  });
  for (let column = 1; column <= 17; column += 1) sheet.getColumn(column).width = column === 9 ? 2.5 : 12;
  const cardColumns = [1, 10];
  const groups = [...new Map(runtimeRows.map((row) => [String(row.values[dimension.id] ?? "未填写").trim() || "未填写", row])).keys()];
  const cardCount = Math.min(groups.length, 40);
  const detailHeaderRow = 9 + Math.ceil(cardCount / 2) * 10 + 2;
  const detailStartRow = detailHeaderRow + 1;
  const detailEndRow = Math.max(detailStartRow, detailStartRow + runtimeRows.length - 1);
  const detailColumn = (column?: ColumnDefinition) => column ? excelLetters(project.columns.findIndex((item) => item.id === column.id) + 1) : "A";
  const dimensionLetter = detailColumn(dimension);
  const primaryLetter = detailColumn(primary);
  const secondaryLetter = detailColumn(secondary);
  const quantityLetter = detailColumn(quantity);
  const range = (letter: string) => `$${letter}$${detailStartRow}:$${letter}$${detailEndRow}`;
  const sumFormula = (letter: string, criteria: string) => `SUMIF(${range(dimensionLetter)},${criteria},${range(letter)})`;
  const countFormula = (criteria: string) => `COUNTIF(${range(dimensionLetter)},${criteria})`;

  for (let row = 1; row <= detailEndRow; row += 1) {
    for (let column = 1; column <= 17; column += 1) sheet.getCell(row, column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F5F7" } };
  }
  stylePanelRange(sheet, 1, 1, 5, 17, "FFFFFF");
  sheet.mergeCells(1, 1, 2, 17);
  const title = sheet.getCell(1, 1);
  title.value = view.name || "卡牌界面";
  title.font = { name: "Microsoft YaHei", bold: true, size: 20, color: { argb: "FF111827" } };
  title.alignment = { vertical: "middle", horizontal: "left" };
  sheet.mergeCells(3, 1, 3, 17);
  const description = sheet.getCell(3, 1);
  description.value = view.draftDescription || "基于当前表格数据生成的卡牌仪表板";
  description.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF64748B" } };
  description.alignment = { vertical: "middle", wrapText: true };
  sheet.mergeCells(4, 1, 4, 17);
  const overview = sheet.getCell(4, 1);
  const totalPrimary = runtimeRows.reduce((total, row) => total + exportNumber(row.values[primary?.id || ""]), 0);
  overview.value = `${dimension.name} ${cardCount} 个 · ${primary?.name || "指标"}总计 ${totalPrimary.toLocaleString("zh-CN")} · ${runtimeRows.length} 条记录`;
  overview.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF64748B" } };
  sheet.getRow(1).height = 28;
  sheet.getRow(2).height = 16;
  sheet.getRow(3).height = 24;
  sheet.getRow(4).height = 22;
  sheet.getRow(5).height = 9;

  const detailHeader = sheet.getRow(detailHeaderRow);
  project.columns.forEach((column, index) => {
    const cell = detailHeader.getCell(index + 1);
    cell.value = column.name;
    cell.font = { name: "Microsoft YaHei", bold: true, size: 9, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
    cell.alignment = { vertical: "middle", horizontal: column.dataType === "number" ? "right" : "left" };
  });
  detailHeader.height = 24;
  runtimeRows.forEach((row, rowIndex) => {
    const excelRow = detailStartRow + rowIndex;
    project.columns.forEach((column, columnIndex) => {
      const cell = sheet.getCell(excelRow, columnIndex + 1);
      cell.value = excelValue(row.values[column.id], column.dataType);
      if (column.dataType === "number") cell.numFmt = dashboardNumberFormat(column);
      if (column.dataType === "date") cell.numFmt = "yyyy-mm-dd";
      cell.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF334155" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowIndex % 2 ? "FFF8FAFC" : "FFFFFFFF" } };
      cell.border = { bottom: { style: "thin", color: { argb: "FFE2E8F0" } } };
      cell.alignment = { vertical: "middle", horizontal: column.dataType === "number" ? "right" : "left" };
    });
    sheet.getRow(excelRow).height = 21;
  });
  if (project.columns.length) sheet.autoFilter = { from: { row: detailHeaderRow, column: 1 }, to: { row: detailHeaderRow, column: project.columns.length } };
  project.columns.forEach((column, index) => {
    const values = runtimeRows.slice(0, 200).map((row) => String(row.values[column.id] ?? ""));
    sheet.getColumn(index + 1).width = Math.min(28, Math.max(11, column.name.length + 3, ...values.map((value) => value.length + 2)));
  });

  groups.slice(0, cardCount).forEach((groupName, index) => {
    const cardRow = 7 + Math.floor(index / 2) * 10;
    const cardColumn = cardColumns[index % 2];
    const cardEndColumn = cardColumn + 7;
    const criteria = `$${excelLetters(cardColumn)}$${cardRow}`;
    const members = runtimeRows.filter((row) => (String(row.values[dimension.id] ?? "").trim() || "未填写") === groupName);
    const primaryValue = members.reduce((total, row) => total + exportNumber(row.values[primary?.id || ""]), 0);
    const secondaryValue = members.reduce((total, row) => total + exportNumber(row.values[secondary?.id || ""]), 0);
    const quantityValue = quantity ? members.reduce((total, row) => total + exportNumber(row.values[quantity.id]), 0) : members.length;
    stylePanelRange(sheet, cardRow, cardColumn, cardRow + 8, cardEndColumn, index === 0 ? "EFF6FF" : "FFFFFF");
    sheet.mergeCells(cardRow, cardColumn, cardRow, cardEndColumn);
    sheet.mergeCells(cardRow + 1, cardColumn, cardRow + 1, cardEndColumn);
    sheet.mergeCells(cardRow + 2, cardColumn, cardRow + 2, cardEndColumn);
    sheet.mergeCells(cardRow + 3, cardColumn, cardRow + 3, cardEndColumn);
    sheet.mergeCells(cardRow + 4, cardColumn, cardRow + 4, cardColumn + 3);
    sheet.mergeCells(cardRow + 4, cardColumn + 4, cardRow + 4, cardEndColumn);
    sheet.mergeCells(cardRow + 5, cardColumn, cardRow + 5, cardColumn + 3);
    sheet.mergeCells(cardRow + 5, cardColumn + 4, cardRow + 5, cardEndColumn);
    sheet.mergeCells(cardRow + 6, cardColumn, cardRow + 6, cardEndColumn);
    sheet.mergeCells(cardRow + 7, cardColumn, cardRow + 7, cardEndColumn);
    sheet.mergeCells(cardRow + 8, cardColumn, cardRow + 8, cardEndColumn);
    const nameCell = sheet.getCell(cardRow, cardColumn);
    nameCell.value = groupName;
    nameCell.font = { name: "Microsoft YaHei", bold: true, size: 15, color: { argb: "FF1F2937" } };
    nameCell.alignment = { vertical: "middle", horizontal: "left" };
    const countCell = sheet.getCell(cardRow + 1, cardColumn);
    countCell.value = `${members.length} 条记录`;
    countCell.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF64748B" } };
    const metricLabel = sheet.getCell(cardRow + 2, cardColumn);
    metricLabel.value = primary?.name || "指标";
    metricLabel.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF64748B" } };
    const valueCell = sheet.getCell(cardRow + 3, cardColumn);
    valueCell.value = { formula: sumFormula(primaryLetter, criteria), result: primaryValue };
    valueCell.numFmt = dashboardNumberFormat(primary);
    valueCell.font = { name: "Microsoft YaHei", bold: true, size: 18, color: { argb: "FF111827" } };
    valueCell.alignment = { vertical: "middle", horizontal: "left" };
    const quantityLabel = sheet.getCell(cardRow + 4, cardColumn);
    quantityLabel.value = quantity?.name || "记录数";
    quantityLabel.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF64748B" } };
    const secondaryLabel = sheet.getCell(cardRow + 4, cardColumn + 4);
    secondaryLabel.value = secondary?.name || "";
    secondaryLabel.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF64748B" } };
    const quantityCell = sheet.getCell(cardRow + 5, cardColumn);
    quantityCell.value = { formula: quantity ? sumFormula(quantityLetter, criteria) : countFormula(criteria), result: quantityValue };
    quantityCell.numFmt = "#,##0";
    quantityCell.font = { name: "Microsoft YaHei", bold: true, size: 11, color: { argb: "FF334155" } };
    quantityCell.alignment = { vertical: "middle", horizontal: "left" };
    const secondaryCell = sheet.getCell(cardRow + 5, cardColumn + 4);
    secondaryCell.value = secondary ? { formula: sumFormula(secondaryLetter, criteria), result: secondaryValue } : "";
    secondaryCell.numFmt = secondary ? dashboardNumberFormat(secondary) : "General";
    secondaryCell.font = { name: "Microsoft YaHei", bold: true, size: 11, color: { argb: "FF334155" } };
    secondaryCell.alignment = { vertical: "middle", horizontal: "left" };
    const rateLabel = sheet.getCell(cardRow + 6, cardColumn);
    rateLabel.value = secondary && secondaryMode === "margin" ? `${secondary.name}率` : `平均${primary?.name || "指标"}`;
    rateLabel.font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF64748B" } };
    const rateCell = sheet.getCell(cardRow + 7, cardColumn);
    rateCell.value = secondary && secondaryMode === "margin"
      ? { formula: `IFERROR(${secondaryCell.address}/${valueCell.address},0)`, result: primaryValue ? secondaryValue / primaryValue : 0 }
      : { formula: `IFERROR(${valueCell.address}/COUNTIF(${range(dimensionLetter)},${criteria}),0)`, result: members.length ? primaryValue / members.length : 0 };
    rateCell.numFmt = secondary && secondaryMode === "margin" ? "0.0%" : dashboardNumberFormat(primary);
    rateCell.font = { name: "Microsoft YaHei", bold: true, size: 11, color: { argb: "FF059669" } };
    rateCell.alignment = { vertical: "middle", horizontal: "left" };
    const progress = sheet.getCell(cardRow + 8, cardColumn);
    progress.value = { formula: `IFERROR(${valueCell.address}/MAX(${range(primaryLetter)}),0)`, result: totalPrimary ? primaryValue / totalPrimary : 0 };
    progress.numFmt = "0%";
    progress.font = { color: { argb: "FF3B82F6" } };
    progress.alignment = { vertical: "middle", horizontal: "left" };
    sheet.addConditionalFormatting({
      ref: progress.address,
      rules: [{ type: "dataBar", priority: index + 1, cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF3B82F6" }, gradient: true, showValue: false } as ExcelJS.DataBarRuleType],
    });
    sheet.getRow(cardRow).height = 25;
    sheet.getRow(cardRow + 1).height = 18;
    sheet.getRow(cardRow + 2).height = 24;
    sheet.getRow(cardRow + 3).height = 18;
    sheet.getRow(cardRow + 6).height = 18;
    sheet.getRow(cardRow + 7).height = 20;
    sheet.getRow(cardRow + 8).height = 16;
  });
  sheet.mergeCells(detailHeaderRow - 2, 1, detailHeaderRow - 2, 17);
  sheet.getCell(detailHeaderRow - 2, 1).value = "可编辑明细数据：修改下方数据后，上方卡片指标会在 Excel 中重新计算";
  sheet.getCell(detailHeaderRow - 2, 1).font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF94A3B8" } };
  sheet.getCell(detailHeaderRow - 2, 1).alignment = { vertical: "middle" };
  sheet.pageSetup.printArea = `A1:Q${detailEndRow}`;
  return [];
}

function addExportCharts(project: Project, workbook: ExcelJS.Workbook) {
  const specs: NativeChartSpec[] = [];
  const views = project.views.filter((view) => view.chartType !== "custom") as Array<ViewDefinition & { chartType: Exclude<ChartType, "custom"> }>;
  if (views.length) {
    const sheet = workbook.addWorksheet("图表", { views: [{ state: "frozen", ySplit: 2, showGridLines: false }] });
    styleExportChartSheet(sheet);
    sheet.mergeCells(1, 1, 1, 9);
    sheet.getCell(1, 1).value = "图表汇总";
    sheet.getCell(1, 1).font = { bold: true, size: 18, color: { argb: "FF18231D" } };
    sheet.getCell(2, 1).value = "按应用标签顺序排列；右侧保留图表数据源，图表对象可在 Excel 中继续编辑。";
    sheet.mergeCells(2, 1, 2, 9);
    sheet.getCell(2, 1).font = { size: 10, color: { argb: "FF66736C" } };

    let sectionRow = 4;
    for (const view of views) {
      const sourceData = resolveViewData(project, view.sourceSheetIds);
      const sourceProject = projectForView(project, view, sourceData);
      const sourceRows = sourceData.rows;
      const dataset = exportChartDataset(sourceProject, view, sourceRows);
      const chartTitle = view.name || "图表";
      sheet.mergeCells(sectionRow, 1, sectionRow, 9);
      const titleCell = sheet.getCell(sectionRow, 1);
      titleCell.value = chartTitle;
      titleCell.font = { bold: true, size: 13, color: { argb: "FF19734A" } };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6F3EC" } };
      sheet.mergeCells(sectionRow + 1, 1, sectionRow + 1, 9);
      const metaCell = sheet.getCell(sectionRow + 1, 1);
      metaCell.value = `${view.chartType} · ${dataset.dimensionName} × ${dataset.metricName} · ${sourceRows.length} 条分析记录`;
      metaCell.font = { size: 10, color: { argb: "FF66736C" } };
      specs.push(addChartSource(sheet, view, sourceProject, sourceRows, sectionRow, 11, { fromRow: sectionRow + 1, fromCol: 0, toRow: sectionRow + 22, toCol: 9 }));
      sectionRow += Math.max(25, dataset.points.length + 4) + 3;
    }
  }
  project.views.filter((view) => view.chartType === "custom").forEach((view) => {
    const sourceData = resolveViewData(project, view.sourceSheetIds);
    const sourceProject = projectForView(project, view, sourceData);
    specs.push(...addDashboardSheet(sourceProject, view, sourceData.rows, workbook));
  });
  return specs;
}

async function worksheetPathByName(zip: JSZip, sheetName: string) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relsXml) throw new Error("Excel 工作簿结构不完整");
  const parser = new DOMParser();
  const workbookDocument = parser.parseFromString(workbookXml, "application/xml");
  const sheet = [...workbookDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "sheet")]
    .find((item) => item.getAttribute("name") === sheetName);
  const relationshipId = sheet?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
    || sheet?.getAttribute("r:id");
  if (!relationshipId) throw new Error(`找不到工作表：${sheetName}`);
  const relationshipsDocument = parser.parseFromString(relsXml, "application/xml");
  const relationship = [...relationshipsDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship")]
    .find((item) => item.getAttribute("Id") === relationshipId);
  const target = relationship?.getAttribute("Target");
  if (!target) throw new Error(`找不到工作表关系：${sheetName}`);
  return target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.\//, "")}`;
}

async function injectNativeCharts(buffer: Uint8Array, specs: NativeChartSpec[]) {
  if (!specs.length) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  const grouped = new Map<string, NativeChartSpec[]>();
  specs.forEach((spec) => {
    const group = grouped.get(spec.sheetName) || [];
    group.push(spec);
    grouped.set(spec.sheetName, group);
  });
  let chartIndex = 0;
  let drawingIndex = 0;
  const contentOverrides: string[] = [];
  for (const [sheetName, sheetSpecs] of grouped) {
    drawingIndex += 1;
    const worksheetPath = await worksheetPathByName(zip, sheetName);
    const worksheet = await zip.file(worksheetPath)?.async("string");
    if (!worksheet) throw new Error(`无法读取工作表：${sheetName}`);
    const slash = worksheetPath.lastIndexOf("/");
    const relationshipPath = `${worksheetPath.slice(0, slash)}/_rels/${worksheetPath.slice(slash + 1)}.rels`;
    const existingRelationships = await zip.file(relationshipPath)?.async("string");
    const usedRelationshipIds = new Set<string>();
    existingRelationships?.replace(/\bId="([^"]+)"/g, (_match, id: string) => {
      usedRelationshipIds.add(id);
      return _match;
    });
    let drawingRelationshipIndex = 1;
    while (usedRelationshipIds.has(`rId${drawingRelationshipIndex}`)) drawingRelationshipIndex += 1;
    const drawingId = `rId${drawingRelationshipIndex}`;
    const drawingMarkup = `<drawing r:id="${drawingId}"/>`;
    const trailingWorksheetNode = /<(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/;
    const withDrawing = trailingWorksheetNode.test(worksheet)
      ? worksheet.replace(trailingWorksheetNode, `${drawingMarkup}$&`)
      : worksheet.replace("</worksheet>", `${drawingMarkup}</worksheet>`);
    zip.file(worksheetPath, withDrawing);
    const drawingRelationship = `<Relationship Id="${drawingId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIndex}.xml"/>`;
    zip.file(
      relationshipPath,
      existingRelationships
        ? existingRelationships.replace("</Relationships>", `${drawingRelationship}</Relationships>`)
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRelationship}</Relationships>`,
    );
    const chartIndices: number[] = [];
    sheetSpecs.forEach((spec) => {
      chartIndex += 1;
      chartIndices.push(chartIndex);
      zip.file(`xl/charts/chart${chartIndex}.xml`, nativeChartXml(spec));
      contentOverrides.push(`<Override PartName="/xl/charts/chart${chartIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
    });
    zip.file(`xl/drawings/drawing${drawingIndex}.xml`, drawingXml(sheetSpecs));
    zip.file(`xl/drawings/_rels/drawing${drawingIndex}.xml.rels`, drawingRelsXml(chartIndices));
    contentOverrides.push(`<Override PartName="/xl/drawings/drawing${drawingIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
  }
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  if (!contentTypes) throw new Error("Excel 内容类型文件缺失");
  zip.file("[Content_Types].xml", contentTypes.replace("</Types>", `${contentOverrides.join("")}</Types>`));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export async function projectToExcel(project: Project, runtimeRows: RuntimeRow[]) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Excel";
  workbook.created = new Date(project.createdAt);
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  const usedNames = new Set<string>();
  const safeSheetName = (name: string) => {
    const base = name.replace(/[\\/*?:\[\]]/g, " ").trim().slice(0, 31) || "数据";
    let candidate = base;
    let index = 2;
    while (usedNames.has(candidate)) {
      const suffix = ` (${index})`;
      candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
      index += 1;
    }
    usedNames.add(candidate);
    return candidate;
  };
  const sourceSheets = project.sheets?.length ? project.sheets.map((sheet) => sheet.id === project.activeSheetId ? {
    ...sheet,
    columns: project.columns,
    rows: project.rows,
    tableRule: project.tableRule,
  } : sheet) : [{
    id: project.activeSheetId || "sheet_data",
    name: "数据",
    columns: project.columns,
    rows: project.rows,
    tableRule: project.tableRule,
  }];
  sourceSheets.forEach((sourceSheet) => {
    const rows = sourceSheet.id === project.activeSheetId
      ? runtimeRows
      : calculateRuntimeRows(sourceSheet.rows, sourceSheet.columns);
    const worksheet = workbook.addWorksheet(safeSheetName(sourceSheets.length === 1 ? "数据" : sourceSheet.name), { views: [{ state: "frozen", ySplit: 1 }] });
    worksheet.addRow(sourceSheet.columns.map((column) => column.name));
    rows.forEach((row) => {
      worksheet.addRow(sourceSheet.columns.map((column) => excelValue(row.values[column.id], column.dataType)));
    });
    const header = worksheet.getRow(1);
    header.height = 24;
    header.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF19734A" } };
      cell.alignment = { vertical: "middle" };
    });
    sourceSheet.columns.forEach((column, index) => {
      const excelColumn = worksheet.getColumn(index + 1);
      const longest = Math.max(column.name.length, ...rows.slice(0, 200).map((row) => String(row.values[column.id] ?? "").length));
      excelColumn.width = Math.min(32, Math.max(11, longest + 3));
      if (column.dataType === "number") excelColumn.numFmt = "#,##0.00";
      if (column.dataType === "date") excelColumn.numFmt = "yyyy-mm-dd";
    });
    if (sourceSheet.columns.length) worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: sourceSheet.columns.length } };
    worksheet.eachRow((row, index) => {
      if (index > 1) row.alignment = { vertical: "middle" };
    });
  });

  const nativeCharts = addExportCharts(project, workbook);

  const metadata = workbook.addWorksheet(META_SHEET);
  metadata.state = "veryHidden";
  const serialized = serializeProject(project);
  for (let index = 0; index < serialized.length; index += META_CHUNK_SIZE) {
    metadata.addRow([serialized.slice(index, index + META_CHUNK_SIZE)]);
  }

  const output = new Uint8Array(await workbook.xlsx.writeBuffer());
  const nativeOutput = await injectNativeCharts(output, nativeCharts);
  const finalBuffer = new Uint8Array(nativeOutput.byteLength);
  finalBuffer.set(nativeOutput);
  return new Blob([finalBuffer.buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
