import { calculateRuntimeRows, filterAnalysisRows } from "./core";
import type {
  ColumnDefinition,
  Project,
  RuntimeRow,
  SheetDefinition,
  SheetRelation,
  ViewDefinition,
} from "./types";

export type ViewSourceSheet = {
  id: string;
  name: string;
  columns: ColumnDefinition[];
  rows: RuntimeRow[];
  allRows: RuntimeRow[];
};

export type ResolvedViewData = {
  sourceSheetIds: string[];
  columns: ColumnDefinition[];
  rows: RuntimeRow[];
  allRows: RuntimeRow[];
  sheets: ViewSourceSheet[];
  relations: SheetRelation[];
};

function projectSheets(project: Project): SheetDefinition[] {
  if (project.sheets?.length) {
    const activeId = project.activeSheetId || project.sheets[0].id;
    return project.sheets.map((sheet) => sheet.id === activeId ? {
      ...sheet,
      columns: project.columns,
      rows: project.rows,
      tableRule: project.tableRule,
    } : sheet);
  }
  return [{
    id: project.activeSheetId || "sheet_data",
    name: "数据",
    columns: project.columns,
    rows: project.rows,
    tableRule: project.tableRule,
  }];
}

export function resolveSourceSheetIds(project: Project, sourceSheetIds?: string[]) {
  const sheets = projectSheets(project);
  const available = new Set(sheets.map((sheet) => sheet.id));
  const selected = (sourceSheetIds || []).filter((id, index, ids) => available.has(id) && ids.indexOf(id) === index);
  if (selected.length) return selected;
  const activeId = project.activeSheetId && available.has(project.activeSheetId)
    ? project.activeSheetId
    : sheets[0]?.id;
  return activeId ? [activeId] : [];
}

function normalizedName(value: string) {
  return value.trim().toLocaleLowerCase().replace(/[\s_\-—（）()【】\[\].]+/g, "");
}

function normalizedValue(value: unknown) {
  if (value == null) return "";
  return String(value).trim().toLocaleLowerCase();
}

function valuesFor(rows: RuntimeRow[], columnId: string) {
  return rows
    .slice(0, 2000)
    .map((row) => normalizedValue(row.values[columnId]))
    .filter(Boolean);
}

function inferRelation(
  leftRows: RuntimeRow[],
  leftColumns: ColumnDefinition[],
  right: ViewSourceSheet,
): SheetRelation | null {
  let best: { score: number; relation: SheetRelation } | null = null;
  for (const leftColumn of leftColumns) {
    const leftValues = valuesFor(leftRows, leftColumn.id);
    if (leftValues.length < 2) continue;
    const leftSet = new Set(leftValues);
    for (const rightColumn of right.columns) {
      const rightValues = valuesFor(right.allRows, rightColumn.id);
      if (rightValues.length < 2) continue;
      const rightSet = new Set(rightValues);
      const intersection = [...leftSet].filter((value) => rightSet.has(value)).length;
      if (!intersection) continue;
      const matchRate = intersection / Math.max(1, Math.min(leftSet.size, rightSet.size));
      const leftName = normalizedName(leftColumn.name);
      const rightName = normalizedName(rightColumn.name);
      const sameName = leftName === rightName;
      const relatedName = sameName || leftName.includes(rightName) || rightName.includes(leftName);
      const keyLike = /(?:^|[^a-z])id$|id|编号|编码|代码|名称|姓名|地区|区域/.test(leftName)
        && /(?:^|[^a-z])id$|id|编号|编码|代码|名称|姓名|地区|区域/.test(rightName);
      if (matchRate < 0.15 || (!relatedName && !keyLike && (matchRate < 0.8 || intersection < 3))) continue;
      const rightUniqueness = rightSet.size / rightValues.length;
      const leftUniqueness = leftSet.size / leftValues.length;
      const leftUnique = leftUniqueness >= 0.9;
      const rightUnique = rightUniqueness >= 0.9;
      const cardinality = leftUnique
        ? (rightUnique ? "one-to-one" : "one-to-many")
        : (rightUnique ? "many-to-one" : "many-to-many");
      const score = (sameName ? 5 : relatedName ? 3 : keyLike ? 1 : 0) + matchRate * 5 + rightUniqueness;
      const relation: SheetRelation = {
        leftSheetId: leftColumn.sourceSheetId || "",
        leftSheetName: leftColumn.sourceSheetName || "",
        leftColumnId: leftColumn.id,
        leftColumnName: leftColumn.name,
        rightSheetId: right.id,
        rightSheetName: right.name,
        rightColumnId: rightColumn.id,
        rightColumnName: rightColumn.name,
        matchRate,
        cardinality,
        joined: rightUnique,
      };
      if (!best || score > best.score) best = { score, relation };
    }
  }
  return best?.relation || null;
}

function joinRows(leftRows: RuntimeRow[], rightRows: RuntimeRow[], relation: SheetRelation) {
  const rightByKey = new Map<string, RuntimeRow>();
  rightRows.forEach((row) => {
    const key = normalizedValue(row.values[relation.rightColumnId]);
    if (key && !rightByKey.has(key)) rightByKey.set(key, row);
  });
  return leftRows.map((row) => {
    const match = rightByKey.get(normalizedValue(row.values[relation.leftColumnId]));
    return match ? {
      id: row.id,
      values: { ...row.values, ...match.values },
      errors: { ...row.errors, ...match.errors },
    } : row;
  });
}

export function resolveViewData(project: Project, sourceSheetIds?: string[]): ResolvedViewData {
  const ids = resolveSourceSheetIds(project, sourceSheetIds);
  const byId = new Map(projectSheets(project).map((sheet) => [sheet.id, sheet]));
  const sheets = ids.flatMap((id) => {
    const sheet = byId.get(id);
    if (!sheet) return [];
    const columns = sheet.columns.map((column) => ({
      ...column,
      sourceSheetId: sheet.id,
      sourceSheetName: sheet.name,
    }));
    const allRows = calculateRuntimeRows(sheet.rows, sheet.columns);
    return [{
      id: sheet.id,
      name: sheet.name,
      columns,
      rows: filterAnalysisRows(allRows, sheet.columns),
      allRows,
    }];
  });
  if (!sheets.length) return { sourceSheetIds: [], columns: [], rows: [], allRows: [], sheets: [], relations: [] };

  const columns = sheets.flatMap((sheet) => sheet.columns);
  let allRows = sheets[0].allRows;
  let rows = sheets[0].rows;
  const joinedColumns = [...sheets[0].columns];
  const relations: SheetRelation[] = [];
  sheets.slice(1).forEach((sheet) => {
    const relation = inferRelation(allRows, joinedColumns, sheet);
    if (relation) {
      relations.push(relation);
      if (relation.joined) {
        allRows = joinRows(allRows, sheet.allRows, relation);
        rows = joinRows(rows, sheet.rows, relation);
      }
    }
    joinedColumns.push(...sheet.columns);
  });
  return { sourceSheetIds: ids, columns, rows, allRows, sheets, relations };
}

export function projectForView(
  project: Project,
  view: Pick<ViewDefinition, "sourceSheetIds">,
  data = resolveViewData(project, view.sourceSheetIds),
): Project {
  return {
    ...project,
    columns: data.columns,
    rows: data.allRows.map((row) => ({ id: row.id, values: row.values })),
  };
}

export function describeRelations(relations: SheetRelation[]) {
  if (!relations.length) return "未发现高可信公共键；请分别分析各工作表，禁止笛卡尔积。";
  const cardinalityNames = {
    "one-to-one": "一对一",
    "many-to-one": "多对一",
    "one-to-many": "一对多",
    "many-to-many": "多对多",
  };
  return relations.map((relation) =>
    `${relation.leftSheetName}.${relation.leftColumnName} = ${relation.rightSheetName}.${relation.rightColumnName}（${cardinalityNames[relation.cardinality]}，样例值匹配 ${Math.round(relation.matchRate * 100)}%，${relation.joined ? "已自动左关联" : "由 AI 按业务口径聚合"}）`,
  ).join("\n");
}
