import { describe, expect, it } from "vitest";
import {
  applyTableRule,
  calculateRuntimeRows,
  cleanModelCode,
  createDefaultProject,
  createGenerationState,
  filterAnalysisRows,
  parseProjectJson,
  projectFromCsv,
  serializeProject,
} from "./core";

describe("计算列运行时", () => {
  it("按列顺序计算，并隔离单行错误", () => {
    const project = createDefaultProject();
    const rows = calculateRuntimeRows(project.rows, project.columns);
    expect(rows[0].values.col_amount).toBe(3072);
    expect(rows[0].errors).toEqual({});

    const amount = project.columns.find((column) => column.id === "col_amount")!;
    amount.generation!.generatedCode = 'if (row["col_quantity"] === 42) throw new Error("bad row"); return 1;';
    const failed = calculateRuntimeRows(project.rows, project.columns);
    expect(failed[1].errors.col_amount).toBe("bad row");
    expect(failed[0].values.col_amount).toBe(1);
  });
});

describe("表格规则", () => {
  it("只接受输入集合中的行", () => {
    const project = createDefaultProject();
    const rows = calculateRuntimeRows(project.rows, project.columns);
    const filtered = applyTableRule(
      rows,
      project.columns,
      'return rows.filter(row => row["col_status"] === "进行中");',
    );
    expect(filtered).toHaveLength(3);
    expect(() => applyTableRule(rows, project.columns, 'return [{ id: "unknown" }];')).toThrow(
      "未知行",
    );
    expect(() => applyTableRule(rows, project.columns, "return 1;")).toThrow("行数组");
  });
});

describe("分析数据口径", () => {
  it("保留原表行但排除合计和空占位行", () => {
    const project = createDefaultProject();
    project.rows.push(
      { id: "summary", values: { col_task: "总计", col_owner: "-", col_status: "-", col_price: 1000, col_quantity: 0 } },
      { id: "placeholder", values: { col_task: "-", col_owner: "-", col_status: "-", col_price: 0, col_quantity: 0 } },
    );
    const allRows = calculateRuntimeRows(project.rows, project.columns);
    const analysisRows = filterAnalysisRows(allRows, project.columns);
    expect(allRows).toHaveLength(8);
    expect(analysisRows.map((row) => row.id)).not.toContain("summary");
    expect(analysisRows.map((row) => row.id)).not.toContain("placeholder");
    expect(analysisRows).toHaveLength(6);
  });
});

describe("CSV 与项目持久化", () => {
  it("解析 BOM、引号内逗号、换行和重复表头", () => {
    const csv = '\uFEFF名称,金额,名称,日期\r\n"第一项, 特别版",12.5,"说明\n第二行",2026-07-19\r\n第二项,8,普通,2026/07/20';
    const project = projectFromCsv(csv, "sample.csv");
    expect(project.name).toBe("sample");
    expect(project.columns.map((column) => column.name)).toEqual(["名称", "金额", "名称 (2)", "日期"]);
    expect(project.columns[1].dataType).toBe("number");
    expect(project.columns[3].dataType).toBe("date");
    expect(Object.values(project.rows[0].values)[0]).toBe("第一项, 特别版");
    expect(Object.values(project.rows[0].values)[2]).toBe("说明\n第二行");
  });

  it("可以完整导出和导入项目", () => {
    const project = createDefaultProject();
    const restored = parseProjectJson(serializeProject(project));
    expect(restored.name).toBe(project.name);
    expect(restored.columns).toHaveLength(project.columns.length);
    expect(restored.rows).toHaveLength(project.rows.length);
  });

  it("把旧视图迁移到 Runtime v2", () => {
    const project = createDefaultProject();
    project.views.push({
      id: "legacy_view",
      name: "旧视图",
      chartType: "custom",
      referenceImages: [],
      runtimeError: null,
      ...createGenerationState(),
    } as unknown as (typeof project.views)[number]);
    const restored = parseProjectJson(serializeProject(project));
    expect(restored.views[0]).toMatchObject({ appState: {}, runtimeVersion: 3, runtimeEpoch: 0 });
  });
});

describe("模型输出清理", () => {
  it("移除 Markdown 代码围栏", () => {
    expect(cleanModelCode("```js\nreturn 1;\n```")) .toBe("return 1;");
  });
});
