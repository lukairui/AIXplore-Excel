import ExcelJS from "exceljs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { calculateRuntimeRows, createDefaultProject, createGenerationState } from "./core";
import { projectFromExcel, projectToExcel } from "./excel";
import type { ChartType } from "./types";

describe("Excel 项目往返", () => {
  it("保留类型、计算列和图表视图元数据", async () => {
    const project = createDefaultProject();
    const chartTypes: Array<Exclude<ChartType, "custom">> = [
      "pie", "bar", "line", "area", "scatter", "bubble", "histogram", "radar", "waterfall", "funnel", "treemap",
    ];
    const viewNames = chartTypes.map((type, index) => `${String(index + 1).padStart(2, "0")}-${type}`);
    chartTypes.forEach((chartType, index) => {
      project.views.push({
        id: `view_${chartType}`,
        name: viewNames[index],
        chartType,
        referenceImages: [],
        appState: {},
        runtimeVersion: 2,
        runtimeEpoch: 0,
        runtimeError: null,
        ...createGenerationState("按任务展示金额"),
      });
    });
    project.views.push({
      id: "view_dashboard",
      name: "经营大屏",
      chartType: "custom",
      referenceImages: [],
      appState: {},
      runtimeVersion: 2,
      runtimeEpoch: 0,
      runtimeError: null,
      ...createGenerationState("按任务和金额生成经营大屏"),
    });
    const runtimeRows = calculateRuntimeRows(project.rows, project.columns);
    const blob = await projectToExcel(
      project,
      runtimeRows,
    );
    expect(blob.type).toContain("spreadsheetml");
    const buffer = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(blob);
    });
    const exportedWorkbook = new ExcelJS.Workbook();
    await exportedWorkbook.xlsx.load(buffer);
    expect((exportedWorkbook.getWorksheet("数据")?.getCell("F2").value as Date).toISOString().slice(0, 10)).toBe("2026-07-24");
    const chartSheet = exportedWorkbook.getWorksheet("图表");
    expect(chartSheet).toBeDefined();
    expect(exportedWorkbook.worksheets.filter((sheet) => sheet.name === "图表")).toHaveLength(1);
    expect(chartSheet?.getCell("A1").value).toBe("图表汇总");
    const firstColumn = (chartSheet?.getColumn(1).values || []).map((value) => String(value ?? ""));
    expect(firstColumn.filter((value) => viewNames.includes(value))).toEqual(viewNames);
    expect(chartSheet?.getCell("K4").value).toBe("分类");
    expect(chartSheet?.getCell("K5").value).toBe("门店陈列更新");
    const dashboardSheet = exportedWorkbook.getWorksheet("大屏-经营大屏");
    expect(dashboardSheet?.getCell("A1").value).toBe("经营大屏");
    expect(dashboardSheet?.getCell("A4").value).toBe("负责人筛选");
    expect(dashboardSheet?.getCell("C4").value).toBe("全部");
    expect(dashboardSheet?.getCell("G4").value).toBe("截止日期区间");
    expect(dashboardSheet?.getCell("A7").value).toBe("当前筛选金额");
    expect(dashboardSheet?.getCell("E7").value).toBe("当前筛选数量");
    expect(dashboardSheet?.getCell("I7").value).toBe("当前筛选记录数");
    expect(dashboardSheet?.getCell("M7").value).toBe("平均金额");
    expect(dashboardSheet?.getCell("A54").value).toBe("任务");
    expect(dashboardSheet?.getColumn(20).hidden).toBe(true);
    expect(dashboardSheet?.getColumn(24).hidden).toBe(true);
    expect(chartSheet?.getImages?.() || []).toHaveLength(0);
    const zip = await JSZip.loadAsync(new Uint8Array(buffer));
    const zipFiles = Object.keys(zip.files);
    expect(zipFiles.filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name))).toHaveLength(14);
    expect(zipFiles.filter((name) => /^xl\/media\//.test(name))).toHaveLength(0);
    expect(zipFiles.filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name))).toHaveLength(2);
    expect(await zip.file("xl/charts/chart1.xml")?.async("string")).toContain("<c:pieChart>");
    expect(await zip.file("xl/charts/chart12.xml")?.async("string")).toContain("<c:barChart>");
    expect(await zip.file("xl/charts/chart13.xml")?.async("string")).toContain("<c:doughnutChart>");
    expect(await zip.file("xl/charts/chart14.xml")?.async("string")).toContain('<c:barDir val="bar"/>');
    const dashboardWorksheetXml = await zip.file("xl/worksheets/sheet3.xml")?.async("string") || "";
    expect(dashboardWorksheetXml.indexOf("<drawing ")).toBeLessThan(dashboardWorksheetXml.indexOf("<legacyDrawing "));
    expect(dashboardWorksheetXml).toContain("<f>SUMPRODUCT");
    expect(dashboardWorksheetXml).not.toContain("<f>=SUMPRODUCT");
    const dashboardRelationships = await zip.file("xl/worksheets/_rels/sheet3.xml.rels")?.async("string") || "";
    expect(dashboardRelationships).toContain("/comments");
    expect(dashboardRelationships).toContain("/vmlDrawing");
    expect(dashboardRelationships).toContain("/drawing");
    const restored = await projectFromExcel(buffer, "roundtrip.xlsx");
    expect(restored.columns.find((column) => column.id === "col_amount")?.kind).toBe("computed");
    expect(restored.columns.find((column) => column.id === "col_due")?.dataType).toBe("date");
    expect(restored.views.map((view) => view.chartType)).toEqual([...chartTypes, "custom"]);
    expect(restored.rows).toHaveLength(project.rows.length);
  });

  it("按业务字段导出带利润率双轴图的原生 Excel 大屏", async () => {
    const project = createDefaultProject();
    project.name = "销售分析";
    project.columns = [
      { id: "order", name: "订单编号", kind: "manual", dataType: "text" },
      { id: "date", name: "订单日期", kind: "manual", dataType: "date" },
      { id: "region", name: "地区", kind: "manual", dataType: "text" },
      { id: "category", name: "产品类别", kind: "manual", dataType: "text" },
      { id: "sales", name: "销售额", kind: "manual", dataType: "number" },
      { id: "quantity", name: "数量", kind: "manual", dataType: "number" },
      { id: "profit", name: "利润", kind: "manual", dataType: "number" },
    ];
    project.rows = [
      { id: "r1", values: { order: "OD001", date: "2025-01-15", region: "华东", category: "电子产品", sales: 6500, quantity: 2, profit: 1300 } },
      { id: "r2", values: { order: "OD002", date: "2025-01-18", region: "华北", category: "办公用品", sales: 1200, quantity: 1, profit: 240 } },
      { id: "r3", values: { order: "OD003", date: "2025-02-02", region: "华东", category: "电子产品", sales: 2800, quantity: 1, profit: -150 } },
      { id: "r4", values: { order: "OD004", date: "2025-02-12", region: "华南", category: "服装鞋帽", sales: 450, quantity: 3, profit: 120 } },
    ];
    project.views = [{
      id: "sales_dashboard",
      name: "各地区销售情况对比分析",
      chartType: "custom",
      referenceImages: [],
      appState: {},
      runtimeVersion: 2,
      runtimeEpoch: 0,
      runtimeError: null,
      ...createGenerationState("按地区对比销售额、利润、利润率、数量，并展示排行榜和交易明细"),
    }];
    const blob = await projectToExcel(project, calculateRuntimeRows(project.rows, project.columns));
    const buffer = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(blob);
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const dashboard = workbook.getWorksheet("大屏-各地区销售情况对比分析");
    expect(dashboard?.getCell("A4").value).toBe("产品类别筛选");
    expect(dashboard?.getCell("C4").value).toBe("全部");
    expect(dashboard?.getCell("C4").dataValidation).toMatchObject({ type: "list", allowBlank: true });
    expect(dashboard?.getCell("G4").value).toBe("订单日期区间");
    expect(dashboard?.getCell("I4").dataValidation).toMatchObject({ type: "date", operator: "between", allowBlank: true });
    expect(dashboard?.getCell("A7").value).toBe("当前筛选销售额");
    expect(dashboard?.getCell("E7").value).toBe("当前筛选利润");
    expect(dashboard?.getCell("I7").value).toBe("当前筛选数量");
    expect(dashboard?.getCell("M7").value).toBe("利润率");
    expect(dashboard?.getCell("A9").value).toMatchObject({ result: 10950 });
    expect(String((dashboard?.getCell("A9").value as ExcelJS.CellFormulaValue)?.formula)).toContain("SUMPRODUCT");
    expect(String((dashboard?.getCell("A9").value as ExcelJS.CellFormulaValue)?.formula)).toContain('$D$55:$D$58=$C$4');
    expect(dashboard?.getCell("E9").value).toMatchObject({ result: 1510 });
    expect(dashboard?.getCell("M9").value).toMatchObject({ formula: "IFERROR($E$9/$A$9,0)" });
    expect((dashboard?.getCell("B55").value as Date).toISOString().slice(0, 10)).toBe("2025-01-15");
    expect(dashboard?.getCell("T4").value).toBe("华东");
    expect(dashboard?.getCell("U4").value).toMatchObject({ result: 9300 });
    expect(dashboard?.getColumn(20).hidden).toBe(true);

    const zip = await JSZip.loadAsync(new Uint8Array(buffer));
    const combo = await zip.file("xl/charts/chart1.xml")?.async("string");
    const share = await zip.file("xl/charts/chart2.xml")?.async("string");
    const ranking = await zip.file("xl/charts/chart3.xml")?.async("string");
    expect(combo).toContain("<c:barChart>");
    expect(combo).toContain("<c:lineChart>");
    expect(combo).toContain("0.0%");
    expect(share).toContain("<c:doughnutChart>");
    expect(share).toContain('val="3B82F6"');
    expect(share).toContain('val="10B981"');
    expect(ranking).toContain('<c:barDir val="bar"/>');
    expect(Object.keys(zip.files).filter((name) => /^xl\/media\//.test(name))).toHaveLength(0);
  });
});
