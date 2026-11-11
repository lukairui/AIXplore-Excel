import type { ChartBindings, ChartType, Project } from "./types";

/**
 * 预置图表模板（Vue 版）。
 *
 * 与 ai_excel 原项目（React + Recharts）对应，但生成的是
 * Vue 3 渲染函数风格的 GeneratedView 源码，使用 ECharts 绘图。
 * 这些源码会交给 GeneratedViewHost 通过 @vue/babel-plugin-jsx 编译。
 */

export const CHART_OPTIONS: Array<{ id: ChartType; name: string; description: string }> = [
  { id: "custom", name: "AI 应用", description: "自由描述数据、布局和交互" },
  { id: "pie", name: "饼图", description: "展示构成与占比" },
  { id: "bar", name: "柱状图", description: "比较不同类别数值" },
  { id: "line", name: "折线图", description: "查看变化趋势" },
  { id: "area", name: "面积图", description: "强调趋势与规模" },
  { id: "scatter", name: "散点图", description: "观察两个指标关系" },
  { id: "bubble", name: "气泡图", description: "用大小表达第三指标" },
  { id: "histogram", name: "直方图", description: "查看数值分布" },
  { id: "radar", name: "雷达图", description: "比较多维轮廓" },
  { id: "waterfall", name: "瀑布图", description: "展示累计变化" },
  { id: "funnel", name: "漏斗图", description: "展示阶段转化" },
  { id: "treemap", name: "树状图", description: "展示层级与面积占比" },
];

const chartName = (type: ChartType) => CHART_OPTIONS.find((item) => item.id === type)?.name || "图表";

export function chartDescription(type: ChartType) {
  const option = CHART_OPTIONS.find((item) => item.id === type);
  return option ? `根据所选数据源生成${option.name}，${option.description}；自动识别工作表关系，并选择合适的分类与数值字段。` : "";
}

export function resolveChartBindings(
  project: Project,
  type: ChartType,
  description = "",
  explicit: ChartBindings = {},
): ChartBindings {
  const byId = new Map(project.columns.map((column) => [column.id, column]));
  const numericColumns = project.columns.filter((column) => column.dataType === "number");
  const dimensionColumns = project.columns.filter((column) => column.dataType !== "number");
  const mentioned = project.columns.filter((column) => description.includes(column.name));
  const mentionedNumeric = mentioned.filter((column) => column.dataType === "number");
  const mentionedDimensions = mentioned.filter((column) => column.dataType !== "number");
  const validNumeric = (id?: string) => id && byId.get(id)?.dataType === "number" ? id : undefined;
  const validColumn = (id?: string) => id && byId.has(id) ? id : undefined;
  const metricColumnId = validNumeric(explicit.metricColumnId)
    || mentionedNumeric[0]?.id
    || numericColumns.at(-1)?.id;
  const secondaryMetricColumnId = validNumeric(explicit.secondaryMetricColumnId)
    || mentionedNumeric.find((column) => column.id !== metricColumnId)?.id
    || numericColumns.find((column) => column.id !== metricColumnId)?.id
    || metricColumnId;
  return {
    dimensionColumnId: validColumn(explicit.dimensionColumnId)
      || mentionedDimensions[0]?.id
      || dimensionColumns[0]?.id
      || project.columns[0]?.id,
    metricColumnId,
    secondaryMetricColumnId,
    sizeColumnId: validNumeric(explicit.sizeColumnId)
      || mentionedNumeric.find((column) => column.id !== metricColumnId && column.id !== secondaryMetricColumnId)?.id
      || numericColumns.find((column) => column.id !== metricColumnId && column.id !== secondaryMetricColumnId)?.id
      || metricColumnId,
  };
}

/**
 * 生成预置图表的 GeneratedView 源码字符串。
 *
 * 输出的代码会被 GeneratedViewHost 用 Babel (@vue/babel-plugin-jsx) 编译。
 * 这里采用「写法 A」：返回一个 () => VNode 渲染函数，使用 JSX。
 * 图表通过 onMounted 中 echarts.init(domEl).setOption(option) 渲染。
 */
export function createChartViewSource(
  project: Project,
  type: ChartType,
  bindings: ChartBindings = {},
  description = "",
) {
  const resolved = resolveChartBindings(project, type, description, bindings);
  const labelColumn = project.columns.find((column) => column.id === resolved.dimensionColumnId);
  const primaryColumn = project.columns.find((column) => column.id === resolved.metricColumnId);
  const secondaryColumn = project.columns.find((column) => column.id === resolved.secondaryMetricColumnId);
  const labelId = JSON.stringify(resolved.dimensionColumnId || "");
  const primaryId = JSON.stringify(resolved.metricColumnId || "");
  const secondaryId = JSON.stringify(resolved.secondaryMetricColumnId || "");
  const sizeId = JSON.stringify(resolved.sizeColumnId || "");
  const primaryName = JSON.stringify(primaryColumn?.name || "数值");
  const secondaryName = JSON.stringify(secondaryColumn?.name || "指标二");
  const dimensionName = JSON.stringify(labelColumn?.name || "分类");
  const title = JSON.stringify(chartName(type));
  const chartType = JSON.stringify(type);
  const colorPalette = ["#19734a", "#d3922e", "#3977a8", "#b85151", "#7a64a3", "#4f8f8a", "#87944c", "#c26f3a"];

  // 根据图表类型生成对应的 ECharts option 构建逻辑
  const optionBuilder = (() => {
    switch (type) {
      case "pie":
        return `option = {
          color: ${JSON.stringify(colorPalette)},
          tooltip: { trigger: "item", formatter: (p) => p.name + " " + p.value + " (" + p.percent + "%)" },
          legend: { bottom: 8, type: "scroll" },
          series: [{
            type: "pie",
            radius: ["36%", "68%"],
            center: ["50%", "46%"],
            itemStyle: { borderColor: "#fff", borderWidth: 2 },
            label: { formatter: "{b} {d}%" },
            data: grouped.map((item, index) => ({ name: item.name, value: Math.max(0, item.primary), itemStyle: { color: colors[index % colors.length] } }))
          }]
        };`;
      case "bar":
        return `option = {
          color: colors,
          tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
          grid: { left: 8, right: 16, top: 16, bottom: 8, containLabel: true },
          xAxis: { type: "category", data: grouped.map(item => item.name), axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", axisLabel: { fontSize: 11 } },
          series: [{ name: ${primaryName}, type: "bar", barMaxWidth: 36, itemStyle: { borderRadius: [4, 4, 0, 0] }, data: grouped.map((item, index) => ({ value: Math.max(0, item.primary), itemStyle: { color: colors[index % colors.length] } })) }]
        };`;
      case "line":
        return `option = {
          color: [colors[0], colors[1]],
          tooltip: { trigger: "axis" },
          legend: { data: [${primaryName}, ${secondaryName}], bottom: 4 },
          grid: { left: 8, right: 20, top: 16, bottom: 36, containLabel: true },
          xAxis: { type: "category", boundaryGap: false, data: grouped.map(item => item.name), axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", axisLabel: { fontSize: 11 } },
          series: [
            { name: ${primaryName}, type: "line", smooth: true, symbol: "circle", symbolSize: 7, lineStyle: { width: 3 }, data: grouped.map(item => Math.max(0, item.primary)) },
            { name: ${secondaryName}, type: "line", smooth: true, symbol: "circle", symbolSize: 6, lineStyle: { width: 2, type: "dashed" }, data: grouped.map(item => Math.max(0, item.secondary)) }
          ]
        };`;
      case "area":
        return `option = {
          color: [colors[0]],
          tooltip: { trigger: "axis" },
          grid: { left: 8, right: 20, top: 16, bottom: 16, containLabel: true },
          xAxis: { type: "category", boundaryGap: false, data: grouped.map(item => item.name), axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", axisLabel: { fontSize: 11 } },
          series: [{ name: ${primaryName}, type: "line", smooth: true, areaStyle: { opacity: 0.35 }, lineStyle: { width: 2 }, data: grouped.map(item => Math.max(0, item.primary)) }]
        };`;
      case "scatter":
        return `option = {
          color: [colors[0]],
          tooltip: { trigger: "item", formatter: (p) => p.value[2] + "<br/>" + ${secondaryName} + ": " + p.value[0] + "<br/>" + ${primaryName} + ": " + p.value[1] },
          grid: { left: 8, right: 18, top: 16, bottom: 16, containLabel: true },
          xAxis: { type: "value", name: ${secondaryName}, nameLocation: "middle", nameGap: 28, axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", name: ${primaryName}, nameLocation: "middle", nameGap: 40, axisLabel: { fontSize: 11 } },
          series: [{ type: "scatter", symbolSize: 9, data: rawData.map(item => [Math.max(0, item.secondary), Math.max(0, item.primary), item.name]) }]
        };`;
      case "bubble":
        return `option = {
          color: [colors[1]],
          tooltip: { trigger: "item", formatter: (p) => p.value[3] + "<br/>" + ${secondaryName} + ": " + p.value[0] + "<br/>" + ${primaryName} + ": " + p.value[1] },
          grid: { left: 8, right: 24, top: 16, bottom: 16, containLabel: true },
          xAxis: { type: "value", name: ${secondaryName}, nameLocation: "middle", nameGap: 28, axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", name: ${primaryName}, nameLocation: "middle", nameGap: 40, axisLabel: { fontSize: 11 } },
          series: [{ type: "scatter", data: rawData.map(item => [Math.max(0, item.secondary), Math.max(0, item.primary), Math.max(8, Math.min(64, item.size)), item.name]) }]
        };`;
      case "histogram":
        return `option = {
          color: [colors[0]],
          tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
          grid: { left: 8, right: 16, top: 16, bottom: 8, containLabel: true },
          xAxis: { type: "category", data: histogramData.map(item => item.name), axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", minInterval: 1, axisLabel: { fontSize: 11 } },
          series: [{ type: "bar", barMaxWidth: 40, itemStyle: { color: colors[0], borderRadius: [3, 3, 0, 0] }, data: histogramData.map(item => item.count) }]
        };`;
      case "radar":
        return `option = {
          color: [colors[0]],
          tooltip: {},
          radar: { radius: "66%", indicator: grouped.map(item => ({ name: item.name, max: Math.max(...grouped.map(g => g.primary)) || 1 })) },
          series: [{ type: "radar", areaStyle: { opacity: 0.35 }, data: [{ value: grouped.map(item => Math.max(0, item.primary)), name: ${primaryName} }] }]
        };`;
      case "waterfall":
        return `option = {
          tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
          grid: { left: 8, right: 16, top: 16, bottom: 8, containLabel: true },
          xAxis: { type: "category", data: waterfallData.map(item => item.name), axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", axisLabel: { fontSize: 11 } },
          series: [
            { name: "占位", type: "bar", stack: "wf", itemStyle: { color: "transparent" }, data: waterfallData.map(item => item.base) },
            { name: "变化值", type: "bar", stack: "wf", itemStyle: { color: colors[0], borderRadius: [3, 3, 0, 0] }, data: waterfallData.map(item => item.delta) }
          ]
        };`;
      case "funnel":
        return `option = {
          color: colors,
          tooltip: { trigger: "item", formatter: "{b} : {c}" },
          series: [{ type: "funnel", left: "10%", width: "80%", gap: 2, label: { show: true, position: "inside", fontSize: 11 }, data: funnelData.map((item, index) => ({ name: item.name, value: Math.max(0, item.value), itemStyle: { color: colors[index % colors.length] } })) }]
        };`;
      case "treemap":
        return `option = {
          tooltip: { formatter: "{b}: {c}" },
          series: [{ type: "treemap", roam: false, nodeClick: false, breadcrumb: { show: false }, label: { fontSize: 11 }, itemStyle: { borderColor: "#fff", borderWidth: 2, gapWidth: 2 }, data: treemapData.map((item, index) => ({ name: item.name, value: Math.max(0, item.value), itemStyle: { color: colors[index % colors.length] } })) }]
        };`;
      default:
        return `option = {
          color: colors,
          tooltip: { trigger: "axis" },
          grid: { left: 8, right: 16, top: 16, bottom: 8, containLabel: true },
          xAxis: { type: "category", data: grouped.map(item => item.name), axisLabel: { fontSize: 11 } },
          yAxis: { type: "value", axisLabel: { fontSize: 11 } },
          series: [{ type: "bar", data: grouped.map(item => Math.max(0, item.primary)) }]
        };`;
    }
  })();

  return `function GeneratedView({ rows, libs }) {
  const Vue = libs.Vue;
  const echarts = libs.echarts;
  const onMounted = Vue.onMounted;
  const onBeforeUnmount = Vue.onBeforeUnmount;
  const ref = Vue.ref;
  const colors = ${JSON.stringify(colorPalette)};
  const el = ref(null);
  let chart = null;
  const rawData = rows.map((row, index) => ({ id: row.id, name: String(row[${labelId}] ?? ("记录 " + (index + 1))), primary: Number(row[${primaryId}] ?? 0), secondary: Number(row[${secondaryId}] ?? 0), size: Math.max(1, Math.abs(Number(row[${sizeId}] ?? 1))) }));
  const grouped = Object.values(rawData.reduce((result, item) => { const current = result[item.name] || { id: item.name, name: item.name, primary: 0, secondary: 0, size: 0 }; current.primary += item.primary; current.secondary += item.secondary; current.size += item.size; result[item.name] = current; return result; }, {}));
  const chartData = ["scatter", "bubble", "histogram", "waterfall"].includes(${chartType}) ? rawData : grouped;
  const values = (chartData.length ? chartData : grouped).map(item => item.primary);
  const minValue = values.length ? Math.min(...values) : 0;
  const maxValue = values.length ? Math.max(...values) : 0;
  const bucketCount = Math.max(1, Math.min(6, Math.ceil(Math.sqrt(values.length || 1))));
  const bucketSize = Math.max(1, (maxValue - minValue) / bucketCount);
  const histogramData = Array.from({ length: bucketCount }, (_, index) => ({ name: Math.round(minValue + index * bucketSize) + "–" + Math.round(minValue + (index + 1) * bucketSize), count: 0 }));
  values.forEach(value => { const index = Math.min(bucketCount - 1, Math.floor((value - minValue) / bucketSize)); if (histogramData[index]) histogramData[index].count += 1; });
  let runningTotal = 0;
  const waterfallData = grouped.map(item => { const base = runningTotal; runningTotal += item.primary; return { name: item.name, base, delta: item.primary }; });
  const funnelData = grouped.map((item, index) => ({ name: item.name, value: Math.max(0, item.primary) })).sort((a, b) => b.value - a.value);
  const treemapData = grouped.map((item, index) => ({ name: item.name, value: Math.max(0, item.primary) }));
  function buildOption() {
    let option;
    ${optionBuilder}
    return option;
  }
  function render() {
    if (chart) chart.setOption(buildOption(), true);
  }
  onMounted(() => {
    if (!el.value) return;
    chart = echarts.init(el.value);
    render();
    const onResize = () => chart && chart.resize();
    window.addEventListener("resize", onResize);
    el.value.__chart_resize = onResize;
  });
  onBeforeUnmount(() => {
    if (el.value && el.value.__chart_resize) window.removeEventListener("resize", el.value.__chart_resize);
    if (chart) { chart.dispose(); chart = null; }
  });
  if (!rows.length) {
    return () => <div style={{ padding: "32px", color: "#718078", textAlign: "center" }}>暂无可绘制的数据</div>;
  }
  return () => (
    <div style={{ minHeight: "100%", padding: "24px", background: "#f5f7f6", color: "#17221c", fontFamily: "Inter, system-ui, sans-serif" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "end", gap: "12px", flexWrap: "wrap", marginBottom: "16px" }}>
        <div>
          <div style={{ color: "#758179", fontSize: "12px", marginBottom: "5px" }}>图表视图</div>
          <h2 style={{ margin: "0", fontSize: "23px" }}>${JSON.parse(title)}</h2>
        </div>
        <div style={{ color: "#758179", fontSize: "12px" }}>{rows.length} 条记录 · ${JSON.parse(dimensionName)} × ${JSON.parse(primaryName)}</div>
      </div>
      <div ref={el} style={{ height: "min(560px, 62vh)", minHeight: "360px", padding: "8px", border: "1px solid #dfe4e1", borderRadius: "8px", background: "#fff" }}></div>
    </div>
  );
}`;
}
