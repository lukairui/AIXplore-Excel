import { ensureProjectSheets } from "./core";
import { createChartViewSource, resolveChartBindings } from "./chart-templates";
import { projectForView, resolveSourceSheetIds } from "./data-sources";
import type { Project } from "./types";

export function synchronizeChartViews(project: Project) {
  const sheets = ensureProjectSheets(project);
  project.views.forEach((view) => {
    const sourceSheetIds = resolveSourceSheetIds(project, view.sourceSheetIds);
    view.sourceSheetIds = sourceSheetIds;
    if (view.chartType === "custom" || sourceSheetIds.length > 1) return;
    const sourceProject = projectForView(project, view);
    view.chartBindings = resolveChartBindings(
      sourceProject,
      view.chartType,
      view.draftDescription,
      view.chartBindingsLocked ? view.chartBindings : {},
    );
    const source = createChartViewSource(
      sourceProject,
      view.chartType,
      view.chartBindings,
      view.draftDescription,
    );
    view.description = view.draftDescription;
    view.generatedCode = source;
    view.lastAttemptCode = source;
    view.generatedVersion = view.sourceVersion;
    view.status = "ready";
    view.error = null;
    view.runtimeError = null;
  });
  if (!project.activeSheetId) project.activeSheetId = sheets[0]?.id;
  return project;
}
