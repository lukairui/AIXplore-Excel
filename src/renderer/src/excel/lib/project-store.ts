import { PROJECT_STORAGE_KEY, parseProjectJson, recoverInterruptedProject } from "./core";
import { synchronizeChartViews } from "./view-sync";
import type { Project, TableRow, ColumnDefinition } from "./types";

const DATABASE_NAME = "nl-sheet.projects";
const DATABASE_VERSION = 1;
const PROJECT_STORE = "projects";
const SUMMARY_STORE = "summaries";

export type ProjectSummary = {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  sheetCount?: number;
  viewCount: number;
  createdAt: number;
  updatedAt: number;
};

let databasePromise: Promise<IDBDatabase> | null = null;

function openDatabase() {
  if (databasePromise) return databasePromise;
  databasePromise = new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(PROJECT_STORE)) {
        database.createObjectStore(PROJECT_STORE, { keyPath: "id" });
      }
      if (!database.objectStoreNames.contains(SUMMARY_STORE)) {
        const summaries = database.createObjectStore(SUMMARY_STORE, { keyPath: "id" });
        summaries.createIndex("updatedAt", "updatedAt");
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("无法打开项目数据库"));
  });
  return databasePromise;
}

function transactionDone(transaction: IDBTransaction) {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error("项目数据库操作失败"));
    transaction.onabort = () => reject(transaction.error || new Error("项目数据库操作已取消"));
  });
}

function requestResult<T>(request: IDBRequest<T>) {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("项目数据库读取失败"));
  });
}

export function summarizeProject(project: Project): ProjectSummary {
  const sheets: Array<{ rows: TableRow[]; columns: ColumnDefinition[] }> = project.sheets?.length
    ? project.sheets
    : [{ rows: project.rows, columns: project.columns }];
  return {
    id: project.id,
    name: project.name,
    rowCount: sheets.reduce((total, sheet) => total + sheet.rows.length, 0),
    columnCount: sheets.reduce((total, sheet) => total + sheet.columns.length, 0),
    sheetCount: sheets.length,
    viewCount: project.views.length,
    createdAt: project.createdAt,
    updatedAt: project.updatedAt,
  };
}

export async function saveProject(project: Project) {
  const database = await openDatabase();
  const persistedProject = JSON.parse(JSON.stringify(project)) as Project;
  const summary = summarizeProject(persistedProject);
  const transaction = database.transaction([PROJECT_STORE, SUMMARY_STORE], "readwrite");
  const done = transactionDone(transaction);
  transaction.objectStore(PROJECT_STORE).put(persistedProject);
  transaction.objectStore(SUMMARY_STORE).put(summary);
  await done;
  return summary;
}

export async function loadProjectById(projectId: string) {
  const database = await openDatabase();
  const transaction = database.transaction(PROJECT_STORE, "readonly");
  const done = transactionDone(transaction);
  const project = await requestResult(transaction.objectStore(PROJECT_STORE).get(projectId)) as Project | undefined;
  await done;
  return project ? synchronizeChartViews(recoverInterruptedProject(project)) : null;
}

export async function listProjectSummaries() {
  const database = await openDatabase();
  const transaction = database.transaction(SUMMARY_STORE, "readonly");
  const done = transactionDone(transaction);
  const summaries = await requestResult(transaction.objectStore(SUMMARY_STORE).getAll()) as ProjectSummary[];
  await done;
  return summaries.sort((left, right) => right.updatedAt - left.updatedAt);
}

export async function deleteStoredProject(projectId: string) {
  const database = await openDatabase();
  const transaction = database.transaction([PROJECT_STORE, SUMMARY_STORE], "readwrite");
  const done = transactionDone(transaction);
  transaction.objectStore(PROJECT_STORE).delete(projectId);
  transaction.objectStore(SUMMARY_STORE).delete(projectId);
  await done;
}

export async function initializeProjectStore() {
  const legacy = localStorage.getItem(PROJECT_STORAGE_KEY);
  if (legacy) {
    try {
      await saveProject(synchronizeChartViews(parseProjectJson(legacy)));
      localStorage.removeItem(PROJECT_STORAGE_KEY);
    } catch {
      // Invalid legacy data is left untouched so the user can recover it manually.
    }
  }
  return listProjectSummaries();
}

export async function clearProjectStore() {
  const database = await openDatabase();
  database.close();
  databasePromise = null;
  await new Promise<void>((resolve, reject) => {
    const request = indexedDB.deleteDatabase(DATABASE_NAME);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error || new Error("无法清空项目数据库"));
    request.onblocked = () => reject(new Error("项目数据库正在使用"));
  });
}
