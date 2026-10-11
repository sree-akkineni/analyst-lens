import type { GraphState } from "../dependencies/graph";
import {
  canonicalCellId,
  type CellFacts,
  type CellRef,
} from "../../shared/types";

export type PersistedWatchItem = Readonly<{
  address: string;
  category: string;
  label: string;
  sheetId: string;
  workbookId: string;
}>;

export type WatchItemStatus = "loading" | "missing" | "ready";

export type WatchItem = Readonly<{
  category: string;
  currentFormula: string | null;
  currentValue: CellFacts["value"];
  formulaChanged: boolean;
  id: string;
  label: string;
  lastFormula: string | null;
  lastValue: CellFacts["value"];
  message?: string;
  ref: CellRef;
  status: WatchItemStatus;
  valueChanged: boolean;
}>;

export type WatchlistPersistence = "failed" | "saved" | "saving" | "session" | "unavailable";

export type WatchlistState = Readonly<{
  affected: readonly string[];
  items: readonly WatchItem[];
  persistence: WatchlistPersistence;
  refreshing: boolean;
}>;

export const initialWatchlistState = (
  persistence: WatchlistPersistence = "session",
): WatchlistState => ({ affected: [], items: [], persistence, refreshing: false });

export const watchDisplayRef = (ref: CellRef): string =>
  `${ref.sheetName}!${ref.address}`;

export const createWatchItem = (
  facts: CellFacts,
  metadata: Readonly<{ category?: string; label?: string }> = {},
): WatchItem => ({
  category: metadata.category ?? "",
  currentFormula: facts.formula,
  currentValue: facts.value,
  formulaChanged: false,
  id: canonicalCellId(facts),
  label: metadata.label ?? "",
  lastFormula: facts.formula,
  lastValue: facts.value,
  ref: facts,
  status: "ready",
  valueChanged: false,
});

export const refreshWatchItem = (
  item: WatchItem,
  facts: CellFacts,
): WatchItem => {
  const firstObservation = item.status === "loading";
  const { message, ...retained } = item;
  void message;
  return {
    ...retained,
    currentFormula: facts.formula,
    currentValue: facts.value,
    formulaChanged: firstObservation ? false : item.currentFormula !== facts.formula,
    id: canonicalCellId(facts),
    lastFormula: firstObservation ? facts.formula : item.currentFormula,
    lastValue: firstObservation ? facts.value : item.currentValue,
    ref: facts,
    status: "ready",
    valueChanged: firstObservation ? false : !Object.is(item.currentValue, facts.value),
  };
};

export const markWatchItemMissing = (
  item: WatchItem,
  message = "The pinned sheet or cell is missing or was renamed.",
): WatchItem => ({ ...item, message, status: "missing" });

export const updateWatchMetadata = (
  item: WatchItem,
  metadata: Readonly<{ category: string; label: string }>,
): WatchItem => ({ ...item, ...metadata });

export const serializeForWorkbook = (
  items: readonly WatchItem[],
): PersistedWatchItem[] => items.map((item) => ({
  address: item.ref.address,
  category: item.category,
  label: item.label,
  sheetId: item.ref.sheetId,
  workbookId: item.ref.workbookId,
}));

const stringField = (
  value: Readonly<Record<string, unknown>>,
  key: keyof PersistedWatchItem,
): string | null => typeof value[key] === "string" ? value[key] : null;

export const parseWatchlistMetadata = (value: unknown): PersistedWatchItem[] => {
  if (!Array.isArray(value)) return [];
  return value.flatMap((candidate) => {
    if (typeof candidate !== "object" || candidate === null) return [];
    const record = candidate as Readonly<Record<string, unknown>>;
    const address = stringField(record, "address");
    const category = stringField(record, "category");
    const label = stringField(record, "label");
    const sheetId = stringField(record, "sheetId");
    const workbookId = stringField(record, "workbookId");
    if (
      address === null || category === null || label === null ||
      sheetId === null || workbookId === null
    ) return [];
    return [{ address, category, label, sheetId, workbookId }];
  });
};

export const hydrateWatchlistMetadata = (
  metadata: readonly PersistedWatchItem[],
): WatchItem[] => metadata.map((item) => {
  const ref: CellRef = {
    address: item.address,
    sheetId: item.sheetId,
    sheetName: "Unknown sheet",
    workbookId: item.workbookId,
  };
  return {
    category: item.category,
    currentFormula: null,
    currentValue: null,
    formulaChanged: false,
    id: canonicalCellId(ref),
    label: item.label,
    lastFormula: null,
    lastValue: null,
    ref,
    status: "loading",
    valueChanged: false,
  };
});

export const affectedPins = (
  graph: GraphState | null,
  assumption: CellRef,
  pins: readonly CellRef[],
): string[] => {
  if (graph === null) return [];
  const adjacency = new Map<string, Set<string>>();
  for (const edge of graph.edges) {
    const targets = adjacency.get(edge.source) ?? new Set<string>();
    targets.add(edge.target);
    adjacency.set(edge.source, targets);
  }
  const start = canonicalCellId(assumption);
  const reachable = new Set<string>();
  const pending = [...(adjacency.get(start) ?? [])];
  while (pending.length > 0) {
    const current = pending.pop();
    if (current === undefined || reachable.has(current)) continue;
    reachable.add(current);
    for (const next of adjacency.get(current) ?? []) pending.push(next);
  }
  return pins.flatMap((pin) => reachable.has(canonicalCellId(pin))
    ? [watchDisplayRef(pin)]
    : []);
};
