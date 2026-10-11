import {
  collapseBranchResult,
  createGraphState,
  expandNode,
  type GraphState,
} from "../domain/dependencies/graph";
import {
  back,
  forward,
  returnToOrigin,
  startHistory,
  visit,
  type NavigationHistory,
} from "../domain/dependencies/navigationHistory";
import {
  analyzeFormula,
  type FormulaRisk,
  type FormulaToken,
} from "../domain/formulas/tokenizeFormula";
import { normalizeFormula } from "../domain/formulas/normalizeFormula";
import {
  createWatchItem,
  hydrateWatchlistMetadata,
  initialWatchlistState,
  markWatchItemMissing,
  refreshWatchItem,
  serializeForWorkbook,
  updateWatchMetadata,
  watchDisplayRef,
  type WatchItem,
  type WatchlistState,
} from "../domain/watchlist/watchlist";
import {
  normalizeOfficeError,
  type NeighborCellFacts,
  type OfficeGateway,
  type Unsubscribe,
} from "../platform/office/officeGateway";
import {
  canonicalCellId,
  type CellFacts,
  type CellRef,
  type DependencyDirection,
  type DependencyNode,
  type DependencyRef,
} from "../shared/types";
import { dependencyRepositoryFor } from "./dependencyRepository";

export type InspectorCellKind =
  | "blank"
  | "error"
  | "formula"
  | "hardcode"
  | "multi";

export type InspectorView =
  | Readonly<{ kind: "failure"; message: string }>
  | Readonly<{ kind: "loading" }>
  | Readonly<{
      cellKind: InspectorCellKind;
      facts: CellFacts;
      kind: "selection";
    }>
  | Readonly<{ kind: "unsupported"; message: string }>;

export type InspectorOverlay = "graph" | "palette" | "settings";

export type InspectorState = Readonly<{
  capped: boolean;
  dependents: readonly DependencyNode[];
  graph: GraphState | null;
  history: NavigationHistory | null;
  neighbors: readonly Readonly<{
    label: string;
    normalizedFormula: string | null;
    status: "different" | "same" | "unavailable";
  }>[];
  notice: string | null;
  overlays: Readonly<Record<InspectorOverlay, boolean>>;
  precedents: readonly DependencyNode[];
  requestId: number;
  risks: readonly FormulaRisk[];
  tokens: readonly FormulaToken[];
  view: InspectorView;
  watchlist: WatchlistState;
}>;

export type InspectorIntent =
  | Readonly<{ direction: DependencyDirection; nodeId: string; type: "collapseBranch" }>
  | Readonly<{ direction: DependencyDirection; nodeId: string; type: "expandNode" }>
  | Readonly<{ overlay: InspectorOverlay; type: "closeOverlay" }>
  | Readonly<{ overlay: InspectorOverlay; type: "openOverlay" }>
  | Readonly<{ ref: CellRef; type: "navigate" }>
  | Readonly<{ text: string; type: "navigateReference" }>
  | Readonly<{ category: string; id: string; label: string; type: "updateWatchMetadata" }>
  | Readonly<{ id: string; type: "unpinWatch" }>
  | Readonly<{
      type:
        | "back"
        | "forward"
        | "pinCurrent"
        | "refreshWatchlist"
        | "returnToOrigin"
        | "saveWatchlistWithWorkbook";
    }>;

export type InspectorController = Readonly<{
  dispatch(intent: InspectorIntent): Promise<void>;
  getState(): InspectorState;
  refresh(): Promise<void>;
  start(): Promise<void>;
  stop(): void;
  subscribe(listener: () => void): () => void;
}>;

export type ControllerOptions = Readonly<{
  debounceMs?: number;
  safeConstants?: readonly number[];
}>;

type A1Bounds = Readonly<{
  bottom: number;
  left: number;
  right: number;
  top: number;
}>;

const columnNumber = (letters: string): number => [...letters.toUpperCase()].reduce(
  (value, letter) => value * 26 + letter.charCodeAt(0) - 64,
  0,
);

const parseA1Bounds = (address: string): A1Bounds | null => {
  const cells = address.split(":");
  if (cells.length > 2) return null;
  const parsed = cells.map((cell) => cell.match(/^\$?([A-Z]{1,3})\$?(\d+)$/iu));
  const start = parsed[0];
  const end = parsed[1] ?? start;
  if (start?.[1] === undefined || start[2] === undefined || end?.[1] === undefined || end[2] === undefined) {
    return null;
  }
  const startColumn = columnNumber(start[1]);
  const endColumn = columnNumber(end[1]);
  const startRow = Number(start[2]);
  const endRow = Number(end[2]);
  return {
    bottom: Math.max(startRow, endRow),
    left: Math.min(startColumn, endColumn),
    right: Math.max(startColumn, endColumn),
    top: Math.min(startRow, endRow),
  };
};

const containsRef = (container: CellRef, candidate: CellRef): boolean => {
  if (container.workbookId !== candidate.workbookId || container.sheetId !== candidate.sheetId) {
    return false;
  }
  const containerBounds = parseA1Bounds(container.address);
  const candidateBounds = parseA1Bounds(candidate.address);
  if (containerBounds === null || candidateBounds === null) return false;
  return candidateBounds.left >= containerBounds.left
    && candidateBounds.right <= containerBounds.right
    && candidateBounds.top >= containerBounds.top
    && candidateBounds.bottom <= containerBounds.bottom;
};

export const initialInspectorState = (): InspectorState => ({
  capped: false,
  dependents: [],
  graph: null,
  history: null,
  neighbors: [],
  notice: null,
  overlays: { graph: false, palette: false, settings: false },
  precedents: [],
  requestId: 0,
  risks: [],
  tokens: [],
  view: { kind: "loading" },
  watchlist: initialWatchlistState(),
});

const classifySelection = (facts: CellFacts): InspectorCellKind => {
  if (facts.rowCount > 1 || facts.columnCount > 1) return "multi";
  if (typeof facts.value === "string" && facts.value.startsWith("#")) return "error";
  if (facts.formula !== null && facts.formula.startsWith("=")) return "formula";
  if (facts.value === null) return "blank";
  return "hardcode";
};

const dependencyNode = (ref: DependencyRef): DependencyNode => ({
  id: canonicalCellId(ref),
  ref,
  ...(ref.unresolved === undefined ? {} : { unresolved: ref.unresolved }),
});

const selectionNode = (facts: CellFacts): DependencyNode => ({
  id: canonicalCellId(facts),
  ref: facts,
  value: facts.value,
});

const sameWorkbook = (
  history: NavigationHistory | null,
  facts: CellFacts,
): history is NavigationHistory =>
  history !== null && history.current.workbookId === facts.workbookId;

export const createController = (
  gateway: OfficeGateway,
  options: ControllerOptions = {},
): InspectorController => {
  let state: InspectorState = {
    ...initialInspectorState(),
    watchlist: initialWatchlistState(
      gateway.saveWatchlistMetadata === undefined ? "unavailable" : "session",
    ),
  };
  const listeners = new Set<() => void>();
  let recalculationSubscription: Unsubscribe | null = null;
  let selectionSubscription: Unsubscribe | null = null;
  let selectionTimer: ReturnType<typeof setTimeout> | null = null;
  let activeSelection: AbortController | null = null;
  let activeExpansion: AbortController | null = null;
  let activeWatchRefresh: AbortController | null = null;
  let expansionRequestId = 0;
  let watchRefreshRequestId = 0;
  let watchSaveRequestId = 0;
  let watchlistRevision = 0;
  let lifecycleId = 0;
  let requestId = 0;
  let started = false;
  const debounceMs = options.debounceMs ?? 100;
  const dependencyRepository = dependencyRepositoryFor(gateway);
  const neighborCache = new Map<string, readonly NeighborCellFacts[]>();
  const cache = <Value,>(store: Map<string, Value>, key: string, value: Value) => {
    if (store.size >= 500) {
      const oldest = store.keys().next().value;
      if (oldest !== undefined) store.delete(oldest);
    }
    store.set(key, value);
  };

  const resolveAffectedWatchlist = async (
    selected: CellRef,
    items: readonly WatchItem[],
    signal: AbortSignal,
  ): Promise<string[]> => {
    if (items.length === 0) return [];
    const targets = new Map(items.map((item) => [canonicalCellId(item.ref), item]));
    targets.delete(canonicalCellId(selected));
    const found = new Set<string>();
    const visited = new Set<string>([canonicalCellId(selected)]);
    const pending: CellRef[] = [selected];
    let index = 0;
    while (index < pending.length && found.size < targets.size) {
      if (signal.aborted) return [];
      const current = pending[index];
      index += 1;
      if (current === undefined) continue;
      const dependents = await dependencyRepository.get(current, "dependent", signal);
      for (const dependent of dependents) {
        if (signal.aborted) return [];
        const id = canonicalCellId(dependent);
        if (targets.has(id)) found.add(id);
        for (const [targetId, target] of targets) {
          if (containsRef(dependent, target.ref)) found.add(targetId);
        }
        if (dependent.unresolved !== undefined || visited.has(id)) continue;
        visited.add(id);
        pending.push(dependent);
      }
    }
    return items.flatMap((item) => found.has(canonicalCellId(item.ref))
      ? [watchDisplayRef(item.ref)]
      : []);
  };

  const loadNeighbors = async (
    facts: CellFacts,
    signal: AbortSignal,
  ): Promise<readonly NeighborCellFacts[]> => {
    const key = `${canonicalCellId(facts)}::${facts.formula ?? ""}`;
    const cached = neighborCache.get(key);
    if (cached !== undefined) return cached;
    const neighbors = gateway.getNeighborFacts === undefined
      ? []
      : await gateway.getNeighborFacts(facts, signal);
    if (!signal.aborted) cache(neighborCache, key, [...neighbors]);
    return neighbors;
  };

  const compareNeighbors = (
    facts: CellFacts,
    neighbors: readonly NeighborCellFacts[],
  ): InspectorState["neighbors"] => {
    const selectedFormula = facts.formula === null
      ? null
      : normalizeFormula(facts.formula, facts);
    return neighbors.map((neighbor) => {
      const normalizedFormula = neighbor.facts.formula === null
        ? null
        : normalizeFormula(neighbor.facts.formula, neighbor.facts);
      const label = neighbor.direction[0]?.toUpperCase() + neighbor.direction.slice(1);
      return {
        label,
        normalizedFormula,
        status: selectedFormula === null || normalizedFormula === null
          ? "unavailable"
          : selectedFormula === normalizedFormula
            ? "same"
            : "different",
      };
    });
  };

  const publish = (next: InspectorState) => {
    state = next;
    listeners.forEach((listener) => listener());
  };

  const publishView = (view: InspectorView, currentRequestId = state.requestId) => {
    publish({ ...state, requestId: currentRequestId, view });
  };

  const selectedFacts = (view: InspectorView = state.view): CellFacts | null =>
    view.kind === "selection" && view.cellKind !== "multi" ? view.facts : null;

  const watchlistWithImpact = (
    items: readonly WatchItem[],
    patch: Partial<WatchlistState> = {},
  ): WatchlistState => {
    const displayed = new Set(items.map((item) => watchDisplayRef(item.ref)));
    return {
      ...state.watchlist,
      ...patch,
      affected: state.watchlist.affected.filter((current) => displayed.has(current)),
      items,
    };
  };

  const pendingWatchlistPersistence = (): WatchlistState["persistence"] =>
    state.watchlist.persistence === "unavailable" ? "unavailable" : "session";

  const resolveCurrentWatchlistImpact = async (
    selected: CellRef,
    signal: AbortSignal,
  ): Promise<string[]> => {
    let items = [...state.watchlist.items];
    const itemIds = items.map((item) => item.id).join("\n");
    let affected = await resolveAffectedWatchlist(selected, items, signal);
    if (signal.aborted) return [];
    const currentIds = state.watchlist.items.map((item) => item.id).join("\n");
    if (currentIds !== itemIds) {
      items = [...state.watchlist.items];
      affected = await resolveAffectedWatchlist(selected, items, signal);
    }
    const displayed = new Set(state.watchlist.items.map((item) => watchDisplayRef(item.ref)));
    return affected.filter((current) => displayed.has(current));
  };

  const refreshWatchlist = async (): Promise<void> => {
    if (state.watchlist.items.length === 0) return;
    const getCellFacts = gateway.getCellFacts;
    const getCellsFacts = gateway.getCellsFacts;
    if (getCellFacts === undefined && getCellsFacts === undefined) {
      publish({
        ...state,
        notice: "This Excel host cannot refresh pinned cells without changing selection.",
      });
      return;
    }
    const currentRequest = ++watchRefreshRequestId;
    activeWatchRefresh?.abort();
    const abort = new AbortController();
    activeWatchRefresh = abort;
    const snapshot = [...state.watchlist.items];
    publish({
      ...state,
      watchlist: watchlistWithImpact(snapshot, { refreshing: true }),
    });
    const refreshed = new Map<
      string,
      Readonly<{ facts: CellFacts }> | Readonly<{ message: string }>
    >();
    let batchCompleted = false;
    if (getCellsFacts !== undefined) {
      try {
        const facts = await getCellsFacts(snapshot.map((item) => item.ref), abort.signal);
        if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
        snapshot.forEach((item, index) => {
          const current = facts[index];
          refreshed.set(item.id, current === null || current === undefined
            ? { message: "The pinned sheet or cell is missing or was renamed." }
            : { facts: current });
        });
        batchCompleted = true;
      } catch (error) {
        if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
        if (getCellFacts === undefined) {
          const message = normalizeOfficeError(error).message;
          snapshot.forEach((item) => refreshed.set(item.id, { message }));
          batchCompleted = true;
        }
      }
    }
    if (!batchCompleted && getCellFacts !== undefined) {
      for (const item of snapshot) {
        if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
        try {
          const facts = await getCellFacts(item.ref, abort.signal);
          if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
          refreshed.set(item.id, { facts });
        } catch (error) {
          if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
          const normalized = normalizeOfficeError(error);
          refreshed.set(item.id, { message: normalized.message });
        }
      }
    }
    if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
    const items = state.watchlist.items.map((item) => {
      const result = refreshed.get(item.id);
      if (result === undefined) return item;
      return "facts" in result
        ? refreshWatchItem(item, result.facts)
        : markWatchItemMissing(item, result.message);
    });
    let affected = watchlistWithImpact(items).affected;
    const selected = selectedFacts();
    if (selected !== null) {
      try {
        affected = await resolveAffectedWatchlist(selected, items, abort.signal);
      } catch {
        affected = watchlistWithImpact(items).affected;
      }
    }
    if (abort.signal.aborted || currentRequest !== watchRefreshRequestId) return;
    publish({
      ...state,
      watchlist: {
        ...watchlistWithImpact(items, { refreshing: false }),
        affected,
      },
    });
  };

  const loadPersistedWatchlist = async (expectedLifecycleId: number): Promise<void> => {
    const load = gateway.loadWatchlistMetadata;
    if (load === undefined) return;
    try {
      const metadata = await load();
      if (!started || expectedLifecycleId !== lifecycleId) return;
      const hydrated = hydrateWatchlistMetadata(metadata);
      const items = new Map(state.watchlist.items.map((item) => [item.id, item]));
      hydrated.forEach((item) => {
        if (!items.has(item.id)) items.set(item.id, item);
      });
      publish({
        ...state,
        watchlist: watchlistWithImpact([...items.values()]),
      });
      await refreshWatchlist();
    } catch (error) {
      if (!started || expectedLifecycleId !== lifecycleId) return;
      publish({
        ...state,
        notice: normalizeOfficeError(error).message,
        watchlist: { ...state.watchlist, persistence: "failed", refreshing: false },
      });
    }
  };

  const saveWatchlistWithWorkbook = async (): Promise<void> => {
    const save = gateway.saveWatchlistMetadata;
    if (save === undefined) {
      publish({
        ...state,
        watchlist: { ...state.watchlist, persistence: "unavailable" },
      });
      return;
    }
    const currentSaveRequest = ++watchSaveRequestId;
    const savedRevision = watchlistRevision;
    const savedLifecycle = lifecycleId;
    const metadata = serializeForWorkbook(state.watchlist.items);
    publish({ ...state, watchlist: { ...state.watchlist, persistence: "saving" } });
    try {
      await save(metadata);
      if (
        !started || savedLifecycle !== lifecycleId ||
        currentSaveRequest !== watchSaveRequestId
      ) return;
      publish({
        ...state,
        watchlist: {
          ...state.watchlist,
          persistence: savedRevision === watchlistRevision ? "saved" : "session",
        },
      });
    } catch (error) {
      if (
        !started || savedLifecycle !== lifecycleId ||
        currentSaveRequest !== watchSaveRequestId
      ) return;
      publish({
        ...state,
        notice: normalizeOfficeError(error).message,
        watchlist: { ...state.watchlist, persistence: "failed" },
      });
    }
  };

  const refresh = async () => {
    const currentRequestId = ++requestId;
    activeSelection?.abort();
    activeExpansion?.abort();
    expansionRequestId += 1;
    const abort = new AbortController();
    activeSelection = abort;
    publishView({ kind: "loading" }, currentRequestId);
    try {
      const facts = await gateway.getSelectionFacts(abort.signal);
      if (abort.signal.aborted || currentRequestId !== requestId) return;
      const cellKind = classifySelection(facts);
      if (cellKind === "multi") {
        const view: InspectorView = { cellKind, facts, kind: "selection" };
        publish({
          ...state,
          capped: false,
          dependents: [],
          graph: null,
          precedents: [],
          requestId: currentRequestId,
          risks: [],
          tokens: [],
          view,
          watchlist: { ...watchlistWithImpact(state.watchlist.items), affected: [] },
        });
        return;
      }
      const [precedentRefs, dependentRefs, neighborFacts] = await Promise.all([
        dependencyRepository.get(facts, "precedent", abort.signal),
        dependencyRepository.get(facts, "dependent", abort.signal),
        loadNeighbors(facts, abort.signal),
      ]);
      if (abort.signal.aborted || currentRequestId !== requestId) return;
      const precedents = precedentRefs.map(dependencyNode);
      const dependents = dependentRefs.map(dependencyNode);
      const origin = selectionNode(facts);
      let graph = createGraphState(origin);
      const precedentExpansion = expandNode(graph, origin.id, "precedent", precedents);
      graph = precedentExpansion.state;
      const dependentExpansion = expandNode(graph, origin.id, "dependent", dependents);
      graph = dependentExpansion.state;
      const analysis = facts.formula === null
        ? { risks: [], tokens: [] }
        : analyzeFormula(
            facts.formula,
            options.safeConstants === undefined
              ? { origin: facts }
              : { origin: facts, safeConstants: options.safeConstants },
          );
      const history = sameWorkbook(state.history, facts)
        ? visit(state.history, facts)
        : startHistory(facts);
      const view: InspectorView = { cellKind, facts, kind: "selection" };
      const affected = await resolveCurrentWatchlistImpact(facts, abort.signal);
      if (abort.signal.aborted || currentRequestId !== requestId) return;
      publish({
        ...state,
        capped: precedentExpansion.capped || dependentExpansion.capped,
        dependents,
        graph,
        history,
        neighbors: compareNeighbors(facts, neighborFacts),
        notice: null,
        precedents,
        requestId: currentRequestId,
        risks: analysis.risks,
        tokens: analysis.tokens,
        view,
        watchlist: { ...watchlistWithImpact(state.watchlist.items), affected },
      });
    } catch (error) {
      if (abort.signal.aborted || currentRequestId !== requestId) return;
      const normalized = normalizeOfficeError(error);
      publishView(
        normalized.code === "UNSUPPORTED_OPERATION"
          ? { kind: "unsupported", message: normalized.message }
          : { kind: "failure", message: normalized.message },
        currentRequestId,
      );
    }
  };

  const scheduleRefresh = () => {
    if (selectionTimer !== null) clearTimeout(selectionTimer);
    selectionTimer = setTimeout(() => {
      selectionTimer = null;
      void refresh();
    }, debounceMs);
  };

  const scheduleRecalculationRefresh = () => {
    dependencyRepository.clear();
    neighborCache.clear();
    void refreshWatchlist();
    scheduleRefresh();
  };

  const navigateHistory = async (nextHistory: NavigationHistory | null) => {
    if (nextHistory === null || nextHistory === state.history) return;
    await gateway.selectCell(nextHistory.current);
    publish({ ...state, history: nextHistory, notice: null });
  };

  const dispatch = async (intent: InspectorIntent): Promise<void> => {
    if (intent.type === "openOverlay" || intent.type === "closeOverlay") {
      publish({
        ...state,
        overlays: {
          ...state.overlays,
          [intent.overlay]: intent.type === "openOverlay",
        },
      });
      return;
    }
    if (intent.type === "pinCurrent") {
      const facts = selectedFacts();
      if (facts === null) {
        publish({ ...state, notice: "Select one cell before pinning a key output." });
        return;
      }
      const item = createWatchItem(facts);
      if (state.watchlist.items.some((current) => current.id === item.id)) {
        publish({ ...state, notice: "That cell is already pinned for this session." });
        return;
      }
      const items = [...state.watchlist.items, item];
      watchlistRevision += 1;
      publish({
        ...state,
        notice: "Pinned for this session.",
        watchlist: watchlistWithImpact(items, {
          persistence: pendingWatchlistPersistence(),
        }),
      });
      return;
    }
    if (intent.type === "refreshWatchlist") {
      await refreshWatchlist();
      return;
    }
    if (intent.type === "saveWatchlistWithWorkbook") {
      await saveWatchlistWithWorkbook();
      return;
    }
    if (intent.type === "unpinWatch") {
      const items = state.watchlist.items.filter((item) => item.id !== intent.id);
      if (items.length !== state.watchlist.items.length) watchlistRevision += 1;
      publish({
        ...state,
        watchlist: watchlistWithImpact(items, {
          persistence: pendingWatchlistPersistence(),
        }),
      });
      return;
    }
    if (intent.type === "updateWatchMetadata") {
      const changed = state.watchlist.items.some((item) =>
        item.id === intent.id &&
        (item.category !== intent.category || item.label !== intent.label));
      const items = state.watchlist.items.map((item) => item.id === intent.id
        ? updateWatchMetadata(item, { category: intent.category, label: intent.label })
        : item);
      if (changed) watchlistRevision += 1;
      publish({
        ...state,
        watchlist: watchlistWithImpact(items, {
          persistence: pendingWatchlistPersistence(),
        }),
      });
      return;
    }
    if (intent.type === "navigate") {
      await gateway.selectCell(intent.ref);
      const history = state.history === null
        ? startHistory(intent.ref)
        : visit(state.history, intent.ref);
      publish({ ...state, history, notice: null });
      return;
    }
    if (intent.type === "navigateReference") {
      const normalized = intent.text.replaceAll("$", "").toLocaleLowerCase();
      const match = [...state.precedents, ...state.dependents].find((current) =>
        `${current.ref.sheetName}!${current.ref.address}`
          .replaceAll("$", "")
          .toLocaleLowerCase()
          .endsWith(normalized.replaceAll("'", "")) ||
        current.ref.address.replaceAll("$", "").toLocaleLowerCase() === normalized);
      if (match === undefined) {
        publish({ ...state, notice: "Excel did not resolve that reference as a direct dependency." });
        return;
      }
      await dispatch({ ref: match.ref, type: "navigate" });
      return;
    }
    if (intent.type === "back") {
      await navigateHistory(state.history === null ? null : back(state.history));
      return;
    }
    if (intent.type === "forward") {
      await navigateHistory(state.history === null ? null : forward(state.history));
      return;
    }
    if (intent.type === "returnToOrigin") {
      await navigateHistory(state.history === null ? null : returnToOrigin(state.history));
      return;
    }
    if (intent.type === "collapseBranch") {
      if (state.graph === null) return;
      const collapsed = collapseBranchResult(
        state.graph,
        intent.nodeId,
        intent.direction,
      );
      publish({
        ...state,
        capped: collapsed.capped,
        graph: collapsed.state,
      });
      return;
    }
    if (intent.type !== "expandNode") return;
    if (state.graph === null) return;
    const parent = state.graph.nodes.find((current) => current.id === intent.nodeId);
    if (parent === undefined || parent.unresolved !== undefined) return;
    const currentExpansionRequest = ++expansionRequestId;
    activeExpansion?.abort();
    const abort = new AbortController();
    activeExpansion = abort;
    try {
      const refs = await dependencyRepository.get(
        parent.ref,
        intent.direction,
        abort.signal,
      );
      if (abort.signal.aborted || currentExpansionRequest !== expansionRequestId) return;
      const expanded = expandNode(
        state.graph,
        parent.id,
        intent.direction,
        refs.map(dependencyNode),
      );
      publish({ ...state, capped: expanded.capped, graph: expanded.state });
    } catch (error) {
      if (abort.signal.aborted || currentExpansionRequest !== expansionRequestId) return;
      publish({ ...state, notice: normalizeOfficeError(error).message });
    }
  };

  return {
    dispatch,
    getState: () => state,
    refresh,
    async start() {
      if (started) return;
      started = true;
      const currentLifecycleId = ++lifecycleId;
      try {
        const subscription = await gateway.subscribeSelection(() => {
          if (started && currentLifecycleId === lifecycleId) scheduleRefresh();
        });
        if (!started || currentLifecycleId !== lifecycleId) {
          subscription();
          return;
        }
        selectionSubscription = subscription;
        try {
          const recalculation = await gateway.subscribeRecalculation(
            () => {
              if (started && currentLifecycleId === lifecycleId) {
                scheduleRecalculationRefresh();
              }
            },
          );
          if (!started || currentLifecycleId !== lifecycleId) {
            recalculation();
            return;
          }
          recalculationSubscription = recalculation;
        } catch (error) {
          const normalized = normalizeOfficeError(error);
          if (normalized.code !== "UNSUPPORTED_OPERATION") throw error;
        }
        await loadPersistedWatchlist(currentLifecycleId);
        if (!started || currentLifecycleId !== lifecycleId) return;
        await refresh();
      } catch (error) {
        if (!started || currentLifecycleId !== lifecycleId) return;
        selectionSubscription?.();
        selectionSubscription = null;
        const normalized = normalizeOfficeError(error);
        publishView(normalized.code === "UNSUPPORTED_OPERATION"
          ? { kind: "unsupported", message: normalized.message }
          : { kind: "failure", message: normalized.message });
      }
    },
    stop() {
      started = false;
      lifecycleId += 1;
      if (selectionTimer !== null) clearTimeout(selectionTimer);
      selectionTimer = null;
      activeSelection?.abort();
      activeExpansion?.abort();
      activeWatchRefresh?.abort();
      watchRefreshRequestId += 1;
      watchSaveRequestId += 1;
      recalculationSubscription?.();
      recalculationSubscription = null;
      selectionSubscription?.();
      selectionSubscription = null;
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
};
