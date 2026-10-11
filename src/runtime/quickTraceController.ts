import {
  createGraphState,
  expandNode,
} from "../domain/dependencies/graph";
import {
  applyQuickTraceExpansion,
  createQuickTraceTree,
  moveQuickTraceLeft,
  moveQuickTraceRight,
  moveQuickTraceSelection,
  projectQuickTraceRows,
  type QuickTraceTreeState,
} from "../domain/dependencies/quickTraceTree";
import {
  returnToOrigin as returnHistoryToOrigin,
  startHistory,
  visit,
  type NavigationHistory,
} from "../domain/dependencies/navigationHistory";
import type {
  QuickTraceDialogHost,
  QuickTraceDialogSession,
} from "../platform/office/quickTraceDialogHost";
import {
  normalizeOfficeError,
  type OfficeGateway,
  type Unsubscribe,
} from "../platform/office/officeGateway";
import {
  parseQuickTraceDialogIntent,
  QUICK_TRACE_PROTOCOL_VERSION,
  serializeQuickTraceMessage,
  type QuickTraceHostEvent,
  type QuickTraceSessionIdentity,
  type QuickTraceSnapshot,
} from "../shared/quickTraceProtocol";
import {
  canonicalCellId,
  type CellFacts,
  type CellRef,
  type DependencyDirection,
  type DependencyNode,
  type DependencyRef,
} from "../shared/types";
import { dependencyRepositoryFor } from "./dependencyRepository";

export type QuickTraceUserIntent =
  | Readonly<{ movement: "previous" | "next"; type: "move" }>
  | Readonly<{ type: "left" | "right" }>
  | Readonly<{ direction: DependencyDirection; type: "setDirection" }>
  | Readonly<{ rowId: string; type: "navigate" | "reroot" | "toggleBranch" }>
  | Readonly<{ type: "returnToOrigin" | "useExcelSelection" | "close" }>;

export type QuickTraceController = Readonly<{
  close(): void;
  dispatch(intent: QuickTraceUserIntent): Promise<void>;
  getState(): QuickTraceSnapshot | null;
  isSupported(): boolean;
  open(direction: DependencyDirection): Promise<void>;
  stop(): void;
  subscribe(listener: () => void): () => void;
}>;

type ActiveSession = {
  activeExpansion: AbortController | null;
  cappedByDirection: Partial<Record<DependencyDirection, boolean>>;
  dialog: QuickTraceDialogSession;
  dialogReady: boolean;
  direction: DependencyDirection;
  expansionRequestId: number;
  history: NavigationHistory;
  hostSequence: number;
  identity: QuickTraceSessionIdentity;
  lastDialogSequence: number;
  navigationBlocked: boolean;
  notice: string | null;
  originalDirection: DependencyDirection;
  originFacts: CellFacts;
  pendingExcelSelection: CellRef | null;
  pendingExpectedSelection: Readonly<{ ref: CellRef; sequence: number }> | null;
  recalculationRequestId: number;
  rootFacts: CellFacts;
  selectionSequence: number;
  status: QuickTraceSnapshot["status"];
  trees: Partial<Record<DependencyDirection, QuickTraceTreeState>>;
  unsubscribeClosed: Unsubscribe;
  unsubscribeMessage: Unsubscribe;
  unsubscribeRecalculation: Unsubscribe;
  unsubscribeSelection: Unsubscribe;
};

type QuickTraceControllerOptions = Readonly<{
  dialogHost: QuickTraceDialogHost;
  dialogUrl: string;
  gateway: OfficeGateway;
  identity?: () => QuickTraceSessionIdentity;
}>;

const defaultIdentity = (): QuickTraceSessionIdentity => ({
  nonce: crypto.randomUUID(),
  sessionId: crypto.randomUUID(),
});

const asRef = (ref: CellRef): CellRef => ({
  address: ref.address,
  sheetId: ref.sheetId,
  sheetName: ref.sheetName,
  workbookId: ref.workbookId,
});

const dependencyNode = (ref: DependencyRef): DependencyNode => ({
  id: canonicalCellId(ref),
  ref: asRef(ref),
  ...(ref.unresolved === undefined ? {} : { unresolved: ref.unresolved }),
});

const selectionNode = (facts: CellFacts): DependencyNode => ({
  id: canonicalCellId(facts),
  ref: asRef(facts),
  value: facts.value,
});

const hydrateFacts = async (
  gateway: OfficeGateway,
  refs: readonly CellRef[],
  signal?: AbortSignal,
): Promise<Readonly<Record<string, CellFacts>>> => {
  signal?.throwIfAborted();
  const unique = [...new Map(refs.map((ref) => [canonicalCellId(ref), ref])).values()];
  if (unique.length === 0) return {};
  const loaded = gateway.getCellsFacts === undefined
    ? await Promise.all(unique.map(async (ref) => gateway.getCellFacts?.(ref, signal) ?? null))
    : await gateway.getCellsFacts(unique, signal);
  signal?.throwIfAborted();
  const facts: Record<string, CellFacts> = {};
  unique.forEach((ref, index) => {
    const current = loaded[index];
    if (current !== null && current !== undefined) facts[canonicalCellId(ref)] = current;
  });
  return facts;
};

const buildTree = async (
  gateway: OfficeGateway,
  root: CellFacts,
  direction: DependencyDirection,
  signal?: AbortSignal,
): Promise<Readonly<{ capped: boolean; tree: QuickTraceTreeState }>> => {
  const dependencies = await dependencyRepositoryFor(gateway).get(root, direction, signal);
  signal?.throwIfAborted();
  const expanded = expandNode(
    createGraphState(selectionNode(root)),
    canonicalCellId(root),
    direction,
    dependencies.map(dependencyNode),
  );
  const facts = await hydrateFacts(
    gateway,
    dependencies.filter((dependency) => dependency.unresolved === undefined),
    signal,
  );
  return {
    capped: expanded.capped,
    tree: createQuickTraceTree(expanded.state, direction, {
      ...facts,
      [canonicalCellId(root)]: root,
    }),
  };
};

export const createQuickTraceController = ({
  dialogHost,
  dialogUrl,
  gateway,
  identity: createIdentity = defaultIdentity,
}: QuickTraceControllerOptions): QuickTraceController => {
  const listeners = new Set<() => void>();
  const repository = dependencyRepositoryFor(gateway);
  let active: ActiveSession | null = null;

  const currentTree = (session: ActiveSession): QuickTraceTreeState => {
    const tree = session.trees[session.direction];
    if (tree === undefined) throw new Error("Quick Trace direction is not loaded");
    return tree;
  };

  const snapshotFor = (session: ActiveSession): QuickTraceSnapshot => {
    const tree = currentTree(session);
    const rows = projectQuickTraceRows(tree);
    const selected = rows.find((row) => row.rowId === tree.selectedRowId);
    const selectedFacts = selected === undefined
      ? undefined
      : tree.factsByNodeId[selected.node.id];
    return {
      capped: session.cappedByDirection[session.direction] ?? false,
      direction: session.direction,
      formula: session.rootFacts.formula,
      navigationBlocked: session.navigationBlocked,
      nodeCount: tree.graph.nodes.length,
      notice: session.notice,
      origin: asRef(session.originFacts),
      pendingExcelSelection: session.pendingExcelSelection === null
        ? null
        : asRef(session.pendingExcelSelection),
      root: asRef(session.rootFacts),
      rows,
      selectedFormula: selectedFacts?.formula ?? null,
      selectedRowId: tree.selectedRowId,
      status: session.status,
    };
  };

  const cleanup = (session: ActiveSession, closeFrame: boolean) => {
    if (active !== session) return;
    active = null;
    session.activeExpansion?.abort();
    session.expansionRequestId += 1;
    session.recalculationRequestId += 1;
    session.unsubscribeMessage();
    session.unsubscribeClosed();
    session.unsubscribeSelection();
    session.unsubscribeRecalculation();
    listeners.forEach((listener) => listener());
    if (closeFrame) session.dialog.close();
  };

  const postState = (
    session: ActiveSession,
    type: "quickTrace.initialize" | "quickTrace.state",
  ) => {
    if (active !== session || !session.dialogReady) return;
    const event: QuickTraceHostEvent = {
      ...session.identity,
      payload: { state: snapshotFor(session) },
      protocol: QUICK_TRACE_PROTOCOL_VERSION,
      sequence: ++session.hostSequence,
      type,
    };
    try {
      session.dialog.post(serializeQuickTraceMessage(event));
    } catch {
      cleanup(session, true);
    }
  };

  const publish = (session: ActiveSession) => {
    if (active !== session) return;
    listeners.forEach((listener) => listener());
    postState(session, "quickTrace.state");
  };

  const startExpansionRequest = (session: ActiveSession): Readonly<{
    abort: AbortController;
    requestId: number;
  }> => {
    session.activeExpansion?.abort();
    const abort = new AbortController();
    session.activeExpansion = abort;
    return { abort, requestId: ++session.expansionRequestId };
  };

  const requestIsActive = (
    session: ActiveSession,
    requestId: number,
    abort: AbortController,
  ): boolean => active === session &&
    session.expansionRequestId === requestId &&
    !abort.signal.aborted;

  const loadDirection = async (
    session: ActiveSession,
    direction: DependencyDirection,
  ) => {
    const request = startExpansionRequest(session);
    session.notice = null;
    const cached = session.trees[direction];
    if (cached !== undefined) {
      session.direction = direction;
      session.status = "ready";
      publish(session);
      return;
    }
    session.status = "loading";
    try {
      const loaded = await buildTree(
        gateway,
        session.rootFacts,
        direction,
        request.abort.signal,
      );
      if (!requestIsActive(session, request.requestId, request.abort)) return;
      session.trees[direction] = loaded.tree;
      session.cappedByDirection[direction] = loaded.capped;
      session.direction = direction;
      session.status = "ready";
      publish(session);
    } catch (error) {
      if (!requestIsActive(session, request.requestId, request.abort)) return;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
    }
  };

  const loadRoot = async (
    session: ActiveSession,
    root: CellFacts,
    direction: DependencyDirection,
    commit?: () => void,
  ): Promise<boolean> => {
    const request = startExpansionRequest(session);
    session.status = "loading";
    session.notice = null;
    publish(session);
    try {
      const loaded = await buildTree(gateway, root, direction, request.abort.signal);
      if (!requestIsActive(session, request.requestId, request.abort)) return false;
      session.rootFacts = root;
      session.direction = direction;
      session.trees = { [direction]: loaded.tree };
      session.cappedByDirection = { [direction]: loaded.capped };
      session.status = "ready";
      commit?.();
      publish(session);
      return true;
    } catch (error) {
      if (!requestIsActive(session, request.requestId, request.abort)) return false;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
      return false;
    }
  };

  const factsForRef = async (ref: CellRef): Promise<CellFacts> => {
    if (gateway.getCellFacts !== undefined) return gateway.getCellFacts(ref);
    const selected = await gateway.getSelectionFacts();
    if (canonicalCellId(selected) === canonicalCellId(ref)) return selected;
    throw new Error("Quick Trace cannot read that cell without changing selection");
  };

  const selectPublishedRow = async (
    session: ActiveSession,
    rowId: string,
  ) => {
    if (active !== session || session.navigationBlocked) return;
    const tree = currentTree(session);
    const row = projectQuickTraceRows(tree).find((candidate) => candidate.rowId === rowId);
    if (row === undefined) return;
    session.trees[session.direction] = { ...tree, selectedRowId: row.rowId };
    const expected = {
      ref: asRef(row.node.ref),
      sequence: ++session.selectionSequence,
    };
    session.pendingExpectedSelection = expected;
    session.pendingExcelSelection = null;
    session.notice = null;
    publish(session);
    try {
      await gateway.selectCell(row.node.ref);
    } catch (error) {
      if (active !== session || session.pendingExpectedSelection !== expected) return;
      session.pendingExpectedSelection = null;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
    }
  };

  const expandSelected = async (session: ActiveSession) => {
    if (session.navigationBlocked) return;
    const tree = currentTree(session);
    const transition = moveQuickTraceRight(tree);
    session.trees[session.direction] = transition.state;
    if (transition.loadNodeId === undefined) {
      publish(session);
      return;
    }
    const node = tree.graph.nodes.find((candidate) =>
      candidate.id === transition.loadNodeId);
    if (node === undefined || node.unresolved !== undefined) return;
    const direction = session.direction;
    const request = startExpansionRequest(session);
    session.status = "loading";
    publish(session);
    try {
      const dependencies = await repository.get(node.ref, direction, request.abort.signal);
      if (!requestIsActive(session, request.requestId, request.abort) ||
        session.direction !== direction) return;
      const expanded = expandNode(
        currentTree(session).graph,
        node.id,
        direction,
        dependencies.map(dependencyNode),
      );
      const facts = await hydrateFacts(
        gateway,
        dependencies.filter((dependency) => dependency.unresolved === undefined),
        request.abort.signal,
      );
      if (!requestIsActive(session, request.requestId, request.abort) ||
        session.direction !== direction) return;
      session.trees[direction] = applyQuickTraceExpansion(
        currentTree(session),
        expanded.state,
        facts,
      );
      session.cappedByDirection[direction] =
        (session.cappedByDirection[direction] ?? false) || expanded.capped;
      session.status = "ready";
      session.notice = null;
      publish(session);
    } catch (error) {
      if (!requestIsActive(session, request.requestId, request.abort)) return;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
    }
  };

  const refreshVisibleFacts = async (session: ActiveSession) => {
    const requestId = ++session.recalculationRequestId;
    repository.clear();
    try {
      const tree = currentTree(session);
      const refs = projectQuickTraceRows(tree)
        .filter((row) => row.node.unresolved === undefined)
        .map((row) => row.node.ref);
      const facts = await hydrateFacts(gateway, refs);
      if (active !== session || requestId !== session.recalculationRequestId) return;
      session.trees[session.direction] = { ...tree, factsByNodeId: facts };
      const rootFacts = facts[canonicalCellId(session.rootFacts)];
      if (rootFacts !== undefined) session.rootFacts = rootFacts;
      session.status = "ready";
      session.notice = null;
      publish(session);
    } catch (error) {
      if (active !== session || requestId !== session.recalculationRequestId) return;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
    }
  };

  const handleSelection = async (session: ActiveSession) => {
    if (active !== session) return;
    let facts: CellFacts;
    try {
      facts = await gateway.getSelectionFacts();
    } catch (error) {
      if (active !== session) return;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
      return;
    }
    if (active !== session) return;
    if (facts.workbookId !== session.rootFacts.workbookId) {
      session.navigationBlocked = true;
      session.pendingExcelSelection = null;
      session.pendingExpectedSelection = null;
      session.notice = "Return to the original workbook to continue Quick Trace.";
      publish(session);
      return;
    }
    const selectedId = canonicalCellId(facts);
    if (session.pendingExpectedSelection !== null &&
      canonicalCellId(session.pendingExpectedSelection.ref) === selectedId) {
      session.pendingExpectedSelection = null;
      session.pendingExcelSelection = null;
      session.navigationBlocked = false;
      return;
    }
    if (selectedId === canonicalCellId(session.rootFacts)) {
      session.pendingExcelSelection = null;
      session.navigationBlocked = false;
      publish(session);
      return;
    }
    session.pendingExpectedSelection = null;
    session.pendingExcelSelection = asRef(facts);
    session.navigationBlocked = false;
    session.notice = null;
    publish(session);
  };

  const dispatch = async (intent: QuickTraceUserIntent): Promise<void> => {
    const session = active;
    if (session === null) return;
    if (intent.type === "close") {
      cleanup(session, true);
      return;
    }
    if (intent.type === "move") {
      if (session.navigationBlocked) return;
      const moved = moveQuickTraceSelection(currentTree(session), intent.movement);
      session.trees[session.direction] = moved;
      await selectPublishedRow(session, moved.selectedRowId);
      return;
    }
    if (intent.type === "navigate") {
      await selectPublishedRow(session, intent.rowId);
      return;
    }
    if (intent.type === "left") {
      if (session.navigationBlocked) return;
      session.trees[session.direction] = moveQuickTraceLeft(currentTree(session));
      publish(session);
      return;
    }
    if (intent.type === "right") {
      await expandSelected(session);
      return;
    }
    if (intent.type === "toggleBranch") {
      if (session.navigationBlocked) return;
      const tree = currentTree(session);
      const row = projectQuickTraceRows(tree).find((candidate) =>
        candidate.rowId === intent.rowId);
      if (row === undefined) return;
      session.trees[session.direction] = { ...tree, selectedRowId: row.rowId };
      if (row.disclosure === "expanded") {
        session.trees[session.direction] = moveQuickTraceLeft(
          currentTree(session),
        );
        publish(session);
        return;
      }
      if (row.disclosure === "collapsed") {
        await expandSelected(session);
        return;
      }
      publish(session);
      return;
    }
    if (intent.type === "setDirection") {
      await loadDirection(session, intent.direction);
      return;
    }
    if (intent.type === "reroot") {
      if (session.navigationBlocked) return;
      const row = projectQuickTraceRows(currentTree(session)).find((candidate) =>
        candidate.rowId === intent.rowId);
      if (row === undefined || row.node.unresolved !== undefined) return;
      try {
        const root = await factsForRef(row.node.ref);
        if (active !== session) return;
        const history = visit(session.history, root);
        await loadRoot(session, root, session.direction, () => {
          session.history = history;
          session.pendingExcelSelection = null;
          session.navigationBlocked = false;
        });
      } catch (error) {
        if (active !== session) return;
        session.notice = normalizeOfficeError(error).message;
        session.status = "error";
        publish(session);
      }
      return;
    }
    if (intent.type === "returnToOrigin") {
      const history = returnHistoryToOrigin(session.history);
      await loadRoot(session, session.originFacts, session.originalDirection, () => {
        session.history = history;
        session.pendingExcelSelection = null;
        session.navigationBlocked = false;
      });
      return;
    }
    if (session.pendingExcelSelection === null) return;
    try {
      const root = await factsForRef(session.pendingExcelSelection);
      if (active !== session) return;
      const direction = session.direction;
      await loadRoot(session, root, direction, () => {
        session.originFacts = root;
        session.originalDirection = direction;
        session.history = startHistory(root);
        session.pendingExcelSelection = null;
        session.pendingExpectedSelection = null;
        session.navigationBlocked = false;
      });
    } catch (error) {
      if (active !== session) return;
      session.notice = normalizeOfficeError(error).message;
      session.status = "error";
      publish(session);
    }
  };

  const handleMessage = async (session: ActiveSession, raw: string) => {
    if (active !== session) return;
    let message;
    try {
      message = parseQuickTraceDialogIntent(raw, session.identity);
    } catch {
      return;
    }
    if (message.sequence <= session.lastDialogSequence) return;
    session.lastDialogSequence = message.sequence;
    if (message.type === "quickTrace.ready") {
      session.dialogReady = true;
      postState(session, "quickTrace.initialize");
      return;
    }
    switch (message.type) {
      case "quickTrace.move":
        await dispatch({ movement: message.payload.movement, type: "move" });
        return;
      case "quickTrace.navigate":
        await dispatch({ rowId: message.payload.rowId, type: "navigate" });
        return;
      case "quickTrace.toggleBranch":
        await dispatch({ rowId: message.payload.rowId, type: "toggleBranch" });
        return;
      case "quickTrace.reroot":
        await dispatch({ rowId: message.payload.rowId, type: "reroot" });
        return;
      case "quickTrace.setDirection":
        await dispatch({ direction: message.payload.direction, type: "setDirection" });
        return;
      case "quickTrace.left":
        await dispatch({ type: "left" });
        return;
      case "quickTrace.right":
        await dispatch({ type: "right" });
        return;
      case "quickTrace.returnToOrigin":
        await dispatch({ type: "returnToOrigin" });
        return;
      case "quickTrace.useExcelSelection":
        await dispatch({ type: "useExcelSelection" });
        return;
      case "quickTrace.close":
        await dispatch({ type: "close" });
    }
  };

  const open = async (direction: DependencyDirection): Promise<void> => {
    if (!dialogHost.isSupported()) {
      throw new Error("Quick Trace requires Office DialogApi 1.2");
    }
    if (active !== null) {
      await loadDirection(active, direction);
      return;
    }

    const root = await gateway.getSelectionFacts();
    if (root.address.includes(",")) {
      throw new Error("Quick Trace does not support a multi-area selection");
    }
    const loaded = await buildTree(gateway, root, direction);
    const identity = createIdentity();
    const url = new URL(dialogUrl);
    url.searchParams.set("sessionId", identity.sessionId);
    url.searchParams.set("nonce", identity.nonce);
    const dialog = await dialogHost.open(url.href);
    const placeholder = () => undefined;
    const session: ActiveSession = {
      activeExpansion: null,
      cappedByDirection: { [direction]: loaded.capped },
      dialog,
      dialogReady: false,
      direction,
      expansionRequestId: 0,
      history: startHistory(root),
      hostSequence: 0,
      identity,
      lastDialogSequence: -1,
      navigationBlocked: false,
      notice: null,
      originalDirection: direction,
      originFacts: root,
      pendingExcelSelection: null,
      pendingExpectedSelection: null,
      recalculationRequestId: 0,
      rootFacts: root,
      selectionSequence: 0,
      status: "ready",
      trees: { [direction]: loaded.tree },
      unsubscribeClosed: placeholder,
      unsubscribeMessage: placeholder,
      unsubscribeRecalculation: placeholder,
      unsubscribeSelection: placeholder,
    };
    active = session;
    try {
      session.unsubscribeMessage = dialog.subscribeMessage((message) => {
        void handleMessage(session, message);
      });
      session.unsubscribeClosed = dialog.subscribeClosed(() => cleanup(session, false));
      session.unsubscribeSelection = await gateway.subscribeSelection(() => {
        void handleSelection(session);
      });
      try {
        session.unsubscribeRecalculation = await gateway.subscribeRecalculation(() => {
          void refreshVisibleFacts(session);
        });
      } catch (error) {
        if (normalizeOfficeError(error).code !== "UNSUPPORTED_OPERATION") throw error;
      }
      listeners.forEach((listener) => listener());
    } catch (error) {
      cleanup(session, true);
      throw error;
    }
  };

  return {
    close: () => {
      if (active !== null) cleanup(active, true);
    },
    dispatch,
    getState: () => active === null ? null : snapshotFor(active),
    isSupported: () => dialogHost.isSupported(),
    open,
    stop: () => {
      if (active !== null) cleanup(active, true);
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
};
