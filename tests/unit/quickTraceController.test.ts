import { describe, expect, it, vi } from "vitest";
import type {
  QuickTraceDialogHost,
  QuickTraceDialogSession,
} from "../../src/platform/office/quickTraceDialogHost";
import type { OfficeGateway } from "../../src/platform/office/officeGateway";
import { createQuickTraceController } from "../../src/runtime/quickTraceController";
import type {
  QuickTraceDialogIntent,
  QuickTraceHostEvent,
  QuickTraceSessionIdentity,
  QuickTraceSnapshot,
} from "../../src/shared/quickTraceProtocol";
import type { CellFacts, DependencyRef } from "../../src/shared/types";
import {
  sampleModelScenario,
  createSimulatedWorkbookFixture,
} from "../fixtures/simulatedWorkbook";

const flush = async () => {
  await new Promise((resolve) => setTimeout(resolve, 0));
};

const createInMemoryDialogHost = () => {
  const messages: string[] = [];
  const urls: string[] = [];
  let current: QuickTraceDialogSession | null = null;
  let closeListeners = new Set<(error: number) => void>();
  let messageListeners = new Set<(message: string, origin?: string) => void>();
  let postError: Error | null = null;
  let opens = 0;

  const finish = (error: number) => {
    if (current === null) return;
    current = null;
    closeListeners.forEach((listener) => listener(error));
    closeListeners = new Set();
    messageListeners = new Set();
  };
  const host: QuickTraceDialogHost = {
    isSupported: () => true,
    open: async (url) => {
      if (current !== null) return current;
      opens += 1;
      urls.push(url);
      current = {
        close: () => finish(0),
        post: (message) => {
          if (postError !== null) {
            const error = postError;
            postError = null;
            throw error;
          }
          messages.push(message);
        },
        subscribeClosed(listener) {
          closeListeners.add(listener);
          return () => closeListeners.delete(listener);
        },
        subscribeMessage(listener) {
          messageListeners.add(listener);
          return () => messageListeners.delete(listener);
        },
      };
      return current;
    },
  };

  return {
    closeNative: (error = 12006) => finish(error),
    failNextPost: (error: Error) => {
      postError = error;
    },
    host,
    get openCount() {
      return opens;
    },
    latestState: (): QuickTraceSnapshot => {
      const event = messages
        .slice()
        .reverse()
        .map((message) => JSON.parse(message) as QuickTraceHostEvent)
        .find((message) =>
          message.type === "quickTrace.initialize" || message.type === "quickTrace.state"
        );
      if (event?.type !== "quickTrace.initialize" && event?.type !== "quickTrace.state") {
        throw new Error("expected Quick Trace state message");
      }
      return event.payload.state;
    },
    messages,
    sendFromChild: (message: QuickTraceDialogIntent) => {
      messageListeners.forEach((listener) => listener(JSON.stringify(message)));
    },
    urls,
  };
};

const createTrackedGateway = (
  initial: CellFacts = sampleModelScenario.states.formula,
) => {
  const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
  const recalculationListeners = new Set<() => void>();
  const selectionListeners = new Set<() => void>();
  let selection = initial;
  const gateway: OfficeGateway = {
    ...fixture.gateway,
    getSelectionFacts: async (signal) => {
      signal?.throwIfAborted();
      return selection;
    },
    selectCell: async (ref) => {
      const facts = await fixture.gateway.getCellFacts?.(ref);
      if (facts === undefined) throw new Error("Expected fixture cell facts");
      selection = facts;
      queueMicrotask(() => selectionListeners.forEach((listener) => listener()));
    },
    subscribeRecalculation: async (listener) => {
      recalculationListeners.add(listener);
      return () => recalculationListeners.delete(listener);
    },
    subscribeSelection: async (listener) => {
      selectionListeners.add(listener);
      return () => selectionListeners.delete(listener);
    },
  };
  return {
    activeRecalculationSubscriptions: () => recalculationListeners.size,
    activeSelectionSubscriptions: () => selectionListeners.size,
    emitRecalculation: () => recalculationListeners.forEach((listener) => listener()),
    gateway,
    manualSelect: (facts: CellFacts) => {
      selection = facts;
      selectionListeners.forEach((listener) => listener());
    },
    selectedAddress: () => selection.address,
  };
};

const identity = (suffix = "1"): QuickTraceSessionIdentity => ({
  nonce: `nonce-${suffix}`,
  sessionId: `session-${suffix}`,
});

const ready = (suffix = "1", sequence = 1): QuickTraceDialogIntent => ({
  ...identity(suffix),
  protocol: 1,
  sequence,
  type: "quickTrace.ready",
});

describe("Quick Trace controller spike", () => {
  it("opens from the active cell and navigates only to published dependency rows", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(),
    });

    await controller.open("precedent");
    const dependencyRowId = controller.getState()?.rows.find((row) => row.depth === 1)?.rowId;
    if (dependencyRowId === undefined) throw new Error("expected dependency row");
    dialog.sendFromChild(ready());
    dialog.sendFromChild({
      ...identity(),
      payload: { rowId: dependencyRowId },
      protocol: 1,
      sequence: 2,
      type: "quickTrace.navigate",
    });

    await vi.waitFor(async () => {
      expect((await fixture.gateway.getSelectionFacts()).address).toBe("$B$7");
    });
    const before = (await fixture.gateway.getSelectionFacts()).address;
    dialog.sendFromChild({
      ...identity(),
      payload: { rowId: "not-published" },
      protocol: 1,
      sequence: 3,
      type: "quickTrace.navigate",
    });
    await flush();
    expect((await fixture.gateway.getSelectionFacts()).address).toBe(before);
  });

  it("accepts a contiguous range root but rejects a multi-area root before opening", async () => {
    const rangeFixture = createSimulatedWorkbookFixture({ initialState: "multi" });
    const rangeDialog = createInMemoryDialogHost();
    const rangeController = createQuickTraceController({
      dialogHost: rangeDialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: rangeFixture.gateway,
      identity: () => identity(),
    });
    await rangeController.open("precedent");
    rangeDialog.sendFromChild(ready());
    expect(rangeDialog.latestState().rows[0]?.displayValue).toBe("2 × 2 range");

    const multiAreaDialog = createInMemoryDialogHost();
    const dependencyRead = vi.fn<OfficeGateway["getDirectDependencies"]>();
    const multiAreaFacts: CellFacts = {
      ...sampleModelScenario.states.multi,
      address: "$A$1,$B$2",
    };
    const multiAreaGateway: OfficeGateway = {
      ...rangeFixture.gateway,
      getDirectDependencies: dependencyRead,
      getSelectionFacts: async () => multiAreaFacts,
    };
    const multiAreaController = createQuickTraceController({
      dialogHost: multiAreaDialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: multiAreaGateway,
      identity: () => identity("2"),
    });

    await expect(multiAreaController.open("precedent")).rejects.toThrow(/multi-area/iu);
    expect(multiAreaDialog.openCount).toBe(0);
    expect(dependencyRead).not.toHaveBeenCalled();
  });

  it("collapses and re-expands the loaded root without moving Excel", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    dialog.sendFromChild(ready());
    const originalCount = dialog.latestState().rows.length;

    dialog.sendFromChild({ ...identity(), protocol: 1, sequence: 2, type: "quickTrace.left" });
    await flush();
    expect(dialog.latestState().rows).toHaveLength(1);
    dialog.sendFromChild({ ...identity(), protocol: 1, sequence: 3, type: "quickTrace.right" });
    await flush();

    expect(dialog.latestState().rows).toHaveLength(originalCount);
    expect((await fixture.gateway.getSelectionFacts()).address).toBe("$F$42");
  });

  it("switches direction in the same dialog without changing origin", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    dialog.sendFromChild(ready());
    dialog.sendFromChild({
      ...identity(),
      payload: { direction: "dependent" },
      protocol: 1,
      sequence: 2,
      type: "quickTrace.setDirection",
    });
    await vi.waitFor(() => expect(dialog.latestState().direction).toBe("dependent"));

    expect(dialog.openCount).toBe(1);
    expect(dialog.latestState()).toMatchObject({
      direction: "dependent",
      origin: { address: "$F$42" },
    });
  });

  it("reuses the active identity when an open command arrives while the dialog is live", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    let identities = 0;
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(String(++identities)),
    });

    await controller.open("precedent");
    dialog.sendFromChild(ready());
    await controller.open("dependent");

    expect(dialog.openCount).toBe(1);
    expect(identities).toBe(1);
    const openedUrl = dialog.urls[0];
    if (openedUrl === undefined) throw new Error("expected Quick Trace URL");
    expect(new URL(openedUrl).searchParams.get("sessionId")).toBe("session-1");
    expect(dialog.messages.map((raw) => JSON.parse(raw) as QuickTraceHostEvent).at(-1))
      .toMatchObject({ nonce: "nonce-1", sessionId: "session-1", type: "quickTrace.state" });
  });

  it("preserves the tree for a correlated dialog selection event", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    const dependencyRowId = controller.getState()?.rows.find((row) => row.depth === 1)?.rowId;
    if (dependencyRowId === undefined) throw new Error("expected dependency row");
    dialog.sendFromChild(ready());
    const rowIds = dialog.latestState().rows.map((row) => row.rowId);
    dialog.sendFromChild({
      ...identity(),
      payload: { rowId: dependencyRowId },
      protocol: 1,
      sequence: 2,
      type: "quickTrace.navigate",
    });
    await flush();

    expect(dialog.latestState()).toMatchObject({
      origin: { address: "$F$42" },
      pendingExcelSelection: null,
      selectedRowId: dependencyRowId,
    });
    expect(dialog.latestState().rows.map((row) => row.rowId)).toEqual(rowIds);
  });

  it("publishes an unrelated Excel selection without replacing the tree", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    dialog.sendFromChild(ready());
    const rowIds = dialog.latestState().rows.map((row) => row.rowId);
    await fixture.gateway.selectCell(sampleModelScenario.states.hardcode);
    await flush();

    expect(dialog.latestState()).toMatchObject({
      origin: { address: "$F$42" },
      pendingExcelSelection: { address: "$B$7" },
    });
    expect(dialog.latestState().rows.map((row) => row.rowId)).toEqual(rowIds);
  });

  it("closes and reopens with fresh identities and ignores stale messages", async () => {
    const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
    const dialog = createInMemoryDialogHost();
    let identities = 0;
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: fixture.gateway,
      identity: () => identity(String(++identities)),
    });

    await controller.open("precedent");
    dialog.sendFromChild(ready("1"));
    dialog.sendFromChild({ ...identity("1"), protocol: 1, sequence: 2, type: "quickTrace.close" });
    await controller.open("precedent");
    dialog.sendFromChild(ready("2"));
    const secondState = dialog.latestState();
    dialog.sendFromChild({ ...identity("1"), protocol: 1, sequence: 3, type: "quickTrace.left" });
    await flush();
    expect(dialog.latestState()).toEqual(secondState);

    dialog.closeNative();
    await controller.open("dependent");
    dialog.sendFromChild(ready("3"));
    expect(dialog.latestState()).toMatchObject({ direction: "dependent" });
    expect(identities).toBe(3);
  });

  it("expands only the selected node one level on right", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    await controller.dispatch({ movement: "next", type: "move" });
    await controller.dispatch({ movement: "next", type: "move" });
    const selected = controller.getState()?.selectedRowId;
    await controller.dispatch({ type: "right" });

    expect(controller.getState()?.rows.length).toBeGreaterThan(6);
    expect(controller.getState()?.selectedRowId).toBe(selected);
  });

  it("switches direction without changing the session origin or root", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    const before = controller.getState();
    await controller.dispatch({ direction: "dependent", type: "setDirection" });

    expect(controller.getState()).toMatchObject({
      direction: "dependent",
      origin: before?.origin,
      root: before?.root,
    });
  });

  it("reroots on the selected row but returns to the original origin and direction", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    await controller.dispatch({ movement: "next", type: "move" });
    const rowId = controller.getState()?.selectedRowId;
    if (rowId === undefined) throw new Error("expected selected row");
    await controller.dispatch({ rowId, type: "reroot" });
    expect(controller.getState()?.root.address).toBe("$B$7");
    await controller.dispatch({ direction: "dependent", type: "setDirection" });
    await controller.dispatch({ type: "returnToOrigin" });

    expect(controller.getState()).toMatchObject({
      direction: "precedent",
      root: { address: "$F$42" },
    });
  });

  it("correlates dialog navigation and preserves the tree", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    const before = controller.getState()?.rows.map((row) => row.rowId);
    await controller.dispatch({ movement: "next", type: "move" });
    await flush();

    expect(controller.getState()?.rows.map((row) => row.rowId)).toEqual(before);
    expect(controller.getState()?.pendingExcelSelection).toBeNull();
  });

  it("offers and accepts an unrelated same-workbook selection as a new root", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    tracked.manualSelect(sampleModelScenario.states.blank);
    await flush();
    expect(controller.getState()).toMatchObject({
      origin: { address: "$F$42" },
      pendingExcelSelection: { address: "$A$1" },
    });
    await controller.dispatch({ type: "useExcelSelection" });

    expect(controller.getState()).toMatchObject({
      origin: { address: "$A$1" },
      pendingExcelSelection: null,
      root: { address: "$A$1" },
    });
  });

  it("freezes stale navigation after the active workbook changes", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    tracked.manualSelect({
      ...sampleModelScenario.states.blank,
      workbookId: "fixture://other-workbook",
    });
    await flush();
    expect(controller.getState()?.navigationBlocked).toBe(true);
    const before = tracked.selectedAddress();
    await controller.dispatch({ movement: "next", type: "move" });

    expect(tracked.selectedAddress()).toBe(before);
  });

  it("stops cleanly when dialog message delivery fails", async () => {
    const tracked = createTrackedGateway();
    const dialog = createInMemoryDialogHost();
    const controller = createQuickTraceController({
      dialogHost: dialog.host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    dialog.failNextPost(new Error("Dialog messaging failed"));
    await controller.open("precedent");
    dialog.sendFromChild(ready());
    await flush();

    expect(controller.getState()).toBeNull();
    expect(tracked.activeSelectionSubscriptions()).toBe(0);
    expect(tracked.activeRecalculationSubscriptions()).toBe(0);
  });

  it("discards a stale expansion after a newer direction request", async () => {
    const tracked = createTrackedGateway();
    const originalRead = tracked.gateway.getDirectDependencies.bind(tracked.gateway);
    let resolveOlder: ((value: DependencyRef[]) => void) | undefined;
    tracked.gateway.getDirectDependencies = async (ref, direction, signal) => {
      if (ref.address === "$D$20" && direction === "precedent") {
        return new Promise<DependencyRef[]>((resolve) => {
          resolveOlder = resolve;
          signal?.addEventListener("abort", () => undefined, { once: true });
        });
      }
      return originalRead(ref, direction, signal);
    };
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    await controller.dispatch({ movement: "next", type: "move" });
    await controller.dispatch({ movement: "next", type: "move" });
    const older = controller.dispatch({ type: "right" });
    await Promise.resolve();
    await controller.dispatch({ direction: "dependent", type: "setDirection" });
    const stale: DependencyRef = {
      ...sampleModelScenario.states.blank,
      address: "$A$99",
    };
    resolveOlder?.([stale]);
    await older;

    expect(controller.getState()?.direction).toBe("dependent");
    expect(controller.getState()?.rows.some((row) => row.node.id.includes("$A$99")))
      .toBe(false);
  });

  it("refreshes visible values on recalculation without changing expansion state", async () => {
    const tracked = createTrackedGateway();
    const originalBatch = tracked.gateway.getCellsFacts?.bind(tracked.gateway);
    if (originalBatch === undefined) throw new Error("expected batch facts reader");
    let rootValue = sampleModelScenario.states.formula.value;
    tracked.gateway.getCellsFacts = async (refs, signal) => {
      const loaded = await originalBatch(refs, signal);
      return loaded.map((facts) => facts?.address === "$F$42"
        ? { ...facts, value: rootValue }
        : facts);
    };
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    const rowIds = controller.getState()?.rows.map((row) => row.rowId);
    rootValue = 0.25;
    tracked.emitRecalculation();

    await vi.waitFor(() => expect(controller.getState()?.rows[0]?.displayValue).toBe("0.25"));
    expect(controller.getState()?.rows.map((row) => row.rowId)).toEqual(rowIds);
  });

  it("toggles a branch occurrence without changing origin or root", async () => {
    const tracked = createTrackedGateway();
    const controller = createQuickTraceController({
      dialogHost: createInMemoryDialogHost().host,
      dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
      gateway: tracked.gateway,
      identity: () => identity(),
    });
    await controller.open("precedent");
    await controller.dispatch({ movement: "next", type: "move" });
    await controller.dispatch({ movement: "next", type: "move" });
    const before = controller.getState();
    const rowId = before?.selectedRowId;
    if (rowId === undefined) throw new Error("expected selected branch");
    await controller.dispatch({ rowId, type: "toggleBranch" });

    expect(controller.getState()).toMatchObject({
      origin: before?.origin,
      root: before?.root,
    });
    expect(controller.getState()?.rows.length).toBeGreaterThan(before?.rows.length ?? 0);
    await controller.dispatch({ rowId, type: "toggleBranch" });
    expect(controller.getState()?.rows).toHaveLength(before?.rows.length ?? 0);
  });
});
