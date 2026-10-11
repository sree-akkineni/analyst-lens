import { afterEach, describe, expect, it, vi } from "vitest";
import {
  createController,
  type InspectorController,
} from "../../src/runtime/controller";
import type { OfficeGateway } from "../../src/platform/office/officeGateway";
import { detectCapabilities } from "../../src/platform/office/capabilityMap";
import type { CellFacts, CellRef } from "../../src/shared/types";

const facts = (address: string, formula: string | null = `=${address}`): CellFacts => ({
  address,
  columnCount: 1,
  formula,
  numberFormat: "0.0x",
  rowCount: 1,
  sheetId: "sheet-1",
  sheetName: "DCF",
  value: 12,
  workbookId: "workbook-1",
});

const dependency = (address: string): CellRef => ({
  address,
  sheetId: "sheet-2",
  sheetName: "Inputs",
  workbookId: "workbook-1",
});

const createGateway = (selection: CellFacts = facts("$F$42")) => {
  let selected = selection;
  let recalculationListener: (() => void) | null = null;
  let selectionListener: (() => void) | null = null;
  const selectedRefs: CellRef[] = [];
  let dependencyReads = 0;
  let selectionReads = 0;
  const gateway = {
    getCapabilities: () => detectCapabilities({ isSetSupported: () => true }),
    getDirectDependencies: async (
      _ref: CellRef,
      direction: "precedent" | "dependent",
    ) => {
      dependencyReads += 1;
      return direction === "precedent" ? [dependency("$B$7")] : [dependency("$F$44")];
    },
    getNeighborFacts: async () => [
      { direction: "left" as const, facts: facts("$E$42", "=$B$7*2") },
      { direction: "right" as const, facts: facts("$G$42", "=$B$7*2") },
      { direction: "above" as const, facts: facts("$F$41", null) },
    ],
    getSelectionFacts: async () => {
      selectionReads += 1;
      return selected;
    },
    getWorkbookFacts: async () => ({
      calculationMode: "automatic" as const,
      id: "workbook-1",
      name: "Model.xlsx",
      worksheets: [],
    }),
    runWrite: async () => ({ ok: true as const }),
    selectCell: async (ref: CellRef) => {
      selectedRefs.push(ref);
      selected = facts(ref.address);
    },
    subscribeRecalculation: async (listener: () => void) => {
      recalculationListener = listener;
      return () => {
        recalculationListener = null;
      };
    },
    subscribeSelection: async (listener: () => void) => {
      selectionListener = listener;
      return () => {
        selectionListener = null;
      };
    },
  } satisfies OfficeGateway;
  return {
    emitSelection: () => selectionListener?.(),
    emitRecalculation: () => recalculationListener?.(),
    gateway,
    get dependencyReads() {
      return dependencyReads;
    },
    get selectionReads() {
      return selectionReads;
    },
    selectedRefs,
    setSelection(next: CellFacts) {
      selected = next;
    },
  };
};

afterEach(() => vi.useRealTimers());

describe("Inspector controller", () => {
  it("loads formula facts, tokens, risks, and both direct dependency directions", async () => {
    const fake = createGateway(facts("$F$42", "=$B$7*2"));
    const controller = createController(fake.gateway);

    await controller.start();
    const state = controller.getState();

    expect(state.view).toMatchObject({ cellKind: "formula", kind: "selection" });
    expect(state.tokens).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: "reference", text: "$B$7" }),
    ]));
    expect(state.risks).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: "embedded-number" }),
    ]));
    expect(state.precedents).toMatchObject([{ ref: { address: "$B$7" } }]);
    expect(state.dependents).toMatchObject([{ ref: { address: "$F$44" } }]);
    expect(state.graph?.nodes).toHaveLength(3);
    expect(state.neighbors).toEqual([
      { label: "Left", normalizedFormula: "=R7C2*2", status: "same" },
      { label: "Right", normalizedFormula: "=R7C2*2", status: "same" },
      { label: "Above", normalizedFormula: null, status: "unavailable" },
    ]);
  });

  it("debounces a 20-event selection burst and renders only the newest request", async () => {
    vi.useFakeTimers();
    const fake = createGateway();
    const controller = createController(fake.gateway, { debounceMs: 100 });
    await controller.start();
    fake.setSelection(facts("$G$20"));

    for (let index = 0; index < 20; index += 1) fake.emitSelection();
    await vi.advanceTimersByTimeAsync(99);
    expect(fake.selectionReads).toBe(1);
    await vi.advanceTimersByTimeAsync(1);
    await vi.waitFor(() => expect(fake.selectionReads).toBe(2));

    expect(controller.getState().view).toMatchObject({
      facts: { address: "$G$20" },
      kind: "selection",
    });
  });

  it("suppresses stale async selection results by request id", async () => {
    let resolveFirst: ((value: CellFacts) => void) | undefined;
    let call = 0;
    const fake = createGateway();
    fake.gateway.getSelectionFacts = async () => {
      call += 1;
      if (call === 1) {
        return new Promise<CellFacts>((resolve) => {
          resolveFirst = resolve;
        });
      }
      return facts("$H$30");
    };
    const controller = createController(fake.gateway);

    const first = controller.refresh();
    await Promise.resolve();
    await controller.refresh();
    resolveFirst?.(facts("$A$1"));
    await first;

    expect(controller.getState().view).toMatchObject({
      facts: { address: "$H$30" },
      kind: "selection",
    });
  });

  it("navigates a direct dependency and preserves return-to-origin history", async () => {
    const fake = createGateway();
    const controller: InspectorController = createController(fake.gateway);
    await controller.start();
    const precedent = controller.getState().precedents[0];
    expect(precedent).toBeDefined();

    if (precedent !== undefined) {
      await controller.dispatch({ ref: precedent.ref, type: "navigate" });
    }
    expect(fake.selectedRefs).toMatchObject([{ address: "$B$7" }]);
    expect(controller.getState().history?.origin.address).toBe("$F$42");

    await controller.dispatch({ type: "returnToOrigin" });
    expect(fake.selectedRefs.at(-1)?.address).toBe("$F$42");
  });

  it("preserves a three-level traversal and returns directly to its origin", async () => {
    const fake = createGateway();
    const controller = createController(fake.gateway);
    await controller.start();

    for (const address of ["$B$7", "$C$8", "$D$9"]) {
      await controller.dispatch({ ref: dependency(address), type: "navigate" });
    }

    expect(controller.getState().history).toMatchObject({
      current: { address: "$D$9" },
      origin: { address: "$F$42" },
      past: [
        { address: "$F$42" },
        { address: "$B$7" },
        { address: "$C$8" },
      ],
    });
    await controller.dispatch({ type: "returnToOrigin" });
    expect(fake.selectedRefs.at(-1)?.address).toBe("$F$42");
  });

  it("drops an older expansion result when a newer expansion supersedes it", async () => {
    let resolveOlder: ((value: CellRef[]) => void) | undefined;
    let dependencyReads = 0;
    const fake = createGateway();
    fake.gateway.getDirectDependencies = async (_ref, direction) => {
      dependencyReads += 1;
      if (dependencyReads <= 2) {
        return direction === "precedent" ? [dependency("$B$7")] : [dependency("$F$44")];
      }
      if (dependencyReads === 3) {
        return new Promise<CellRef[]>((resolve) => {
          resolveOlder = resolve;
        });
      }
      return [dependency("$J$9")];
    };
    const controller = createController(fake.gateway);
    await controller.start();
    const parent = controller.getState().precedents[0]?.id;
    expect(parent).toBeDefined();
    if (parent === undefined) return;

    const older = controller.dispatch({ direction: "precedent", nodeId: parent, type: "expandNode" });
    await Promise.resolve();
    await controller.dispatch({ direction: "dependent", nodeId: parent, type: "expandNode" });
    resolveOlder?.([dependency("$A$9")]);
    await older;

    expect(controller.getState().graph?.nodes).toEqual(expect.arrayContaining([
      expect.objectContaining({ ref: expect.objectContaining({ address: "$J$9" }) }),
    ]));
    expect(controller.getState().graph?.nodes).not.toEqual(expect.arrayContaining([
      expect.objectContaining({ ref: expect.objectContaining({ address: "$A$9" }) }),
    ]));
  });

  it("aborts an in-flight expansion when the selected cell changes", async () => {
    let rejectExpansion: ((reason: Error) => void) | undefined;
    let expansionSignal: AbortSignal | undefined;
    let dependencyReads = 0;
    const fake = createGateway();
    const gateway: OfficeGateway = fake.gateway;
    gateway.getDirectDependencies = async (_ref, direction, signal) => {
      dependencyReads += 1;
      if (dependencyReads <= 2 || dependencyReads > 3) {
        return direction === "precedent" ? [dependency("$B$7")] : [dependency("$F$44")];
      }
      expansionSignal = signal;
      return new Promise<CellRef[]>((_resolve, reject) => {
        rejectExpansion = reject;
      });
    };
    const controller = createController(gateway);
    await controller.start();
    const parent = controller.getState().precedents[0]?.id;
    expect(parent).toBeDefined();
    if (parent === undefined) return;

    const expansion = controller.dispatch({ direction: "precedent", nodeId: parent, type: "expandNode" });
    await Promise.resolve();
    fake.setSelection(facts("$H$30"));
    const refresh = controller.refresh();
    expect(expansionSignal?.aborted).toBe(true);
    rejectExpansion?.(new Error("stale expansion failed"));
    await expansion;
    await refresh;

    expect(controller.getState().view).toMatchObject({
      facts: { address: "$H$30" },
      kind: "selection",
    });
    expect(controller.getState().notice).toBeNull();
  });

  it("renders a warm cached selection within the 150 ms interaction budget", async () => {
    const fake = createGateway();
    const controller = createController(fake.gateway);
    await controller.start();
    const root = controller.getState().graph?.origin.id;
    expect(fake.dependencyReads).toBe(2);

    const startedAt = performance.now();
    await controller.refresh();
    if (root !== undefined) {
      await controller.dispatch({ direction: "precedent", nodeId: root, type: "expandNode" });
    }

    expect(performance.now() - startedAt).toBeLessThan(150);
    expect(fake.dependencyReads).toBe(2);
    expect(controller.getState().view.kind).toBe("selection");
  });

  it("invalidates formula-scoped caches after workbook recalculation", async () => {
    vi.useFakeTimers();
    const fake = createGateway();
    const controller = createController(fake.gateway, { debounceMs: 100 });
    await controller.start();
    await controller.refresh();
    expect(fake.dependencyReads).toBe(2);

    fake.emitRecalculation();
    await vi.advanceTimersByTimeAsync(100);
    await vi.waitFor(() => expect(fake.dependencyReads).toBe(4));

    expect(controller.getState().view.kind).toBe("selection");
  });

  it("propagates the graph cap from oversized direct dependency reads", async () => {
    const fake = createGateway();
    fake.gateway.getDirectDependencies = async (_ref, direction) =>
      direction === "precedent"
        ? Array.from({ length: 600 }, (_, index) => dependency(`$A$${index + 1}`))
        : [];
    const controller = createController(fake.gateway);

    await controller.start();

    expect(controller.getState().graph?.nodes).toHaveLength(500);
    expect(controller.getState().capped).toBe(true);

    const root = controller.getState().graph?.origin.id;
    if (root !== undefined) {
      await controller.dispatch({ direction: "precedent", nodeId: root, type: "collapseBranch" });
    }
    expect(controller.getState().capped).toBe(false);
  });

  it("does not call an exactly 500-node graph capped after collapsing an empty branch", async () => {
    const fake = createGateway();
    fake.gateway.getDirectDependencies = async (_ref, direction) =>
      direction === "precedent"
        ? Array.from({ length: 499 }, (_, index) => dependency(`$A$${index + 1}`))
        : [];
    const controller = createController(fake.gateway);
    await controller.start();
    const root = controller.getState().graph?.origin.id;

    expect(controller.getState().graph?.nodes).toHaveLength(500);
    expect(controller.getState().capped).toBe(false);
    if (root !== undefined) {
      await controller.dispatch({ direction: "dependent", nodeId: root, type: "collapseBranch" });
    }
    expect(controller.getState().graph?.nodes).toHaveLength(500);
    expect(controller.getState().capped).toBe(false);
  });

  it("removes a late selection subscription after stop", async () => {
    let releaseSubscription: ((unsubscribe: () => void) => void) | undefined;
    let unsubscribes = 0;
    const fake = createGateway();
    fake.gateway.subscribeSelection = async () =>
      new Promise<() => void>((resolve) => {
        releaseSubscription = resolve;
      });
    const controller = createController(fake.gateway);

    const starting = controller.start();
    controller.stop();
    releaseSubscription?.(() => {
      unsubscribes += 1;
    });
    await starting;

    expect(unsubscribes).toBe(1);
    expect(fake.selectionReads).toBe(0);
  });

  it("ignores an Office event delivered while asynchronous unsubscribe is draining", async () => {
    vi.useFakeTimers();
    let listener: (() => void) | undefined;
    const fake = createGateway();
    fake.gateway.subscribeSelection = async (next) => {
      listener = next;
      return () => undefined;
    };
    const controller = createController(fake.gateway, { debounceMs: 100 });
    await controller.start();
    expect(fake.selectionReads).toBe(1);

    controller.stop();
    listener?.();
    await vi.advanceTimersByTimeAsync(100);

    expect(fake.selectionReads).toBe(1);
  });

  it("normalizes unsupported and multi-selection states without dependency reads", async () => {
    const multi = createGateway({ ...facts("$A$1:$B$2"), columnCount: 2, rowCount: 2 });
    multi.gateway.getDirectDependencies = async () => {
      throw new Error("dependencies should not load");
    };
    const controller = createController(multi.gateway);
    await controller.start();
    expect(controller.getState().view).toMatchObject({ cellKind: "multi", kind: "selection" });

    const unsupported = createGateway();
    unsupported.gateway.getSelectionFacts = async () => {
      throw { code: "UnsupportedOperation" };
    };
    const unsupportedController = createController(unsupported.gateway);
    await unsupportedController.start();
    expect(unsupportedController.getState().view).toMatchObject({ kind: "unsupported" });
  });

  it("pins the current cell and refreshes value and formula changes independently", async () => {
    const fake = createGateway(facts("$F$42", "=IRR(B2:B8)"));
    let observed = { ...facts("$F$42", "=IRR(B2:B8)"), value: 18.4 };
    Object.assign(fake.gateway, {
      getCellFacts: async () => observed,
    });
    const controller = createController(fake.gateway);
    await controller.start();

    await controller.dispatch({ type: "pinCurrent" });
    expect(controller.getState().watchlist.items).toMatchObject([{
      currentValue: 12,
      formulaChanged: false,
      valueChanged: false,
    }]);

    observed = { ...observed, value: 19.1 };
    await controller.dispatch({ type: "refreshWatchlist" });
    expect(controller.getState().watchlist.items[0]).toMatchObject({
      currentValue: 19.1,
      formulaChanged: false,
      lastValue: 12,
      valueChanged: true,
    });

    observed = { ...observed, formula: "=XIRR(B2:B8,A2:A8)" };
    await controller.dispatch({ type: "refreshWatchlist" });
    expect(controller.getState().watchlist.items[0]).toMatchObject({
      formulaChanged: true,
      valueChanged: false,
    });
  });

  it("refreshes pins after recalculation and exposes exact current-selection impact", async () => {
    vi.useFakeTimers();
    const output = {
      ...facts("$F$44", "=$B$7*2"),
      sheetId: "sheet-2",
      sheetName: "Inputs",
    };
    const fake = createGateway(output);
    let pinnedReadCount = 0;
    Object.assign(fake.gateway, {
      getCellFacts: async (current: CellRef) => {
        pinnedReadCount += 1;
        return {
          ...facts(current.address, "=$B$7*2"),
          ...current,
          value: 20 + pinnedReadCount,
        };
      },
    });
    const controller = createController(fake.gateway, { debounceMs: 100 });
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    fake.setSelection(facts("$F$42", "=$B$7*2"));
    await controller.refresh();

    expect(controller.getState().watchlist.affected).toEqual(["Inputs!$F$44"]);

    fake.emitRecalculation();
    await vi.waitFor(() =>
      expect(controller.getState().watchlist.items[0]?.currentValue).toBe(21));
    expect(pinnedReadCount).toBe(1);
    expect(controller.getState().watchlist.affected).toEqual(["Inputs!$F$44"]);
  });

  it("persists only metadata after explicit save intent", async () => {
    const saved: unknown[] = [];
    const fake = createGateway();
    Object.assign(fake.gateway, {
      loadWatchlistMetadata: async () => [],
      saveWatchlistMetadata: async (metadata: unknown) => {
        saved.push(metadata);
      },
    });
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    const item = controller.getState().watchlist.items[0];
    expect(item).toBeDefined();
    if (item === undefined) return;
    await controller.dispatch({
      category: "Returns",
      id: item.id,
      label: "IRR",
      type: "updateWatchMetadata",
    });

    await controller.dispatch({ type: "saveWatchlistWithWorkbook" });

    expect(saved).toEqual([[{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "sheet-1",
      workbookId: "workbook-1",
    }]]);
    expect(JSON.stringify(saved)).not.toMatch(/value|formula|numberFormat|sheetName/u);
    expect(controller.getState().watchlist.persistence).toBe("saved");
  });

  it("rehydrates a renamed sheet by stable ID and preserves missing pins", async () => {
    const fake = createGateway();
    Object.assign(fake.gateway, {
      getCellFacts: async (current: CellRef) => {
        if (current.address === "$A$1") throw { code: "ItemNotFound" };
        return {
          ...facts(current.address, "=B1"),
          sheetId: current.sheetId,
          sheetName: "Renamed Returns",
        };
      },
      loadWatchlistMetadata: async () => [
        {
          address: "$F$42",
          category: "Returns",
          label: "IRR",
          sheetId: "stable-sheet-id",
          workbookId: "workbook-1",
        },
        {
          address: "$A$1",
          category: "Checks",
          label: "Missing check",
          sheetId: "missing-sheet-id",
          workbookId: "workbook-1",
        },
      ],
      saveWatchlistMetadata: async () => undefined,
    });
    const controller = createController(fake.gateway);

    await controller.start();

    expect(controller.getState().watchlist.items).toEqual(expect.arrayContaining([
      expect.objectContaining({
        label: "IRR",
        ref: expect.objectContaining({ sheetName: "Renamed Returns" }),
        status: "ready",
      }),
      expect.objectContaining({
        label: "Missing check",
        status: "missing",
      }),
    ]));
  });

  it("preserves metadata edits made during an in-flight watch refresh", async () => {
    let release: ((next: CellFacts) => void) | undefined;
    const fake = createGateway();
    Object.assign(fake.gateway, {
      getCellFacts: async () => new Promise<CellFacts>((resolve) => {
        release = resolve;
      }),
    });
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    const item = controller.getState().watchlist.items[0];
    expect(item).toBeDefined();
    if (item === undefined) return;

    const refreshing = controller.dispatch({ type: "refreshWatchlist" });
    await vi.waitFor(() => expect(release).toBeDefined());
    await controller.dispatch({
      category: "Returns",
      id: item.id,
      label: "Edited during refresh",
      type: "updateWatchMetadata",
    });
    release?.({ ...facts("$F$42"), value: 25 });
    await refreshing;

    expect(controller.getState().watchlist.items[0]).toMatchObject({
      category: "Returns",
      currentValue: 25,
      label: "Edited during refresh",
    });
  });

  it("does not publish persisted metadata that resolves after stop", async () => {
    let release: ((metadata: readonly {
      address: string;
      category: string;
      label: string;
      sheetId: string;
      workbookId: string;
    }[]) => void) | undefined;
    const fake = createGateway();
    Object.assign(fake.gateway, {
      loadWatchlistMetadata: async () => new Promise((resolve) => {
        release = resolve;
      }),
      saveWatchlistMetadata: async () => undefined,
    });
    const controller = createController(fake.gateway);

    const starting = controller.start();
    await vi.waitFor(() => expect(release).toBeDefined());
    controller.stop();
    release?.([{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "sheet-1",
      workbookId: "workbook-1",
    }]);
    await starting;

    expect(controller.getState().watchlist.items).toEqual([]);
  });

  it("keeps explicit watch refresh available when recalculation events are unsupported", async () => {
    const fake = createGateway();
    Object.assign(fake.gateway, {
      getCellFacts: async (current: CellRef) => ({
        ...facts(current.address),
        value: 30,
      }),
      subscribeRecalculation: async () => {
        throw { code: "UnsupportedOperation" };
      },
    });
    const controller = createController(fake.gateway);

    await controller.start();
    expect(controller.getState().view.kind).toBe("selection");
    await controller.dispatch({ type: "pinCurrent" });
    await controller.dispatch({ type: "refreshWatchlist" });

    expect(controller.getState().watchlist.items[0]?.currentValue).toBe(30);
  });

  it("refreshes multiple pins through one bounded gateway batch", async () => {
    const fake = createGateway();
    let batchReads = 0;
    Object.assign(fake.gateway, {
      getCellsFacts: async (refs: readonly CellRef[]) => {
        batchReads += 1;
        return refs.map((ref, index) => ({
          ...facts(ref.address),
          ...ref,
          value: 40 + index,
        }));
      },
    });
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    fake.setSelection(facts("$G$43"));
    await controller.refresh();
    await controller.dispatch({ type: "pinCurrent" });

    await controller.dispatch({ type: "refreshWatchlist" });

    expect(batchReads).toBe(1);
    expect(controller.getState().watchlist.items.map((item) => item.currentValue))
      .toEqual([40, 41]);
  });

  it("does not advertise workbook persistence when document settings are unavailable", async () => {
    const fake = createGateway();
    const controller = createController(fake.gateway);
    await controller.start();

    await controller.dispatch({ type: "pinCurrent" });

    expect(controller.getState().watchlist.persistence).toBe("unavailable");
  });

  it("resolves multi-hop affected pins without requiring visible graph expansion", async () => {
    const output = {
      ...facts("$F$42"),
      sheetId: "output-sheet",
      sheetName: "Returns",
    };
    const bridge = {
      address: "$D$20",
      sheetId: "bridge-sheet",
      sheetName: "DCF",
      workbookId: "workbook-1",
    };
    const fake = createGateway(output);
    fake.gateway.getDirectDependencies = async (current, direction) => {
      if (direction === "precedent") return [];
      if (current.sheetId === "sheet-1" && current.address === "$B$7") return [bridge];
      if (current.sheetId === "bridge-sheet" && current.address === "$D$20") return [output];
      return [];
    };
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    fake.setSelection(facts("$B$7"));

    await controller.refresh();

    expect(controller.getState().graph?.nodes).not.toEqual(expect.arrayContaining([
      expect.objectContaining({ ref: expect.objectContaining({ sheetId: "output-sheet" }) }),
    ]));
    expect(controller.getState().watchlist.affected).toEqual(["Returns!$F$42"]);
  });

  it("matches pinned cells contained by Excel's coalesced dependent ranges", async () => {
    const output = {
      ...facts("$AE$16"),
      sheetId: "output-sheet",
      sheetName: "OpEx & Headcount Schedule",
    };
    const assumption = {
      ...facts("$C$22", null),
      sheetId: "assumption-sheet",
      sheetName: "Model",
    };
    const fake = createGateway(output);
    fake.gateway.getDirectDependencies = async (current, direction) => {
      if (direction === "precedent") return [];
      if (current.sheetId === "assumption-sheet" && current.address === "$C$22") {
        return [{
          address: "$H$10:$AE$30",
          sheetId: "output-sheet",
          sheetName: "OpEx & Headcount Schedule",
          workbookId: "workbook-1",
        }];
      }
      return [];
    };
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    fake.setSelection(assumption);

    await controller.refresh();

    expect(controller.getState().watchlist.affected)
      .toEqual(["OpEx & Headcount Schedule!$AE$16"]);
  });

  it("does not silently truncate exact impact traversal after 500 cells", async () => {
    const output = {
      ...facts("$F$42"),
      sheetId: "output-sheet",
      sheetName: "Returns",
    };
    const fake = createGateway(output);
    fake.gateway.getDirectDependencies = async (current, direction) => {
      if (direction === "precedent") return [];
      if (current.sheetId === "sheet-1" && current.address === "$B$7") {
        return [{
          address: "$A$1",
          sheetId: "chain-sheet",
          sheetName: "Chain",
          workbookId: "workbook-1",
        }];
      }
      if (current.sheetId !== "chain-sheet") return [];
      const row = Number.parseInt(current.address.replace(/\D/gu, ""), 10);
      if (row < 505) {
        return [{
          address: `$A$${row + 1}`,
          sheetId: "chain-sheet",
          sheetName: "Chain",
          workbookId: "workbook-1",
        }];
      }
      return [output];
    };
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    fake.setSelection(facts("$B$7"));

    await controller.refresh();

    expect(controller.getState().watchlist.affected).toEqual(["Returns!$F$42"]);
  });

  it("does not mark metadata edits made during save as persisted", async () => {
    let releaseSave: (() => void) | undefined;
    const fake = createGateway();
    Object.assign(fake.gateway, {
      loadWatchlistMetadata: async () => [],
      saveWatchlistMetadata: async () => new Promise<void>((resolve) => {
        releaseSave = resolve;
      }),
    });
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });
    const item = controller.getState().watchlist.items[0];
    expect(item).toBeDefined();
    if (item === undefined) return;

    const saving = controller.dispatch({ type: "saveWatchlistWithWorkbook" });
    await vi.waitFor(() => expect(releaseSave).toBeDefined());
    await controller.dispatch({
      category: "Returns",
      id: item.id,
      label: "Edited after save started",
      type: "updateWatchMetadata",
    });
    releaseSave?.();
    await saving;

    expect(controller.getState().watchlist.persistence).toBe("session");
    expect(controller.getState().watchlist.items[0]?.label).toBe("Edited after save started");
  });

  it("ignores a save completion delivered after stop", async () => {
    let releaseSave: (() => void) | undefined;
    const fake = createGateway();
    Object.assign(fake.gateway, {
      loadWatchlistMetadata: async () => [],
      saveWatchlistMetadata: async () => new Promise<void>((resolve) => {
        releaseSave = resolve;
      }),
    });
    const controller = createController(fake.gateway);
    await controller.start();
    await controller.dispatch({ type: "pinCurrent" });

    const saving = controller.dispatch({ type: "saveWatchlistWithWorkbook" });
    await vi.waitFor(() => expect(releaseSave).toBeDefined());
    controller.stop();
    releaseSave?.();
    await saving;

    expect(controller.getState().watchlist.persistence).not.toBe("saved");
  });
});
