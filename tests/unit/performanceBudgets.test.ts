import { createElement, type ComponentType } from "react";
import {
  cleanup,
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { createController } from "../../src/runtime/controller";
import type { QuickTraceDialogHost } from "../../src/platform/office/quickTraceDialogHost";
import type { OfficeGateway } from "../../src/platform/office/officeGateway";
import { createQuickTraceController } from "../../src/runtime/quickTraceController";
import { canonicalCellId, type CellFacts, type CellRef } from "../../src/shared/types";
import { App, type AppProps } from "../../src/ui/App";
import {
  sampleModelScenario,
  createSimulatedWorkbookFixture,
} from "../fixtures/simulatedWorkbook";

const percentile95 = (samples: readonly number[]): number => {
  const ordered = [...samples].sort((left, right) => left - right);
  return ordered[Math.ceil(ordered.length * 0.95) - 1] ?? Number.POSITIVE_INFINITY;
};

const headlessDialogHost = (): QuickTraceDialogHost => ({
  isSupported: () => true,
  open: async () => ({
    close: () => undefined,
    post: () => undefined,
    subscribeClosed: () => () => undefined,
    subscribeMessage: () => () => undefined,
  }),
});

const quickTraceControllerFor = (gateway: OfficeGateway) => createQuickTraceController({
  dialogHost: headlessDialogHost(),
  dialogUrl: "https://127.0.0.1:43119/quick-trace.html",
  gateway,
});

const fiftyNodeGateway = (): OfficeGateway => {
  const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
  const root = sampleModelScenario.states.formula;
  const branch: CellFacts = {
    ...root,
    address: "$D$20",
    formula: "=SUM($A$1:$A$50)",
    value: 1275,
  };
  const leaves = Array.from({ length: 50 }, (_, index): CellFacts => ({
    ...root,
    address: `$A$${index + 1}`,
    formula: null,
    value: index + 1,
  }));
  const facts = new Map([root, branch, ...leaves].map((cell) => [canonicalCellId(cell), cell]));
  let selection = root;
  const selectionListeners = new Set<() => void>();
  return {
    ...fixture.gateway,
    getCellFacts: async (ref, signal) => {
      signal?.throwIfAborted();
      const cell = facts.get(canonicalCellId(ref));
      if (cell === undefined) throw new Error("missing performance fixture cell");
      return cell;
    },
    getCellsFacts: async (refs, signal) => {
      signal?.throwIfAborted();
      return refs.map((ref) => facts.get(canonicalCellId(ref)) ?? null);
    },
    getDirectDependencies: async (ref, direction, signal) => {
      signal?.throwIfAborted();
      if (direction !== "precedent") return [];
      if (ref.address === root.address) return [branch];
      if (ref.address === branch.address) return leaves;
      return [];
    },
    getSelectionFacts: async () => selection,
    selectCell: async (ref: CellRef) => {
      const cell = facts.get(canonicalCellId(ref));
      if (cell === undefined) throw new Error("missing performance fixture selection");
      selection = cell;
      queueMicrotask(() => selectionListeners.forEach((listener) => listener()));
    },
    subscribeSelection: async (listener) => {
      selectionListeners.add(listener);
      return () => selectionListeners.delete(listener);
    },
  };
};

afterEach(cleanup);

describe("standalone product performance budgets", () => {
  it("commits every qualification state and workbook-risk fixture", async () => {
    expect(Object.keys(sampleModelScenario.states).sort()).toEqual([
      "blank",
      "error",
      "formula",
      "hardcode",
      "multi",
      "protected",
    ]);
    expect(sampleModelScenario.unresolved.map((ref) => ref.unresolved)).toEqual([
      "external",
      "dynamic",
    ]);
    expect(sampleModelScenario.groupedRows).not.toHaveLength(0);
    expect(sampleModelScenario.groupedColumns).not.toHaveLength(0);
    const protectedFixture = createSimulatedWorkbookFixture({ initialState: "protected" });
    await expect(protectedFixture.runCommand("format.input")).resolves.toMatchObject({
      code: "PROTECTED",
      ok: false,
    });
  });

  it("keeps warm selection-to-inspector p95 within 600 ms", async () => {
    const fixture = createSimulatedWorkbookFixture();
    const controller = createController(fixture.gateway);
    await controller.start();
    render(createElement(App as ComponentType<AppProps>, {
      controller,
      runCommand: fixture.runCommand,
    }));
    await screen.findByRole("heading", { name: "DCF $F$42" });
    const samples: number[] = [];
    for (let index = 0; index < 8; index += 1) {
      const target = index % 2 === 0
        ? sampleModelScenario.states.hardcode
        : sampleModelScenario.states.formula;
      const heading = index % 2 === 0 ? "Inputs $B$7" : "DCF $F$42";
      const startedAt = performance.now();
      await fixture.gateway.selectCell(target);
      await waitFor(() =>
        expect(screen.getByRole("heading", { name: heading })).toBeTruthy(), {
        timeout: 600,
      });
      samples.push(performance.now() - startedAt);
    }
    controller.stop();

    expect(percentile95(samples)).toBeLessThanOrEqual(600);
  });

  it("keeps loaded direct-dependency navigation p95 within 500 ms", async () => {
    const fixture = createSimulatedWorkbookFixture();
    const controller = createController(fixture.gateway);
    await controller.start();
    render(createElement(App as ComponentType<AppProps>, {
      controller,
      runCommand: fixture.runCommand,
    }));
    await screen.findByRole("heading", { name: "DCF $F$42" });
    const samples: number[] = [];
    for (let index = 0; index < 8; index += 1) {
      const navigatingToInput = index % 2 === 0;
      const dependencyList = screen.getByRole("list", {
        name: navigatingToInput ? "Direct precedents" : "Direct dependents",
      });
      const targetHeading = navigatingToInput ? "Inputs $B$7" : "DCF $F$42";
      const startedAt = performance.now();
      fireEvent.click(within(dependencyList).getByRole("button", {
        name: targetHeading,
      }));
      await waitFor(() =>
        expect(screen.getByRole("heading", { name: targetHeading })).toBeTruthy(), {
        timeout: 500,
      });
      samples.push(performance.now() - startedAt);
    }
    controller.stop();

    expect(percentile95(samples)).toBeLessThanOrEqual(500);
  });

  it("keeps loaded Quick Trace row movement p95 below 100 ms", async () => {
    const fixture = createSimulatedWorkbookFixture();
    const controller = quickTraceControllerFor(fixture.gateway);
    await controller.open("precedent");
    const samples: number[] = [];
    for (let index = 0; index < 100; index += 1) {
      const started = performance.now();
      await controller.dispatch({
        movement: index % 2 === 0 ? "next" : "previous",
        type: "move",
      });
      samples.push(performance.now() - started);
    }
    controller.stop();

    expect(percentile95(samples)).toBeLessThanOrEqual(100);
  });

  it("keeps cached one-level expansion p95 at or below 250 ms", async () => {
    const fixture = createSimulatedWorkbookFixture();
    const baseRead = fixture.gateway.getDirectDependencies.bind(fixture.gateway);
    let reads = 0;
    const gateway: OfficeGateway = {
      ...fixture.gateway,
      getDirectDependencies: async (ref, direction, signal) => {
        reads += 1;
        return baseRead(ref, direction, signal);
      },
    };
    const controller = quickTraceControllerFor(gateway);
    await controller.open("precedent");
    await controller.dispatch({ movement: "next", type: "move" });
    await controller.dispatch({ movement: "next", type: "move" });
    await controller.dispatch({ type: "right" });
    await controller.dispatch({ type: "left" });
    const readsBeforeSamples = reads;
    const samples: number[] = [];
    for (let index = 0; index < 100; index += 1) {
      const started = performance.now();
      await controller.dispatch({ type: index % 2 === 0 ? "right" : "left" });
      samples.push(performance.now() - started);
    }
    controller.stop();

    expect(reads).toBe(readsBeforeSamples);
    expect(percentile95(samples)).toBeLessThanOrEqual(250);
  });

  it("keeps an uncached one-level expansion of 50 nodes p95 at or below 1 second", async () => {
    const samples: number[] = [];
    for (let index = 0; index < 20; index += 1) {
      const controller = quickTraceControllerFor(fiftyNodeGateway());
      await controller.open("precedent");
      await controller.dispatch({ movement: "next", type: "move" });
      const started = performance.now();
      await controller.dispatch({ type: "right" });
      samples.push(performance.now() - started);
      expect(controller.getState()?.rows).toHaveLength(52);
      controller.stop();
    }

    expect(percentile95(samples)).toBeLessThanOrEqual(1000);
  });

  it("keeps cached direction toggle p95 at or below 250 ms", async () => {
    const fixture = createSimulatedWorkbookFixture();
    const baseRead = fixture.gateway.getDirectDependencies.bind(fixture.gateway);
    let reads = 0;
    const gateway: OfficeGateway = {
      ...fixture.gateway,
      getDirectDependencies: async (ref, direction, signal) => {
        reads += 1;
        return baseRead(ref, direction, signal);
      },
    };
    const controller = quickTraceControllerFor(gateway);
    await controller.open("precedent");
    await controller.dispatch({ direction: "dependent", type: "setDirection" });
    await controller.dispatch({ direction: "precedent", type: "setDirection" });
    const readsBeforeSamples = reads;
    const samples: number[] = [];
    for (let index = 0; index < 100; index += 1) {
      const started = performance.now();
      await controller.dispatch({
        direction: index % 2 === 0 ? "dependent" : "precedent",
        type: "setDirection",
      });
      samples.push(performance.now() - started);
    }
    controller.stop();

    expect(reads).toBe(readsBeforeSamples);
    expect(percentile95(samples)).toBeLessThanOrEqual(250);
  });
});
