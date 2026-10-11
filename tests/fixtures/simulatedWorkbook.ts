import { detectCapabilities } from "../../src/platform/office/capabilityMap";
import {
  OfficeGatewayError,
  type NeighborCellFacts,
  type OfficeGateway,
  type WorkbookFacts,
} from "../../src/platform/office/officeGateway";
import type { HealthWorkbookFacts } from "../../src/domain/health/rules";
import type { PersistedWatchItem } from "../../src/domain/watchlist/watchlist";
import {
  canonicalCellId,
  type CellFacts,
  type CellRef,
  type CommandId,
  type CommandResult,
  type DependencyDirection,
  type DependencyRef,
} from "../../src/shared/types";

const workbookId = "fixture://sample-model";

const cell = (
  sheetId: string,
  sheetName: string,
  address: string,
  value: CellFacts["value"],
  formula: string | null,
  numberFormat = "General",
): CellFacts => ({
  address,
  columnCount: 1,
  formula,
  numberFormat,
  rowCount: 1,
  sheetId,
  sheetName,
  value,
  workbookId,
});

const formulaCell = cell(
  "dcf",
  "DCF",
  "$F$42",
  0.184,
  "=XIRR(Inputs!$B$7:$B$12,Inputs!$A$7:$A$12)",
  "0.0%",
);
const hardcodeCell = cell("inputs", "Inputs", "$B$7", 0.1, null, "0.0%");
const blankCell = cell("dcf", "DCF", "$A$1", null, null);
const errorCell = cell("checks", "Checks", "$H$5", "#REF!", "=#REF!");
const multiSelection: CellFacts = {
  ...cell("dcf", "DCF", "$A$1:$B$2", null, null),
  columnCount: 2,
  rowCount: 2,
};
const branchCell = cell("dcf", "DCF", "$D$20", 0.2, "=Inputs!$B$7*2", "0.0%");
const cycleA = cell("dcf", "DCF", "$C$10", 2, "=$C$11+1");
const cycleB = cell("dcf", "DCF", "$C$11", 1, "=$C$10-1");
const protectedCell = cell("protected", "Protected", "$B$2", 7, "=Inputs!$B$7*70");

const unresolved = (
  kind: "dynamic" | "external",
  address: string,
  sheetId: string,
  sheetName: string,
): DependencyRef => ({ address, sheetId, sheetName, unresolved: kind, workbookId });

const externalLeaf = unresolved(
  "external",
  "[Debt.xlsx]Terms!$B$4",
  "external-debt",
  "Debt.xlsx",
);
const dynamicLeaf = unresolved(
  "dynamic",
  "INDIRECT(DriverCell)",
  "dynamic-indirect",
  "DCF",
);

const cells = [
  formulaCell,
  hardcodeCell,
  blankCell,
  errorCell,
  branchCell,
  cycleA,
  cycleB,
  protectedCell,
];

const cellMap = new Map(cells.map((facts) => [canonicalCellId(facts), facts]));
const dependencyKey = (
  ref: CellRef,
  direction: DependencyDirection,
): string => `${canonicalCellId(ref)}::${direction}`;

const dependencies = new Map<string, readonly DependencyRef[]>([
  [dependencyKey(formulaCell, "precedent"), [
    hardcodeCell,
    branchCell,
    cycleA,
    externalLeaf,
    dynamicLeaf,
  ]],
  [dependencyKey(formulaCell, "dependent"), []],
  [dependencyKey(hardcodeCell, "precedent"), []],
  [dependencyKey(hardcodeCell, "dependent"), [branchCell, formulaCell, protectedCell]],
  [dependencyKey(branchCell, "precedent"), [hardcodeCell]],
  [dependencyKey(branchCell, "dependent"), [formulaCell]],
  [dependencyKey(cycleA, "precedent"), [cycleB]],
  [dependencyKey(cycleA, "dependent"), [cycleB]],
  [dependencyKey(cycleB, "precedent"), [cycleA]],
  [dependencyKey(cycleB, "dependent"), [cycleA]],
  [dependencyKey(protectedCell, "precedent"), [hardcodeCell]],
  [dependencyKey(protectedCell, "dependent"), []],
]);

const stateFacts = {
  blank: blankCell,
  error: errorCell,
  formula: formulaCell,
  hardcode: hardcodeCell,
  multi: multiSelection,
  protected: protectedCell,
} as const satisfies Readonly<Record<string, CellFacts>>;

export const sampleModelScenario = {
  groupedColumns: [{ end: 8, sheetId: "dcf", start: 6 }],
  groupedRows: [{ end: 32, level: 2, sheetId: "dcf", start: 20 }],
  states: stateFacts,
  unresolved: [externalLeaf, dynamicLeaf],
} as const;

const throwIfAborted = (signal?: AbortSignal) => {
  if (signal?.aborted === true) {
    throw new OfficeGatewayError("ABORTED", "The Office request was cancelled");
  }
};

const fixtureDelay = async (signal?: AbortSignal): Promise<void> => {
  throwIfAborted(signal);
  await new Promise<void>((resolve, reject) => {
    const onAbort = () => {
      clearTimeout(timer);
      reject(new OfficeGatewayError("ABORTED", "The Office request was cancelled"));
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, 75);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
  throwIfAborted(signal);
};

const factsFor = (ref: CellRef): CellFacts => {
  const facts = cellMap.get(canonicalCellId(ref));
  if (facts === undefined) {
    throw new OfficeGatewayError("NOT_FOUND", "The simulated cell was not found");
  }
  return facts;
};

const healthFacts = (): HealthWorkbookFacts => ({
  calculationMode: "manual",
  cells: [
    ...Array.from({ length: 5_000 }, (_, index) => ({
      ...cell("operations", "Operations", `$A$${index + 1}`, index + 1, null, "0"),
      fontColor: "#000000",
    })),
    { ...errorCell, fontColor: "#C00000" },
    {
      ...cell(
        "checks",
        "Checks",
        "$I$5",
        12,
        "='[Debt.xlsx]Terms'!$B$4*2",
        "0.0x",
      ),
      fontColor: "#008000",
    },
  ],
  workbookRef: formulaCell,
});

export type SimulatedWorkbookFixture = Readonly<{
  gateway: OfficeGateway;
  runCommand(id: CommandId): Promise<CommandResult>;
}>;

export const createSimulatedWorkbookFixture = (
  options: Readonly<{ initialState?: string }> = {},
): SimulatedWorkbookFixture => {
  const requestedState = options.initialState;
  let selection: CellFacts = requestedState !== undefined && requestedState in stateFacts
    ? stateFacts[requestedState as keyof typeof stateFacts]
    : formulaCell;
  let metadata: PersistedWatchItem[] = [];
  let lastWrite:
    | { id: CommandId; kind: "format" }
    | { inverseId: CommandId; inverseLabel: string; kind: "structural" }
    | null = null;
  const recalculationListeners = new Set<() => void>();
  const selectionListeners = new Set<() => void>();

  const runCommand = async (id: CommandId): Promise<CommandResult> => {
    if (id === "system.acknowledgeSafetyLock") return { ok: true };
    if (id === "system.revertLastAction") {
      if (lastWrite === null) {
        return { code: "NO_REVERT", message: "There is no Analyst Lens action to revert", ok: false };
      }
      if (lastWrite.kind === "structural") {
        return {
          code: "STRUCTURAL_REVERT_REQUIRES_INVERSE",
          message: `Automatic Revert is unavailable. Select the affected range and run ${lastWrite.inverseLabel} once.`,
          ok: false,
        };
      }
      lastWrite = null;
      return { ok: true };
    }
    if (id.startsWith("structure.")) {
      if (selection.sheetId === "protected") {
        return {
          code: "PROTECTED",
          message: "Structure changes are unavailable on the protected fixture sheet",
          ok: false,
        };
      }
      if (id === "structure.showDetails" || id === "structure.hideDetails") {
        return { ok: true };
      }
      const inverses: Record<string, { id: CommandId; label: string }> = {
        "structure.groupRows": { id: "structure.ungroupRows", label: "Ungroup rows" },
        "structure.ungroupRows": { id: "structure.groupRows", label: "Group rows" },
        "structure.groupColumns": { id: "structure.ungroupColumns", label: "Ungroup columns" },
        "structure.ungroupColumns": { id: "structure.groupColumns", label: "Group columns" },
      };
      const inverse = inverses[id];
      if (inverse === undefined) {
        return { code: "UNKNOWN_COMMAND", message: `Unknown command: ${id}`, ok: false };
      }
      lastWrite = { inverseId: inverse.id, inverseLabel: inverse.label, kind: "structural" };
      return { ok: true };
    }
    if (selection.sheetId === "protected") {
      return {
        code: "PROTECTED",
        message: "Formatting is unavailable on the protected fixture sheet",
        ok: false,
      };
    }
    lastWrite = { id, kind: "format" };
    return { ok: true };
  };

  const gateway: OfficeGateway = {
    getCapabilities: () => detectCapabilities({ isSetSupported: () => true }),
    getCellFacts: async (ref, signal) => {
      throwIfAborted(signal);
      return factsFor(ref);
    },
    getCellsFacts: async (refs, signal) => {
      throwIfAborted(signal);
      return refs.map((ref) => cellMap.get(canonicalCellId(ref)) ?? null);
    },
    getDirectDependencies: async (ref, direction, signal) => {
      throwIfAborted(signal);
      return [...(dependencies.get(dependencyKey(ref, direction)) ?? [])];
    },
    getHealthWorkbookFacts: async (signal) => {
      await fixtureDelay(signal);
      return healthFacts();
    },
    getNeighborFacts: async (ref, signal): Promise<NeighborCellFacts[]> => {
      throwIfAborted(signal);
      if (ref.sheetId !== "dcf" || ref.address !== "$F$42") return [];
      return [
        { direction: "left", facts: cell("dcf", "DCF", "$E$42", 0.17, "=IRR(Inputs!$B$7:$B$12)", "0.0%") },
        { direction: "right", facts: cell("dcf", "DCF", "$G$42", 0.19, "=IRR(Inputs!$B$7:$B$12)", "0.0%") },
      ];
    },
    getSelectionFacts: async (signal) => {
      throwIfAborted(signal);
      return selection;
    },
    getWorkbookFacts: async (signal): Promise<WorkbookFacts> => {
      throwIfAborted(signal);
      return {
        calculationMode: "manual",
        id: workbookId,
        name: "Sample-model.xlsx",
        worksheets: [
          { id: "inputs", name: "Inputs", visibility: "Visible" },
          { id: "dcf", name: "DCF", visibility: "Visible" },
          { id: "protected", name: "Protected", visibility: "Visible" },
        ],
      };
    },
    loadWatchlistMetadata: async () => metadata,
    runWrite: async ({ id }) => runCommand(id),
    saveWatchlistMetadata: async (next) => {
      metadata = next.map((item) => ({ ...item }));
    },
    selectCell: async (ref) => {
      selection = factsFor(ref);
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

  return { gateway, runCommand };
};
