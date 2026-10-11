import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { analyzeFormula } from "../../src/domain/formulas/tokenizeFormula";
import { createGraphState } from "../../src/domain/dependencies/graph";
import { startHistory, visit } from "../../src/domain/dependencies/navigationHistory";
import { canonicalCellId, type CellFacts, type DependencyNode } from "../../src/shared/types";
import {
  initialInspectorState,
  type InspectorState,
} from "../../src/runtime/controller";
import { Inspector } from "../../src/ui/Inspector";

afterEach(cleanup);

const facts: CellFacts = {
  address: "$F$42",
  columnCount: 1,
  formula: "=$B$7*2",
  numberFormat: "0.0x",
  rowCount: 1,
  sheetId: "sheet-1",
  sheetName: "DCF",
  value: 12,
  workbookId: "workbook-1",
};

const dependency: DependencyNode = {
  id: "workbook-1::sheet-2::$B$7",
  ref: {
    address: "$B$7",
    sheetId: "sheet-2",
    sheetName: "Inputs",
    workbookId: "workbook-1",
  },
  value: 100,
};

const dependent: DependencyNode = {
  id: "workbook-1::sheet-3::$G$43",
  ref: {
    address: "$G$43",
    sheetId: "sheet-3",
    sheetName: "Outputs",
    workbookId: "workbook-1",
  },
  value: 24,
};

const origin: DependencyNode = {
  id: canonicalCellId(facts),
  ref: facts,
  value: facts.value,
};

const analysis = analyzeFormula(facts.formula ?? "", { origin: facts });
const formulaState: InspectorState = {
  ...initialInspectorState(),
  dependents: [dependent],
  graph: createGraphState(origin),
  history: startHistory(facts),
  precedents: [dependency],
  risks: analysis.risks,
  tokens: analysis.tokens,
  view: { cellKind: "formula", facts, kind: "selection" },
};

describe("Hybrid Formula Inspector", () => {
  it("opens floating Trace In and Trace Out from visible controls", async () => {
    const opened: string[] = [];
    render(<Inspector
      dispatch={() => undefined}
      openQuickTrace={async (direction) => {
        opened.push(direction);
      }}
      quickTraceSupported
      runCommand={async () => ({ ok: true })}
      state={formulaState}
    />);

    fireEvent.click(screen.getByRole("button", { name: /Open Quick Trace In/u }));
    await waitFor(() => expect(opened).toEqual(["precedent"]));
    fireEvent.click(screen.getByRole("button", { name: /Open Quick Trace Out/u }));
    await waitFor(() => expect(opened).toEqual(["precedent", "dependent"]));
  });

  it("prevents overlapping Quick Trace launches", async () => {
    let release: (() => void) | undefined;
    const pending = new Promise<void>((resolve) => {
      release = resolve;
    });
    const opened: string[] = [];
    render(<Inspector
      dispatch={() => undefined}
      openQuickTrace={async (direction) => {
        opened.push(direction);
        await pending;
      }}
      quickTraceSupported
      runCommand={async () => ({ ok: true })}
      state={formulaState}
    />);

    fireEvent.click(screen.getByRole("button", { name: /Open Quick Trace In/u }));
    const traceOut = screen.getByRole("button", { name: /Open Quick Trace Out/u });
    expect(traceOut).toHaveProperty("disabled", true);
    fireEvent.click(traceOut);
    expect(opened).toEqual(["precedent"]);
    release?.();
    await waitFor(() => expect(traceOut).toHaveProperty("disabled", false));
  });

  it("disables Quick Trace controls when DialogApi 1.2 is unavailable", () => {
    render(<Inspector
      dispatch={() => undefined}
      openQuickTrace={async () => undefined}
      quickTraceSupported={false}
      runCommand={async () => ({ ok: true })}
      state={formulaState}
    />);

    expect(screen.getByRole("button", { name: /Open Quick Trace In/u }))
      .toHaveProperty("disabled", true);
    expect(screen.getByText(/requires DialogApi 1.2/iu)).toBeTruthy();
  });

  it("renders formula context and direct dependencies without pointer-only controls", () => {
    const dispatch = vi.fn();
    render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={formulaState}
    />);

    expect(screen.getByRole("heading", { name: "DCF $F$42" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Return to origin" }).hasAttribute("disabled")).toBe(false);
    expect(screen.getByRole("list", { name: "Direct precedents" })).toBeTruthy();
    expect(screen.getByRole("list", { name: "Direct dependents" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Inputs $B$7" })).toBeTruthy();
    expect(screen.getByRole("button", {
      name: /Open commands for this cell.*Command-K/u,
    })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Open dependency map" })).toBeTruthy();
    expect(screen.getByText("embedded-number")).toBeTruthy();
    expect(screen.queryByText("Ask about this cell…")).toBeNull();

    fireEvent.click(screen.getByRole("button", { name: "Inputs $B$7" }));
    expect(dispatch).toHaveBeenCalledWith({ ref: dependency.ref, type: "navigate" });
  });

  it("opens the command palette with Command-K", () => {
    const dispatch = vi.fn();
    render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={formulaState}
    />);

    fireEvent.keyDown(window, { key: "k", metaKey: true });

    expect(dispatch).toHaveBeenCalledWith({ overlay: "palette", type: "openOverlay" });
  });

  it("does not hijack Command-K from an editable field", () => {
    const dispatch = vi.fn();
    render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={formulaState}
    />);
    const input = document.createElement("input");
    document.body.append(input);

    fireEvent.keyDown(input, { key: "k", metaKey: true });

    expect(dispatch).not.toHaveBeenCalledWith({ overlay: "palette", type: "openOverlay" });
    input.remove();
  });

  it.each([
    {
      closeName: "Close dependency map",
      dialogName: "Dependency map",
      openerName: "Open dependency map",
      overlay: "graph",
    },
    {
      closeName: "Close command palette",
      dialogName: "Command palette",
      openerName: /Open commands for this cell.*Command-K/u,
      overlay: "palette",
    },
    {
      closeName: "Close settings",
      dialogName: "Inspector settings",
      openerName: "Settings",
      overlay: "settings",
    },
  ] as const)(
    "keeps Command-K inside the active $overlay overlay and restores its invoker",
    ({ closeName, dialogName, openerName, overlay }) => {
      const dispatch = vi.fn();
      render(<Inspector
        dispatch={dispatch}
        runCommand={async () => ({ ok: true })}
        state={{
          ...formulaState,
          overlays: {
            graph: overlay === "graph",
            palette: overlay === "palette",
            settings: overlay === "settings",
          },
        }}
      />);
      const dialog = screen.getByRole("dialog", { name: dialogName });
      const opener = screen.getByRole("button", { name: openerName });
      screen.getByRole("button", { name: closeName }).focus();

      fireEvent.keyDown(window, { key: "k", metaKey: true });

      expect(dispatch).not.toHaveBeenCalledWith({ overlay: "palette", type: "openOverlay" });
      expect(screen.getAllByRole("dialog")).toHaveLength(1);
      expect(dialog.contains(document.activeElement)).toBe(true);

      fireEvent.click(screen.getByRole("button", { name: closeName }));
      expect(dispatch).toHaveBeenCalledWith({ overlay, type: "closeOverlay" });
      expect(document.activeElement).toBe(opener);
    },
  );

  it("renders a navigable breadcrumb and normalized neighbor comparison", () => {
    const dispatch = vi.fn();
    const next = { ...dependency.ref, address: "$C$8" };
    render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={{
        ...formulaState,
        history: visit(visit(startHistory(facts), dependency.ref), next),
        neighbors: [
          { label: "Left", normalizedFormula: "=R[-1]C[0]", status: "same" },
          { label: "Right", normalizedFormula: "=R[-2]C[1]", status: "different" },
          { label: "Above", normalizedFormula: null, status: "unavailable" },
        ],
      }}
    />);

    expect(screen.getByRole("list", { name: "Navigation breadcrumb" })).toBeTruthy();
    expect(screen.getByRole("list", { name: "Navigation breadcrumb" }).querySelectorAll("li"))
      .toHaveLength(3);
    fireEvent.click(screen.getByRole("button", { name: "Go to origin DCF $F$42" }));
    expect(dispatch).toHaveBeenCalledWith({ ref: facts, type: "navigate" });
    expect(screen.getByText("Matches selected formula")).toBeTruthy();
    expect(screen.getByText("Different pattern")).toBeTruthy();
    expect(screen.getByText("No formula")).toBeTruthy();
  });

  it.each([
    ["loading", { ...initialInspectorState(), view: { kind: "loading" as const } }, "Loading selection"],
    ["hardcode", { ...formulaState, view: { cellKind: "hardcode" as const, facts: { ...facts, formula: null }, kind: "selection" as const } }, "Hardcoded value"],
    ["blank", { ...formulaState, view: { cellKind: "blank" as const, facts: { ...facts, formula: null, value: null }, kind: "selection" as const } }, "Blank cell"],
    ["error", { ...formulaState, view: { cellKind: "error" as const, facts: { ...facts, value: "#REF!" }, kind: "selection" as const } }, "Excel error"],
    ["multi", { ...formulaState, view: { cellKind: "multi" as const, facts: { ...facts, columnCount: 2, rowCount: 3 }, kind: "selection" as const } }, "3 × 2 selection"],
    ["unsupported", { ...initialInspectorState(), view: { kind: "unsupported" as const, message: "Dependency APIs unavailable" } }, "Dependency APIs unavailable"],
    ["failure", { ...initialInspectorState(), view: { kind: "failure" as const, message: "Workbook closed" } }, "Workbook closed"],
  ])("renders the accessible %s state", (_label, state, expected) => {
    render(<Inspector
      dispatch={() => undefined}
      runCommand={async () => ({ ok: true })}
      state={state}
    />);
    expect(screen.getByText(new RegExp(`${expected}$`, "u"))).toBeTruthy();
    if (state.view.kind === "selection" && state.view.cellKind !== "multi") {
      expect(screen.getByRole("button", {
        name: /Open commands for this cell.*Command-K/u,
      })).toBeTruthy();
    }
  });

  it("returns focus to the dependency-map invoker after closing the overlay", () => {
    const dispatch = vi.fn();
    render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={{ ...formulaState, overlays: { graph: true, palette: false, settings: false } }}
    />);
    const opener = screen.getByRole("button", { name: "Open dependency map" });
    fireEvent.click(screen.getByRole("button", { name: "Close dependency map" }));

    expect(dispatch).toHaveBeenCalledWith({ overlay: "graph", type: "closeOverlay" });
    expect(document.activeElement).toBe(opener);
  });

  it("returns to the origin from inside the full graph", () => {
    const dispatch = vi.fn();
    render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={{ ...formulaState, overlays: { graph: true, palette: false, settings: false } }}
    />);

    fireEvent.click(screen.getByRole("button", { name: "Return to origin from dependency map" }));
    expect(dispatch).toHaveBeenCalledWith({ type: "returnToOrigin" });
  });

  it("moves focus into the full graph and contains keyboard traversal", () => {
    render(<Inspector
      dispatch={() => undefined}
      runCommand={async () => ({ ok: true })}
      state={{ ...formulaState, overlays: { graph: true, palette: false, settings: false } }}
    />);
    const dialog = screen.getByRole("dialog", { name: "Dependency map" });
    const firstControl = screen.getByRole("button", { name: "Return to origin from dependency map" });

    expect(document.activeElement).toBe(firstControl);
    fireEvent.keyDown(dialog, { key: "Tab", shiftKey: true });
    expect(dialog.contains(document.activeElement)).toBe(true);
    expect(document.activeElement).not.toBe(firstControl);
  });

  it("closes graph and palette overlays with Escape and restores focus", () => {
    const dispatch = vi.fn();
    const { rerender } = render(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={{ ...formulaState, overlays: { graph: true, palette: false, settings: false } }}
    />);
    const graphOpener = screen.getByRole("button", { name: "Open dependency map" });
    fireEvent.keyDown(screen.getByRole("dialog", { name: "Dependency map" }), { key: "Escape" });
    expect(dispatch).toHaveBeenLastCalledWith({ overlay: "graph", type: "closeOverlay" });
    expect(document.activeElement).toBe(graphOpener);

    rerender(<Inspector
      dispatch={dispatch}
      runCommand={async () => ({ ok: true })}
      state={{ ...formulaState, overlays: { graph: false, palette: true, settings: false } }}
    />);
    const paletteOpener = screen.getByRole("button", {
      name: /Open commands for this cell.*Command-K/u,
    });
    fireEvent.keyDown(screen.getByRole("dialog", { name: "Command palette" }), { key: "Escape" });
    expect(dispatch).toHaveBeenLastCalledWith({ overlay: "palette", type: "closeOverlay" });
    expect(document.activeElement).toBe(paletteOpener);
  });
});
