import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import type {
  QuickTraceIssue,
  QuickTraceRow,
} from "../../src/domain/dependencies/quickTraceTree";
import type { QuickTraceUserIntent } from "../../src/runtime/quickTraceController";
import type { QuickTraceSnapshot } from "../../src/shared/quickTraceProtocol";
import { canonicalCellId, type CellRef } from "../../src/shared/types";
import { QuickTraceApp } from "../../src/ui/quick-trace/QuickTraceApp";

afterEach(cleanup);

const makeRef = (sheetName: string, address: string): CellRef => ({
  address,
  sheetId: sheetName.toLocaleLowerCase(),
  sheetName,
  workbookId: "fixture://quick-trace",
});
const origin = makeRef("DCF", "$F$42");
const input = makeRef("Inputs", "$B$7");
const other = makeRef("Summary", "$G$12");
const rootRowId = "root";
const inputRowId = "root/input";
const snapshot: QuickTraceSnapshot = {
  capped: false,
  direction: "precedent",
  formula: "=XIRR(Inputs!$B$7:$B$12,Inputs!$A$7:$A$12)",
  navigationBlocked: false,
  nodeCount: 2,
  notice: null,
  origin,
  pendingExcelSelection: null,
  root: origin,
  rows: [
    {
      depth: 0,
      disclosure: "expanded",
      displayValue: "18.4%",
      issue: null,
      node: { id: canonicalCellId(origin), ref: origin },
      parentRowId: null,
      rowId: rootRowId,
    },
    {
      depth: 1,
      disclosure: "collapsed",
      displayValue: "10.0%",
      issue: null,
      node: { id: canonicalCellId(input), ref: input },
      parentRowId: rootRowId,
      rowId: inputRowId,
    },
  ],
  selectedFormula: "=XIRR(Inputs!$B$7:$B$12,Inputs!$A$7:$A$12)",
  selectedRowId: rootRowId,
  status: "ready",
};
const snapshotRootRow = snapshot.rows[0];
if (snapshotRootRow === undefined) throw new Error("expected snapshot root row");
const snapshotWithIssue = (
  issue: Exclude<QuickTraceIssue, null>,
): QuickTraceSnapshot => ({
  ...snapshot,
  rows: snapshot.rows.map((row) => row.rowId === inputRowId
    ? { ...row, disclosure: "leaf", issue }
    : row),
});

describe("Quick Trace dialog", () => {
  it("renders only Reference and Value columns with accessible tree metadata", () => {
    render(<QuickTraceApp send={() => undefined} state={snapshot} />);

    expect(screen.getByText("Reference")).toBeTruthy();
    expect(screen.getByText("Value")).toBeTruthy();
    expect(screen.queryByText("Argument")).toBeNull();
    const root = screen.getByRole("treeitem", { name: /DCF.*\$F\$42.*18.4%/u });
    expect(root.getAttribute("aria-level")).toBe("1");
    expect(root.getAttribute("aria-expanded")).toBe("true");
    expect(root.getAttribute("aria-posinset")).toBe("1");
    expect(root.getAttribute("aria-setsize")).toBe("1");
    const tree = screen.getByRole("tree");
    expect(tree.getAttribute("aria-activedescendant")).toBe(root.getAttribute("id"));
    expect(document.activeElement).toBe(tree);
    expect(screen.getByRole("region", { name: "Current root DCF $F$42" })).toBeTruthy();
  });

  it.each([
    ["ArrowUp", { movement: "previous", type: "move" }],
    ["ArrowDown", { movement: "next", type: "move" }],
    ["ArrowLeft", { type: "left" }],
    ["ArrowRight", { type: "right" }],
    ["Enter", { rowId: snapshot.selectedRowId, type: "reroot" }],
    ["o", { type: "returnToOrigin" }],
    ["Escape", { type: "close" }],
  ] as const)("maps %s to one typed intent", (key, expected) => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={snapshot} />);
    fireEvent.keyDown(screen.getByRole("tree"), { key });

    expect(sent).toEqual([expected]);
  });

  it("does not consume Command+W or letter commands from editable targets", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<>
      <QuickTraceApp send={(intent) => sent.push(intent)} state={snapshot} />
      <input aria-label="Test input" />
    </>);
    const tree = screen.getByRole("tree");
    const commandW = new KeyboardEvent("keydown", { bubbles: true, key: "w", metaKey: true });
    tree.dispatchEvent(commandW);
    fireEvent.keyDown(screen.getByRole("textbox", { name: "Test input" }), { key: "t" });

    expect(commandW.defaultPrevented).toBe(false);
    expect(sent).toEqual([]);
  });

  it("toggles direction with T and the visible segmented controls", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={snapshot} />);
    expect(screen.getByRole("button", { name: "Trace In" }).getAttribute("aria-pressed"))
      .toBe("true");
    expect(screen.getByRole("button", { name: "Trace Out" }).getAttribute("aria-pressed"))
      .toBe("false");
    fireEvent.keyDown(screen.getByRole("tree"), { key: "t" });
    fireEvent.click(screen.getByRole("button", { name: "Trace Out" }));

    expect(sent).toEqual([
      { direction: "dependent", type: "setDirection" },
      { direction: "dependent", type: "setDirection" },
    ]);
  });

  it("expands from the chevron without rerooting", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={snapshot} />);
    fireEvent.click(screen.getByRole("button", { name: "Expand Inputs $B$7" }));

    expect(sent).toEqual([{ rowId: inputRowId, type: "toggleBranch" }]);
  });

  it("maps row, return-to-origin, and close pointer controls to typed intents", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={snapshot} />);
    fireEvent.click(screen.getByRole("treeitem", { name: /Inputs.*\$B\$7.*10.0%/u }));
    fireEvent.click(screen.getByRole("button", { name: "Return to origin" }));
    fireEvent.click(screen.getByRole("button", { name: "Close" }));

    expect(sent).toEqual([
      { rowId: inputRowId, type: "navigate" },
      { type: "returnToOrigin" },
      { type: "close" },
    ]);
  });

  it("offers pointer re-root and retry without making the formula editable", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={{
      ...snapshot,
      notice: "Navigation failed",
      status: "error",
    }} />);
    expect(screen.getByRole("alert").textContent).toContain("Navigation failed");
    fireEvent.click(screen.getByRole("button", { name: "Make selected reference the root" }));
    fireEvent.click(screen.getByRole("button", { name: "Retry selected reference" }));

    expect(sent).toEqual([
      { rowId: rootRowId, type: "reroot" },
      { rowId: rootRowId, type: "navigate" },
    ]);
    expect(screen.queryByRole("textbox")).toBeNull();
  });

  it("shows an unrelated Excel selection without silently replacing the root", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={{
      ...snapshot,
      pendingExcelSelection: other,
    }} />);
    expect(screen.getByRole("status").textContent).toContain("Excel selection changed");
    fireEvent.click(screen.getByRole("button", { name: "Use Summary $G$12 as new root" }));

    expect(sent).toEqual([{ type: "useExcelSelection" }]);
    expect(screen.getByText(/Origin: DCF.*\$F\$42/u)).toBeTruthy();
  });

  it.each([
    ["cycle", "Cycle"],
    ["dynamic", "Dynamic"],
    ["external", "External"],
    ["unavailable", "Unavailable"],
  ] as const)("renders the %s issue without color-only meaning", (issue, label) => {
    render(<QuickTraceApp send={() => undefined} state={snapshotWithIssue(issue)} />);
    expect(screen.getByText(label)).toBeTruthy();
  });

  it("renders empty, loading, capped, and workbook-changed states explicitly", () => {
    const { rerender } = render(<QuickTraceApp send={() => undefined} state={{
      ...snapshot,
      nodeCount: 1,
      rows: [snapshotRootRow],
    }} />);
    expect(screen.getByText("No direct precedents")).toBeTruthy();
    rerender(<QuickTraceApp send={() => undefined} state={{ ...snapshot, status: "loading" }} />);
    expect(screen.getByRole("status").textContent).toContain("Loading");
    rerender(<QuickTraceApp send={() => undefined} state={{ ...snapshot, capped: true }} />);
    expect(screen.getByText(/500-node limit/u)).toBeTruthy();
    rerender(<QuickTraceApp send={() => undefined} state={{
      ...snapshot,
      navigationBlocked: true,
      notice: "Workbook changed",
    }} />);
    expect(screen.getByRole("tree").getAttribute("aria-disabled")).toBe("true");
    expect(screen.getByText("Workbook changed")).toBeTruthy();
  });

  it("suppresses navigation and expansion while the workbook is blocked", () => {
    const sent: QuickTraceUserIntent[] = [];
    render(<QuickTraceApp send={(intent) => sent.push(intent)} state={{
      ...snapshot,
      navigationBlocked: true,
    }} />);
    fireEvent.keyDown(screen.getByRole("tree"), { key: "ArrowDown" });
    fireEvent.click(screen.getByRole("treeitem", { name: /Inputs/u }));
    fireEvent.click(screen.getByRole("button", { name: "Expand Inputs $B$7" }));

    expect(sent).toEqual([]);
  });

  it("windows a large flattened tree while keeping the selected row mounted", () => {
    const rows = Array.from({ length: 300 }, (_, index): QuickTraceRow => ({
      depth: 1,
      disclosure: "leaf",
      displayValue: String(index),
      issue: null,
      node: { id: `node-${index}`, ref: makeRef("Inputs", `$B$${index + 1}`) },
      parentRowId: rootRowId,
      rowId: `root/node-${index}`,
    }));
    render(<QuickTraceApp send={() => undefined} state={{
      ...snapshot,
      nodeCount: 301,
      rows: [snapshotRootRow, ...rows],
      selectedRowId: "root/node-250",
    }} />);

    expect(screen.getAllByRole("treeitem").length).toBeLessThanOrEqual(80);
    expect(screen.getByRole("treeitem", { selected: true }).textContent).toContain("$B$251");
  });
});
