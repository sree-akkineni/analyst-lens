import { describe, expect, it } from "vitest";
import {
  collapseBranch,
  createGraphState,
  expandNode,
  searchGraph,
} from "../../src/domain/dependencies/graph";
import {
  canonicalCellId,
  type CellRef,
  type DependencyNode,
} from "../../src/shared/types";

const ref = (
  address: string,
  sheetName = "DCF",
  workbookId = "workbook-1",
): CellRef => ({
  address,
  sheetId: `sheet-${sheetName}`,
  sheetName,
  workbookId,
});

const node = (
  address: string,
  sheetName = "DCF",
  workbookId = "workbook-1",
): DependencyNode => {
  const cell = ref(address, sheetName, workbookId);
  return { id: canonicalCellId(cell), ref: cell };
};

const unresolved = (
  id: string,
  kind: "dynamic" | "external",
): DependencyNode => ({
  id,
  ref: ref(id, kind === "external" ? "External" : "DCF"),
  unresolved: kind,
});

describe("dependency graph", () => {
  it("deduplicates nodes, orients precedent edges, and marks the closing cycle", () => {
    const origin = node("$F$42");
    const input = node("$B$7", "Inputs");
    let state = expandNode(
      createGraphState(origin),
      origin.id,
      "precedent",
      [input, input],
    ).state;

    state = expandNode(state, input.id, "precedent", [origin]).state;

    expect(new Set(state.nodes.map((current) => current.id)).size).toBe(state.nodes.length);
    expect(state.edges).toEqual(expect.arrayContaining([
      expect.objectContaining({ source: input.id, target: origin.id }),
      expect.objectContaining({ cycle: true, source: origin.id, target: input.id }),
    ]));
  });

  it("orients dependent edges from the selected node to each dependent", () => {
    const origin = node("$F$42");
    const output = node("$F$44");
    const { state } = expandNode(
      createGraphState(origin),
      origin.id,
      "dependent",
      [output],
    );

    expect(state.edges).toMatchObject([{
      cycle: false,
      source: origin.id,
      target: output.id,
    }]);
  });

  it("adds only one requested level and keeps unresolved leaves bounded", () => {
    const origin = node("$F$42");
    const direct = node("$F$40");
    const external = unresolved("external::Book.xlsx::A1", "external");
    const dynamic = unresolved("dynamic::INDIRECT", "dynamic");
    const first = expandNode(
      createGraphState(origin),
      origin.id,
      "precedent",
      [direct, external, dynamic],
    );
    const attemptedLeafExpansion = expandNode(
      first.state,
      external.id,
      "precedent",
      [node("$A$1", "Should Not Appear")],
    );

    expect(first.state.nodes.map((current) => current.id)).toEqual([
      origin.id,
      direct.id,
      external.id,
      dynamic.id,
    ]);
    expect(attemptedLeafExpansion.state).toEqual(first.state);
  });

  it("collapses an orphaned branch but preserves nodes shared by another branch", () => {
    const origin = node("$F$42");
    const left = node("$F$40");
    const right = node("$F$41");
    const shared = node("$B$7", "Inputs");
    const leaf = node("$B$3", "Assumptions");
    let state = expandNode(createGraphState(origin), origin.id, "precedent", [left, right]).state;
    state = expandNode(state, left.id, "precedent", [shared]).state;
    state = expandNode(state, right.id, "precedent", [shared]).state;
    state = expandNode(state, shared.id, "precedent", [leaf]).state;

    const collapsed = collapseBranch(state, left.id, "precedent");

    expect(collapsed.nodes.map((current) => current.id)).toEqual(expect.arrayContaining([
      origin.id,
      left.id,
      right.id,
      shared.id,
      leaf.id,
    ]));
    expect(collapsed.edges).not.toEqual(expect.arrayContaining([
      expect.objectContaining({ source: shared.id, target: left.id }),
    ]));
    expect(collapsed.edges).toEqual(expect.arrayContaining([
      expect.objectContaining({ source: shared.id, target: right.id }),
    ]));
  });

  it("preserves reachable descendant expansions when a parent level refreshes", () => {
    const origin = node("$F$42");
    const direct = node("$F$40");
    const leaf = node("$B$7", "Inputs");
    const newDirect = node("$F$41");
    let state = expandNode(createGraphState(origin), origin.id, "precedent", [direct]).state;
    state = expandNode(state, direct.id, "precedent", [leaf]).state;

    const refreshed = expandNode(
      state,
      origin.id,
      "precedent",
      [direct, newDirect],
    ).state;

    expect(refreshed.nodes.map((current) => current.id)).toEqual(expect.arrayContaining([
      origin.id,
      direct.id,
      newDirect.id,
      leaf.id,
    ]));
  });

  it("preserves a shared descendant after collapse regardless of expansion order", () => {
    const origin = node("$F$42");
    const left = node("$F$40");
    const right = node("$F$41");
    const shared = node("$B$7", "Inputs");
    const leaf = node("$B$3", "Assumptions");
    let state = expandNode(createGraphState(origin), origin.id, "precedent", [left, right]).state;
    state = expandNode(state, left.id, "precedent", [shared]).state;
    state = expandNode(state, shared.id, "precedent", [leaf]).state;
    state = expandNode(state, right.id, "precedent", [shared]).state;

    const collapsed = collapseBranch(state, left.id, "precedent");

    expect(collapsed.nodes.map((current) => current.id)).toEqual(expect.arrayContaining([
      shared.id,
      leaf.id,
    ]));
  });

  it.each([50, 100, 250, 500])(
    "keeps a %i-node fixture within the visible-node cap and performance budget",
    (size) => {
      const origin = node("$A$1");
      const dependencies = Array.from({ length: size - 1 }, (_, index) =>
        node(`$A$${index + 2}`));
      const started = performance.now();
      const result = expandNode(
        createGraphState(origin),
        origin.id,
        "dependent",
        dependencies,
        { maxVisibleNodes: 500 },
      );

      expect(result.state.nodes).toHaveLength(size);
      expect(result.capped).toBe(false);
      expect(performance.now() - started).toBeLessThan(250);
    },
  );

  it("expands a lazy 500-node chain within the domain performance budget", () => {
    const origin = node("$A$1");
    let state = createGraphState(origin);
    let parent = origin;
    const started = performance.now();
    for (let row = 2; row <= 500; row += 1) {
      const next = node(`$A$${row}`);
      state = expandNode(state, parent.id, "dependent", [next]).state;
      parent = next;
    }

    expect(state.nodes).toHaveLength(500);
    expect(performance.now() - started).toBeLessThan(250);
  });

  it("stops at the cap without throwing and leaves the input state immutable", () => {
    const origin = node("$A$1");
    const initial = createGraphState(origin);
    const dependencies = Array.from({ length: 600 }, (_, index) =>
      node(`$A$${index + 2}`));
    const result = expandNode(
      initial,
      origin.id,
      "dependent",
      dependencies,
      { maxVisibleNodes: 500 },
    );

    expect(result.capped).toBe(true);
    expect(result.state.nodes).toHaveLength(500);
    expect(initial.nodes).toEqual([origin]);
    expect(initial.edges).toEqual([]);
  });

  it("searches visible graph nodes across workbook, sheet, address, and value", () => {
    const origin = { ...node("$F$42"), value: 12 };
    const crossWorkbook = { ...node("$B$7", "Inputs", "workbook-2"), value: "Revenue" };
    const state = expandNode(
      createGraphState(origin),
      origin.id,
      "precedent",
      [crossWorkbook],
    ).state;

    expect(searchGraph(state, "workbook-2")).toEqual([crossWorkbook]);
    expect(searchGraph(state, "inputs")).toEqual([crossWorkbook]);
    expect(searchGraph(state, "revenue")).toEqual([crossWorkbook]);
    expect(searchGraph(state, "f42")).toEqual([origin]);
  });
});
