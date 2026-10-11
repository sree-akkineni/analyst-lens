import { describe, expect, it } from "vitest";
import {
  createGraphState,
  expandNode,
} from "../../src/domain/dependencies/graph";
import {
  applyQuickTraceExpansion,
  createQuickTraceTree,
  moveQuickTraceLeft,
  moveQuickTraceRight,
  moveQuickTraceSelection,
  projectQuickTraceRows,
  type QuickTraceTreeState,
} from "../../src/domain/dependencies/quickTraceTree";
import {
  canonicalCellId,
  type CellFacts,
  type DependencyNode,
} from "../../src/shared/types";

const makeNode = (sheetName: string, address: string): DependencyNode => {
  const ref = {
    address,
    sheetId: sheetName.toLocaleLowerCase(),
    sheetName,
    workbookId: "fixture://quick-trace",
  };
  return { id: canonicalCellId(ref), ref };
};

const root = makeNode("DCF", "$F$42");
const left = makeNode("Inputs", "$B$7");
const right = makeNode("DCF", "$D$20");
const leaf = makeNode("Model", "$C$22");
const range = makeNode("Schedule", "$H$12:$K$14");
const factsById: Record<string, CellFacts> = {
  [root.id]: {
    ...root.ref,
    columnCount: 1,
    formula: "=Inputs!$B$7",
    numberFormat: "0.0%",
    rowCount: 1,
    value: 0.184,
  },
  [left.id]: {
    ...left.ref,
    columnCount: 1,
    formula: null,
    numberFormat: "0.0%",
    rowCount: 1,
    value: 0.1,
  },
  [right.id]: {
    ...right.ref,
    columnCount: 1,
    formula: "=Inputs!$B$7*2",
    numberFormat: "0.0%",
    rowCount: 1,
    value: 0.2,
  },
  [leaf.id]: {
    ...leaf.ref,
    columnCount: 1,
    formula: null,
    numberFormat: "0.0%",
    rowCount: 1,
    value: 0.2,
  },
  [range.id]: {
    ...range.ref,
    columnCount: 4,
    formula: null,
    numberFormat: "General",
    rowCount: 3,
    value: null,
  },
};

const graph = expandNode(
  createGraphState(root),
  root.id,
  "precedent",
  [left, right],
).state;
const state = createQuickTraceTree(graph, "precedent", factsById);

const selectRow = (
  current: QuickTraceTreeState,
  nodeId: string,
): QuickTraceTreeState => {
  const row = projectQuickTraceRows(current).find((candidate) =>
    candidate.node.id === nodeId);
  if (row === undefined) throw new Error(`Missing test row ${nodeId}`);
  return { ...current, selectedRowId: row.rowId };
};

describe("Quick Trace tree", () => {
  it("projects root and expanded precedents in Excel-returned order", () => {
    expect(projectQuickTraceRows(state).map((row) => [row.depth, row.node.id]))
      .toEqual([[0, root.id], [1, left.id], [1, right.id]]);
  });

  it("moves down and up without crossing the visible bounds", () => {
    const down = moveQuickTraceSelection(state, "next");
    expect(projectQuickTraceRows(down).find((row) =>
      row.rowId === down.selectedRowId)?.node.id).toBe(left.id);
    expect(moveQuickTraceSelection(state, "previous").selectedRowId)
      .toBe(state.selectedRowId);
    const bottom = moveQuickTraceSelection(
      moveQuickTraceSelection(state, "next"),
      "next",
    );
    expect(moveQuickTraceSelection(bottom, "next").selectedRowId)
      .toBe(bottom.selectedRowId);
  });

  it("requests one lazy load on right and expands after the graph arrives", () => {
    const selected = selectRow(state, left.id);
    expect(moveQuickTraceRight(selected).loadNodeId).toBe(left.id);
    const expandedGraph = expandNode(graph, left.id, "precedent", [leaf]).state;
    const expanded = applyQuickTraceExpansion(selected, expandedGraph, factsById);
    expect(projectQuickTraceRows(expanded).map((row) => row.node.id))
      .toContain(leaf.id);
  });

  it("collapses on left, then moves to the parent when already collapsed", () => {
    const selected = selectRow(state, left.id);
    const expandedGraph = expandNode(graph, left.id, "precedent", [leaf]).state;
    const expanded = applyQuickTraceExpansion(selected, expandedGraph, factsById);
    const collapsed = moveQuickTraceLeft(expanded);
    expect(projectQuickTraceRows(collapsed).some((row) => row.node.id === leaf.id))
      .toBe(false);
    expect(moveQuickTraceLeft(collapsed).selectedRowId)
      .toBe(projectQuickTraceRows(collapsed)[0]?.rowId);
  });

  it("repairs a hidden child selection to its collapsed parent", () => {
    const expandedGraph = expandNode(graph, left.id, "precedent", [leaf]).state;
    const expanded = applyQuickTraceExpansion(
      selectRow(state, left.id),
      expandedGraph,
      factsById,
    );
    const childSelected = selectRow(expanded, leaf.id);
    const parentSelected = moveQuickTraceLeft(childSelected);
    const collapsed = moveQuickTraceLeft(parentSelected);

    expect(projectQuickTraceRows(collapsed).find((row) =>
      row.rowId === collapsed.selectedRowId)?.node.id).toBe(left.id);
  });

  it("renders a cycle occurrence as a leaf with a unique row id", () => {
    const cycleGraph = expandNode(
      expandNode(createGraphState(root), root.id, "precedent", [left]).state,
      left.id,
      "precedent",
      [root],
    ).state;
    const cycleState = createQuickTraceTree(cycleGraph, "precedent", factsById);
    const rows = projectQuickTraceRows({
      ...cycleState,
      expandedNodeIds: [root.id, left.id],
    });
    expect(rows.at(-1)).toMatchObject({ disclosure: "leaf", issue: "cycle" });
    expect(new Set(rows.map((row) => row.rowId)).size).toBe(rows.length);
  });

  it("gives shared nodes unique occurrence ids in separate branches", () => {
    const shared = makeNode("Inputs", "$B$9");
    let sharedGraph = graph;
    sharedGraph = expandNode(sharedGraph, left.id, "precedent", [shared]).state;
    sharedGraph = expandNode(sharedGraph, right.id, "precedent", [shared]).state;
    const rows = projectQuickTraceRows({
      ...createQuickTraceTree(sharedGraph, "precedent", factsById),
      expandedNodeIds: [root.id, left.id, right.id],
    });
    const occurrences = rows.filter((row) => row.node.id === shared.id);

    expect(occurrences).toHaveLength(2);
    expect(occurrences[0]?.rowId).not.toBe(occurrences[1]?.rowId);
  });

  it("shows dimensions for a multi-cell dependency instead of a scalar", () => {
    const rangeGraph = expandNode(
      createGraphState(root),
      root.id,
      "precedent",
      [range],
    ).state;
    const rangeState = createQuickTraceTree(rangeGraph, "precedent", factsById);
    expect(projectQuickTraceRows(rangeState).find((row) =>
      row.node.id === range.id)?.displayValue).toBe("3 × 4 range");
  });

  it("renders Blank for null scalar values and unavailable for missing facts", () => {
    const blank = makeNode("Inputs", "$B$8");
    const blankFacts: CellFacts = {
      ...blank.ref,
      columnCount: 1,
      formula: null,
      numberFormat: "General",
      rowCount: 1,
      value: null,
    };
    const valueGraph = expandNode(
      createGraphState(root),
      root.id,
      "precedent",
      [blank, leaf],
    ).state;
    const rows = projectQuickTraceRows(createQuickTraceTree(
      valueGraph,
      "precedent",
      { ...factsById, [blank.id]: blankFacts, [leaf.id]: undefined as never },
    ));

    expect(rows.find((row) => row.node.id === blank.id)?.displayValue).toBe("Blank");
    expect(rows.find((row) => row.node.id === leaf.id)).toMatchObject({
      displayValue: "Unavailable",
      issue: "unavailable",
    });
  });

  it("moves to the first child when Right is pressed on an expanded row", () => {
    const selected = selectRow(state, root.id);
    const moved = moveQuickTraceRight(selected).state;
    expect(projectQuickTraceRows(moved).find((row) =>
      row.rowId === moved.selectedRowId)?.node.id).toBe(left.id);
  });

  it.each(["dynamic", "external"] as const)(
    "never requests a load for %s leaves",
    (issue) => {
      const unresolvedNode = {
        ...leaf,
        id: `${leaf.id}::${issue}`,
        unresolved: issue,
      };
      const unresolvedGraph = expandNode(
        createGraphState(root),
        root.id,
        "precedent",
        [unresolvedNode],
      ).state;
      const unresolved = selectRow(
        createQuickTraceTree(unresolvedGraph, "precedent", factsById),
        unresolvedNode.id,
      );
      expect(moveQuickTraceRight(unresolved).loadNodeId).toBeUndefined();
      expect(projectQuickTraceRows(unresolved).at(-1)?.issue).toBe(issue);
    },
  );

  it("treats a known empty expansion as a leaf and does not reload it", () => {
    const emptyGraph = expandNode(graph, left.id, "precedent", []).state;
    const empty = selectRow(
      createQuickTraceTree(emptyGraph, "precedent", factsById),
      left.id,
    );
    expect(projectQuickTraceRows(empty).find((row) =>
      row.node.id === left.id)?.disclosure).toBe("leaf");
    expect(moveQuickTraceRight(empty).loadNodeId).toBeUndefined();
  });

  it("projects no more canonical nodes than the graph cap", () => {
    const capped = expandNode(
      createGraphState(root, { maxVisibleNodes: 3 }),
      root.id,
      "precedent",
      [left, right, leaf],
    );
    expect(capped.capped).toBe(true);
    expect(projectQuickTraceRows(
      createQuickTraceTree(capped.state, "precedent", factsById),
    )).toHaveLength(3);
  });
});
