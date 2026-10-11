import type { GraphState } from "./graph";
import type {
  CellFacts,
  DependencyDirection,
  DependencyNode,
} from "../../shared/types";

export type QuickTraceDisclosure = "collapsed" | "expanded" | "leaf";
export type QuickTraceIssue =
  | "cycle"
  | "dynamic"
  | "external"
  | "unavailable"
  | null;

export type QuickTraceRow = Readonly<{
  depth: number;
  disclosure: QuickTraceDisclosure;
  displayValue: string;
  issue: QuickTraceIssue;
  node: DependencyNode;
  parentRowId: string | null;
  rowId: string;
}>;

export type QuickTraceTreeState = Readonly<{
  direction: DependencyDirection;
  expandedNodeIds: readonly string[];
  factsByNodeId: Readonly<Record<string, CellFacts>>;
  graph: GraphState;
  selectedRowId: string;
}>;

export type QuickTraceTreeTransition = Readonly<{
  loadNodeId?: string;
  state: QuickTraceTreeState;
}>;

const rowIdForPath = (path: readonly string[]): string =>
  path.map((nodeId) => encodeURIComponent(nodeId)).join("/");

const factsDisplayValue = (facts: CellFacts | undefined): string => {
  if (facts === undefined) return "Unavailable";
  if (facts.rowCount > 1 || facts.columnCount > 1) {
    return `${facts.rowCount} × ${facts.columnCount} range`;
  }
  return facts.value === null ? "Blank" : String(facts.value);
};

const matchingExpansion = (
  state: QuickTraceTreeState,
  nodeId: string,
) => state.graph.expansions.find((expansion) =>
  expansion.parentId === nodeId && expansion.direction === state.direction);

const canonicalDependencies = (
  state: QuickTraceTreeState,
  nodeId: string,
): readonly DependencyNode[] | undefined => {
  const expansion = matchingExpansion(state, nodeId);
  if (expansion === undefined) return undefined;
  const visibleNodeIds = new Set(state.graph.nodes.map((node) => node.id));
  return expansion.dependencies.filter((node) => visibleNodeIds.has(node.id));
};

export const projectQuickTraceRows = (
  state: QuickTraceTreeState,
): QuickTraceRow[] => {
  const rows: QuickTraceRow[] = [];
  const expanded = new Set(state.expandedNodeIds);

  const append = (
    node: DependencyNode,
    path: readonly string[],
    parentRowId: string | null,
    ancestorNodeIds: ReadonlySet<string>,
  ) => {
    const rowId = rowIdForPath(path);
    const cycle = ancestorNodeIds.has(node.id);
    const facts = state.factsByNodeId[node.id];
    const dependencies = cycle
      ? undefined
      : canonicalDependencies(state, node.id);
    const issue: QuickTraceIssue = cycle
      ? "cycle"
      : node.unresolved ?? (facts === undefined ? "unavailable" : null);
    const disclosure: QuickTraceDisclosure = cycle || node.unresolved !== undefined
      ? "leaf"
      : dependencies?.length === 0
        ? "leaf"
        : dependencies !== undefined && expanded.has(node.id)
          ? "expanded"
          : "collapsed";

    rows.push({
      depth: path.length - 1,
      disclosure,
      displayValue: factsDisplayValue(facts),
      issue,
      node,
      parentRowId,
      rowId,
    });

    if (disclosure !== "expanded" || dependencies === undefined) return;
    const descendants = new Set(ancestorNodeIds);
    descendants.add(node.id);
    dependencies.forEach((dependency) => {
      append(dependency, [...path, dependency.id], rowId, descendants);
    });
  };

  append(state.graph.origin, [state.graph.origin.id], null, new Set());
  return rows;
};

const repairSelection = (state: QuickTraceTreeState): QuickTraceTreeState => {
  const rows = projectQuickTraceRows(state);
  if (rows.some((row) => row.rowId === state.selectedRowId)) return state;
  const ancestor = rows
    .filter((row) => state.selectedRowId.startsWith(`${row.rowId}/`))
    .at(-1);
  return {
    ...state,
    selectedRowId: ancestor?.rowId ?? rows[0]?.rowId ?? state.selectedRowId,
  };
};

export const createQuickTraceTree = (
  graph: GraphState,
  direction: DependencyDirection,
  factsByNodeId: Readonly<Record<string, CellFacts>> = {},
): QuickTraceTreeState => {
  const hasLoadedRoot = graph.expansions.some((expansion) =>
    expansion.parentId === graph.origin.id && expansion.direction === direction);
  const selectedRowId = rowIdForPath([graph.origin.id]);
  return {
    direction,
    expandedNodeIds: hasLoadedRoot ? [graph.origin.id] : [],
    factsByNodeId,
    graph,
    selectedRowId,
  };
};

export const moveQuickTraceSelection = (
  state: QuickTraceTreeState,
  movement: "previous" | "next",
): QuickTraceTreeState => {
  const rows = projectQuickTraceRows(state);
  const selectedIndex = rows.findIndex((row) => row.rowId === state.selectedRowId);
  if (selectedIndex < 0) return repairSelection(state);
  const offset = movement === "previous" ? -1 : 1;
  const nextIndex = Math.max(0, Math.min(rows.length - 1, selectedIndex + offset));
  return { ...state, selectedRowId: rows[nextIndex]?.rowId ?? state.selectedRowId };
};

export const moveQuickTraceLeft = (
  state: QuickTraceTreeState,
): QuickTraceTreeState => {
  const selected = projectQuickTraceRows(state).find((row) =>
    row.rowId === state.selectedRowId);
  if (selected === undefined) return repairSelection(state);
  if (selected.disclosure === "expanded") {
    return repairSelection({
      ...state,
      expandedNodeIds: state.expandedNodeIds.filter((id) => id !== selected.node.id),
    });
  }
  return selected.parentRowId === null
    ? state
    : { ...state, selectedRowId: selected.parentRowId };
};

export const moveQuickTraceRight = (
  state: QuickTraceTreeState,
): QuickTraceTreeTransition => {
  const rows = projectQuickTraceRows(state);
  const selectedIndex = rows.findIndex((row) => row.rowId === state.selectedRowId);
  const selected = rows[selectedIndex];
  if (selected === undefined || selected.disclosure === "leaf") return { state };

  if (selected.disclosure === "expanded") {
    const child = rows[selectedIndex + 1];
    return child?.parentRowId === selected.rowId
      ? { state: { ...state, selectedRowId: child.rowId } }
      : { state };
  }

  const dependencies = canonicalDependencies(state, selected.node.id);
  if (dependencies === undefined) return { loadNodeId: selected.node.id, state };
  if (dependencies.length === 0) return { state };
  return {
    state: repairSelection({
      ...state,
      expandedNodeIds: [...new Set([...state.expandedNodeIds, selected.node.id])],
    }),
  };
};

export const applyQuickTraceExpansion = (
  state: QuickTraceTreeState,
  graph: GraphState,
  factsByNodeId: Readonly<Record<string, CellFacts>>,
): QuickTraceTreeState => {
  const selected = projectQuickTraceRows(state).find((row) =>
    row.rowId === state.selectedRowId);
  const expandedNodeIds = selected === undefined
    ? state.expandedNodeIds
    : [...new Set([...state.expandedNodeIds, selected.node.id])];
  return repairSelection({
    ...state,
    expandedNodeIds,
    factsByNodeId: { ...state.factsByNodeId, ...factsByNodeId },
    graph,
  });
};
