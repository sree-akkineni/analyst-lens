import type {
  DependencyDirection,
  DependencyEdge,
  DependencyNode,
} from "../../shared/types";

export type GraphExpansion = Readonly<{
  dependencies: readonly DependencyNode[];
  direction: DependencyDirection;
  key: string;
  parentId: string;
}>;

export type GraphState = Readonly<{
  edges: readonly DependencyEdge[];
  expansions: readonly GraphExpansion[];
  maxVisibleNodes: number;
  nodes: readonly DependencyNode[];
  origin: DependencyNode;
}>;

export type GraphOptions = Readonly<{ maxVisibleNodes?: number }>;

const DEFAULT_MAX_VISIBLE_NODES = 500;

const expansionKey = (
  parentId: string,
  direction: DependencyDirection,
): string => `${parentId}::${direction}`;

const edgeId = (source: string, target: string): string =>
  `${source}->${target}`;

const hasPath = (
  adjacency: ReadonlyMap<string, ReadonlySet<string>>,
  start: string,
  target: string,
): boolean => {
  if (start === target) return true;
  const pending = [start];
  const visited = new Set<string>();
  while (pending.length > 0) {
    const current = pending.pop();
    if (current === undefined || visited.has(current)) continue;
    visited.add(current);
    for (const next of adjacency.get(current) ?? []) {
      if (next === target) return true;
      if (!visited.has(next)) pending.push(next);
    }
  }
  return false;
};

const rebuild = (
  origin: DependencyNode,
  candidates: readonly GraphExpansion[],
  maxVisibleNodes: number,
): Readonly<{ capped: boolean; state: GraphState }> => {
  const nodes = new Map<string, DependencyNode>([[origin.id, origin]]);
  const edges = new Map<string, DependencyEdge>();
  const adjacency = new Map<string, Set<string>>();
  const expansions: GraphExpansion[] = [];
  let capped = false;

  let pending = [...candidates];
  while (pending.length > 0) {
    const deferred: GraphExpansion[] = [];
    let progressed = false;
    for (const expansion of pending) {
      const parent = nodes.get(expansion.parentId);
      if (parent === undefined) {
        deferred.push(expansion);
        continue;
      }
      if (parent.unresolved !== undefined) continue;
      progressed = true;
      expansions.push(expansion);
      const seenDependencies = new Set<string>();
      for (const dependency of expansion.dependencies) {
        if (seenDependencies.has(dependency.id)) continue;
        seenDependencies.add(dependency.id);
        if (!nodes.has(dependency.id)) {
          if (nodes.size >= maxVisibleNodes) {
            capped = true;
            continue;
          }
          nodes.set(dependency.id, dependency);
        }
        const source = expansion.direction === "precedent"
          ? dependency.id
          : expansion.parentId;
        const target = expansion.direction === "precedent"
          ? expansion.parentId
          : dependency.id;
        const id = edgeId(source, target);
        if (edges.has(id)) continue;
        const cycle = hasPath(adjacency, target, source);
        edges.set(id, { cycle, id, source, target });
        const targets = adjacency.get(source) ?? new Set<string>();
        targets.add(target);
        adjacency.set(source, targets);
      }
    }
    if (!progressed) break;
    pending = deferred;
  }

  return {
    capped,
    state: {
      edges: [...edges.values()],
      expansions,
      maxVisibleNodes,
      nodes: [...nodes.values()],
      origin,
    },
  };
};

export const createGraphState = (
  origin: DependencyNode,
  options: GraphOptions = {},
): GraphState => ({
  edges: [],
  expansions: [],
  maxVisibleNodes: options.maxVisibleNodes ?? DEFAULT_MAX_VISIBLE_NODES,
  nodes: [origin],
  origin,
});

export const expandNode = (
  state: GraphState,
  parentId: string,
  direction: DependencyDirection,
  dependencies: readonly DependencyNode[],
  options: GraphOptions = {},
): Readonly<{ capped: boolean; state: GraphState }> => {
  const parent = state.nodes.find((current) => current.id === parentId);
  if (parent === undefined || parent.unresolved !== undefined) {
    return { capped: false, state };
  }
  const key = expansionKey(parentId, direction);
  const expansion: GraphExpansion = {
    dependencies: [...dependencies],
    direction,
    key,
    parentId,
  };
  const existingIndex = state.expansions.findIndex((current) => current.key === key);
  const candidates = existingIndex < 0
    ? [...state.expansions, expansion]
    : state.expansions.map((current, index) =>
        index === existingIndex ? expansion : current);
  return rebuild(
    state.origin,
    candidates,
    options.maxVisibleNodes ?? state.maxVisibleNodes,
  );
};

export const collapseBranch = (
  state: GraphState,
  parentId: string,
  direction: DependencyDirection,
): GraphState => collapseBranchResult(state, parentId, direction).state;

export const collapseBranchResult = (
  state: GraphState,
  parentId: string,
  direction: DependencyDirection,
): Readonly<{ capped: boolean; state: GraphState }> => rebuild(
  state.origin,
  state.expansions.filter((current) =>
    current.key !== expansionKey(parentId, direction)),
  state.maxVisibleNodes,
);

export const searchGraph = (
  state: GraphState,
  query: string,
): DependencyNode[] => {
  const normalized = query.trim().toLocaleLowerCase().replaceAll("$", "");
  if (normalized.length === 0) return [];
  return state.nodes.filter((current) => [
    current.id,
    current.ref.address,
    current.ref.sheetName,
    current.ref.workbookId,
    current.value === undefined ? "" : String(current.value),
  ].some((candidate) => candidate
    .toLocaleLowerCase()
    .replaceAll("$", "")
    .includes(normalized)));
};
