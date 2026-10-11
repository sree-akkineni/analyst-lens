import { useState } from "react";
import { searchGraph, type GraphState } from "../domain/dependencies/graph";
import type { InspectorIntent } from "../runtime/controller";

type DependencyGraphProps = {
  capped?: boolean;
  dispatch(intent: InspectorIntent): void;
  state: GraphState;
};

const nodeLabel = (state: GraphState, id: string): string => {
  const node = state.nodes.find((current) => current.id === id);
  return node === undefined ? id : `${node.ref.sheetName} ${node.ref.address}`;
};

export function DependencyGraph({
  capped = false,
  dispatch,
  state,
}: DependencyGraphProps) {
  const [query, setQuery] = useState("");
  const visibleNodes = query.trim().length === 0 ? state.nodes : searchGraph(state, query);
  return (
    <section aria-label="Dependency map" className="dependency-graph">
      {capped ? (
        <p className="graph-cap" role="status">
          500-node limit reached. Collapse a branch or continue from a narrower node.
        </p>
      ) : null}
      <label className="graph-search">
        Search dependency map
        <input
          onChange={(event) => setQuery(event.currentTarget.value)}
          type="search"
          value={query}
        />
      </label>
      <ul aria-label="Dependency nodes" className="graph-node-list">
        {visibleNodes.map((node) => (
          <li className="graph-node" key={node.id}>
            <div>
              {node.unresolved === undefined ? (
                <button
                  aria-label={`Navigate to ${node.ref.sheetName} ${node.ref.address}`}
                  className="graph-navigate"
                  onClick={() => dispatch({ ref: node.ref, type: "navigate" })}
                  type="button"
                >{node.ref.sheetName} {node.ref.address}</button>
              ) : <strong>{node.ref.sheetName} {node.ref.address}</strong>}
              {node.unresolved === "external" ? <span>External leaf</span> : null}
              {node.unresolved === "dynamic" ? <span>Dynamic unresolved leaf</span> : null}
            </div>
            {node.unresolved === undefined ? (
              <div className="graph-node-actions">
                <button
                  aria-label={`Expand precedents for ${node.ref.sheetName} ${node.ref.address}`}
                  onClick={() => dispatch({
                    direction: "precedent",
                    nodeId: node.id,
                    type: "expandNode",
                  })}
                  type="button"
                >
                  ↑
                </button>
                <button
                  aria-label={`Expand dependents for ${node.ref.sheetName} ${node.ref.address}`}
                  onClick={() => dispatch({
                    direction: "dependent",
                    nodeId: node.id,
                    type: "expandNode",
                  })}
                  type="button"
                >
                  ↓
                </button>
                {(["precedent", "dependent"] as const).map((direction) =>
                  state.expansions.some((current) =>
                    current.parentId === node.id && current.direction === direction) ? (
                      <button
                        aria-label={`Collapse ${direction}s for ${node.ref.sheetName} ${node.ref.address}`}
                        key={direction}
                        onClick={() => dispatch({
                          direction,
                          nodeId: node.id,
                          type: "collapseBranch",
                        })}
                        type="button"
                      >×{direction === "precedent" ? "↑" : "↓"}</button>
                    ) : null)}
              </div>
            ) : null}
          </li>
        ))}
      </ul>
      <ul aria-label="Dependency edges" className="graph-edge-list">
        {state.edges.map((edge) => (
          <li key={edge.id}>
            {nodeLabel(state, edge.source)} → {nodeLabel(state, edge.target)}
            {edge.cycle ? <strong> Cycle</strong> : null}
          </li>
        ))}
      </ul>
    </section>
  );
}
