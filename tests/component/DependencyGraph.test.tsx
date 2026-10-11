import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createGraphState, expandNode } from "../../src/domain/dependencies/graph";
import { canonicalCellId, type DependencyNode } from "../../src/shared/types";
import { DependencyGraph } from "../../src/ui/DependencyGraph";

afterEach(cleanup);

const node = (address: string): DependencyNode => {
  const ref = {
    address,
    sheetId: "sheet-1",
    sheetName: "DCF",
    workbookId: "workbook-1",
  };
  return { id: canonicalCellId(ref), ref };
};

describe("DependencyGraph", () => {
  it("announces the 500-node cap and exposes keyboard expansion controls", () => {
    const origin = node("$A$1");
    const dependencies = Array.from({ length: 499 }, (_, index) => node(`$A$${index + 2}`));
    const state = expandNode(
      createGraphState(origin),
      origin.id,
      "dependent",
      dependencies,
    ).state;
    const dispatch = vi.fn();
    render(<DependencyGraph capped dispatch={dispatch} state={state} />);

    expect(screen.getByRole("status").textContent).toContain("500-node limit reached");
    fireEvent.click(screen.getByRole("button", { name: "Expand precedents for DCF $A$1" }));
    expect(dispatch).toHaveBeenCalledWith({
      direction: "precedent",
      nodeId: origin.id,
      type: "expandNode",
    });
  });

  it("renders cycle and unresolved evidence without inventing edges", () => {
    const origin = node("$A$1");
    const external: DependencyNode = {
      ...node("external::Book.xlsx"),
      unresolved: "external",
    };
    const state = expandNode(
      createGraphState(origin),
      origin.id,
      "precedent",
      [external],
    ).state;
    render(<DependencyGraph dispatch={() => undefined} state={state} />);

    expect(screen.getByText("External leaf")).toBeTruthy();
    expect(screen.queryByRole("button", { name: /Navigate to DCF external/u })).toBeNull();
    expect(screen.getByRole("list", { name: "Dependency edges" })).toBeTruthy();
  });

  it("supports keyboard navigation, collapse, and search on visible graph nodes", () => {
    const origin = node("$A$1");
    const child = node("$B$2");
    const state = expandNode(
      createGraphState(origin),
      origin.id,
      "dependent",
      [child],
    ).state;
    const dispatch = vi.fn();
    render(<DependencyGraph dispatch={dispatch} state={state} />);

    fireEvent.click(screen.getByRole("button", { name: "Navigate to DCF $B$2" }));
    expect(dispatch).toHaveBeenCalledWith({ ref: child.ref, type: "navigate" });
    fireEvent.click(screen.getByRole("button", { name: "Collapse dependents for DCF $A$1" }));
    expect(dispatch).toHaveBeenCalledWith({
      direction: "dependent",
      nodeId: origin.id,
      type: "collapseBranch",
    });

    fireEvent.change(screen.getByRole("searchbox", { name: "Search dependency map" }), {
      target: { value: "$B$2" },
    });
    expect(screen.getByRole("button", { name: "Navigate to DCF $B$2" })).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Navigate to DCF $A$1" })).toBeNull();
  });
});
