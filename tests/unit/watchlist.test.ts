import { describe, expect, it } from "vitest";
import { createGraphState, expandNode } from "../../src/domain/dependencies/graph";
import {
  affectedPins,
  createWatchItem,
  hydrateWatchlistMetadata,
  refreshWatchItem,
  serializeForWorkbook,
  watchDisplayRef,
} from "../../src/domain/watchlist/watchlist";
import { canonicalCellId, type CellFacts, type CellRef, type DependencyNode } from "../../src/shared/types";

const ref = (sheetName: string, address: string, sheetId = sheetName.toLowerCase()): CellRef => ({
  address,
  sheetId,
  sheetName,
  workbookId: "workbook-1",
});

const facts = (
  sheetName: string,
  address: string,
  value: CellFacts["value"],
  formula: string | null,
): CellFacts => ({
  ...ref(sheetName, address),
  columnCount: 1,
  formula,
  numberFormat: "0.0x",
  rowCount: 1,
  value,
});

const node = (current: CellRef): DependencyNode => ({
  id: canonicalCellId(current),
  ref: current,
});

describe("watchlist domain", () => {
  it("stores no value, formula, sheet name, or graph state in workbook metadata", () => {
    const item = {
      ...createWatchItem(facts("Returns", "$F$42", 18.4, "=IRR(B2:B8)")),
      category: "Returns",
      label: "IRR",
    };

    expect(serializeForWorkbook([item])).toEqual([{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "returns",
      workbookId: "workbook-1",
    }]);
    expect(JSON.stringify(serializeForWorkbook([item]))).not.toMatch(/18\.4|IRR\(|sheetName|graph/u);
  });

  it("tracks value and formula changes independently", () => {
    const pinned = createWatchItem(facts("Returns", "$F$42", 18.4, "=IRR(B2:B8)"));
    const valueChanged = refreshWatchItem(
      pinned,
      facts("Returns", "$F$42", 19.1, "=IRR(B2:B8)"),
    );
    const formulaChanged = refreshWatchItem(
      valueChanged,
      facts("Returns", "$F$42", 19.1, "=XIRR(B2:B8,A2:A8)"),
    );

    expect(valueChanged).toMatchObject({
      currentValue: 19.1,
      formulaChanged: false,
      lastValue: 18.4,
      valueChanged: true,
    });
    expect(formulaChanged).toMatchObject({
      currentFormula: "=XIRR(B2:B8,A2:A8)",
      formulaChanged: true,
      lastFormula: "=IRR(B2:B8)",
      valueChanged: false,
    });
  });

  it("returns exact pinned outputs downstream of the current assumption", () => {
    const assumption = ref("Inputs", "$B$7");
    const bridge = ref("DCF", "$D$20");
    const irr = ref("Returns", "$F$42");
    const revenue = ref("Summary", "$C$12");
    let graph = createGraphState(node(assumption));
    graph = expandNode(graph, canonicalCellId(assumption), "dependent", [node(bridge)]).state;
    graph = expandNode(graph, canonicalCellId(bridge), "dependent", [node(irr)]).state;

    expect(affectedPins(graph, assumption, [irr, revenue])).toEqual(["Returns!$F$42"]);
  });

  it("hydrates metadata without inventing sensitive values and preserves missing refs", () => {
    const [item] = hydrateWatchlistMetadata([{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "renamed-sheet-id",
      workbookId: "workbook-1",
    }]);

    expect(item).toMatchObject({
      category: "Returns",
      currentFormula: null,
      currentValue: null,
      label: "IRR",
      status: "loading",
    });
    expect(item === undefined ? "" : watchDisplayRef(item.ref)).toBe("Unknown sheet!$F$42");
  });
});
