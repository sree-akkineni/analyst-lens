import { describe, expect, it } from "vitest";
import {
  back,
  forward,
  returnToOrigin,
  startHistory,
  visit,
} from "../../src/domain/dependencies/navigationHistory";
import type { CellRef } from "../../src/shared/types";

const cell = (address: string, workbookId = "workbook-1"): CellRef => ({
  address,
  sheetId: address.includes("!") ? address.split("!")[0] ?? "sheet-1" : "sheet-1",
  sheetName: address.includes("!") ? address.split("!")[0] ?? "DCF" : "DCF",
  workbookId,
});

describe("navigation history", () => {
  it("preserves a branch and returns to origin", () => {
    const history = visit(
      visit(startHistory(cell("$F$42")), cell("$F$40")),
      cell("Inputs!$B$7"),
    );

    expect(returnToOrigin(history).current.address).toBe("$F$42");
    expect(back(history).current.address).toBe("$F$40");
  });

  it("supports lossless back and forward navigation", () => {
    const visited = visit(visit(startHistory(cell("$F$42")), cell("$F$40")), cell("$F$38"));
    const oneBack = back(visited);
    const twoBack = back(oneBack);

    expect(oneBack.current.address).toBe("$F$40");
    expect(twoBack.current.address).toBe("$F$42");
    expect(forward(twoBack).current.address).toBe("$F$40");
    expect(forward(forward(twoBack)).current.address).toBe("$F$38");
  });

  it("clears the forward branch after a new visit", () => {
    const history = visit(visit(startHistory(cell("$F$42")), cell("$F$40")), cell("$F$38"));
    const branched = visit(back(history), cell("$G$40"));

    expect(branched.current.address).toBe("$G$40");
    expect(forward(branched)).toEqual(branched);
  });

  it("keeps at most 200 visits while retaining an explicit origin", () => {
    let history = startHistory(cell("$A$1"));
    for (let row = 2; row <= 250; row += 1) {
      history = visit(history, cell(`$A$${row}`));
    }

    expect(history.past.length + 1).toBe(200);
    expect(history.origin.address).toBe("$A$1");
    expect(returnToOrigin(history).current.address).toBe("$A$1");
  });

  it("does not add duplicate visits and starts a new window for another workbook", () => {
    const initial = startHistory(cell("$F$42"));
    expect(visit(initial, cell("$f$42"))).toEqual(initial);

    const otherWorkbook = visit(initial, cell("$A$1", "workbook-2"));
    expect(otherWorkbook).toEqual(startHistory(cell("$A$1", "workbook-2")));
  });
});
