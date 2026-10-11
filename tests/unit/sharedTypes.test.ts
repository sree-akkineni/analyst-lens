import { describe, expect, it } from "vitest";
import { canonicalCellId, type CellRef } from "../../src/shared/types";

describe("canonicalCellId", () => {
  it("separates workbook, sheet id, and canonical address", () => {
    const ref: CellRef = { workbookId: "wb-1", sheetId: "sheet-7", sheetName: "DCF", address: "$F$42" };
    expect(canonicalCellId(ref)).toBe("wb-1::sheet-7::$F$42");
  });
});
