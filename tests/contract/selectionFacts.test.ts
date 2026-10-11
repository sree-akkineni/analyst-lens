import { describe, expect, it } from "vitest";
import { readSelectionFacts } from "../../src/platform/office/selectionFacts";
import { createFakeOfficeContext } from "./fakeOfficeContext";

describe("selection facts Office contract", () => {
  it("loads a single-cell selection in one sync", async () => {
    const fake = createFakeOfficeContext().withSelection({
      address: "DCF!F42",
      formula: "=F40-F41",
      numberFormat: "$#,##0",
      value: 12,
    });

    const facts = await readSelectionFacts(fake.context);

    expect(facts).toMatchObject({
      address: "$F$42",
      formula: "=F40-F41",
      numberFormat: "$#,##0",
      sheetId: "sheet-1",
      sheetName: "DCF",
      value: 12,
      workbookId: "workbook-1",
    });
    expect(fake.selectionLoads).toEqual([
      ["address", "columnCount", "formulas", "numberFormat", "rowCount", "values"],
    ]);
    expect(fake.worksheetLoads).toEqual([["id", "name"]]);
    expect(fake.syncCount).toBe(1);
  });
});
