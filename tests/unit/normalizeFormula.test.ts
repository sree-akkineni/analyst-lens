import { describe, expect, it } from "vitest";
import type { CellRef } from "../../src/shared/types";
import { normalizeFormula } from "../../src/domain/formulas/normalizeFormula";

const cell = (address: string): CellRef => ({
  address,
  sheetId: "sheet-1",
  sheetName: "DCF",
  workbookId: "workbook-1",
});

describe("formula normalization", () => {
  it("normalizes copied formulas relative to their origin", () => {
    expect(normalizeFormula("=F10*$B$2", cell("F12"))).toBe("=R[-2]C[0]*R2C2");
    expect(normalizeFormula("=G10*$B$2", cell("G12"))).toBe("=R[-2]C[0]*R2C2");
  });

  it.each([
    ["=$A1+B$2+$C$3", "B4", "=R[-3]C1+R2C[0]+R3C3"],
    ["=SUM(A1:B3)", "C4", "=SUM(R[-3]C[-2]:R[-1]C[-1])"],
    ["='Revenue Build'!A1:B2", "C3", "='Revenue Build'!R[-2]C[-2]:R[-1]C[-1]"],
    ["='O''Brien'!$F10", "H12", "='O''Brien'!R[-2]C6"],
    ["=[Book.xlsx]Sheet1!A1+$B$2", "C3", "=[Book.xlsx]Sheet1!R[-2]C[-2]+R2C2"],
  ])("normalizes absolute, mixed, range, sheet, and external fixture %s", (
    formula,
    origin,
    expected,
  ) => {
    expect(normalizeFormula(formula, cell(origin))).toBe(expected);
  });

  it("leaves strings, names, and structured references unchanged", () => {
    expect(normalizeFormula(
      "=IF(A1=\"B2\",Tax_Rate,Table1[Revenue])",
      cell("C3"),
    )).toBe("=IF(R[-2]C[-2]=\"B2\",Tax_Rate,Table1[Revenue])");
  });

  it("leaves A1-shaped names and out-of-grid identifiers unchanged", () => {
    expect(normalizeFormula("=A1_Forecast+XFE1+A1048577", cell("B2")))
      .toBe("=A1_Forecast+XFE1+A1048577");
  });

  it("normalizes whole-column and whole-row ranges", () => {
    expect(normalizeFormula("=SUM(A:C)+SUM($1:$3)", cell("C3")))
      .toBe("=SUM(C[-2]:C[0])+SUM(R1:R3)");
    expect(normalizeFormula("=SUM($A:$C)+SUM(1:3)", cell("C3")))
      .toBe("=SUM(C1:C3)+SUM(R[-2]:R[0])");
  });

  it("does not normalize an external defined name as a cell", () => {
    expect(normalizeFormula("='[Book.xlsx]Sheet1'!Tax_Rate", cell("C3")))
      .toBe("='[Book.xlsx]Sheet1'!Tax_Rate");
  });

  it("uses invariant formulas for logic even when a localized display fixture exists", () => {
    const fixture = {
      display: "=SOMME(A1;1,5)",
      invariant: "=SUM(A1,1.5)",
    };

    expect(fixture.display).toContain("SOMME");
    expect(normalizeFormula(fixture.invariant, cell("B2")))
      .toBe("=SUM(R[-1]C[-1],1.5)");
  });

  it("rejects a non-cell origin rather than guessing", () => {
    expect(() => normalizeFormula("=A1", cell("not-an-address")))
      .toThrow("single A1 cell");
  });
});
