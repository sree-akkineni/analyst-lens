import { describe, expect, it } from "vitest";
import {
  runRules,
  type HealthCellFacts,
  type HealthWorkbookFacts,
} from "../../src/domain/health/rules";

const ref = (address = "$A$1") => ({
  address,
  sheetId: "sheet-1",
  sheetName: "Model",
  workbookId: "workbook-1",
});

const cell = (
  address: string,
  formula: string | null,
  value: HealthCellFacts["value"],
  numberFormat = "General",
  fontColor = "#000000",
): HealthCellFacts => ({
  ...ref(address),
  columnCount: 1,
  fontColor,
  formula,
  numberFormat,
  rowCount: 1,
  value,
});

const workbook = (
  cells: readonly HealthCellFacts[],
  calculationMode: HealthWorkbookFacts["calculationMode"] = "automatic",
): HealthWorkbookFacts => ({ calculationMode, cells, workbookRef: ref() });

const fixtures: Readonly<Record<string, HealthWorkbookFacts>> = {
  EMBEDDED_CONSTANT: workbook([cell("$A$1", "=B1*2", 10)]),
  ERR_VALUE: workbook([cell("$A$1", "=B1", "#REF!")]),
  EXTERNAL_LINK: workbook([cell("$A$1", "='[Other.xlsx]Inputs'!A1", 10)]),
  FORMULA_DISCONTINUITY: workbook([
    cell("$A$1", "=A10", 1),
    cell("$B$1", "=B10", 2),
    cell("$C$1", "=SUM(C10:C11)", 3),
  ]),
  HARDCODE_IN_RUN: workbook([
    cell("$A$2", "=A1", 1),
    cell("$B$2", null, 5),
    cell("$C$2", "=C1", 3),
  ]),
  INCONSISTENT_FORMAT: workbook([
    cell("$A$3", "=A2", 1, "#,##0"),
    cell("$B$3", "=B2", 2, "#,##0"),
    cell("$C$3", "=C2", 3, "0.0%"),
  ]),
  MANUAL_CALCULATION: workbook([cell("$A$1", "=B1", 1)], "manual"),
  VOLATILE_FUNCTION: workbook([cell("$A$1", "=OFFSET(B1,2,0)", 1)]),
};

const exclusions: Readonly<Record<string, HealthWorkbookFacts>> = {
  EMBEDDED_CONSTANT: workbook([cell("$A$1", "=B1+0+1-1", 1)]),
  ERR_VALUE: workbook([cell("$A$1", "=B1", 1)]),
  EXTERNAL_LINK: workbook([cell("$A$1", "='Inputs'!A1", 1)]),
  FORMULA_DISCONTINUITY: workbook([
    cell("$A$1", "=A10", 1), cell("$B$1", "=B10", 2), cell("$C$1", "=C10", 3),
  ]),
  HARDCODE_IN_RUN: workbook([
    cell("$A$2", "=A1", 1), cell("$B$2", "=B1", 2), cell("$C$2", "=C1", 3),
  ]),
  INCONSISTENT_FORMAT: workbook([
    cell("$A$3", null, "Revenue", "General"),
    cell("$B$3", null, "Q1", "General"),
    cell("$C$3", null, "Q2", "0.0%"),
  ]),
  MANUAL_CALCULATION: workbook([cell("$A$1", "=B1", 1)], "automatic"),
  VOLATILE_FUNCTION: workbook([cell("$A$1", "=OFFSET_RATE+B1", 1)]),
};

describe("model health rules", () => {
  it.each(Object.entries(fixtures))("emits exact evidence for %s", (ruleId, fixture) => {
    expect(runRules(fixture)).toContainEqual(expect.objectContaining({
      evidence: expect.any(String),
      ref: expect.objectContaining({ workbookId: "workbook-1" }),
      remediation: expect.any(String),
      ruleId,
      severity: expect.stringMatching(/error|warning|info/u),
    }));
  });

  it.each(Object.entries(exclusions))("does not emit the %s false-positive fixture", (ruleId, fixture) => {
    expect(runRules(fixture).some((finding) => finding.ruleId === ruleId)).toBe(false);
  });

  it("excludes safe embedded constants and identifier lookalikes", () => {
    const findings = runRules(workbook([
      cell("$A$1", "=B1+0+1-1", 1),
      cell("$A$2", "=OFFSET_RATE+B2", 2),
    ]));

    expect(findings.some(({ ruleId }) => ruleId === "EMBEDDED_CONSTANT")).toBe(false);
    expect(findings.some(({ ruleId }) => ruleId === "VOLATILE_FUNCTION")).toBe(false);
  });

  it("does not call isolated values, tied formulas, or uniform formats inconsistent", () => {
    const findings = runRules(workbook([
      cell("$A$1", null, 5),
      cell("$A$3", "=SUM(A1:A2)", 1, "#,##0"),
      cell("$B$3", "=AVERAGE(B1:B2)", 2, "#,##0"),
      cell("$C$3", "=MAX(C1:C2)", 3, "#,##0"),
    ]));

    expect(findings.some(({ ruleId }) => ruleId === "HARDCODE_IN_RUN")).toBe(false);
    expect(findings.some(({ ruleId }) => ruleId === "FORMULA_DISCONTINUITY")).toBe(false);
    expect(findings.some(({ ruleId }) => ruleId === "INCONSISTENT_FORMAT")).toBe(false);
  });

  it("matches only finite Excel error literals", () => {
    const findings = runRules(workbook([
      cell("$A$1", null, "# Revenue"),
      cell("$A$2", null, "#REF!"),
      cell("$A$3", null, "#N/A"),
      cell("$A$4", null, "#EXTERNAL!"),
    ])).filter(({ ruleId }) => ruleId === "ERR_VALUE");

    expect(findings.map(({ ref: findingRef }) => findingRef.address)).toEqual([
      "$A$2",
      "$A$3",
      "$A$4",
    ]);
  });

  it("reports the formula error token when the cached value is not an error", () => {
    const finding = runRules(workbook([
      cell("$A$1", "=#EXTERNAL!", 0),
    ])).find(({ ruleId }) => ruleId === "ERR_VALUE");

    expect(finding?.evidence).toContain("#EXTERNAL!");
    expect(finding?.evidence).not.toBe("Excel error observed: 0");
  });

  it("breaks repeated formula and format runs at blank separators", () => {
    const findings = runRules(workbook([
      cell("$A$1", "=A10", 1, "#,##0"),
      cell("$B$1", "=B10", 2, "#,##0"),
      cell("$C$1", null, null, "General", "unknown"),
      cell("$D$1", "=SUM(D10:D11)", 3, "0.0%"),
    ]));

    expect(findings.some(({ ruleId }) => ruleId === "FORMULA_DISCONTINUITY")).toBe(false);
    expect(findings.some(({ ruleId }) => ruleId === "INCONSISTENT_FORMAT")).toBe(false);
  });

  it("requires matching neighboring formula patterns for a hardcode run", () => {
    const findings = runRules(workbook([
      cell("$A$2", "=A1", 1),
      cell("$B$2", null, 5),
      cell("$C$2", "=SUM(C1:C2)", 3),
    ]));

    expect(findings.some(({ ruleId }) => ruleId === "HARDCODE_IN_RUN")).toBe(false);
  });
});
