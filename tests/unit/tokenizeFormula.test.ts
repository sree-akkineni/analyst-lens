import { describe, expect, it } from "vitest";
import {
  analyzeFormula,
  tokenizeFormula,
} from "../../src/domain/formulas/tokenizeFormula";

describe("formula tokenization", () => {
  it("tokenizes quoted sheets, names, literals, and external refs", () => {
    expect(tokenizeFormula(
      "='Revenue Build'!$F$12+Tax_Rate*1.5+[Book2.xlsx]Sheet1!A1",
    )).toMatchObject([
      { kind: "reference", text: "'Revenue Build'!$F$12" },
      { kind: "operator", text: "+" },
      { kind: "name", text: "Tax_Rate" },
      { kind: "operator", text: "*" },
      { kind: "number", text: "1.5" },
      { kind: "operator", text: "+" },
      { kind: "external-reference", text: "[Book2.xlsx]Sheet1!A1" },
    ]);
  });

  it("preserves exact source spans for escaped sheets, ranges, and strings", () => {
    const formula = "='O''Brien Plan'!A1:B2&\"A1 says \"\"go\"\"\"";
    const tokens = tokenizeFormula(formula);

    expect(tokens).toMatchObject([
      { kind: "reference", text: "'O''Brien Plan'!A1:B2" },
      { kind: "operator", text: "&" },
      { kind: "string", text: "\"A1 says \"\"go\"\"\"" },
    ]);
    for (const token of tokens) {
      expect(formula.slice(token.start, token.end)).toBe(token.text);
    }
  });

  it("distinguishes functions, sheet-scoped names, structured names, and errors", () => {
    expect(tokenizeFormula(
      "=IF('Tax Sheet'!Tax_Rate>0,Table1[Revenue],#REF!)",
    )).toMatchObject([
      { kind: "function", text: "IF" },
      { kind: "punctuation", text: "(" },
      { kind: "name", text: "'Tax Sheet'!Tax_Rate" },
      { kind: "operator", text: ">" },
      { kind: "number", text: "0" },
      { kind: "punctuation", text: "," },
      { kind: "name", text: "Table1[Revenue]" },
      { kind: "punctuation", text: "," },
      { kind: "error", text: "#REF!" },
      { kind: "punctuation", text: ")" },
    ]);
  });

  it("does not misclassify an A1-shaped function name as a cell reference", () => {
    expect(tokenizeFormula("=LOG10(A1)")).toMatchObject([
      { kind: "function", text: "LOG10" },
      { kind: "punctuation", text: "(" },
      { kind: "reference", text: "A1" },
      { kind: "punctuation", text: ")" },
    ]);
  });

  it("requires cell boundaries and validates Excel grid bounds", () => {
    expect(tokenizeFormula("=A1_Forecast+XFE1+A1048577")).toMatchObject([
      { kind: "name", text: "A1_Forecast" },
      { kind: "operator", text: "+" },
      { kind: "name", text: "XFE1" },
      { kind: "operator", text: "+" },
      { kind: "name", text: "A1048577" },
    ]);
  });

  it("tokenizes whole-column and whole-row ranges as references", () => {
    expect(tokenizeFormula("=SUM($A:$C)+SUM(1:3)")).toMatchObject([
      { kind: "function", text: "SUM" },
      { kind: "punctuation", text: "(" },
      { kind: "reference", text: "$A:$C" },
      { kind: "punctuation", text: ")" },
      { kind: "operator", text: "+" },
      { kind: "function", text: "SUM" },
      { kind: "punctuation", text: "(" },
      { kind: "reference", text: "1:3" },
      { kind: "punctuation", text: ")" },
    ]);
  });

  it("marks an external defined name as external but not navigable", () => {
    const analysis = analyzeFormula("='[Book.xlsx]Sheet1'!Tax_Rate");

    expect(analysis.tokens).toMatchObject([
      { kind: "external-name", text: "'[Book.xlsx]Sheet1'!Tax_Rate" },
    ]);
    expect(analysis.risks).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: "external-link" }),
      expect.objectContaining({ kind: "named-range" }),
    ]));
  });

  it("does not invent named-range risks for LET or LAMBDA bindings", () => {
    const analysis = analyzeFormula("=LET(x,1,LAMBDA(y,x+y+A1)(2))");
    const namedRisks = analysis.risks.filter((risk) => risk.kind === "named-range");

    expect(namedRisks).toEqual([]);
  });

  it("keeps LET argument parity across array constants", () => {
    const namedRisks = analyzeFormula("=LET(x,{1,2},y,3,x+y)").risks
      .filter((risk) => risk.kind === "named-range");

    expect(namedRisks).toEqual([]);
  });

  it("does not label an explicit spill reference as dynamically unresolved", () => {
    expect(analyzeFormula("=A1#").risks).not.toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: "dynamic-reference" }),
    ]));
  });

  it("reports all required formula risks with configurable safe constants", () => {
    const formula = "=OFFSET('Build'!A1,0,1)+INDIRECT(\"B2\")+[Book.xlsx]Sheet1!A1+Tax_Rate*2+#N/A";
    const analysis = analyzeFormula(formula);

    expect(new Set(analysis.risks.map((risk) => risk.kind))).toEqual(new Set([
      "cross-sheet",
      "dynamic-reference",
      "embedded-number",
      "external-link",
      "formula-error",
      "named-range",
      "volatile-function",
    ]));
    expect(analysis.risks.filter((risk) => risk.kind === "embedded-number")).toHaveLength(1);
    expect(analysis.risks.find((risk) => risk.kind === "embedded-number")?.text).toBe("2");

    expect(analyzeFormula("=A1*2-1", { safeConstants: [0, 1, 2] }).risks)
      .not.toEqual(expect.arrayContaining([expect.objectContaining({
        kind: "embedded-number",
      })]));
  });

  it("recognizes Excel's external error literal", () => {
    expect(analyzeFormula("=#EXTERNAL!").risks).toContainEqual(expect.objectContaining({
      kind: "formula-error",
      text: "#EXTERNAL!",
    }));
  });

  it("treats negative one as a safe default constant", () => {
    const embedded = analyzeFormula("=A1*-1+A2*3").risks
      .filter((risk) => risk.kind === "embedded-number");

    expect(embedded).toMatchObject([{ text: "3" }]);
  });
});
