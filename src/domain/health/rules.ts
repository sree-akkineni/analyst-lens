import { normalizeFormula } from "../formulas/normalizeFormula";
import { analyzeFormula, type FormulaRiskKind } from "../formulas/tokenizeFormula";
import type { CellFacts, CellRef, HealthFinding } from "../../shared/types";

export type HealthCellFacts = CellFacts & Readonly<{ fontColor: string }>;

export type HealthWorkbookFacts = Readonly<{
  calculationMode: "automatic" | "automaticExceptTables" | "manual" | "unknown";
  cells: readonly HealthCellFacts[];
  workbookRef: CellRef;
}>;

export type HealthRule = Readonly<{
  applies(facts: HealthWorkbookFacts): boolean;
  defaultSeverity: HealthFinding["severity"];
  evaluate(facts: HealthWorkbookFacts): HealthFinding[];
  id: string;
}>;

type Position = Readonly<{
  cell: HealthCellFacts;
  column: number;
  row: number;
}>;

const columnNumber = (letters: string): number => {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + letter.charCodeAt(0) - 64;
  }
  return value;
};

const position = (cell: HealthCellFacts): Position | null => {
  const match = cell.address.match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/u);
  if (match?.[1] === undefined || match[2] === undefined) return null;
  return { cell, column: columnNumber(match[1]), row: Number(match[2]) };
};

const hasContent = (cell: HealthCellFacts): boolean =>
  cell.formula !== null || (cell.value !== null && cell.value !== "");

const excelErrorValues = new Set([
  "#BLOCKED!",
  "#BUSY!",
  "#CALC!",
  "#CONNECT!",
  "#DIV/0!",
  "#EXTERNAL!",
  "#FIELD!",
  "#GETTING_DATA",
  "#N/A",
  "#NAME?",
  "#NULL!",
  "#NUM!",
  "#PYTHON!",
  "#REF!",
  "#SPILL!",
  "#UNKNOWN!",
  "#VALUE!",
]);

const finding = (
  ruleId: string,
  severity: HealthFinding["severity"],
  ref: CellRef,
  evidence: string,
  remediation: string,
): HealthFinding => ({ evidence, ref, remediation, ruleId, severity });

const formulaRiskFindings = (
  facts: HealthWorkbookFacts,
  kind: FormulaRiskKind,
  ruleId: string,
  severity: HealthFinding["severity"],
  remediation: string,
): HealthFinding[] => facts.cells.flatMap((cell) => {
  if (cell.formula === null) return [];
  const risks = analyzeFormula(cell.formula, { origin: cell }).risks.filter(
    (risk) => risk.kind === kind,
  );
  if (risks.length === 0) return [];
  const observed = [...new Set(risks.map((risk) => risk.text))].join(", ");
  return [finding(
    ruleId,
    severity,
    cell,
    `${cell.formula} contains ${observed}`,
    remediation,
  )];
});

const segments = (
  facts: HealthWorkbookFacts,
  orientation: "column" | "row",
): Position[][] => {
  const groups = new Map<string, Position[]>();
  for (const cell of facts.cells) {
    if (!hasContent(cell)) continue;
    const current = position(cell);
    if (current === null) continue;
    const key = orientation === "row"
      ? `${cell.sheetId}::row::${current.row}`
      : `${cell.sheetId}::column::${current.column}`;
    const group = groups.get(key) ?? [];
    group.push(current);
    groups.set(key, group);
  }
  const result: Position[][] = [];
  for (const group of groups.values()) {
    group.sort((left, right) => orientation === "row"
      ? left.column - right.column
      : left.row - right.row);
    let segment: Position[] = [];
    for (const current of group) {
      const previous = segment.at(-1);
      const adjacent = previous === undefined || (orientation === "row"
        ? current.column === previous.column + 1
        : current.row === previous.row + 1);
      if (!adjacent) {
        if (segment.length > 0) result.push(segment);
        segment = [];
      }
      segment.push(current);
    }
    if (segment.length > 0) result.push(segment);
  }
  return result;
};

const allSegments = (facts: HealthWorkbookFacts): Position[][] => [
  ...segments(facts, "row"),
  ...segments(facts, "column"),
];

const dominantValue = (values: readonly string[]): string | null => {
  const counts = new Map<string, number>();
  values.forEach((value) => counts.set(value, (counts.get(value) ?? 0) + 1));
  const ranked = [...counts].sort((left, right) => right[1] - left[1]);
  const [first, second] = ranked;
  if (first === undefined || first[1] < 2 || first[1] === second?.[1]) return null;
  return first[0];
};

const deduplicate = (findings: readonly HealthFinding[]): HealthFinding[] => {
  const unique = new Map<string, HealthFinding>();
  for (const current of findings) {
    const key = `${current.ruleId}::${current.ref.workbookId}::${current.ref.sheetId}::${current.ref.address}`;
    if (!unique.has(key)) unique.set(key, current);
  }
  return [...unique.values()];
};

const formulaDiscontinuities = (facts: HealthWorkbookFacts): HealthFinding[] => {
  const findings: HealthFinding[] = [];
  for (const segment of allSegments(facts)) {
    const formulas = segment.filter(({ cell }) => cell.formula !== null).map(({ cell }) => ({
      cell,
      normalized: normalizeFormula(cell.formula ?? "", cell),
    }));
    if (formulas.length < 3) continue;
    const dominant = dominantValue(formulas.map(({ normalized }) => normalized));
    if (dominant === null) continue;
    const counts = new Map<string, number>();
    formulas.forEach(({ normalized }) => counts.set(normalized, (counts.get(normalized) ?? 0) + 1));
    for (const current of formulas) {
      if (current.normalized === dominant || counts.get(current.normalized) !== 1) continue;
      findings.push(finding(
        "FORMULA_DISCONTINUITY",
        "warning",
        current.cell,
        `Normalized formula ${current.normalized} differs from repeated pattern ${dominant}`,
        "Confirm the exception or restore the repeated formula pattern.",
      ));
    }
  }
  return deduplicate(findings);
};

const hardcodesInRuns = (facts: HealthWorkbookFacts): HealthFinding[] => {
  const findings: HealthFinding[] = [];
  for (const segment of allSegments(facts)) {
    for (let index = 1; index < segment.length - 1; index += 1) {
      const previous = segment[index - 1]?.cell;
      const current = segment[index]?.cell;
      const next = segment[index + 1]?.cell;
      if (
        current !== undefined && current.formula === null && current.value !== null &&
        previous !== undefined && previous.formula !== null &&
        next !== undefined && next.formula !== null
      ) {
        const previousPattern = normalizeFormula(previous.formula, previous);
        const nextPattern = normalizeFormula(next.formula, next);
        if (previousPattern !== nextPattern) continue;
        findings.push(finding(
          "HARDCODE_IN_RUN",
          "warning",
          current,
          `Hardcoded value ${String(current.value)} interrupts adjacent formulas`,
          "Confirm the exception or replace the hardcode with the expected formula.",
        ));
      }
    }
  }
  return deduplicate(findings);
};

const inconsistentFormats = (facts: HealthWorkbookFacts): HealthFinding[] => {
  const findings: HealthFinding[] = [];
  for (const segment of segments(facts, "row")) {
    if (
      segment.length < 3 ||
      segment.filter(({ cell }) => cell.formula !== null).length < 2
    ) continue;
    const formats = segment.map(({ cell }) => `${cell.numberFormat}::${cell.fontColor}`);
    const dominant = dominantValue(formats);
    if (dominant === null) continue;
    const counts = new Map<string, number>();
    formats.forEach((format) => counts.set(format, (counts.get(format) ?? 0) + 1));
    segment.forEach(({ cell }, index) => {
      const observed = formats[index];
      if (observed === undefined || observed === dominant || counts.get(observed) !== 1) return;
      findings.push(finding(
        "INCONSISTENT_FORMAT",
        "warning",
        cell,
        `Format ${observed} differs from repeated row format ${dominant}`,
        "Confirm the exception or apply the row's repeated number and font format.",
      ));
    });
  }
  return deduplicate(findings);
};

export const healthRules: readonly HealthRule[] = [
  {
    applies: () => true,
    defaultSeverity: "error",
    evaluate: (facts) => facts.cells.flatMap((cell) => {
      const formulaErrors = cell.formula === null
        ? []
        : analyzeFormula(cell.formula, { origin: cell }).risks.filter((risk) =>
            risk.kind === "formula-error");
      const hasErrorValue = typeof cell.value === "string" &&
        excelErrorValues.has(cell.value.trim().toUpperCase());
      if (!hasErrorValue && formulaErrors.length === 0) {
        return [];
      }
      return [finding(
        "ERR_VALUE",
        "error",
        cell,
        `Excel error observed: ${String(hasErrorValue ? cell.value : formulaErrors[0]?.text)}`,
        "Inspect the broken reference or upstream error and restore a valid formula result.",
      )];
    }),
    id: "ERR_VALUE",
  },
  {
    applies: (facts) => facts.cells.length >= 3,
    defaultSeverity: "warning",
    evaluate: formulaDiscontinuities,
    id: "FORMULA_DISCONTINUITY",
  },
  {
    applies: (facts) => facts.cells.length >= 3,
    defaultSeverity: "warning",
    evaluate: hardcodesInRuns,
    id: "HARDCODE_IN_RUN",
  },
  {
    applies: () => true,
    defaultSeverity: "warning",
    evaluate: (facts) => formulaRiskFindings(
      facts,
      "embedded-number",
      "EMBEDDED_CONSTANT",
      "warning",
      "Move the constant to an explicit assumption cell or add it to configured safe constants.",
    ),
    id: "EMBEDDED_CONSTANT",
  },
  {
    applies: () => true,
    defaultSeverity: "warning",
    evaluate: (facts) => formulaRiskFindings(
      facts,
      "external-link",
      "EXTERNAL_LINK",
      "warning",
      "Confirm the external workbook dependency is intentional and available.",
    ),
    id: "EXTERNAL_LINK",
  },
  {
    applies: () => true,
    defaultSeverity: "warning",
    evaluate: (facts) => formulaRiskFindings(
      facts,
      "volatile-function",
      "VOLATILE_FUNCTION",
      "warning",
      "Confirm volatility is required or replace it with a stable reference.",
    ),
    id: "VOLATILE_FUNCTION",
  },
  {
    applies: (facts) => facts.calculationMode === "manual",
    defaultSeverity: "warning",
    evaluate: (facts) => facts.calculationMode === "manual"
      ? [finding(
          "MANUAL_CALCULATION",
          "warning",
          facts.workbookRef,
          "Workbook calculation mode is manual",
          "Confirm manual mode is intentional before relying on displayed results.",
        )]
      : [],
    id: "MANUAL_CALCULATION",
  },
  {
    applies: (facts) => facts.cells.length >= 3,
    defaultSeverity: "warning",
    evaluate: inconsistentFormats,
    id: "INCONSISTENT_FORMAT",
  },
];

export const runRules = (
  facts: HealthWorkbookFacts,
  options: Readonly<{ ruleIds?: readonly string[] }> = {},
): HealthFinding[] => {
  const selected = options.ruleIds === undefined
    ? healthRules
    : healthRules.filter((rule) => options.ruleIds?.includes(rule.id) === true);
  return deduplicate(selected.flatMap((rule) =>
    rule.applies(facts) ? rule.evaluate(facts) : []))
    .sort((left, right) =>
      left.ref.sheetName.localeCompare(right.ref.sheetName) ||
      left.ref.address.localeCompare(right.ref.address, undefined, { numeric: true }) ||
      left.ruleId.localeCompare(right.ruleId));
};
