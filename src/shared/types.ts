export type CellRef = {
  workbookId: string;
  sheetId: string;
  sheetName: string;
  address: string;
};

export const canonicalCellId = (ref: CellRef): string =>
  `${ref.workbookId}::${ref.sheetId}::${ref.address.toUpperCase()}`;

export type CellFacts = CellRef & {
  value: string | number | boolean | null;
  formula: string | null;
  numberFormat: string;
  rowCount: number;
  columnCount: number;
};

export type DependencyDirection = "precedent" | "dependent";
export type DependencyRef = CellRef & {
  unresolved?: "dynamic" | "external";
};
export type DependencyNode = {
  id: string;
  ref: CellRef;
  value?: CellFacts["value"];
  unresolved?: "dynamic" | "external";
};
export type DependencyEdge = { id: string; source: string; target: string; cycle: boolean };
export type HealthFinding = {
  ruleId: string;
  severity: "error" | "warning" | "info";
  ref: CellRef;
  evidence: string;
  remediation: string;
};
export type CommandId = string;
export type CommandResult = { ok: true } | { ok: false; code: string; message: string };
