import type { CellRef } from "../../shared/types";

export type FormulaTokenKind =
  | "boolean"
  | "error"
  | "external-name"
  | "external-reference"
  | "function"
  | "name"
  | "number"
  | "operator"
  | "punctuation"
  | "reference"
  | "string"
  | "unknown";

export type FormulaToken = Readonly<{
  end: number;
  kind: FormulaTokenKind;
  start: number;
  text: string;
}>;

export type FormulaRiskKind =
  | "cross-sheet"
  | "dynamic-reference"
  | "embedded-number"
  | "external-link"
  | "formula-error"
  | "named-range"
  | "volatile-function";

export type FormulaRisk = Readonly<{
  end: number;
  kind: FormulaRiskKind;
  message: string;
  start: number;
  text: string;
}>;

export type FormulaAnalysisOptions = Readonly<{
  origin?: Pick<CellRef, "sheetName">;
  safeConstants?: readonly number[];
}>;

const column = String.raw`\$?[A-Za-z]{1,3}`;
const row = String.raw`\$?\d+`;
const cell = String.raw`${column}${row}`;
const cellRange = String.raw`${cell}(?::${cell})?`;
const columnRange = String.raw`${column}:${column}`;
const rowRange = String.raw`${row}:${row}`;
const a1Reference = String.raw`(?:${cellRange}|${columnRange}|${rowRange})`;
const referenceBoundary = String.raw`(?![A-Za-z0-9_.\\])`;
const quotedQualifier = String.raw`'(?:[^']|'')+'`;
const unquotedQualifier = String.raw`[A-Za-z_\\][A-Za-z0-9_.\\]*`;
const externalQualifier = String.raw`(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'|\[[^\]]+\][^!]+)`;
const identifier = String.raw`[A-Za-z_\\][A-Za-z0-9_.\\]*`;

const externalReferencePattern = new RegExp(
  String.raw`^${externalQualifier}!${a1Reference}${referenceBoundary}`,
);
const externalNamePattern = new RegExp(
  String.raw`^${externalQualifier}!${identifier}`,
);
const qualifiedReferencePattern = new RegExp(
  String.raw`^(?:${quotedQualifier}|${unquotedQualifier})!${a1Reference}${referenceBoundary}`,
);
const localReferencePattern = new RegExp(
  String.raw`^${a1Reference}${referenceBoundary}`,
);
const qualifiedNamePattern = new RegExp(
  String.raw`^(?:${quotedQualifier}|${unquotedQualifier})!${identifier}`,
);
const identifierPattern = new RegExp(String.raw`^${identifier}`);
const numberPattern = /^(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][+-]?\d+)?/;
const errorPattern = /^#(?:GETTING_DATA|DIV\/0!|VALUE!|REF!|NAME\?|NULL!|NUM!|N\/A|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!|EXTERNAL!)/i;

const columnNumber = (letters: string): number => {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + letter.charCodeAt(0) - 64;
  }
  return value;
};

const validReferenceEndpoint = (endpoint: string): boolean => {
  const cellMatch = endpoint.match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/);
  if (cellMatch?.[1] !== undefined && cellMatch[2] !== undefined) {
    const rowNumber = Number(cellMatch[2]);
    return columnNumber(cellMatch[1]) <= 16_384 &&
      rowNumber >= 1 && rowNumber <= 1_048_576;
  }
  const columnMatch = endpoint.match(/^\$?([A-Za-z]{1,3})$/);
  if (columnMatch?.[1] !== undefined) {
    return columnNumber(columnMatch[1]) <= 16_384;
  }
  const rowMatch = endpoint.match(/^\$?(\d+)$/);
  if (rowMatch?.[1] !== undefined) {
    const rowNumber = Number(rowMatch[1]);
    return rowNumber >= 1 && rowNumber <= 1_048_576;
  }
  return false;
};

const validReference = (text: string): boolean => {
  const address = text.slice(text.lastIndexOf("!") + 1);
  return address.split(":").every(validReferenceEndpoint);
};

const token = (
  formula: string,
  kind: FormulaTokenKind,
  start: number,
  text: string,
): FormulaToken => ({
  end: start + text.length,
  kind,
  start,
  text: formula.slice(start, start + text.length),
});

const readString = (formula: string, start: number): string => {
  let end = start + 1;
  while (end < formula.length) {
    if (formula[end] !== "\"") {
      end += 1;
      continue;
    }
    if (formula[end + 1] === "\"") {
      end += 2;
      continue;
    }
    return formula.slice(start, end + 1);
  }
  return formula.slice(start);
};

const readStructuredName = (source: string): string | null => {
  const name = source.match(identifierPattern)?.[0];
  if (name === undefined || source[name.length] !== "[") return null;
  let depth = 0;
  for (let index = name.length; index < source.length; index += 1) {
    if (source[index] === "[") depth += 1;
    if (source[index] === "]") {
      depth -= 1;
      if (depth === 0) return source.slice(0, index + 1);
    }
  }
  return source.slice(0, name.length);
};

export const tokenizeFormula = (formula: string): FormulaToken[] => {
  const tokens: FormulaToken[] = [];
  let index = formula.startsWith("=") ? 1 : 0;

  while (index < formula.length) {
    const current = formula[index];
    if (current === undefined) break;
    if (/\s/.test(current)) {
      index += 1;
      continue;
    }
    if (current === "\"") {
      const text = readString(formula, index);
      tokens.push(token(formula, "string", index, text));
      index += text.length;
      continue;
    }

    const source = formula.slice(index);
    const error = source.match(errorPattern)?.[0];
    if (error !== undefined) {
      tokens.push(token(formula, "error", index, error));
      index += error.length;
      continue;
    }
    const externalReference = source.match(externalReferencePattern)?.[0];
    if (externalReference !== undefined && validReference(externalReference)) {
      tokens.push(token(formula, "external-reference", index, externalReference));
      index += externalReference.length;
      continue;
    }
    const externalName = source.match(externalNamePattern)?.[0];
    if (externalName !== undefined) {
      tokens.push(token(formula, "external-name", index, externalName));
      index += externalName.length;
      continue;
    }
    const leadingIdentifier = source.match(identifierPattern)?.[0];
    if (
      leadingIdentifier !== undefined &&
      source.slice(leadingIdentifier.length).match(/^\s*\(/) !== null
    ) {
      tokens.push(token(formula, "function", index, leadingIdentifier));
      index += leadingIdentifier.length;
      continue;
    }
    const qualifiedReference = source.match(qualifiedReferencePattern)?.[0];
    if (qualifiedReference !== undefined && validReference(qualifiedReference)) {
      tokens.push(token(formula, "reference", index, qualifiedReference));
      index += qualifiedReference.length;
      continue;
    }
    const localReference = source.match(localReferencePattern)?.[0];
    if (localReference !== undefined && validReference(localReference)) {
      tokens.push(token(formula, "reference", index, localReference));
      index += localReference.length;
      continue;
    }
    const qualifiedName = source.match(qualifiedNamePattern)?.[0];
    if (qualifiedName !== undefined) {
      tokens.push(token(formula, "name", index, qualifiedName));
      index += qualifiedName.length;
      continue;
    }
    const structuredName = readStructuredName(source);
    if (structuredName !== null) {
      tokens.push(token(formula, "name", index, structuredName));
      index += structuredName.length;
      continue;
    }
    const numeric = source.match(numberPattern)?.[0];
    if (numeric !== undefined) {
      tokens.push(token(formula, "number", index, numeric));
      index += numeric.length;
      continue;
    }
    const name = leadingIdentifier;
    if (name !== undefined) {
      const upper = name.toUpperCase();
      const kind: FormulaTokenKind = upper === "TRUE" || upper === "FALSE"
        ? "boolean"
        : "name";
      tokens.push(token(formula, kind, index, name));
      index += name.length;
      continue;
    }

    const twoCharacterOperator = source.match(/^(?:<=|>=|<>)/)?.[0];
    if (twoCharacterOperator !== undefined) {
      tokens.push(token(formula, "operator", index, twoCharacterOperator));
      index += twoCharacterOperator.length;
      continue;
    }
    if ("+-*/^&=<>%@#".includes(current)) {
      tokens.push(token(formula, "operator", index, current));
      index += 1;
      continue;
    }
    if ("(),;:{}".includes(current)) {
      tokens.push(token(formula, "punctuation", index, current));
      index += 1;
      continue;
    }
    tokens.push(token(formula, "unknown", index, current));
    index += 1;
  }

  return tokens;
};

const volatileFunctions = new Set([
  "CELL",
  "INDIRECT",
  "INFO",
  "NOW",
  "OFFSET",
  "RAND",
  "RANDBETWEEN",
  "TODAY",
]);
const dynamicFunctions = new Set(["INDIRECT", "OFFSET"]);

type LocalBinding = Readonly<{
  end: number;
  name: string;
  start: number;
}>;

const functionArguments = (
  tokens: readonly FormulaToken[],
  functionIndex: number,
): Readonly<{ end: number; groups: FormulaToken[][] }> | null => {
  const open = tokens[functionIndex + 1];
  if (open?.kind !== "punctuation" || open.text !== "(") return null;
  const groups: FormulaToken[][] = [];
  let current: FormulaToken[] = [];
  let arrayDepth = 0;
  let depth = 1;
  for (let index = functionIndex + 2; index < tokens.length; index += 1) {
    const next = tokens[index];
    if (next === undefined) break;
    if (next.kind === "punctuation" && next.text === "(") {
      depth += 1;
      current.push(next);
      continue;
    }
    if (next.kind === "punctuation" && next.text === ")") {
      depth -= 1;
      if (depth === 0) {
        groups.push(current);
        return { end: next.end, groups };
      }
      current.push(next);
      continue;
    }
    if (next.kind === "punctuation" && next.text === "{") {
      arrayDepth += 1;
      current.push(next);
      continue;
    }
    if (next.kind === "punctuation" && next.text === "}") {
      arrayDepth = Math.max(0, arrayDepth - 1);
      current.push(next);
      continue;
    }
    if (
      depth === 1 &&
      arrayDepth === 0 &&
      next.kind === "punctuation" &&
      (next.text === "," || next.text === ";")
    ) {
      groups.push(current);
      current = [];
      continue;
    }
    current.push(next);
  }
  return null;
};

const localBindings = (tokens: readonly FormulaToken[]): LocalBinding[] => {
  const bindings: LocalBinding[] = [];
  tokens.forEach((current, functionIndex) => {
    if (current.kind !== "function") return;
    const functionName = current.text.toUpperCase();
    if (functionName !== "LET" && functionName !== "LAMBDA") return;
    const parsed = functionArguments(tokens, functionIndex);
    if (parsed === null) return;
    const candidates = parsed.groups.slice(0, -1).filter((_group, index) =>
      functionName === "LAMBDA" || index % 2 === 0);
    for (const candidate of candidates) {
      const binding = candidate.length === 1 ? candidate[0] : undefined;
      if (binding?.kind === "name") {
        bindings.push({
          end: parsed.end,
          name: binding.text.toUpperCase(),
          start: current.start,
        });
      }
    }
  });
  return bindings;
};

const risk = (
  kind: FormulaRiskKind,
  source: FormulaToken,
  message: string,
): FormulaRisk => ({
  end: source.end,
  kind,
  message,
  start: source.start,
  text: source.text,
});

const qualifierSheetName = (text: string): string | null => {
  const separator = text.lastIndexOf("!");
  if (separator < 0) return null;
  const qualifier = text.slice(0, separator);
  if (qualifier.includes("[")) return null;
  return qualifier.startsWith("'") && qualifier.endsWith("'")
    ? qualifier.slice(1, -1).replaceAll("''", "'")
    : qualifier;
};

const isUnaryMinus = (
  tokens: readonly FormulaToken[],
  numberIndex: number,
): boolean => {
  const minus = tokens[numberIndex - 1];
  if (minus?.kind !== "operator" || minus.text !== "-") return false;
  const beforeMinus = tokens[numberIndex - 2];
  return beforeMinus === undefined ||
    beforeMinus.kind === "operator" ||
    (beforeMinus.kind === "punctuation" && "(,;{".includes(beforeMinus.text));
};

export const analyzeFormula = (
  formula: string,
  options: FormulaAnalysisOptions = {},
): Readonly<{ risks: FormulaRisk[]; tokens: FormulaToken[] }> => {
  const tokens = tokenizeFormula(formula);
  const risks: FormulaRisk[] = [];
  const safeConstants = new Set(options.safeConstants ?? [0, 1, -1]);
  const bindings = localBindings(tokens);

  tokens.forEach((current, index) => {
    if (
      current.kind === "external-name" ||
      current.kind === "external-reference"
    ) {
      risks.push(risk("external-link", current, "Formula references an external workbook"));
    }
    if (current.kind === "error") {
      risks.push(risk("formula-error", current, "Formula contains an Excel error literal"));
    }
    if (
      (current.kind === "name" || current.kind === "external-name") &&
      !bindings.some((binding) =>
        current.start >= binding.start &&
        current.end <= binding.end &&
        current.text.toUpperCase() === binding.name)
    ) {
      risks.push(risk("named-range", current, "Formula uses a defined or structured name"));
    }
    const sheetName = qualifierSheetName(current.text);
    if (
      sheetName !== null &&
      (options.origin === undefined ||
        sheetName.toLocaleLowerCase() !== options.origin.sheetName.toLocaleLowerCase())
    ) {
      risks.push(risk("cross-sheet", current, "Formula references another worksheet"));
    }
    if (current.kind === "function") {
      const functionName = current.text.toUpperCase();
      if (volatileFunctions.has(functionName)) {
        risks.push(risk("volatile-function", current, "Formula uses a volatile function"));
      }
      if (dynamicFunctions.has(functionName)) {
        risks.push(risk("dynamic-reference", current, "Formula resolves a reference dynamically"));
      }
    }
    if (current.kind === "number") {
      const numeric = Number(current.text) * (isUnaryMinus(tokens, index) ? -1 : 1);
      if (!safeConstants.has(numeric)) {
        risks.push(risk("embedded-number", current, "Formula contains a non-safe numeric literal"));
      }
    }
  });

  return { risks, tokens };
};
