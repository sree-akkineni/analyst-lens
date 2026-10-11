import type { CellRef } from "../../shared/types";
import { tokenizeFormula } from "./tokenizeFormula";

type GridPosition = Readonly<{ column: number; row: number }>;

const columnNumber = (letters: string): number => {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + letter.charCodeAt(0) - 64;
  }
  return value;
};

const parseOrigin = (origin: CellRef): GridPosition => {
  const unqualified = origin.address.slice(origin.address.lastIndexOf("!") + 1);
  const match = unqualified.match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/);
  if (match === null || match[1] === undefined || match[2] === undefined) {
    throw new Error("Formula normalization requires an origin with a single A1 cell address");
  }
  return { column: columnNumber(match[1]), row: Number(match[2]) };
};

const normalizeCell = (address: string, origin: GridPosition): string => {
  const match = address.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
  if (match === null) return address;
  const [, columnAbsolute, columnLetters, rowAbsolute, rowDigits] = match;
  if (columnLetters === undefined || rowDigits === undefined) return address;
  const column = columnNumber(columnLetters);
  const row = Number(rowDigits);
  const rowPart = rowAbsolute === "$" ? `R${row}` : `R[${row - origin.row}]`;
  const columnPart = columnAbsolute === "$"
    ? `C${column}`
    : `C[${column - origin.column}]`;
  return `${rowPart}${columnPart}`;
};

const normalizeEndpoint = (address: string, origin: GridPosition): string => {
  const column = address.match(/^(\$?)([A-Za-z]{1,3})$/);
  if (column?.[2] !== undefined) {
    const number = columnNumber(column[2]);
    return column[1] === "$" ? `C${number}` : `C[${number - origin.column}]`;
  }
  const row = address.match(/^(\$?)(\d+)$/);
  if (row?.[2] !== undefined) {
    const number = Number(row[2]);
    return row[1] === "$" ? `R${number}` : `R[${number - origin.row}]`;
  }
  return normalizeCell(address, origin);
};

const normalizeReference = (text: string, origin: GridPosition): string => {
  const separator = text.lastIndexOf("!");
  const qualifier = separator < 0 ? "" : text.slice(0, separator + 1);
  const address = separator < 0 ? text : text.slice(separator + 1);
  return qualifier + address
    .split(":")
    .map((endpoint) => normalizeEndpoint(endpoint, origin))
    .join(":");
};

export const normalizeFormula = (formula: string, origin: CellRef): string => {
  const position = parseOrigin(origin);
  const references = tokenizeFormula(formula).filter((current) =>
    current.kind === "reference" || current.kind === "external-reference");
  let cursor = 0;
  let normalized = "";
  for (const reference of references) {
    normalized += formula.slice(cursor, reference.start);
    normalized += normalizeReference(reference.text, position);
    cursor = reference.end;
  }
  return normalized + formula.slice(cursor);
};
