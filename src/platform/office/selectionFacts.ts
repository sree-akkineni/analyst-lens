import type { CellFacts } from "../../shared/types";

type CellValue = CellFacts["value"];

type SelectionWorksheetPort = {
  id: string;
  load(properties: string[]): void;
  name: string;
};

export type CellFactsRangePort = {
  address: string;
  columnCount: number;
  formulas: unknown[][];
  load(properties: string[]): void;
  numberFormat: unknown[][];
  rowCount: number;
  values: unknown[][];
  worksheet: SelectionWorksheetPort;
};

export type SelectionReadContext = {
  getSelection(): CellFactsRangePort;
  sync(): Promise<void>;
  workbookId: string;
};

const throwIfAborted = (signal: AbortSignal | undefined) => {
  if (signal?.aborted === true) {
    throw signal.reason ?? new DOMException("The Office request was aborted", "AbortError");
  }
};

export const normalizeCellValue = (value: unknown): CellValue => {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  return value === undefined ? null : String(value);
};

const first = (values: unknown[][]): unknown => values[0]?.[0];

export const cellFactsLoadProperties = [
  "address",
  "columnCount",
  "formulas",
  "numberFormat",
  "rowCount",
  "values",
] as const;

export const queueCellFactsLoad = (range: CellFactsRangePort): void => {
  range.load([...cellFactsLoadProperties]);
  range.worksheet.load(["id", "name"]);
};

export const factsFromLoadedRange = (
  range: CellFactsRangePort,
  workbookId: string,
): CellFacts => {
  const formulaValue = first(range.formulas);
  const numberFormatValue = first(range.numberFormat);
  return {
    address: canonicalA1Address(range.address),
    columnCount: range.columnCount,
    formula:
      typeof formulaValue === "string" && formulaValue.startsWith("=")
        ? formulaValue
        : null,
    numberFormat:
      typeof numberFormatValue === "string"
        ? numberFormatValue
        : String(numberFormatValue ?? "General"),
    rowCount: range.rowCount,
    sheetId: range.worksheet.id,
    sheetName: range.worksheet.name,
    value: normalizeCellValue(first(range.values)),
    workbookId,
  };
};

export const canonicalA1Address = (address: string): string => {
  const separator = address.lastIndexOf("!");
  const localAddress = separator === -1 ? address : address.slice(separator + 1);
  return localAddress.toUpperCase().replace(
    /\$?([A-Z]{1,3})\$?([0-9]+)/gu,
    (_match, column: string, row: string) => `$${column}$${row}`,
  );
};

export const readSelectionFacts = async (
  context: SelectionReadContext,
  signal?: AbortSignal,
): Promise<CellFacts> => {
  throwIfAborted(signal);
  const selection = context.getSelection();
  queueCellFactsLoad(selection);

  await context.sync();
  throwIfAborted(signal);

  return factsFromLoadedRange(selection, context.workbookId);
};
