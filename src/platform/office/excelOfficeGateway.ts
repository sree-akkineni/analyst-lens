import type {
  CellFacts,
  CellRef,
  CommandResult,
  DependencyDirection,
} from "../../shared/types";
import { getCommand } from "../../domain/commands/commandRegistry";
import {
  adjustNumberFormatDecimals,
  formatSelection,
  type FormatMutation,
  type SelectionFormatSnapshot,
  type SelectionFormatWritePort,
} from "../../domain/commands/formatCommands";
import {
  detectCapabilities,
  type RequirementSetSupport,
} from "./capabilityMap";
import {
  normalizeOfficeError,
  type OfficeGateway,
  OfficeGatewayError,
  type WorkbookFacts,
} from "./officeGateway";
import {
  canonicalA1Address,
  factsFromLoadedRange,
  normalizeCellValue,
  queueCellFactsLoad,
  readSelectionFacts,
  type CellFactsRangePort,
  type SelectionReadContext,
} from "./selectionFacts";
import { createWriteExecutor } from "./writeExecutor";
import type { HealthCellFacts } from "../../domain/health/rules";
import {
  parseWatchlistMetadata,
  type PersistedWatchItem,
} from "../../domain/watchlist/watchlist";

type DependencyRangePort = {
  address: string;
  worksheet: { id: string; name: string };
};

type DependencyRangesPort = {
  items: DependencyRangePort[];
  load(properties: string): void;
};

type DependencyResultPort = {
  ranges: DependencyRangesPort;
};

type OutlineGroupOption = "ByRows" | "ByColumns";

type StructuralRangePort = {
  group(option: OutlineGroupOption): void;
  ungroup(option: OutlineGroupOption): void;
};

type StructuralWorksheetPort = {
  id: string;
  name: string;
  protection: {
    load(properties: string): void;
    protected: boolean;
  };
  showOutlineLevels(rowLevels: number, columnLevels: number): void;
};

type TableCollectionPort = {
  items: unknown[];
  load(properties: string): void;
};

type SelectedRangeAreasPort = {
  areaCount: number;
  getTables(fullyContained?: boolean): TableCollectionPort;
  isEntireColumn: boolean;
  isEntireRow: boolean;
  load(properties: string | string[]): void;
  worksheet: StructuralWorksheetPort;
};

type ExcelRangePort = ReturnType<SelectionReadContext["getSelection"]> & {
  getDirectDependents(): DependencyResultPort;
  getDirectPrecedents(): DependencyResultPort;
  getEntireColumn(): StructuralRangePort;
  getEntireRow(): StructuralRangePort;
  getOffsetRange(rowOffset: number, columnOffset: number): ExcelRangePort;
  select(): void;
};

type FormattingBorderPort = {
  color: string;
  style: string;
  tintAndShade: number;
  weight: string;
};

type FormattingCellPort = {
  format: {
    borders: {
      getItem(index: string): FormattingBorderPort;
    };
  };
};

type FormattingRangePort = ExcelRangePort & FormattingCellPort & {
  clear(applyTo: "Formats"): void;
  getCell(rowIndex: number, columnIndex: number): FormattingCellPort;
  format: {
    borders: {
      getItem(index: string): FormattingBorderPort;
    };
    font: { color: string };
  };
  getCellProperties(
    options: Excel.CellPropertiesLoadOptions,
  ): { value: unknown[][] };
  numberFormat: string[][];
  setCellProperties(properties: Excel.SettableCellProperties[][]): void;
  worksheet: ExcelRangePort["worksheet"] & {
    protection: {
      load(properties: string): void;
      protected: boolean;
    };
  };
};

type EventRegistrationPort = { remove(): void };
type EventPort = {
  add(handler: () => Promise<void>): EventRegistrationPort;
};

type WorksheetFactsPort = {
  id: string;
  name: string;
  visibility: string;
};

type CellFactsWorksheetPort = {
  getRange(address: string): CellFactsRangePort;
  id: string;
  isNullObject: boolean;
  load(properties: string[]): void;
  name: string;
};

type HealthUsedRangePort = {
  address: string;
  columnCount: number;
  columnIndex: number;
  isNullObject: boolean;
  load(properties: string): void;
  rowCount: number;
  rowIndex: number;
};

type HealthDataRangePort = {
  formulas: unknown[][];
  getCellProperties(options: Excel.CellPropertiesLoadOptions): { value: unknown[][] };
  load(properties: string): void;
  numberFormat: unknown[][];
  values: unknown[][];
};

type HealthWorksheetPort = WorksheetFactsPort & {
  getRangeByIndexes(
    startRow: number,
    startColumn: number,
    rowCount: number,
    columnCount: number,
  ): HealthDataRangePort;
  getUsedRangeOrNullObject(valuesOnly?: boolean): HealthUsedRangePort;
};

type ExcelContextPort = {
  sync(): Promise<void>;
  workbook: {
    application: {
      calculationMode: string;
      load(properties: string): void;
    };
    getSelectedRange(): ExcelRangePort;
    getSelectedRanges(): SelectedRangeAreasPort;
    load(properties: string): void;
    name: string;
    worksheets: {
      getActiveWorksheet(): StructuralWorksheetPort;
      getItem(key: string): {
        getRange(address: string): ExcelRangePort;
      };
      getItemOrNullObject(key: string): CellFactsWorksheetPort;
      items: HealthWorksheetPort[];
      load(properties: string): void;
      onCalculated: EventPort;
      onSelectionChanged: EventPort;
    };
  };
};

type DocumentSettingsPort = {
  get(key: string): unknown;
  saveAsync(callback: (result: Readonly<{
    error?: Readonly<{ message?: string }>;
    status: string;
  }>) => void): void;
  set(key: string, value: unknown): void;
};

export type ExcelGatewayRuntime = {
  documentSettings?: DocumentSettingsPort;
  getDocumentUrl(): string;
  requirements: RequirementSetSupport;
  run<T>(batch: (context: unknown) => Promise<T>): Promise<T>;
  runWithContext?<T>(
    context: unknown,
    batch: (context: unknown) => Promise<T>,
  ): Promise<T>;
};

const productionRuntime = (): ExcelGatewayRuntime => ({
  documentSettings: Office.context.document.settings as unknown as DocumentSettingsPort,
  getDocumentUrl: () => Office.context.document.url,
  requirements: Office.context.requirements,
  run: <T>(batch: (context: unknown) => Promise<T>) =>
    Excel.run((context) => batch(context)),
  runWithContext: <T>(
    existingContext: unknown,
    batch: (context: unknown) => Promise<T>,
  ) =>
    Excel.run(
      existingContext as OfficeExtension.ClientRequestContext,
      (context) => batch(context),
    ),
});

const throwIfAborted = (signal: AbortSignal | undefined) => {
  if (signal?.aborted === true) {
    throw new OfficeGatewayError("ABORTED", "The Office request was cancelled");
  }
};

const normalizeCalculationMode = (
  value: string,
): WorkbookFacts["calculationMode"] => {
  switch (value.toLowerCase()) {
    case "automatic":
      return "automatic";
    case "automaticexcepttables":
      return "automaticExceptTables";
    case "manual":
      return "manual";
    default:
      return "unknown";
  }
};

const columnLetters = (zeroBasedColumn: number): string => {
  let value = zeroBasedColumn + 1;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
};

const healthScanCellLimit = 100_000;
const healthReadChunkSize = 5_000;
const watchlistReadChunkSize = 100;
const watchlistSettingsKey = "analystLens.watchlist.v1";

const failure = (code: string, message: string): CommandResult => ({
  code,
  message,
  ok: false,
});

const officeErrorDetails = (error: unknown): { code: string; message: string } => {
  const record = recordValue(error);
  return {
    code: typeof record?.code === "string" ? record.code : "",
    message: typeof record?.message === "string" ? record.message : "",
  };
};

const structuralFailure = (error: unknown): CommandResult => {
  const details = officeErrorDetails(error);
  if (
    details.code === "InvalidOperation" &&
    /outline|maximum|level/u.test(details.message.toLocaleLowerCase())
  ) {
    return failure(
      "OUTLINE_LIMIT_REACHED",
      "Excel could not add or remove the requested outline layer because the outline limit was reached",
    );
  }
  if (
    details.code === "InvalidArgument" ||
    details.code === "InvalidReference" ||
    details.code === "InvalidObjectPath"
  ) {
    return failure(
      "INVALID_STRUCTURE_SELECTION",
      "Select one valid worksheet range before changing an outline layer",
    );
  }
  const normalized = normalizeOfficeError(error);
  return failure(normalized.code, normalized.message);
};

const deepClone = <T>(value: T): T =>
  JSON.parse(JSON.stringify(value)) as T;

const captureScope = (
  mutation: FormatMutation,
): SelectionFormatSnapshot["captured"] => {
  if ("fontColor" in mutation) return "fontColor";
  if ("numberFormat" in mutation || "decimalDelta" in mutation) {
    return "numberFormat";
  }
  if ("border" in mutation || "clearBorders" in mutation) return "borders";
  return "allFormats";
};

const borderLoadOptions: Excel.CellPropertiesBorderLoadOptions = {
  color: true,
  style: true,
  tintAndShade: true,
  weight: true,
};

const fontLoadOptions: Excel.CellPropertiesFontLoadOptions = {
  bold: true,
  color: true,
  italic: true,
  name: true,
  size: true,
  strikethrough: true,
  subscript: true,
  superscript: true,
  tintAndShade: true,
  underline: true,
};

const allFormatLoadOptions: Excel.CellPropertiesLoadOptions = {
  format: {
    autoIndent: true,
    borders: borderLoadOptions,
    fill: {
      color: true,
      pattern: true,
      patternColor: true,
      patternTintAndShade: true,
      tintAndShade: true,
    },
    font: fontLoadOptions,
    horizontalAlignment: true,
    indentLevel: true,
    protection: true,
    readingOrder: true,
    shrinkToFit: true,
    textOrientation: true,
    useStandardHeight: true,
    useStandardWidth: true,
    verticalAlignment: true,
    wrapText: true,
  },
  style: true,
};

const cellLoadOptions = (
  scope: SelectionFormatSnapshot["captured"],
): Excel.CellPropertiesLoadOptions | null => {
  switch (scope) {
    case "fontColor":
      return { format: { font: { color: true } } };
    case "borders":
      return { format: { borders: borderLoadOptions } };
    case "allFormats":
      return allFormatLoadOptions;
    case "numberFormat":
      return null;
  }
};

const recordValue = (value: unknown): Record<string, unknown> | null =>
  typeof value === "object" && value !== null
    ? value as Record<string, unknown>
    : null;

const readPath = (value: unknown, path: readonly string[]): unknown => {
  let current: unknown = value;
  for (const key of path) {
    const record = recordValue(current);
    if (record === null) return undefined;
    current = record[key];
  }
  return current;
};

const everyCell = (
  cells: unknown[][],
  predicate: (cell: unknown, rowIndex: number, columnIndex: number) => boolean,
): boolean => cells.every((row, rowIndex) =>
  row.every((cell, columnIndex) => predicate(cell, rowIndex, columnIndex)));

const numberFormatsEqual = (left: string[][], right: string[][]): boolean =>
  JSON.stringify(left) === JSON.stringify(right);

const selectionTarget = (
  context: ExcelContextPort,
  target?: Pick<SelectionFormatSnapshot, "address" | "sheetId">,
): FormattingRangePort => target === undefined
  ? context.workbook.getSelectedRange() as FormattingRangePort
  : context.workbook.worksheets
      .getItem(target.sheetId)
      .getRange(target.address) as FormattingRangePort;

const readSelectionFormat = async (
  runtime: ExcelGatewayRuntime,
  scope: SelectionFormatSnapshot["captured"],
  target?: Pick<SelectionFormatSnapshot, "address" | "sheetId">,
): Promise<SelectionFormatSnapshot> => runtime.run(async (rawContext) => {
  const context = rawContext as ExcelContextPort;
  const range = selectionTarget(context, target);
  const includeNumberFormat = scope === "allFormats" || scope === "numberFormat";
  range.load(includeNumberFormat
    ? ["address", "numberFormat", "worksheet/id"]
    : ["address", "worksheet/id"]);
  range.worksheet.protection.load("protected");
  const options = cellLoadOptions(scope);
  const properties = options === null ? null : range.getCellProperties(options);
  await context.sync();
  return {
    address: canonicalA1Address(range.address),
    captured: scope,
    cellProperties: properties === null ? [] : deepClone(properties.value),
    numberFormat: includeNumberFormat ? deepClone(range.numberFormat) : [],
    sheetProtected: range.worksheet.protection.protected,
    sheetId: range.worksheet.id,
    workbookId: runtime.getDocumentUrl() || "active-workbook",
  };
});

const selectionMatchesSnapshot = async (
  runtime: ExcelGatewayRuntime,
  snapshot: SelectionFormatSnapshot,
): Promise<boolean> => {
  const current = await readSelectionFormat(runtime, snapshot.captured, snapshot);
  return current.address === snapshot.address &&
    current.sheetId === snapshot.sheetId &&
    current.workbookId === snapshot.workbookId &&
    JSON.stringify(current.cellProperties) === JSON.stringify(snapshot.cellProperties) &&
    numberFormatsEqual(current.numberFormat, snapshot.numberFormat);
};

const borderIndexes = [
  "EdgeTop",
  "EdgeBottom",
  "EdgeLeft",
  "EdgeRight",
  "InsideHorizontal",
  "InsideVertical",
  "DiagonalDown",
  "DiagonalUp",
] as const;

const borderProperties = [
  ["bottom", "EdgeBottom"],
  ["diagonalDown", "DiagonalDown"],
  ["diagonalUp", "DiagonalUp"],
  ["horizontal", "InsideHorizontal"],
  ["left", "EdgeLeft"],
  ["right", "EdgeRight"],
  ["top", "EdgeTop"],
  ["vertical", "InsideVertical"],
] as const;

const restoreCellBorders = (
  range: FormattingRangePort,
  cellProperties: unknown[][],
) => {
  for (const [rowIndex, row] of cellProperties.entries()) {
    for (const [columnIndex, cell] of row.entries()) {
      for (const [property, index] of borderProperties) {
        const captured = recordValue(readPath(cell, ["format", "borders", property]));
        if (captured === null) continue;
        const border = range.getCell(rowIndex, columnIndex).format.borders.getItem(index);
        const style = captured.style;
        if (typeof style === "string") border.style = style;
        if (style === "None") continue;
        if (typeof captured.color === "string") border.color = captured.color;
        if (typeof captured.weight === "string") border.weight = captured.weight;
        if (typeof captured.tintAndShade === "number") {
          border.tintAndShade = captured.tintAndShade;
        }
      }
    }
  }
};

const createSelectionFormatWritePort = (
  runtime: ExcelGatewayRuntime,
): SelectionFormatWritePort => {
  return {
    async applySelectionFormat(mutation, target) {
      await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const range = selectionTarget(context, target);
        if ("fontColor" in mutation) {
          range.format.font.color = mutation.fontColor;
        } else if ("numberFormat" in mutation) {
          range.load(["columnCount", "rowCount"]);
          await context.sync();
          range.numberFormat = Array.from({ length: range.rowCount }, () =>
            Array.from({ length: range.columnCount }, () => mutation.numberFormat));
        } else if ("decimalDelta" in mutation) {
          range.numberFormat = target.numberFormat.map((row) => row.map((value) =>
            adjustNumberFormatDecimals(
              typeof value === "string" ? value : String(value ?? "General"),
              mutation.decimalDelta,
            )));
        } else if ("border" in mutation) {
          const border = range.format.borders.getItem(
            mutation.border.side === "top" ? "EdgeTop" : "EdgeBottom",
          );
          border.color = "#000000";
          border.style = mutation.border.style;
          border.weight = mutation.border.weight;
        } else if ("clearBorders" in mutation) {
          for (const index of borderIndexes) {
            range.format.borders.getItem(index).style = "None";
          }
        } else {
          range.clear("Formats");
        }
        await context.sync();
      });
    },
    captureSelectionFormat(mutation) {
      return readSelectionFormat(runtime, captureScope(mutation));
    },
    async preflightSelection(_mutation, target) {
      return target.sheetProtected
        ? failure("PROTECTED", "Formatting is unavailable on a protected sheet")
        : { ok: true };
    },
    async restoreSelectionFormat(snapshot) {
      await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const range = selectionTarget(context, snapshot);
        if (snapshot.captured === "borders") {
          restoreCellBorders(range, snapshot.cellProperties);
        } else if (snapshot.cellProperties.length > 0) {
          range.setCellProperties(
            snapshot.cellProperties as Excel.SettableCellProperties[][],
          );
        }
        if (snapshot.numberFormat.length > 0) {
          range.numberFormat = deepClone(snapshot.numberFormat);
        }
        await context.sync();
      });
    },
    async verifySelectionFormat(mutation, target) {
      const scope = captureScope(mutation);
      const current = await readSelectionFormat(runtime, scope, target);
      if ("fontColor" in mutation) {
        return everyCell(current.cellProperties, (cell) =>
          readPath(cell, ["format", "font", "color"]) === mutation.fontColor);
      }
      if ("numberFormat" in mutation) {
        return everyCell(current.numberFormat, (value) =>
          value === mutation.numberFormat);
      }
      if ("decimalDelta" in mutation) {
        const expected = target.numberFormat.map((row) => row.map((value) =>
          adjustNumberFormatDecimals(
            typeof value === "string" ? value : String(value ?? "General"),
            mutation.decimalDelta,
          )));
        return numberFormatsEqual(current.numberFormat, expected);
      }
      if ("border" in mutation) {
        const rowIndex = mutation.border.side === "top"
          ? 0
          : current.cellProperties.length - 1;
        return current.cellProperties[rowIndex]?.every((cell) =>
          readPath(cell, ["format", "borders", mutation.border.side, "style"]) ===
            mutation.border.style &&
          readPath(cell, ["format", "borders", mutation.border.side, "weight"]) ===
            mutation.border.weight) === true;
      }
      if ("clearBorders" in mutation) {
        const sides = ["top", "bottom", "left", "right", "horizontal", "vertical"];
        return everyCell(current.cellProperties, (cell) => sides.every((side) => {
          const style = readPath(cell, ["format", "borders", side, "style"]);
          return style === undefined || style === "None";
        }));
      }
      return everyCell(current.cellProperties, (cell) =>
        readPath(cell, ["style"]) === "Normal") &&
        everyCell(current.numberFormat, (value) => value === "General");
    },
    verifySelectionRestore(snapshot) {
      return selectionMatchesSnapshot(runtime, snapshot);
    },
  };
};

export const createExcelOfficeGateway = (
  injectedRuntime?: ExcelGatewayRuntime,
): OfficeGateway => {
  const runtime = injectedRuntime ?? productionRuntime();
  const capabilities = detectCapabilities(runtime.requirements);
  let selectionRequestId = 0;
  const dependencyRequestIds: Record<DependencyDirection, number> = {
    dependent: 0,
    precedent: 0,
  };
  let workbookRequestId = 0;
  let neighborRequestId = 0;
  const formatExecutor = createWriteExecutor(
    createSelectionFormatWritePort(runtime),
  );

  const getSelectionFacts: OfficeGateway["getSelectionFacts"] = async (signal) => {
    const requestId = ++selectionRequestId;
    throwIfAborted(signal);
    try {
      const facts = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        return readSelectionFacts(
          {
            getSelection: () => context.workbook.getSelectedRange(),
            sync: () => context.sync(),
            workbookId: runtime.getDocumentUrl() || "active-workbook",
          },
          signal,
        );
      });
      throwIfAborted(signal);
      if (requestId !== selectionRequestId) {
        throw new OfficeGatewayError(
          "STALE_REQUEST",
          "A newer Office request superseded this result",
        );
      }
      return facts;
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const getCellFacts: NonNullable<OfficeGateway["getCellFacts"]> = async (
    ref,
    signal,
  ) => {
    throwIfAborted(signal);
    try {
      const facts = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const range = context.workbook.worksheets
          .getItem(ref.sheetId)
          .getRange(ref.address);
        return readSelectionFacts(
          {
            getSelection: () => range,
            sync: () => context.sync(),
            workbookId: runtime.getDocumentUrl() || ref.workbookId,
          },
          signal,
        );
      });
      throwIfAborted(signal);
      return facts;
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const getCellsFacts: NonNullable<OfficeGateway["getCellsFacts"]> = async (
    refs,
    signal,
  ) => {
    throwIfAborted(signal);
    if (refs.length === 0) return [];
    try {
      const results = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const worksheets = new Map<string, CellFactsWorksheetPort>();
        for (const ref of refs) {
          if (worksheets.has(ref.sheetId)) continue;
          const worksheet = context.workbook.worksheets.getItemOrNullObject(ref.sheetId);
          worksheet.load(["id", "name", "isNullObject"]);
          worksheets.set(ref.sheetId, worksheet);
        }
        await context.sync();
        throwIfAborted(signal);

        const loaded: (CellFacts | null)[] = refs.map(() => null);
        const available = refs.flatMap((ref, index) => {
          const worksheet = worksheets.get(ref.sheetId);
          if (worksheet === undefined || worksheet.isNullObject) return [];
          return [{ index, range: worksheet.getRange(ref.address), ref }];
        });
        for (let offset = 0; offset < available.length; offset += watchlistReadChunkSize) {
          throwIfAborted(signal);
          const chunk = available.slice(offset, offset + watchlistReadChunkSize);
          chunk.forEach(({ range }) => queueCellFactsLoad(range));
          await context.sync();
          throwIfAborted(signal);
          chunk.forEach(({ index, range, ref }) => {
            loaded[index] = factsFromLoadedRange(
              range,
              runtime.getDocumentUrl() || ref.workbookId,
            );
          });
        }
        return loaded;
      });
      throwIfAborted(signal);
      return results;
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const getDirectDependencies: OfficeGateway["getDirectDependencies"] = async (
    ref,
    direction,
    signal,
  ) => {
    const capability =
      direction === "precedent"
        ? capabilities.directPrecedents
        : capabilities.directDependents;
    if (!capability.supported) {
      throw new OfficeGatewayError(
        "UNSUPPORTED_OPERATION",
        `${capability.set} ${capability.version} is required`,
      );
    }
    const requestId = ++dependencyRequestIds[direction];
    throwIfAborted(signal);
    try {
      const dependencies = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const range = context.workbook.worksheets
          .getItem(ref.sheetId)
          .getRange(ref.address);
        const result =
          direction === "precedent"
            ? range.getDirectPrecedents()
            : range.getDirectDependents();
        result.ranges.load(
          "items/address,items/worksheet/id,items/worksheet/name",
        );
        await context.sync();
        throwIfAborted(signal);
        return result.ranges.items.map(
          (item): CellRef => ({
            address: canonicalA1Address(item.address),
            sheetId: item.worksheet.id,
            sheetName: item.worksheet.name,
            workbookId: ref.workbookId,
          }),
        );
      });
      if (requestId !== dependencyRequestIds[direction]) {
        throw new OfficeGatewayError(
          "STALE_REQUEST",
          "A newer Office request superseded this result",
        );
      }
      return dependencies;
    } catch (error) {
      const normalized = normalizeOfficeError(error);
      if (normalized.code === "NOT_FOUND") {
        return [];
      }
      throw normalized;
    }
  };

  const getNeighborFacts: OfficeGateway["getNeighborFacts"] = async (ref, signal) => {
    const requestId = ++neighborRequestId;
    throwIfAborted(signal);
    const address = canonicalA1Address(ref.address);
    const match = address.match(/^\$([A-Z]{1,3})\$(\d+)$/u);
    if (match?.[1] === undefined || match[2] === undefined) {
      throw new OfficeGatewayError("INVALID_REFERENCE", "Neighbor comparison requires one cell");
    }
    const column = [...match[1]].reduce(
      (value, letter) => value * 26 + letter.charCodeAt(0) - 64,
      0,
    );
    const row = Number(match[2]);
    const offsets = [
      { column: -1, direction: "left" as const, row: 0, valid: column > 1 },
      { column: 1, direction: "right" as const, row: 0, valid: column < 16_384 },
      { column: 0, direction: "above" as const, row: -1, valid: row > 1 },
      { column: 0, direction: "below" as const, row: 1, valid: row < 1_048_576 },
    ].filter((offset) => offset.valid);
    try {
      const neighbors = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const origin = context.workbook.worksheets
          .getItem(ref.sheetId)
          .getRange(ref.address);
        const ranges = offsets.map((offset) => ({
          direction: offset.direction,
          range: origin.getOffsetRange(offset.row, offset.column),
        }));
        ranges.forEach(({ range }) => queueCellFactsLoad(range as CellFactsRangePort));
        await context.sync();
        throwIfAborted(signal);
        return ranges.map(({ direction, range }) => ({
          direction,
          facts: factsFromLoadedRange(range, ref.workbookId),
        }));
      });
      if (requestId !== neighborRequestId) {
        throw new OfficeGatewayError(
          "STALE_REQUEST",
          "A newer Office request superseded this result",
        );
      }
      return neighbors;
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const selectCell: OfficeGateway["selectCell"] = async (ref) => {
    try {
      await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        context.workbook.worksheets
          .getItem(ref.sheetId)
          .getRange(ref.address)
          .select();
        await context.sync();
      });
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const getWorkbookFacts: OfficeGateway["getWorkbookFacts"] = async (signal) => {
    const requestId = ++workbookRequestId;
    throwIfAborted(signal);
    try {
      const facts = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        context.workbook.load("name");
        context.workbook.application.load("calculationMode");
        context.workbook.worksheets.load(
          "items/id,items/name,items/visibility",
        );
        await context.sync();
        throwIfAborted(signal);
        return {
          calculationMode: normalizeCalculationMode(
            context.workbook.application.calculationMode,
          ),
          id: runtime.getDocumentUrl() || context.workbook.name,
          name: context.workbook.name,
          worksheets: context.workbook.worksheets.items.map((worksheet) => ({
            id: worksheet.id,
            name: worksheet.name,
            visibility: worksheet.visibility,
          })),
        } satisfies WorkbookFacts;
      });
      if (requestId !== workbookRequestId) {
        throw new OfficeGatewayError(
          "STALE_REQUEST",
          "A newer Office request superseded this result",
        );
      }
      return facts;
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const getHealthWorkbookFacts: NonNullable<OfficeGateway["getHealthWorkbookFacts"]> = async (
    signal,
  ) => {
    throwIfAborted(signal);
    try {
      return await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        context.workbook.application.load("calculationMode");
        context.workbook.worksheets.load(
          "items/id,items/name,items/visibility",
        );
        await context.sync();
        throwIfAborted(signal);

        const ranges = context.workbook.worksheets.items.map((worksheet) => ({
          range: worksheet.getUsedRangeOrNullObject(true),
          worksheet,
        }));
        ranges.forEach(({ range }) => range.load(
          "address,columnCount,columnIndex,isNullObject,rowCount,rowIndex",
        ));
        await context.sync();
        throwIfAborted(signal);

        const loaded = ranges.filter(({ range }) => !range.isNullObject);
        const usedCellCount = loaded.reduce(
          (total, { range }) => total + range.rowCount * range.columnCount,
          0,
        );
        if (usedCellCount > healthScanCellLimit) {
          throw new OfficeGatewayError(
            "UNSUPPORTED_OPERATION",
            "Model health scans support up to 100,000 used cells. Narrow the workbook before scanning.",
          );
        }

        const workbookId = runtime.getDocumentUrl() || "active-workbook";
        const firstWorksheet = context.workbook.worksheets.items[0];
        const cells: HealthCellFacts[] = [];
        for (const { range: usedRange, worksheet } of loaded) {
          const columnStep = Math.min(usedRange.columnCount, healthReadChunkSize);
          const rowStep = Math.max(1, Math.floor(healthReadChunkSize / columnStep));
          for (let rowOffset = 0; rowOffset < usedRange.rowCount; rowOffset += rowStep) {
            const rowCount = Math.min(rowStep, usedRange.rowCount - rowOffset);
            for (
              let columnOffset = 0;
              columnOffset < usedRange.columnCount;
              columnOffset += healthReadChunkSize
            ) {
              throwIfAborted(signal);
              const columnCount = Math.min(
                healthReadChunkSize,
                usedRange.columnCount - columnOffset,
              );
              const dataRange = worksheet.getRangeByIndexes(
                usedRange.rowIndex + rowOffset,
                usedRange.columnIndex + columnOffset,
                rowCount,
                columnCount,
              );
              dataRange.load("formulas,numberFormat,values");
              const properties = dataRange.getCellProperties({
                format: { font: { color: true } },
              });
              await context.sync();
              throwIfAborted(signal);

              for (let chunkRow = 0; chunkRow < rowCount; chunkRow += 1) {
                for (let chunkColumn = 0; chunkColumn < columnCount; chunkColumn += 1) {
                  const formulaValue = dataRange.formulas[chunkRow]?.[chunkColumn];
                  const formula = typeof formulaValue === "string" && formulaValue.startsWith("=")
                    ? formulaValue
                    : null;
                  const value = normalizeCellValue(dataRange.values[chunkRow]?.[chunkColumn]);
                  if (formula === null && (value === null || value === "")) continue;
                  const numberFormatValue = dataRange.numberFormat[chunkRow]?.[chunkColumn];
                  const fontColorValue = readPath(
                    properties.value[chunkRow]?.[chunkColumn],
                    ["format", "font", "color"],
                  );
                  cells.push({
                    address: `$${columnLetters(usedRange.columnIndex + columnOffset + chunkColumn)}$${usedRange.rowIndex + rowOffset + chunkRow + 1}`,
                    columnCount: 1,
                    fontColor: typeof fontColorValue === "string" ? fontColorValue : "unknown",
                    formula,
                    numberFormat: typeof numberFormatValue === "string"
                      ? numberFormatValue
                      : String(numberFormatValue ?? "General"),
                    rowCount: 1,
                    sheetId: worksheet.id,
                    sheetName: worksheet.name,
                    value,
                    workbookId,
                  });
                }
              }
            }
          }
        }
        return {
          calculationMode: normalizeCalculationMode(
            context.workbook.application.calculationMode,
          ),
          cells,
          workbookRef: {
            address: "$A$1",
            sheetId: firstWorksheet?.id ?? "workbook",
            sheetName: firstWorksheet?.name ?? "Workbook",
            workbookId,
          },
        };
      });
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const documentSettings = (): DocumentSettingsPort => {
    if (runtime.documentSettings === undefined) {
      throw new OfficeGatewayError(
        "UNSUPPORTED_OPERATION",
        "This Excel host does not expose workbook document settings",
      );
    }
    return runtime.documentSettings;
  };

  const loadWatchlistMetadata = async (): Promise<readonly PersistedWatchItem[]> =>
    parseWatchlistMetadata(documentSettings().get(watchlistSettingsKey));

  const saveWatchlistMetadata = async (
    metadata: readonly PersistedWatchItem[],
  ): Promise<void> => {
    const settings = documentSettings();
    settings.set(watchlistSettingsKey, parseWatchlistMetadata(metadata));
    await new Promise<void>((resolve, reject) => {
      settings.saveAsync((result) => {
        if (result.status.toLocaleLowerCase() === "succeeded") {
          resolve();
          return;
        }
        reject(new OfficeGatewayError(
          "OFFICE_ERROR",
          result.error?.message ?? "Excel could not save watchlist metadata",
        ));
      });
    });
  };

  const subscribe = async (
    capability: "selectionEvents" | "recalculationEvents",
    event: "onSelectionChanged" | "onCalculated",
    listener: () => void,
  ) => {
    const requirement = capabilities[capability];
    if (!requirement.supported) {
      throw new OfficeGatewayError(
        "UNSUPPORTED_OPERATION",
        `${requirement.set} ${requirement.version} is required`,
      );
    }
    try {
      const subscription = await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const registration = context.workbook.worksheets[event].add(async () => {
          listener();
        });
        await context.sync();
        return { context, registration };
      });
      return () => {
        const remove = async () => {
          const runWithContext =
            runtime.runWithContext ??
            (async <T>(
              context: unknown,
              batch: (context: unknown) => Promise<T>,
            ) => batch(context));
          await runWithContext(subscription.context, async (rawContext) => {
            subscription.registration.remove();
            await (rawContext as ExcelContextPort).sync();
          });
        };
        void remove().catch(() => undefined);
      };
    } catch (error) {
      throw normalizeOfficeError(error);
    }
  };

  const outlineCapabilityFailure = (): CommandResult | null => {
    const requirement = capabilities.outlineGroups;
    return requirement.supported
      ? null
      : failure(
          "UNSUPPORTED_OPERATION",
          `${requirement.set} ${requirement.version} is required`,
        );
  };

  const runOutlineVisibility = async (
    id: "structure.showDetails" | "structure.hideDetails",
  ): Promise<CommandResult> => {
    const unsupported = outlineCapabilityFailure();
    if (unsupported !== null) return unsupported;
    try {
      await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const levels: readonly [number, number] = id === "structure.showDetails"
          ? [8, 8]
          : [1, 1];
        worksheet.showOutlineLevels(levels[0], levels[1]);
        await context.sync();
      });
      return { ok: true };
    } catch (error) {
      return structuralFailure(error);
    }
  };

  const structuralConfig = (id: string): Readonly<{
    dimension: "columns" | "rows";
    inverseId: string;
    inverseLabel: string;
    operation: "group" | "ungroup";
    option: OutlineGroupOption;
  }> | null => {
    switch (id) {
      case "structure.groupRows":
        return {
          dimension: "rows",
          inverseId: "structure.ungroupRows",
          inverseLabel: "Ungroup rows",
          operation: "group",
          option: "ByRows",
        };
      case "structure.ungroupRows":
        return {
          dimension: "rows",
          inverseId: "structure.groupRows",
          inverseLabel: "Group rows",
          operation: "ungroup",
          option: "ByRows",
        };
      case "structure.groupColumns":
        return {
          dimension: "columns",
          inverseId: "structure.ungroupColumns",
          inverseLabel: "Ungroup columns",
          operation: "group",
          option: "ByColumns",
        };
      case "structure.ungroupColumns":
        return {
          dimension: "columns",
          inverseId: "structure.groupColumns",
          inverseLabel: "Group columns",
          operation: "ungroup",
          option: "ByColumns",
        };
      default:
        return null;
    }
  };

  const runStructuralMutation = async (
    id: string,
  ): Promise<CommandResult> => {
    const unsupported = outlineCapabilityFailure();
    if (unsupported !== null) return unsupported;
    if (formatExecutor.isSafetyLocked()) {
      return failure(
        "SAFETY_LOCKED",
        "Writes are paused until the previous verification failure is acknowledged",
      );
    }
    const config = structuralConfig(id);
    if (config === null) {
      return failure("UNKNOWN_COMMAND", `Unknown structural command: ${id}`);
    }

    try {
      await runtime.run(async (rawContext) => {
        const context = rawContext as ExcelContextPort;
        const selectedAreas = context.workbook.getSelectedRanges();
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const intersectingTables = selectedAreas.getTables(false);
        const containedTables = selectedAreas.getTables(true);
        selectedAreas.load(["areaCount", "isEntireColumn", "isEntireRow"]);
        worksheet.protection.load("protected");
        intersectingTables.load("items/id");
        containedTables.load("items/id");
        await context.sync();

        if (selectedAreas.areaCount !== 1) {
          throw new OfficeGatewayError(
            "INVALID_REFERENCE",
            "MULTI_AREA_SELECTION",
          );
        }
        if (
          (config.dimension === "rows" && selectedAreas.isEntireColumn) ||
          (config.dimension === "columns" && selectedAreas.isEntireRow)
        ) {
          throw new OfficeGatewayError(
            "INVALID_REFERENCE",
            "MIXED_STRUCTURE_SELECTION",
          );
        }
        if (worksheet.protection.protected) {
          throw new OfficeGatewayError("ACCESS_DENIED", "PROTECTED");
        }
        if (intersectingTables.items.length > containedTables.items.length) {
          throw new OfficeGatewayError(
            "INVALID_REFERENCE",
            "PARTIAL_TABLE_SELECTION",
          );
        }

        const selection = context.workbook.getSelectedRange();
        const target = config.dimension === "rows"
          ? selection.getEntireRow()
          : selection.getEntireColumn();
        target[config.operation](config.option);
        await context.sync();
      });
    } catch (error) {
      if (error instanceof OfficeGatewayError) {
        switch (error.message) {
          case "MULTI_AREA_SELECTION":
            return failure(
              "MULTI_AREA_SELECTION",
              "Grouping requires one contiguous selection",
            );
          case "MIXED_STRUCTURE_SELECTION":
            return failure(
              "MIXED_STRUCTURE_SELECTION",
              `The selected ${config.dimension === "rows" ? "columns" : "rows"} cannot be changed with this command`,
            );
          case "PROTECTED":
            return failure(
              "PROTECTED",
              "Structure changes are unavailable on a protected sheet",
            );
          case "PARTIAL_TABLE_SELECTION":
            return failure(
              "PARTIAL_TABLE_SELECTION",
              "Select the complete intersecting Excel table before changing outline structure",
            );
        }
      }
      return structuralFailure(error);
    }

    formatExecutor.blockRevert({
      commandId: id,
      inverseCommandId: config.inverseId,
      message: `Automatic Revert is unavailable. Select the affected ${config.dimension} and run ${config.inverseLabel} once.`,
    });
    return { ok: true };
  };

  const runWrite: OfficeGateway["runWrite"] = async (request) => {
    const command = getCommand(request.id);
    if (command === undefined) {
      return failure("UNKNOWN_COMMAND", `Unknown command: ${request.id}`);
    }
    if (request.id === "system.revertLastAction") {
      return formatExecutor.revertLast();
    }
    if (request.id === "system.acknowledgeSafetyLock") {
      formatExecutor.acknowledgeSafetyLock();
      return { ok: true };
    }
    if (request.id === "structure.showDetails" || request.id === "structure.hideDetails") {
      return runOutlineVisibility(request.id);
    }
    if (command.effect === "structural-write") {
      return runStructuralMutation(request.id);
    }
    if (command.category !== "format") {
      return failure(
        "WRITE_ADAPTER_NOT_READY",
        `Write ${request.id} is not available yet`,
      );
    }
    if (!capabilities.formatCellProperties.supported) {
      return failure(
        "UNSUPPORTED_OPERATION",
        `${capabilities.formatCellProperties.set} ${capabilities.formatCellProperties.version} is required`,
      );
    }
    return formatExecutor.execute(formatSelection(request.id));
  };

  return {
    getCapabilities: () => capabilities,
    getCellFacts,
    getCellsFacts,
    getDirectDependencies,
    getHealthWorkbookFacts,
    getNeighborFacts,
    getSelectionFacts,
    getWorkbookFacts,
    ...(runtime.documentSettings === undefined
      ? {}
      : { loadWatchlistMetadata, saveWatchlistMetadata }),
    runWrite,
    selectCell,
    subscribeRecalculation: (listener) =>
      subscribe("recalculationEvents", "onCalculated", listener),
    subscribeSelection: (listener) =>
      subscribe("selectionEvents", "onSelectionChanged", listener),
  };
};
