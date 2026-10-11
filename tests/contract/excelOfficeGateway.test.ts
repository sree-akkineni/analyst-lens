import { describe, expect, it } from "vitest";
import { createExcelOfficeGateway } from "../../src/platform/office/excelOfficeGateway";
import { createFakeOfficeContext } from "./fakeOfficeContext";

const supportedRequirements = {
  isSetSupported: () => true,
};

const createRuntime = (contexts: unknown[], documentSettings?: {
  get(key: string): unknown;
  saveAsync(callback: (result: { error?: { message?: string }; status: string }) => void): void;
  set(key: string, value: unknown): void;
}) => ({
  ...(documentSettings === undefined ? {} : { documentSettings }),
  getDocumentUrl: () => "file:///tmp/model.xlsx",
  requirements: supportedRequirements,
  async run<T>(batch: (context: unknown) => Promise<T>): Promise<T> {
    const context = contexts.shift();
    if (context === undefined) {
      throw new Error("No fake Office context was queued");
    }
    return batch(context);
  },
});

const selectionAdapterContext = (
  fake: ReturnType<typeof createFakeOfficeContext>,
) => ({
  sync: fake.context.sync,
  workbook: {
    getSelectedRange: fake.context.getSelection,
  },
});

type BorderSide =
  | "bottom"
  | "diagonalDown"
  | "diagonalUp"
  | "horizontal"
  | "left"
  | "right"
  | "top"
  | "vertical";
type BorderState = { color: string; style: string; weight: string };
type FormatCellState = {
  format: {
    autoIndent: boolean;
    borders: Record<BorderSide, BorderState>;
    fill: {
      color: string;
      pattern: string;
      patternColor: string;
      patternTintAndShade: number;
      tintAndShade: number;
    };
    font: {
      bold: boolean;
      color: string;
      italic: boolean;
      name: string;
      size: number;
      strikethrough: boolean;
      subscript: boolean;
      superscript: boolean;
      tintAndShade: number;
      underline: string;
    };
    horizontalAlignment: string;
    indentLevel: number;
    protection: { formulaHidden: boolean; locked: boolean };
    readingOrder: string;
    shrinkToFit: boolean;
    textOrientation: number;
    useStandardHeight: boolean;
    useStandardWidth: boolean;
    verticalAlignment: string;
    wrapText: boolean;
  };
  style: string;
};

const emptyBorders = (): Record<BorderSide, BorderState> => ({
  bottom: { color: "#000000", style: "None", weight: "Thin" },
  diagonalDown: { color: "#000000", style: "None", weight: "Thin" },
  diagonalUp: { color: "#000000", style: "None", weight: "Thin" },
  horizontal: { color: "#000000", style: "None", weight: "Thin" },
  left: { color: "#000000", style: "None", weight: "Thin" },
  right: { color: "#000000", style: "None", weight: "Thin" },
  top: { color: "#000000", style: "None", weight: "Thin" },
  vertical: { color: "#000000", style: "None", weight: "Thin" },
});

const formatCell = (
  fontColor: string,
  style: string,
  variant = 0,
): FormatCellState => ({
  format: {
    autoIndent: variant === 1,
    borders: emptyBorders(),
    fill: {
      color: variant === 0 ? "#FFF2CC" : "#DDEBF7",
      pattern: variant === 0 ? "Solid" : "Gray25",
      patternColor: variant === 0 ? "#AA5500" : "#225588",
      patternTintAndShade: variant === 0 ? 0.1 : -0.2,
      tintAndShade: variant === 0 ? 0.2 : -0.1,
    },
    font: {
      bold: variant === 1,
      color: fontColor,
      italic: variant === 0,
      name: variant === 0 ? "Arial" : "Aptos Narrow",
      size: variant === 0 ? 10 : 12,
      strikethrough: variant === 1,
      subscript: false,
      superscript: variant === 0,
      tintAndShade: variant === 0 ? 0.15 : -0.25,
      underline: variant === 0 ? "Single" : "None",
    },
    horizontalAlignment: variant === 0 ? "Left" : "Right",
    indentLevel: variant,
    protection: { formulaHidden: variant === 1, locked: variant === 0 },
    readingOrder: variant === 0 ? "LeftToRight" : "RightToLeft",
    shrinkToFit: variant === 1,
    textOrientation: variant === 0 ? 0 : 45,
    useStandardHeight: variant === 0,
    useStandardWidth: variant === 1,
    verticalAlignment: variant === 0 ? "Top" : "Bottom",
    wrapText: variant === 1,
  },
  style,
});

const cloneFormatCells = (cells: FormatCellState[][]): FormatCellState[][] =>
  JSON.parse(JSON.stringify(cells)) as FormatCellState[][];

const createFormattingRuntime = (
  options: {
    formatApiSupported?: boolean;
    ignoreCellPropertyBorderWrites?: boolean;
    ignoreWrite?: boolean;
    protectedSheet?: boolean;
  } = {},
) => {
  const originalCells: FormatCellState[][] = [[
    formatCell("#112233", "Input"),
    formatCell("#445566", "Output", 1),
  ]];
  const originalNumberFormat = [["General", "0.00"]];
  let cells = cloneFormatCells(originalCells);
  let numberFormat = originalNumberFormat.map((row) => [...row]);
  let runCount = 0;
  let syncCount = 0;
  let ignoreWrite = options.ignoreWrite === true;
  const structuralCalls: Array<{ method: string; option?: string; levels?: [number, number] }> = [];
  const outlineLayers = { columns: 0, rows: 0 };
  const borderIndex: Record<string, BorderSide> = {
    EdgeBottom: "bottom",
    EdgeLeft: "left",
    EdgeRight: "right",
    EdgeTop: "top",
    DiagonalDown: "diagonalDown",
    DiagonalUp: "diagonalUp",
    InsideHorizontal: "horizontal",
    InsideVertical: "vertical",
  };
  const borderCells = (
    side: BorderSide,
    scopedCells?: FormatCellState[],
  ): FormatCellState[] => {
    if (scopedCells !== undefined) return scopedCells;
    if (side === "top") return cells[0] ?? [];
    if (side === "bottom") return cells.at(-1) ?? [];
    if (side === "left") return cells.flatMap((row) => row[0] === undefined ? [] : [row[0]]);
    if (side === "right") return cells.flatMap((row) => {
      const last = row.at(-1);
      return last === undefined ? [] : [last];
    });
    return cells.flat();
  };
  const borderPort = (index: string, scopedCells?: FormatCellState[]) => {
    const side = borderIndex[index] ?? "top";
    const first = () =>
      borderCells(side, scopedCells)[0]?.format.borders[side] ?? emptyBorders()[side];
    return {
      get color() {
        return first().color;
      },
      set color(color: string) {
        for (const cell of borderCells(side, scopedCells)) cell.format.borders[side].color = color;
      },
      get style() {
        return first().style;
      },
      set style(style: string) {
        for (const cell of borderCells(side, scopedCells)) cell.format.borders[side].style = style;
      },
      get weight() {
        return first().weight;
      },
      set weight(weight: string) {
        for (const cell of borderCells(side, scopedCells)) cell.format.borders[side].weight = weight;
      },
    };
  };
  const worksheet = {
    id: "sheet-1",
    name: "DCF",
    protection: {
      protected: options.protectedSheet === true,
      load: () => undefined,
    },
    showOutlineLevels(rowLevels: number, columnLevels: number) {
      structuralCalls.push({ levels: [rowLevels, columnLevels], method: "showOutlineLevels" });
    },
  };
  const entireRows = {
    group(option: string) {
      outlineLayers.rows += 1;
      structuralCalls.push({ method: "groupRows", option });
    },
    ungroup(option: string) {
      outlineLayers.rows = Math.max(0, outlineLayers.rows - 1);
      structuralCalls.push({ method: "ungroupRows", option });
    },
  };
  const entireColumns = {
    group(option: string) {
      outlineLayers.columns += 1;
      structuralCalls.push({ method: "groupColumns", option });
    },
    ungroup(option: string) {
      outlineLayers.columns = Math.max(0, outlineLayers.columns - 1);
      structuralCalls.push({ method: "ungroupColumns", option });
    },
  };
  const range = {
    address: "DCF!F42:G42",
    clear() {
      cells = cells.map((row) => row.map(() => ({
        ...formatCell("#000000", "Normal"),
      })));
      numberFormat = numberFormat.map((row) => row.map(() => "General"));
    },
    columnCount: 2,
    format: {
      borders: { getItem: borderPort },
      font: {
        get color() {
          return cells[0]?.[0]?.format.font.color ?? "#000000";
        },
        set color(color: string) {
          if (ignoreWrite) return;
          cells = cells.map((row) => row.map((cell) => ({
            ...cell,
            format: { ...cell.format, font: { ...cell.format.font, color } },
          })));
        },
      },
    },
    getCellProperties: (loadOptions: Excel.CellPropertiesLoadOptions) => ({
      value: cells.map((row) => row.map((cell) => {
        const result: {
          format?: Record<string, unknown>;
          style?: string;
        } = {};
        if (loadOptions.style === true) result.style = cell.style;
        if (loadOptions.format !== undefined) {
          const requested = loadOptions.format as Record<string, unknown>;
          const loaded: Record<string, unknown> = {};
          const scalarKeys = [
            "autoIndent",
            "horizontalAlignment",
            "indentLevel",
            "readingOrder",
            "shrinkToFit",
            "textOrientation",
            "useStandardHeight",
            "useStandardWidth",
            "verticalAlignment",
            "wrapText",
          ] as const;
          for (const key of scalarKeys) {
            if (requested[key] === true) loaded[key] = cell.format[key];
          }
          if (loadOptions.format.font !== undefined) {
            const requestedFont = loadOptions.format.font as Record<string, unknown>;
            loaded.font = Object.fromEntries(
              Object.entries(cell.format.font).filter(([key]) => requestedFont[key] === true),
            );
          }
          if (loadOptions.format.fill !== undefined) {
            const requestedFill = loadOptions.format.fill as Record<string, unknown>;
            loaded.fill = Object.fromEntries(
              Object.entries(cell.format.fill).filter(([key]) => requestedFill[key] === true),
            );
          }
          if (loadOptions.format.protection !== undefined) {
            loaded.protection = { ...cell.format.protection };
          }
          if (loadOptions.format.borders !== undefined) {
            loaded.borders = Object.fromEntries(
              Object.entries(cell.format.borders).map(([side, border]) => [side, { ...border }]),
            );
          }
          result.format = loaded;
        }
        return result;
      })),
    }),
    getCell(rowIndex: number, columnIndex: number) {
      const cell = cells[rowIndex]?.[columnIndex];
      return {
        format: {
          borders: {
            getItem: (index: string) => borderPort(
              index,
              cell === undefined ? [] : [cell],
            ),
          },
        },
      };
    },
    getEntireColumn() {
      structuralCalls.push({ method: "getEntireColumn" });
      return entireColumns;
    },
    getEntireRow() {
      structuralCalls.push({ method: "getEntireRow" });
      return entireRows;
    },
    load: () => undefined,
    get numberFormat() {
      return numberFormat;
    },
    set numberFormat(value: string[][]) {
      numberFormat = value.map((row) => [...row]);
    },
    rowCount: 1,
    setCellProperties(value: unknown[][]) {
      cells = cells.map((row, rowIndex) => row.map((cell, columnIndex) => {
        const update = value[rowIndex]?.[columnIndex] as {
          format?: {
            borders?: Partial<Record<BorderSide, Partial<BorderState>>>;
            fill?: Partial<FormatCellState["format"]["fill"]>;
            font?: Partial<FormatCellState["format"]["font"]>;
            protection?: Partial<FormatCellState["format"]["protection"]>;
          } & Partial<Omit<FormatCellState["format"], "borders" | "fill" | "font" | "protection">>;
          style?: string;
        } | undefined;
        if (update === undefined) return cell;
        const next = cloneFormatCells([[cell]])[0]?.[0];
        if (next === undefined) return cell;
        if (update.style !== undefined) next.style = update.style;
        Object.assign(next.format.font, update.format?.font);
        Object.assign(next.format.fill, update.format?.fill);
        Object.assign(next.format.protection, update.format?.protection);
        const scalarKeys = [
          "autoIndent",
          "horizontalAlignment",
          "indentLevel",
          "readingOrder",
          "shrinkToFit",
          "textOrientation",
          "useStandardHeight",
          "useStandardWidth",
          "verticalAlignment",
          "wrapText",
        ] as const;
        for (const key of scalarKeys) {
          const value = update.format?.[key];
          if (value !== undefined) next.format[key] = value as never;
        }
        if (options.ignoreCellPropertyBorderWrites !== true) {
          for (const [side, border] of Object.entries(update.format?.borders ?? {})) {
            Object.assign(next.format.borders[side as BorderSide], border);
          }
        }
        return next;
      }));
    },
    worksheet,
  };
  const context = {
    async sync() {
      syncCount += 1;
    },
    workbook: {
      getSelectedRange: () => range,
      getSelectedRanges: () => ({
        areaCount: 1,
        getTables: () => ({ items: [], load: () => undefined }),
        isEntireColumn: false,
        isEntireRow: false,
        load: () => undefined,
        worksheet,
      }),
      worksheets: {
        getActiveWorksheet: () => worksheet,
        getItem: () => ({ getRange: () => range }),
      },
    },
  };
  const runtime = {
    getDocumentUrl: () => "file:///tmp/model.xlsx",
    requirements: {
      isSetSupported(_set: string, version: string) {
        return version !== "1.9" || options.formatApiSupported !== false;
      },
    },
    async run<T>(batch: (rawContext: unknown) => Promise<T>) {
      runCount += 1;
      return batch(context);
    },
  };

  return {
    get cells() {
      return cells;
    },
    get numberFormat() {
      return numberFormat;
    },
    originalCells,
    originalNumberFormat,
    get runCount() {
      return runCount;
    },
    outlineLayers,
    runtime,
    setIgnoreWrite(value: boolean) {
      ignoreWrite = value;
    },
    setCellStyle(row: number, column: number, style: string) {
      const cell = cells[row]?.[column];
      if (cell !== undefined) cell.style = style;
    },
    get syncCount() {
      return syncCount;
    },
    structuralCalls,
  };
};

const createDriftingFormattingRuntime = (
  options: { driftAfterReads?: number; unintendedProtected?: boolean } = {},
) => {
  let selectedRangeReads = 0;
  const driftAfterReads = options.driftAfterReads ?? 2;
  const colors = {
    original: ["#112233", "#445566"],
    unintended: ["#778899", "#AABBCC"],
  };
  function makeRange(
    address: string,
    sheetId: string,
    key: keyof typeof colors,
    protectedSheet = false,
  ) {
    const range = {
      address,
      clear: () => undefined,
      columnCount: 2,
      format: {
        borders: { getItem: () => ({ color: "#000000", style: "None", weight: "Thin" }) },
        font: {
          get color() {
            return colors[key][0] ?? "#000000";
          },
          set color(value: string) {
            colors[key] = colors[key].map(() => value);
          },
        },
      },
      getCellProperties: () => ({
        value: [colors[key].map((color) => ({ format: { font: { color } } }))],
      }),
      load: () => undefined,
      numberFormat: [["General", "General"]],
      rowCount: 1,
      setCellProperties: () => undefined,
      worksheet: {
        id: sheetId,
        name: "DCF",
        protection: { load: () => undefined, protected: protectedSheet },
      },
    };
    return range;
  }

  const original = makeRange("DCF!F42:G42", "sheet-1", "original");
  const unintended = makeRange(
    "DCF!J10:K10",
    "sheet-2",
    "unintended",
    options.unintendedProtected,
  );
  const context = {
    sync: async () => undefined,
    workbook: {
      getSelectedRange: () => {
        selectedRangeReads += 1;
        return selectedRangeReads <= driftAfterReads ? original : unintended;
      },
      worksheets: {
        getItem: () => ({
          getRange: (address: string) => address === "$F$42:$G$42" ? original : unintended,
        }),
      },
    },
  };

  return {
    colors,
    runtime: {
      getDocumentUrl: () => "file:///tmp/model.xlsx",
      requirements: supportedRequirements,
      async run<T>(batch: (rawContext: unknown) => Promise<T>) {
        return batch(context);
      },
    },
  };
};

type StructuralRuntimeOptions = {
  areaCount?: number;
  containedTables?: number;
  intersectingTables?: number;
  isEntireColumn?: boolean;
  isEntireRow?: boolean;
  outlineSupported?: boolean;
  protectedSheet?: boolean;
  selectionError?: { code: string; message: string };
  structureError?: { code: string; message: string };
};

const createStructuralRuntime = (options: StructuralRuntimeOptions = {}) => {
  const calls: Array<{ method: string; option?: string; levels?: [number, number] }> = [];
  const layers = { columns: 0, rows: 0 };
  let pendingError: StructuralRuntimeOptions["structureError"];
  let selectionError = options.selectionError;
  const tableCollection = (count: number) => ({
    items: Array.from({ length: count }, (_, index) => ({ id: `table-${index}` })),
    load: () => undefined,
  });
  const structuralRange = (dimension: "rows" | "columns") => ({
    group(option: string) {
      calls.push({ method: dimension === "rows" ? "groupRows" : "groupColumns", option });
      pendingError = options.structureError;
      if (pendingError === undefined) layers[dimension] += 1;
    },
    ungroup(option: string) {
      calls.push({ method: dimension === "rows" ? "ungroupRows" : "ungroupColumns", option });
      pendingError = options.structureError;
      if (pendingError === undefined) layers[dimension] = Math.max(0, layers[dimension] - 1);
    },
  });
  const rows = structuralRange("rows");
  const columns = structuralRange("columns");
  const worksheet = {
    id: "sheet-1",
    name: "DCF",
    protection: {
      load: () => undefined,
      protected: options.protectedSheet === true,
    },
    showOutlineLevels(rowLevels: number, columnLevels: number) {
      calls.push({ levels: [rowLevels, columnLevels], method: "showOutlineLevels" });
      pendingError = options.structureError;
    },
  };
  const selection = {
    getEntireColumn() {
      calls.push({ method: "getEntireColumn" });
      return columns;
    },
    getEntireRow() {
      calls.push({ method: "getEntireRow" });
      return rows;
    },
    worksheet,
  };
  const rangeAreas = {
    areaCount: options.areaCount ?? 1,
    getTables(fullyContained = false) {
      return tableCollection(
        fullyContained
          ? options.containedTables ?? 0
          : options.intersectingTables ?? 0,
      );
    },
    isEntireColumn: options.isEntireColumn === true,
    isEntireRow: options.isEntireRow === true,
    load: () => undefined,
    worksheet,
  };
  const context = {
    async sync() {
      if (selectionError !== undefined) {
        const error = selectionError;
        selectionError = undefined;
        throw error;
      }
      if (pendingError !== undefined) {
        const error = pendingError;
        pendingError = undefined;
        throw error;
      }
    },
    workbook: {
      getSelectedRange: () => selection,
      getSelectedRanges: () => rangeAreas,
      worksheets: {
        getActiveWorksheet: () => worksheet,
      },
    },
  };
  return {
    calls,
    layers,
    runtime: {
      getDocumentUrl: () => "file:///tmp/model.xlsx",
      requirements: {
        isSetSupported(_set: string, version: string) {
          return version !== "1.10" || options.outlineSupported !== false;
        },
      },
      async run<T>(batch: (rawContext: unknown) => Promise<T>) {
        return batch(context);
      },
    },
  };
};

describe("Excel Office gateway", () => {
  it("performs and exactly reverts a mixed-selection format through Office cell properties", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toEqual({ ok: true });
    expect(fake.cells.map((row) => row.map((cell) => cell.format.font.color))).toEqual([
      ["#0000FF", "#0000FF"],
    ]);
    expect(fake.numberFormat).toEqual(fake.originalNumberFormat);

    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.originalCells);
    expect(fake.numberFormat).toEqual(fake.originalNumberFormat);
  });

  it("pins apply and verification to the range captured before selection drift", async () => {
    const fake = createDriftingFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toEqual({ ok: true });
    expect(fake.colors.original).toEqual(["#0000FF", "#0000FF"]);
    expect(fake.colors.unintended).toEqual(["#778899", "#AABBCC"]);
  });

  it("cannot bypass protected-sheet preflight when selection drifts before capture", async () => {
    const fake = createDriftingFormattingRuntime({
      driftAfterReads: 1,
      unintendedProtected: true,
    });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toEqual({ ok: true });
    expect(fake.colors.original).toEqual(["#0000FF", "#0000FF"]);
    expect(fake.colors.unintended).toEqual(["#778899", "#AABBCC"]);
  });

  it("applies an exact finance number format to every selected cell and reverts it", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.integer" })).resolves.toEqual({ ok: true });
    expect(fake.numberFormat).toEqual([
      ["#,##0;(#,##0);-", "#,##0;(#,##0);-"],
    ]);
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.numberFormat).toEqual(fake.originalNumberFormat);
  });

  it("increases decimals relative to each selected cell and verifies the resulting matrix", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.decimals.increase" })).resolves.toEqual({ ok: true });
    expect(fake.numberFormat).toEqual([["General", "0.000"]]);
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.numberFormat).toEqual(fake.originalNumberFormat);
  });

  it("applies and exactly reverts a selected-range border", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.border.top" })).resolves.toEqual({ ok: true });
    expect(fake.cells[0]?.map((cell) => cell.format.borders.top)).toEqual([
      { color: "#000000", style: "Continuous", weight: "Thin" },
      { color: "#000000", style: "Continuous", weight: "Thin" },
    ]);
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.originalCells);
  });

  it("directly restores borders when Excel ignores setCellProperties border writes", async () => {
    const fake = createFormattingRuntime({
      ignoreCellPropertyBorderWrites: true,
    });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.border.top" })).resolves.toEqual({ ok: true });
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.originalCells);
  });

  it("clears selected-range borders and reverts the prior border state", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);
    await expect(gateway.runWrite({ id: "format.border.top" })).resolves.toEqual({ ok: true });

    await expect(gateway.runWrite({ id: "format.border.clear" })).resolves.toEqual({ ok: true });
    expect(fake.cells.flat().every((cell) =>
      Object.values(cell.format.borders).every((border) => border.style === "None"),
    )).toBe(true);
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.cells[0]?.every((cell) => cell.format.borders.top.style === "Continuous")).toBe(true);
  });

  it("clears formats without changing values and restores the complete captured format", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.clear" })).resolves.toEqual({ ok: true });
    expect(fake.cells.flat().every((cell) =>
      cell.style === "Normal" && cell.format.font.color === "#000000",
    )).toBe(true);
    expect(fake.numberFormat).toEqual([["General", "General"]]);
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.originalCells);
    expect(fake.numberFormat).toEqual(fake.originalNumberFormat);
  });

  it.each([
    ["structure.groupRows", "getEntireRow", "groupRows", "ByRows"],
    ["structure.groupColumns", "getEntireColumn", "groupColumns", "ByColumns"],
  ] as const)("normalizes selected cells and performs %s exactly once", async (
    id,
    normalizeMethod,
    mutationMethod,
    option,
  ) => {
    const fake = createStructuralRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id })).resolves.toEqual({ ok: true });
    expect(fake.calls.filter((call) => call.method === normalizeMethod)).toHaveLength(1);
    expect(fake.calls.filter((call) => call.method === mutationMethod)).toEqual([
      { method: mutationMethod, option },
    ]);
  });

  it("does not de-duplicate repeated Group commands, allowing nested row layers", async () => {
    const fake = createStructuralRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "structure.groupRows" })).resolves.toEqual({ ok: true });
    await expect(gateway.runWrite({ id: "structure.groupRows" })).resolves.toEqual({ ok: true });

    expect(fake.calls.filter((call) => call.method === "groupRows")).toHaveLength(2);
    expect(fake.layers.rows).toBe(2);
  });

  it("removes exactly one nested outline layer per Ungroup invocation", async () => {
    const fake = createStructuralRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);
    await gateway.runWrite({ id: "structure.groupColumns" });
    await gateway.runWrite({ id: "structure.groupColumns" });

    await expect(gateway.runWrite({ id: "structure.ungroupColumns" })).resolves.toEqual({ ok: true });
    expect(fake.layers.columns).toBe(1);
    await expect(gateway.runWrite({ id: "structure.ungroupColumns" })).resolves.toEqual({ ok: true });
    expect(fake.layers.columns).toBe(0);
    expect(fake.calls.filter((call) => call.method === "ungroupColumns")).toHaveLength(2);
  });

  it.each([
    ["multi-area", { areaCount: 2 }, "structure.groupRows", "MULTI_AREA_SELECTION"],
    ["row command on entire columns", { isEntireColumn: true }, "structure.groupRows", "MIXED_STRUCTURE_SELECTION"],
    ["column command on entire rows", { isEntireRow: true }, "structure.groupColumns", "MIXED_STRUCTURE_SELECTION"],
    ["protected", { protectedSheet: true }, "structure.groupRows", "PROTECTED"],
    ["partial table", { intersectingTables: 1 }, "structure.groupRows", "PARTIAL_TABLE_SELECTION"],
  ] as const)("fails closed before mutation for %s selection", async (
    _case,
    options,
    id,
    code,
  ) => {
    const fake = createStructuralRuntime(options);
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id })).resolves.toMatchObject({ code, ok: false });
    expect(fake.calls.filter((call) => /^(group|ungroup)/u.test(call.method))).toHaveLength(0);
  });

  it("normalizes invalid Office selections without attempting a mutation", async () => {
    const fake = createStructuralRuntime({
      selectionError: { code: "InvalidArgument", message: "The selection is invalid" },
    });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "structure.groupRows" })).resolves.toMatchObject({
      code: "INVALID_STRUCTURE_SELECTION",
      ok: false,
    });
    expect(fake.calls.filter((call) => /^(group|ungroup)/u.test(call.method))).toHaveLength(0);
  });

  it("types an Excel outline-limit failure and does not install a false success barrier", async () => {
    const fake = createStructuralRuntime({
      structureError: {
        code: "InvalidOperation",
        message: "Cannot add another outline level because the maximum is 8",
      },
    });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "structure.groupRows" })).resolves.toMatchObject({
      code: "OUTLINE_LIMIT_REACHED",
      ok: false,
    });
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toMatchObject({
      code: "NO_REVERT",
      ok: false,
    });
  });

  it("fails unsupported outline commands before entering Office", async () => {
    const fake = createStructuralRuntime({ outlineSupported: false });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "structure.groupRows" })).resolves.toMatchObject({
      code: "UNSUPPORTED_OPERATION",
      ok: false,
    });
    expect(fake.calls).toHaveLength(0);
  });

  it.each([
    ["structure.showDetails", [8, 8]],
    ["structure.hideDetails", [1, 1]],
  ] as const)("uses only showOutlineLevels for %s", async (id, levels) => {
    const fake = createStructuralRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id })).resolves.toEqual({ ok: true });
    expect(fake.calls).toEqual([{ levels, method: "showOutlineLevels" }]);
  });

  it("keeps Open/Close out of mutation history but blocks stale history after Group", async () => {
    const fake = createFormattingRuntime();
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toEqual({ ok: true });
    await expect(gateway.runWrite({ id: "structure.showDetails" })).resolves.toEqual({ ok: true });
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.originalCells);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toEqual({ ok: true });
    await expect(gateway.runWrite({ id: "structure.groupRows" })).resolves.toEqual({ ok: true });
    await expect(gateway.runWrite({ id: "system.revertLastAction" })).resolves.toEqual({
      code: "STRUCTURAL_REVERT_REQUIRES_INVERSE",
      message: expect.stringContaining("Ungroup rows once"),
      ok: false,
    });
    expect(fake.cells.flat().every((cell) => cell.format.font.color === "#0000FF")).toBe(true);
  });

  it("disables formatting when per-cell property readback is unavailable", async () => {
    const fake = createFormattingRuntime({ formatApiSupported: false });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toMatchObject({
      code: "UNSUPPORTED_OPERATION",
      ok: false,
    });
    expect(fake.runCount).toBe(0);
  });

  it("rejects a protected sheet before snapshot or mutation", async () => {
    const fake = createFormattingRuntime({ protectedSheet: true });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toMatchObject({
      code: "PROTECTED",
      ok: false,
    });
    expect(fake.cells).toEqual(fake.originalCells);
    expect(fake.runCount).toBe(1);
  });

  it("acknowledges the safety lock before allowing a later write", async () => {
    const fake = createFormattingRuntime({ ignoreWrite: true });
    const gateway = createExcelOfficeGateway(fake.runtime);

    await expect(gateway.runWrite({ id: "format.input" })).resolves.toMatchObject({
      code: "WRITE_VERIFY_FAILED_REVERTED",
      ok: false,
    });
    await expect(gateway.runWrite({ id: "format.formula" })).resolves.toMatchObject({
      code: "SAFETY_LOCKED",
      ok: false,
    });
    await expect(gateway.runWrite({ id: "system.acknowledgeSafetyLock" })).resolves.toEqual({ ok: true });

    fake.setIgnoreWrite(false);
    await expect(gateway.runWrite({ id: "format.formula" })).resolves.toEqual({ ok: true });
  });

  it("returns canonical direct dependencies in one batch", async () => {
    let syncCount = 0;
    const loaded: string[] = [];
    const dependencyRanges = {
      items: [
        {
          address: "Assumptions!B3",
          worksheet: { id: "sheet-2", name: "Assumptions" },
        },
        {
          address: "DCF!F40",
          worksheet: { id: "sheet-1", name: "DCF" },
        },
      ],
      load(properties: string) {
        loaded.push(properties);
      },
    };
    const range = {
      getDirectDependents() {
        throw new Error("Wrong dependency direction");
      },
      getDirectPrecedents: () => ({ ranges: dependencyRanges }),
    };
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        worksheets: {
          getItem: (key: string) => {
            expect(key).toBe("sheet-1");
            return {
              getRange: (address: string) => {
                expect(address).toBe("$F$42");
                return range;
              },
            };
          },
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));

    const dependencies = await gateway.getDirectDependencies(
      {
        address: "$F$42",
        sheetId: "sheet-1",
        sheetName: "DCF",
        workbookId: "file:///tmp/model.xlsx",
      },
      "precedent",
    );

    expect(dependencies).toEqual([
      {
        address: "$B$3",
        sheetId: "sheet-2",
        sheetName: "Assumptions",
        workbookId: "file:///tmp/model.xlsx",
      },
      {
        address: "$F$40",
        sheetId: "sheet-1",
        sheetName: "DCF",
        workbookId: "file:///tmp/model.xlsx",
      },
    ]);
    expect(loaded).toEqual([
      "items/address,items/worksheet/id,items/worksheet/name",
    ]);
    expect(syncCount).toBe(1);
  });

  it("supports parallel precedent and dependent reads without cross-direction staleness", async () => {
    const dependencyContext = (address: string) => ({
      sync: async () => undefined,
      workbook: {
        worksheets: {
          getItem: () => ({
            getRange: () => ({
              getDirectDependents: () => ({
                ranges: {
                  items: [{ address, worksheet: { id: "sheet-2", name: "Inputs" } }],
                  load: () => undefined,
                },
              }),
              getDirectPrecedents: () => ({
                ranges: {
                  items: [{ address, worksheet: { id: "sheet-2", name: "Inputs" } }],
                  load: () => undefined,
                },
              }),
            }),
          }),
        },
      },
    });
    const gateway = createExcelOfficeGateway(createRuntime([
      dependencyContext("Inputs!B7"),
      dependencyContext("Inputs!B8"),
    ]));
    const selected = {
      address: "$F$42",
      sheetId: "sheet-1",
      sheetName: "DCF",
      workbookId: "file:///tmp/model.xlsx",
    };

    await expect(Promise.all([
      gateway.getDirectDependencies(selected, "precedent"),
      gateway.getDirectDependencies(selected, "dependent"),
    ])).resolves.toMatchObject([
      [{ address: "$B$7" }],
      [{ address: "$B$8" }],
    ]);
  });

  it("loads the four adjacent formula cells in one Office batch", async () => {
    let syncCount = 0;
    const loads: string[] = [];
    const worksheet = {
      id: "sheet-1",
      load: (properties: string[]) => loads.push(`sheet:${properties.join(",")}`),
      name: "DCF",
    };
    const neighborRange = (address: string, formula: string, value: number) => ({
      address: `DCF!${address}`,
      columnCount: 1,
      formulas: [[formula]],
      load: (properties: string[]) => loads.push(`${address}:${properties.join(",")}`),
      numberFormat: [["0.0x"]],
      rowCount: 1,
      values: [[value]],
      worksheet,
    });
    const neighbors = new Map([
      ["0,-1", neighborRange("E42", "=E40-E41", 11)],
      ["0,1", neighborRange("G42", "=G40-G41", 13)],
      ["-1,0", neighborRange("F41", "=F39-F40", 10)],
      ["1,0", neighborRange("F43", "=F41-F42", 14)],
    ]);
    const origin = {
      getOffsetRange: (rowOffset: number, columnOffset: number) => {
        const neighbor = neighbors.get(`${rowOffset},${columnOffset}`);
        if (neighbor === undefined) throw new Error("Unexpected neighbor offset");
        return neighbor;
      },
    };
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        worksheets: {
          getItem: () => ({ getRange: () => origin }),
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));
    const getNeighborFacts = gateway.getNeighborFacts;
    expect(getNeighborFacts).toBeDefined();
    if (getNeighborFacts === undefined) return;

    await expect(getNeighborFacts({
      address: "$F$42",
      sheetId: "sheet-1",
      sheetName: "DCF",
      workbookId: "file:///tmp/model.xlsx",
    })).resolves.toMatchObject([
      { direction: "left", facts: { address: "$E$42", formula: "=E40-E41" } },
      { direction: "right", facts: { address: "$G$42", formula: "=G40-G41" } },
      { direction: "above", facts: { address: "$F$41", formula: "=F39-F40" } },
      { direction: "below", facts: { address: "$F$43", formula: "=F41-F42" } },
    ]);
    expect(syncCount).toBe(1);
    expect(loads.filter((entry) => !entry.startsWith("sheet:"))).toHaveLength(4);
  });

  it("treats ItemNotFound as an empty dependency state", async () => {
    const context = {
      sync: async () => undefined,
      workbook: {
        worksheets: {
          getItem: () => ({
            getRange: () => ({
              getDirectPrecedents() {
                throw { code: "ItemNotFound" };
              },
            }),
          }),
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));

    await expect(
      gateway.getDirectDependencies(
        {
          address: "$F$42",
          sheetId: "sheet-1",
          sheetName: "DCF",
          workbookId: "file:///tmp/model.xlsx",
        },
        "precedent",
      ),
    ).resolves.toEqual([]);
  });

  it("rejects a superseded selection result", async () => {
    let releaseFirst: (() => void) | undefined;
    const firstSync = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });
    const first = createFakeOfficeContext().withSelection({
      address: "DCF!A1",
      value: 1,
    });
    first.context.sync = async () => firstSync;
    const second = createFakeOfficeContext().withSelection({
      address: "DCF!B2",
      value: 2,
    });
    const gateway = createExcelOfficeGateway(
      createRuntime([
        selectionAdapterContext(first),
        selectionAdapterContext(second),
      ]),
    );

    const staleResult = gateway.getSelectionFacts();
    await expect(gateway.getSelectionFacts()).resolves.toMatchObject({
      address: "$B$2",
      value: 2,
    });
    releaseFirst?.();

    await expect(staleResult).rejects.toMatchObject({ code: "STALE_REQUEST" });
  });

  it("does not enter an Office batch for an already-aborted request", async () => {
    const controller = new AbortController();
    controller.abort();
    const gateway = createExcelOfficeGateway(createRuntime([]));

    await expect(gateway.getSelectionFacts(controller.signal)).rejects.toMatchObject({
      code: "ABORTED",
    });
  });

  it("reads a pinned cell by stable sheet ID without changing selection", async () => {
    const fake = createFakeOfficeContext().withSelection({
      address: "Renamed Returns!F42",
      formula: "=XIRR(B2:B8,A2:A8)",
      sheetId: "stable-sheet-id",
      sheetName: "Renamed Returns",
      value: 19.1,
    });
    const context = {
      sync: fake.context.sync,
      workbook: {
        worksheets: {
          getItem: (sheetId: string) => {
            expect(sheetId).toBe("stable-sheet-id");
            return {
              getRange: (address: string) => {
                expect(address).toBe("$F$42");
                return fake.context.getSelection();
              },
            };
          },
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));
    const getCellFacts = gateway.getCellFacts;
    expect(getCellFacts).toBeDefined();
    if (getCellFacts === undefined) return;

    await expect(getCellFacts({
      address: "$F$42",
      sheetId: "stable-sheet-id",
      sheetName: "Returns",
      workbookId: "file:///tmp/model.xlsx",
    })).resolves.toMatchObject({
      address: "$F$42",
      formula: "=XIRR(B2:B8,A2:A8)",
      sheetId: "stable-sheet-id",
      sheetName: "Renamed Returns",
      value: 19.1,
    });
    expect(fake.syncCount).toBe(1);
  });

  it("batches pinned-cell reads and preserves missing sheets", async () => {
    let syncCount = 0;
    let selectedRangeReads = 0;
    const validSheet = {
      id: "stable-sheet-id",
      isNullObject: false,
      load: () => undefined,
      name: "Renamed Returns",
      getRange: (address: string) => ({
        address: `Renamed Returns!${address.replaceAll("$", "")}`,
        columnCount: 1,
        formulas: [[address === "$F$42" ? "=IRR(B2:B8)" : "=MOIC(B2:B8)"]],
        load: () => undefined,
        numberFormat: [["0.0x"]],
        rowCount: 1,
        values: [[address === "$F$42" ? 19.1 : 2.4]],
        worksheet: validSheet,
      }),
    };
    const missingSheet = {
      id: "missing-sheet-id",
      isNullObject: true,
      load: () => undefined,
      name: "",
      getRange: () => {
        throw new Error("A missing worksheet must not be dereferenced");
      },
    };
    const context = {
      sync: async () => {
        syncCount += 1;
      },
      workbook: {
        getSelectedRange: () => {
          selectedRangeReads += 1;
          throw new Error("Batch pin refresh must not inspect the selection");
        },
        worksheets: {
          getItemOrNullObject: (sheetId: string) =>
            sheetId === validSheet.id ? validSheet : missingSheet,
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));
    const getCellsFacts = gateway.getCellsFacts;
    expect(getCellsFacts).toBeDefined();
    if (getCellsFacts === undefined) return;

    await expect(getCellsFacts([
      {
        address: "$F$42",
        sheetId: "stable-sheet-id",
        sheetName: "Returns",
        workbookId: "file:///tmp/model.xlsx",
      },
      {
        address: "$G$43",
        sheetId: "stable-sheet-id",
        sheetName: "Returns",
        workbookId: "file:///tmp/model.xlsx",
      },
      {
        address: "$A$1",
        sheetId: "missing-sheet-id",
        sheetName: "Deleted",
        workbookId: "file:///tmp/model.xlsx",
      },
    ])).resolves.toEqual([
      expect.objectContaining({
        address: "$F$42",
        formula: "=IRR(B2:B8)",
        sheetId: "stable-sheet-id",
        sheetName: "Renamed Returns",
        value: 19.1,
      }),
      expect.objectContaining({
        address: "$G$43",
        formula: "=MOIC(B2:B8)",
        sheetId: "stable-sheet-id",
        sheetName: "Renamed Returns",
        value: 2.4,
      }),
      null,
    ]);
    expect(syncCount).toBe(2);
    expect(selectedRangeReads).toBe(0);
  });

  it("enforces metadata-only document settings persistence", async () => {
    let persisted: unknown = [{
      address: "$F$42",
      category: "Returns",
      formula: "=SECRET()",
      label: "IRR",
      sheetId: "sheet-1",
      value: 19.1,
      workbookId: "workbook-1",
    }];
    const settings = {
      get: (key: string) => {
        expect(key).toBe("analystLens.watchlist.v1");
        return persisted;
      },
      saveAsync: (callback: (result: { status: string }) => void) => {
        callback({ status: "succeeded" });
      },
      set: (key: string, value: unknown) => {
        expect(key).toBe("analystLens.watchlist.v1");
        persisted = value;
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([], settings));
    const load = gateway.loadWatchlistMetadata;
    const save = gateway.saveWatchlistMetadata;
    expect(load).toBeDefined();
    expect(save).toBeDefined();
    if (load === undefined || save === undefined) return;

    await expect(load()).resolves.toEqual([{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "sheet-1",
      workbookId: "workbook-1",
    }]);
    await save([{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "sheet-1",
      workbookId: "workbook-1",
      formula: "=SECRET()",
      value: 19.1,
    } as never]);

    expect(persisted).toEqual([{
      address: "$F$42",
      category: "Returns",
      label: "IRR",
      sheetId: "sheet-1",
      workbookId: "workbook-1",
    }]);
  });

  it("selects a canonical cell in one sync", async () => {
    let selected = false;
    let syncCount = 0;
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        worksheets: {
          getItem(key: string) {
            expect(key).toBe("sheet-2");
            return {
              getRange(address: string) {
                expect(address).toBe("$C$9");
                return {
                  select() {
                    selected = true;
                  },
                };
              },
            };
          },
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));

    await gateway.selectCell({
      address: "$C$9",
      sheetId: "sheet-2",
      sheetName: "Assumptions",
      workbookId: "file:///tmp/model.xlsx",
    });

    expect(selected).toBe(true);
    expect(syncCount).toBe(1);
  });

  it("loads workbook identity, calculation mode, and sheet facts in one sync", async () => {
    let syncCount = 0;
    const loads: string[] = [];
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        application: {
          calculationMode: "Manual",
          load: (properties: string) => loads.push(`application:${properties}`),
        },
        load: (properties: string) => loads.push(`workbook:${properties}`),
        name: "Model.xlsx",
        worksheets: {
          items: [
            { id: "sheet-1", name: "DCF", visibility: "Visible" },
            { id: "sheet-2", name: "Checks", visibility: "Hidden" },
          ],
          load: (properties: string) => loads.push(`worksheets:${properties}`),
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));

    const facts = await gateway.getWorkbookFacts();

    expect(facts).toEqual({
      calculationMode: "manual",
      id: "file:///tmp/model.xlsx",
      name: "Model.xlsx",
      worksheets: [
        { id: "sheet-1", name: "DCF", visibility: "Visible" },
        { id: "sheet-2", name: "Checks", visibility: "Hidden" },
      ],
    });
    expect(loads).toEqual([
      "workbook:name",
      "application:calculationMode",
      "worksheets:items/id,items/name,items/visibility",
    ]);
    expect(syncCount).toBe(1);
  });

  it("extracts used-range health facts without mutating the workbook", async () => {
    let syncCount = 0;
    let propertyReads = 0;
    const loads: string[] = [];
    let valuesOnly: boolean | undefined;
    const usedRange = {
      address: "Model!C5:D6",
      columnCount: 2,
      columnIndex: 2,
      isNullObject: false,
      load: (properties: string) => loads.push(`used-range:${properties}`),
      rowCount: 2,
      rowIndex: 4,
    };
    const loadedRange = {
      ...usedRange,
      formulas: [["=A1*2", 2], [3, "=C5"]],
      getCellProperties: () => {
        propertyReads += 1;
        return { value: [
          [{ format: { font: { color: "#0000FF" } } }, { format: { font: { color: "#000000" } } }],
          [{ format: { font: { color: "#000000" } } }, { format: { font: { color: "#008000" } } }],
        ] };
      },
      load: (properties: string) => loads.push(`data-range:${properties}`),
      numberFormat: [["0.0x", "General"], ["#,##0", "#,##0"]],
      values: [[1, 2], [3, "#REF!"]],
    };
    const worksheet = {
      getRangeByIndexes: (...indexes: number[]) => {
        expect(indexes).toEqual([4, 2, 2, 2]);
        return loadedRange;
      },
      getUsedRangeOrNullObject: (requestedValuesOnly?: boolean) => {
        valuesOnly = requestedValuesOnly;
        return usedRange;
      },
      id: "sheet-1",
      name: "Model",
      visibility: "Visible",
    };
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        application: {
          calculationMode: "Manual",
          load: (properties: string) => loads.push(`application:${properties}`),
        },
        name: "Model.xlsx",
        worksheets: {
          items: [worksheet],
          load: (properties: string) => loads.push(`worksheets:${properties}`),
        },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));
    const getHealthWorkbookFacts = gateway.getHealthWorkbookFacts;
    expect(getHealthWorkbookFacts).toBeDefined();
    if (getHealthWorkbookFacts === undefined) return;

    await expect(getHealthWorkbookFacts()).resolves.toMatchObject({
      calculationMode: "manual",
      cells: [
        { address: "$C$5", fontColor: "#0000FF", formula: "=A1*2", value: 1 },
        { address: "$D$5", fontColor: "#000000", formula: null, value: 2 },
        { address: "$C$6", fontColor: "#000000", formula: null, value: 3 },
        { address: "$D$6", fontColor: "#008000", formula: "=C5", value: "#REF!" },
      ],
      workbookRef: { address: "$A$1", sheetId: "sheet-1", sheetName: "Model" },
    });
    expect({ propertyReads, syncCount }).toEqual({ propertyReads: 1, syncCount: 3 });
    expect(valuesOnly).toBe(true);
    expect(loads).toContain(
      "used-range:address,columnCount,columnIndex,isNullObject,rowCount,rowIndex",
    );
    expect(loads).toContain("data-range:formulas,numberFormat,values");
  });

  it("rejects oversized used ranges before loading cell payloads", async () => {
    let dataRangeReads = 0;
    let syncCount = 0;
    const usedRange = {
      address: "Model!A1:A100001",
      columnCount: 1,
      columnIndex: 0,
      isNullObject: false,
      load: () => undefined,
      rowCount: 100_001,
      rowIndex: 0,
    };
    const worksheet = {
      getRangeByIndexes: () => {
        dataRangeReads += 1;
        throw new Error("Cell payload must not be requested");
      },
      getUsedRangeOrNullObject: () => usedRange,
      id: "sheet-1",
      name: "Model",
      visibility: "Visible",
    };
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        application: { calculationMode: "Automatic", load: () => undefined },
        name: "Model.xlsx",
        worksheets: { items: [worksheet], load: () => undefined },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));
    const getHealthWorkbookFacts = gateway.getHealthWorkbookFacts;
    expect(getHealthWorkbookFacts).toBeDefined();
    if (getHealthWorkbookFacts === undefined) return;

    await expect(getHealthWorkbookFacts()).rejects.toMatchObject({
      code: "UNSUPPORTED_OPERATION",
      message: expect.stringContaining("100,000"),
    });
    expect({ dataRangeReads, syncCount }).toEqual({ dataRangeReads: 0, syncCount: 2 });
  });

  it("bounds health fact payloads to 5,000-cell subranges", async () => {
    const reads: number[][] = [];
    let syncCount = 0;
    const worksheet = {
      getRangeByIndexes: (
        startRow: number,
        startColumn: number,
        rowCount: number,
        columnCount: number,
      ) => {
        reads.push([startRow, startColumn, rowCount, columnCount]);
        const rows = Array.from({ length: rowCount }, () => [1]);
        return {
          formulas: rows,
          getCellProperties: () => ({
            value: rows.map(() => [{ format: { font: { color: "#000000" } } }]),
          }),
          load: () => undefined,
          numberFormat: rows.map(() => ["General"]),
          values: rows,
        };
      },
      getUsedRangeOrNullObject: () => ({
        address: "Model!A1:A5001",
        columnCount: 1,
        columnIndex: 0,
        isNullObject: false,
        load: () => undefined,
        rowCount: 5_001,
        rowIndex: 0,
      }),
      id: "sheet-1",
      name: "Model",
      visibility: "Visible",
    };
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        application: { calculationMode: "Automatic", load: () => undefined },
        name: "Model.xlsx",
        worksheets: { items: [worksheet], load: () => undefined },
      },
    };
    const gateway = createExcelOfficeGateway(createRuntime([context]));
    const getHealthWorkbookFacts = gateway.getHealthWorkbookFacts;
    expect(getHealthWorkbookFacts).toBeDefined();
    if (getHealthWorkbookFacts === undefined) return;

    const facts = await getHealthWorkbookFacts();

    expect(facts.cells).toHaveLength(5_001);
    expect(reads).toEqual([[0, 0, 5_000, 1], [5_000, 0, 1, 1]]);
    expect(syncCount).toBe(4);
  });

  it("registers and removes selection and recalculation handlers", async () => {
    let selectionHandler: (() => Promise<void>) | undefined;
    let recalculationHandler: (() => Promise<void>) | undefined;
    let removes = 0;
    let syncCount = 0;
    const eventPort = (kind: "selection" | "recalculation") => ({
      add(handler: () => Promise<void>) {
        if (kind === "selection") selectionHandler = handler;
        else recalculationHandler = handler;
        return { remove: () => { removes += 1; } };
      },
    });
    const context = {
      async sync() {
        syncCount += 1;
      },
      workbook: {
        worksheets: {
          onCalculated: eventPort("recalculation"),
          onSelectionChanged: eventPort("selection"),
        },
      },
    };
    let contextRuns = 0;
    const gateway = createExcelOfficeGateway({
      ...createRuntime([context, context]),
      async runWithContext<T>(
        existingContext: unknown,
        batch: (context: unknown) => Promise<T>,
      ) {
        contextRuns += 1;
        return batch(existingContext);
      },
    });
    let selectionEvents = 0;
    let recalculationEvents = 0;

    const removeSelection = await gateway.subscribeSelection(() => {
      selectionEvents += 1;
    });
    const removeRecalculation = await gateway.subscribeRecalculation(() => {
      recalculationEvents += 1;
    });
    await selectionHandler?.();
    await recalculationHandler?.();
    removeSelection();
    removeRecalculation();

    expect({ contextRuns, selectionEvents, recalculationEvents, removes, syncCount }).toEqual({
      contextRuns: 2,
      recalculationEvents: 1,
      removes: 2,
      selectionEvents: 1,
      syncCount: 4,
    });
  });
});
