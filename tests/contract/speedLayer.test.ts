import { describe, expect, it } from "vitest";
import {
  formatSelection,
  type FormatMutation,
  type SelectionFormatSnapshot,
  type SelectionFormatWritePort,
} from "../../src/domain/commands/formatCommands";
import { createWriteExecutor } from "../../src/platform/office/writeExecutor";

type CellState = {
  fontColor: string;
  numberFormat: string;
  style: string;
};

const copyCells = (cells: CellState[][]): CellState[][] =>
  cells.map((row) => row.map((cell) => ({ ...cell })));

const createFormatPort = (options: { ignoreWrite?: boolean } = {}) => {
  const original: CellState[][] = [[
    { fontColor: "#112233", numberFormat: "General", style: "Input" },
    { fontColor: "#445566", numberFormat: "0.00", style: "Output" },
  ]];
  let cells = copyCells(original);
  let lastMutation: FormatMutation | null = null;

  const apply = (mutation: FormatMutation) => {
    if (options.ignoreWrite === true) return;
    cells = cells.map((row) => row.map((cell) => {
      if ("fontColor" in mutation) return { ...cell, fontColor: mutation.fontColor };
      if ("numberFormat" in mutation) return { ...cell, numberFormat: mutation.numberFormat };
      if ("clearFormats" in mutation) {
        return { fontColor: "#000000", numberFormat: "General", style: "Normal" };
      }
      return cell;
    }));
  };

  const mutationApplied = (mutation: FormatMutation) => cells.every((row) => row.every((cell) => {
    if ("fontColor" in mutation) return cell.fontColor === mutation.fontColor;
    if ("numberFormat" in mutation) return cell.numberFormat === mutation.numberFormat;
    if ("clearFormats" in mutation) return cell.style === "Normal";
    return true;
  }));

  const port: SelectionFormatWritePort = {
    async applySelectionFormat(mutation) {
      lastMutation = mutation;
      apply(mutation);
    },
    async captureSelectionFormat(mutation) {
      return {
        address: "$F$42:$G$42",
        captured: "fontColor" in mutation
          ? "fontColor"
          : "numberFormat" in mutation || "decimalDelta" in mutation
            ? "numberFormat"
            : "border" in mutation || "clearBorders" in mutation
              ? "borders"
              : "allFormats",
        cellProperties: copyCells(cells),
        numberFormat: cells.map((row) => row.map((cell) => cell.numberFormat)),
        sheetProtected: false,
        sheetId: "sheet-1",
        workbookId: "workbook-1",
      } satisfies SelectionFormatSnapshot;
    },
    async preflightSelection() {
      return { ok: true };
    },
    async restoreSelectionFormat(snapshot) {
      cells = copyCells(snapshot.cellProperties as CellState[][]);
    },
    async verifySelectionFormat(mutation) {
      return mutationApplied(mutation);
    },
    async verifySelectionRestore(snapshot) {
      return JSON.stringify(cells) === JSON.stringify(snapshot.cellProperties);
    },
  };

  return {
    get cells() {
      return cells;
    },
    get lastMutation() {
      return lastMutation;
    },
    original,
    port,
  };
};

describe("analyst speed-layer writes", () => {
  it("changes only the declared font property and restores a mixed selection exactly", async () => {
    const fake = createFormatPort();
    const executor = createWriteExecutor(fake.port);

    await expect(executor.execute(formatSelection("format.input"))).resolves.toEqual({ ok: true });
    expect(fake.lastMutation).toEqual({ fontColor: "#0000FF" });
    expect(fake.cells).toEqual([[
      { fontColor: "#0000FF", numberFormat: "General", style: "Input" },
      { fontColor: "#0000FF", numberFormat: "0.00", style: "Output" },
    ]]);

    await expect(executor.revertLast()).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.original);
  });

  it("restores every captured property after clearing formats", async () => {
    const fake = createFormatPort();
    const executor = createWriteExecutor(fake.port);

    await expect(executor.execute(formatSelection("format.clear"))).resolves.toEqual({ ok: true });
    expect(fake.cells.every((row) => row.every((cell) => cell.style === "Normal"))).toBe(true);
    await expect(executor.revertLast()).resolves.toEqual({ ok: true });
    expect(fake.cells).toEqual(fake.original);
  });

  it("automatically restores and safety-locks when Office ignores the requested mutation", async () => {
    const fake = createFormatPort({ ignoreWrite: true });
    const executor = createWriteExecutor(fake.port);

    await expect(executor.execute(formatSelection("format.input"))).resolves.toMatchObject({
      code: "WRITE_VERIFY_FAILED_REVERTED",
      ok: false,
    });
    expect(fake.cells).toEqual(fake.original);
    expect(executor.isSafetyLocked()).toBe(true);
  });
});
