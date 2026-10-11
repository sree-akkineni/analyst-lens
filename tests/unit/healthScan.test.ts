import { describe, expect, it } from "vitest";
import type { HealthCellFacts, HealthWorkbookFacts } from "../../src/domain/health/rules";
import {
  createScanStore,
  exportHealthMarkdown,
  scanWorkbook,
} from "../../src/domain/health/scan";

const generatedWorkbook = (count: number): HealthWorkbookFacts => {
  const workbookRef = {
    address: "$A$1",
    sheetId: "sheet-1",
    sheetName: "Generated",
    workbookId: "workbook-1",
  };
  const cells: HealthCellFacts[] = Array.from({ length: count }, (_, index) => ({
    ...workbookRef,
    address: `$A$${index + 1}`,
    columnCount: 1,
    fontColor: "#000000",
    formula: null,
    numberFormat: "General",
    rowCount: 1,
    value: index,
  }));
  return { calculationMode: "automatic", cells, workbookRef };
};

const generatedFormulaWorkbook = (count: number): HealthWorkbookFacts => {
  const generated = generatedWorkbook(count);
  return {
    ...generated,
    cells: generated.cells.map((current, index) => ({
      ...current,
      formula: `=B${index + 1}`,
    })),
  };
};

describe("health scan coordinator", () => {
  it("evaluates each cell chunk before yielding control", async () => {
    const generated = generatedWorkbook(1);
    const first = generated.cells[0];
    if (first === undefined) throw new Error("Generated fixture did not contain a cell");
    let formulaRead = false;
    const observed = {
      ...first,
      get formula() {
        formulaRead = true;
        return "=B1*2";
      },
    };

    await scanWorkbook({ ...generated, cells: [observed] }, {
      chunkSize: 1,
      yieldControl: async () => {
        expect(formulaRead).toBe(true);
      },
    });
  });

  it("cancels between chunks without publishing a partial result", async () => {
    const controller = new AbortController();
    const store = createScanStore();
    const scan = scanWorkbook(generatedWorkbook(10_000), {
      chunkSize: 1_000,
      signal: controller.signal,
      store,
      yieldControl: async () => {
        controller.abort();
      },
    });

    await expect(scan).rejects.toMatchObject({ name: "AbortError" });
    expect(store.current).toBeNull();
  });

  it("reports progress only after the 500 ms threshold and publishes atomically", async () => {
    let clock = 0;
    const progress: number[] = [];
    const store = createScanStore();
    const findings = await scanWorkbook(generatedWorkbook(3), {
      chunkSize: 1,
      now: () => clock,
      onProgress: ({ completed }) => progress.push(completed),
      progressDelayMs: 500,
      store,
      yieldControl: async () => {
        clock += 300;
      },
    });

    expect(progress).toEqual([2, 3]);
    expect(store.current).toEqual(findings);
  });

  it("does not let an older concurrent scan overwrite a newer result", async () => {
    let releaseFirst: (() => void) | undefined;
    const firstYield = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });
    const store = createScanStore();
    const firstWorkbook = generatedWorkbook(1);
    const firstCell = firstWorkbook.cells[0];
    const secondWorkbook = generatedWorkbook(1);
    const secondCell = secondWorkbook.cells[0];
    if (firstCell === undefined || secondCell === undefined) {
      throw new Error("Generated fixture did not contain a cell");
    }
    const older = scanWorkbook({
      ...firstWorkbook,
      cells: [{ ...firstCell, formula: "=B1*2" }],
    }, { store, yieldControl: () => firstYield });
    const newer = await scanWorkbook({
      ...secondWorkbook,
      cells: [{ ...secondCell, value: "#REF!" }],
    }, { store, yieldControl: async () => undefined });
    releaseFirst?.();
    await older;

    expect(store.current).toEqual(newer);
    expect(store.current?.map(({ ruleId }) => ruleId)).toEqual(["ERR_VALUE"]);
  });

  it.each([10_000, 50_000, 100_000])(
    "scans a generated %i-cell workbook within the 20 second budget",
    async (count) => {
      const startedAt = performance.now();
      await scanWorkbook(generatedFormulaWorkbook(count), { chunkSize: 1_000 });
      expect(performance.now() - startedAt).toBeLessThan(20_000);
    },
    20_000,
  );

  it("exports exact local Markdown evidence without a transport dependency", async () => {
    const generated = generatedWorkbook(1);
    const first = generated.cells[0];
    if (first === undefined) throw new Error("Generated fixture did not contain a cell");
    const workbook: HealthWorkbookFacts = {
      ...generated,
      cells: [{ ...first, formula: "=B1*2", value: 2 }],
    };
    const findings = await scanWorkbook(workbook);
    const markdown = exportHealthMarkdown(findings, "Generated model");

    expect(markdown).toContain("EMBEDDED_CONSTANT");
    expect(markdown).toContain("Generated!$A$1");
    expect(markdown).toContain("Evidence:");
    expect(markdown).toContain("Remediation:");
    expect(markdown).not.toMatch(/https?:\/\//u);
  });
});
