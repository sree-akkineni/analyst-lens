import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { App } from "../../src/ui/App";
import { detectCapabilities } from "../../src/platform/office/capabilityMap";
import type { OfficeGateway } from "../../src/platform/office/officeGateway";

afterEach(cleanup);

describe("Analyst Lens shell", () => {
  it("shows a recognizable, ready task pane", () => {
    render(<App />);

    expect(screen.getByRole("heading", { name: "Analyst Lens" })).toBeTruthy();
    expect(screen.getByRole("status").textContent).toContain("Local and ready");
  });

  it("places visible speed tools in the working task pane", () => {
    render(<App runCommand={async () => ({ ok: true })} />);

    expect(screen.getByRole("button", { name: /Speed tools/u })).toBeTruthy();
    expect(screen.getByRole("button", { name: /^Input blue / })).toBeTruthy();
  });

  it("boots the live Inspector from a shared lazy Office gateway", async () => {
    const opened: string[] = [];
    const gateway: OfficeGateway = {
      getCapabilities: () => detectCapabilities({ isSetSupported: () => true }),
      getDirectDependencies: async () => [],
      getHealthWorkbookFacts: async () => ({
        calculationMode: "automatic",
        cells: [],
        workbookRef: {
          address: "$A$1",
          sheetId: "sheet-1",
          sheetName: "DCF",
          workbookId: "workbook-1",
        },
      }),
      getSelectionFacts: async () => ({
        address: "$F$42",
        columnCount: 1,
        formula: "=$B$7*2",
        numberFormat: "0.0x",
        rowCount: 1,
        sheetId: "sheet-1",
        sheetName: "DCF",
        value: 12,
        workbookId: "workbook-1",
      }),
      getWorkbookFacts: async () => ({
        calculationMode: "automatic",
        id: "workbook-1",
        name: "Model.xlsx",
        worksheets: [],
      }),
      runWrite: async () => ({ ok: true }),
      selectCell: async () => undefined,
      subscribeRecalculation: async () => () => undefined,
      subscribeSelection: async () => () => undefined,
    };

    render(<App
      getGateway={async () => gateway}
      openQuickTrace={async (direction) => {
        opened.push(direction);
      }}
      quickTraceSupported
    />);

    expect(await screen.findByRole("heading", { name: "DCF $F$42" })).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: /Open Quick Trace In/u }));
    await waitFor(() => expect(opened).toEqual(["precedent"]));
    expect(screen.getByRole("button", { name: /Speed tools/u })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Model health/u })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Key outputs/u })).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Run model health scan" })).toBeNull();
  });
});
