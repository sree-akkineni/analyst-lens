import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { HealthFinding } from "../../src/shared/types";
import { HealthScan, type HealthScanRunner } from "../../src/ui/HealthScan";

afterEach(cleanup);

const finding = (ruleId: string, severity: HealthFinding["severity"]): HealthFinding => ({
  evidence: `${ruleId} observed`,
  ref: {
    address: ruleId === "ERR_VALUE" ? "$A$1" : "$B$2",
    sheetId: "sheet-1",
    sheetName: "Model",
    workbookId: "workbook-1",
  },
  remediation: `Review ${ruleId}`,
  ruleId,
  severity,
});

describe("HealthScan", () => {
  it("runs, filters, navigates, and exports exact findings", async () => {
    const navigate = vi.fn();
    const exportLocal = vi.fn();
    const runScan: HealthScanRunner = async ({ onProgress }) => {
      onProgress({ completed: 2, total: 2 });
      return [finding("ERR_VALUE", "error"), finding("EXTERNAL_LINK", "warning")];
    };
    render(<HealthScan exportLocal={exportLocal} navigate={navigate} runScan={runScan} />);

    fireEvent.click(screen.getByRole("button", { name: "Run model health scan" }));
    expect(await screen.findByRole("button", { name: "Model $A$1 ERR_VALUE" })).toBeTruthy();
    fireEvent.change(screen.getByRole("combobox", { name: "Severity" }), {
      target: { value: "warning" },
    });
    expect(screen.queryByRole("button", { name: "Model $A$1 ERR_VALUE" })).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "Model $B$2 EXTERNAL_LINK" }));
    expect(navigate).toHaveBeenCalledWith(expect.objectContaining({ address: "$B$2" }));
    expect(screen.getByText("Review EXTERNAL_LINK")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "Export local Markdown" }));
    expect(exportLocal).toHaveBeenCalledWith(expect.stringContaining("EXTERNAL_LINK"));

    fireEvent.click(screen.getByRole("button", { name: "Run model health scan" }));
    expect(await screen.findByRole("button", { name: "Model $A$1 ERR_VALUE" })).toBeTruthy();
  });

  it("cancels without displaying partial findings", async () => {
    const runScan: HealthScanRunner = ({ signal }) => new Promise((_resolve, reject) => {
      signal.addEventListener("abort", () => reject(new DOMException("cancelled", "AbortError")));
    });
    render(<HealthScan exportLocal={() => undefined} navigate={() => undefined} runScan={runScan} />);

    fireEvent.click(screen.getByRole("button", { name: "Run model health scan" }));
    fireEvent.click(await screen.findByRole("button", { name: "Cancel health scan" }));

    await waitFor(() => expect(screen.getByRole("status").textContent).toBe(
      "Scan cancelled. No partial findings were kept.",
    ));
    expect(screen.queryByRole("list", { name: "Health findings" })).toBeNull();
  });

  it("treats a normalized Office abort as cancellation", async () => {
    const runScan: HealthScanRunner = async () => {
      throw Object.assign(new Error("The Office request was cancelled"), { code: "ABORTED" });
    };
    render(<HealthScan exportLocal={() => undefined} navigate={() => undefined} runScan={runScan} />);

    fireEvent.click(screen.getByRole("button", { name: "Run model health scan" }));

    await waitFor(() => expect(screen.getByRole("status").textContent).toBe(
      "Scan cancelled. No partial findings were kept.",
    ));
    expect(screen.queryByRole("list", { name: "Health findings" })).toBeNull();
  });

  it("aborts the active scan when the panel unmounts", () => {
    let observedSignal: AbortSignal | undefined;
    const runScan: HealthScanRunner = ({ signal }) => {
      observedSignal = signal;
      return new Promise(() => undefined);
    };
    const { unmount } = render(
      <HealthScan exportLocal={() => undefined} navigate={() => undefined} runScan={runScan} />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Run model health scan" }));
    unmount();

    expect(observedSignal?.aborted).toBe(true);
  });

  it("paginates a 10,000-finding result without rendering every row", async () => {
    const findings = Array.from({ length: 10_000 }, (_, index) => ({
      ...finding("ERR_VALUE", "error"),
      ref: {
        ...finding("ERR_VALUE", "error").ref,
        address: `$A$${index + 1}`,
      },
    }));
    const startedAt = performance.now();
    render(
      <HealthScan
        exportLocal={() => undefined}
        navigate={() => undefined}
        runScan={async () => findings}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Run model health scan" }));

    expect(await screen.findByText(
      "Showing 1–100 of 10000 filtered findings (10000 total).",
    )).toBeTruthy();
    expect(screen.getAllByRole("button", { name: /ERR_VALUE$/u })).toHaveLength(100);
    expect(performance.now() - startedAt).toBeLessThan(2_000);
    fireEvent.click(screen.getByRole("button", { name: "Next findings page" }));
    expect(await screen.findByRole("button", { name: "Model $A$101 ERR_VALUE" })).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Model $A$1 ERR_VALUE" })).toBeNull();
  });
});
