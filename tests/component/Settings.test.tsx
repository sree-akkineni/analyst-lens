import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { Settings } from "../../src/ui/Settings";

afterEach(cleanup);

describe("Settings", () => {
  it("documents the Macabacus-compatible default shortcut map", () => {
    render(<Settings onClose={() => undefined} open />);

    expect(screen.getByText("Macabacus-compatible Mac defaults")).toBeTruthy();
    expect(screen.getByText("Ctrl+Shift+8")).toBeTruthy();
    expect(screen.getByText("Ctrl+Option+Shift+Up")).toBeTruthy();
    expect(screen.getByText("Ctrl+Option+Shift+N")).toBeTruthy();
    expect(screen.getByText("Ctrl+Option+Shift+C")).toBeTruthy();
    expect(screen.getByText("Ctrl+Option+P")).toBeTruthy();
    expect(screen.getByText("Ctrl+Option+D")).toBeTruthy();
    expect(screen.getByText(/punctuation.*P.*D/iu)).toBeTruthy();
    expect(screen.queryByText(/remains available as a button/)).toBeNull();
  });

  it("requires an explicit second action before replacing shortcut conflicts", async () => {
    const calls: boolean[] = [];
    render(<Settings
      onClose={() => undefined}
      open
      updateShortcuts={async ({ replaceConflicts = false } = {}) => {
        calls.push(replaceConflicts);
        return replaceConflicts
          ? { current: { FormatInteger: "Ctrl+Shift+1" }, ok: true }
          : {
              conflicts: [{ actionId: "FormatInteger", shortcut: "Ctrl+Shift+1" }],
              ok: false,
            };
      }}
    />);

    fireEvent.click(screen.getByRole("button", { name: "Apply Macabacus defaults" }));
    expect((await screen.findByRole("alert")).textContent).toContain("Ctrl+Shift+1");
    expect(calls).toEqual([false]);

    fireEvent.click(screen.getByRole("button", { name: "Use Analyst Lens anyway" }));
    await waitFor(() => expect(calls).toEqual([false, true]));
    expect(screen.getByRole("status").textContent).toContain("1 shortcut is active");
  });
});
