import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { SpeedLayer } from "../../src/ui/SpeedLayer";

afterEach(cleanup);

describe("SpeedLayer", () => {
  it("runs a verified format from a visible keyboard-accessible control", async () => {
    const calls: string[] = [];
    render(<SpeedLayer runCommand={async (id) => {
      calls.push(id);
      return { ok: true };
    }} />);

    fireEvent.click(screen.getByRole("button", { name: /^Input blue / }));

    await waitFor(() => expect(calls).toEqual(["format.input"]));
    expect(screen.getByRole("status").textContent).toContain("Input blue applied");
    expect(screen.getByRole("button", { name: "Revert last Analyst Lens action" })).toBeTruthy();
  });

  it("enables structure controls with honest one-layer and automatic-Revert copy", () => {
    render(<SpeedLayer runCommand={async () => ({ ok: true })} />);

    expect(screen.getByRole("button", { name: "Group rows" }).hasAttribute("disabled")).toBe(false);
    expect(screen.getByRole("button", { name: "Open All Groups" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Close All Groups" })).toBeTruthy();
    expect(screen.getByText(/adds or removes exactly one outline layer/i)).toBeTruthy();
    expect(screen.getByText(/automatic Revert is unavailable/i)).toBeTruthy();
    expect(screen.getByText(/Command\+Option\+L/)).toBeTruthy();
  });

  it("directs structural success to the one-layer opposite command", async () => {
    render(<SpeedLayer runCommand={async () => ({ ok: true })} />);

    fireEvent.click(screen.getByRole("button", { name: "Group rows" }));

    await waitFor(() => expect(screen.getByRole("status").textContent).toContain(
      "Ungroup rows once",
    ));
    expect(screen.getByRole("status").textContent).toContain("Automatic Revert is unavailable");
  });

  it("reports Open All as visibility-only without offering Revert", async () => {
    render(<SpeedLayer runCommand={async () => ({ ok: true })} />);

    fireEvent.click(screen.getByRole("button", { name: "Open All Groups" }));

    await waitFor(() => expect(screen.getByRole("status").textContent).toContain(
      "Group definitions were not changed",
    ));
    expect(screen.getByRole("status").textContent).not.toContain("Revert is available");
  });

  it("shows the Macabacus-compatible Mac shortcut beside each speed tool", () => {
    render(<SpeedLayer runCommand={async () => ({ ok: true })} />);

    expect(screen.getByRole("button", { name: /Integer Ctrl\+Shift\+1/ })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Currency Ctrl\+Shift\+4/ })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Top border Ctrl\+Option\+Shift\+Up/ })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Clear borders Ctrl\+Option\+Shift\+N/ })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Clear formats Ctrl\+Option\+Shift\+C/ })).toBeTruthy();
    expect(screen.queryByText(/accepts at most 18 add-in shortcut actions/)).toBeNull();
  });

  it("shows a normalized write failure without reporting success", async () => {
    render(<SpeedLayer runCommand={async () => ({
      code: "PROTECTED",
      message: "Formatting is unavailable on a protected sheet",
      ok: false,
    })} />);

    fireEvent.click(screen.getByRole("button", { name: /^Formula black / }));

    await waitFor(() => expect(screen.getByRole("alert").textContent).toContain("protected sheet"));
    expect(screen.queryByText("Formula black applied")).toBeNull();
  });

  it("recovers from a rejected command runner without leaving controls pending", async () => {
    render(<SpeedLayer runCommand={async () => {
      throw new Error("Office failed to initialize");
    }} />);

    fireEvent.click(screen.getByRole("button", { name: /^Input blue / }));

    await waitFor(() => expect(screen.getByRole("alert").textContent).toContain(
      "Analyst Lens could not connect to Excel",
    ));
    expect(screen.getByRole("button", { name: /^Input blue / }).hasAttribute("disabled")).toBe(false);
  });

  it("offers and runs a visible safety-lock acknowledgement", async () => {
    const calls: string[] = [];
    render(<SpeedLayer runCommand={async (id) => {
      calls.push(id);
      return id === "system.acknowledgeSafetyLock"
        ? { ok: true }
        : {
            code: "WRITE_VERIFY_FAILED_REVERTED",
            message: "Write verification failed and the captured state was restored",
            ok: false,
          };
    }} />);

    fireEvent.click(screen.getByRole("button", { name: /^Input blue / }));
    const acknowledge = await screen.findByRole("button", { name: "Acknowledge safety lock" });
    fireEvent.click(acknowledge);

    await waitFor(() => expect(calls).toEqual([
      "format.input",
      "system.acknowledgeSafetyLock",
    ]));
    expect(screen.queryByRole("button", { name: "Acknowledge safety lock" })).toBeNull();
  });
});
