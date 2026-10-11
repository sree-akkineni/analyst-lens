import { waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("../../src/ui/App", () => ({
  App: () => <div>Task pane rendered</div>,
}));

describe("task pane bootstrap", () => {
  afterEach(() => {
    document.body.innerHTML = "";
    vi.resetModules();
    vi.unstubAllGlobals();
  });

  it("does not read Office runtime enums before Office is ready", async () => {
    let releaseReady: (() => void) | undefined;
    let ready = false;
    const readyPromise = new Promise<void>((resolve) => {
      releaseReady = resolve;
    });
    const enumAccessPhases: string[] = [];
    const office = {
      actions: {
        areShortcutsInUse: vi.fn(),
        associate: vi.fn(),
        getShortcuts: vi.fn(),
        replaceShortcuts: vi.fn(),
      },
      addin: { showAsTaskpane: vi.fn() },
      context: {
        requirements: { isSetSupported: vi.fn(() => true) },
        ui: { displayDialogAsync: vi.fn() },
      },
      get AsyncResultStatus() {
        enumAccessPhases.push(ready ? "ready" : "early");
        return { Succeeded: "succeeded" };
      },
      EventType: {
        DialogEventReceived: "dialog-event",
        DialogMessageReceived: "dialog-message",
      },
      onReady: vi.fn(() => readyPromise),
    };
    vi.stubGlobal("__ANALYST_LENS_FIXTURES__", false);
    vi.stubGlobal("Office", office);
    document.body.innerHTML = '<div id="root"></div>';

    await import("../../src/taskpane");

    expect(document.getElementById("root")?.childElementCount).toBe(0);
    expect(enumAccessPhases).toEqual([]);

    ready = true;
    releaseReady?.();

    await waitFor(() => expect(document.getElementById("root")?.textContent).toContain(
      "Task pane rendered",
    ));
    expect(enumAccessPhases).toEqual(["ready"]);
  });
});
