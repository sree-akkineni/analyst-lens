import { afterEach, describe, expect, it, vi } from "vitest";
import { allCommands } from "../../src/domain/commands/commandRegistry";
import {
  applyMacShortcutDefaults,
  initializeOfficeActionBindings,
} from "../../src/platform/office/officeActionPort";

afterEach(() => vi.unstubAllGlobals());

describe("Office action initialization", () => {
  it("waits for Office before associating shared-runtime actions", async () => {
    const order: string[] = [];
    vi.stubGlobal("Office", {
      actions: {
        areShortcutsInUse: async () => [],
        associate: (actionId: string) => order.push(`associate:${actionId}`),
        getShortcuts: async () => ({}),
        replaceShortcuts: async () => undefined,
      },
      addin: { showAsTaskpane: async () => undefined },
      async onReady() {
        order.push("ready");
      },
    });

    await initializeOfficeActionBindings(async () => ({ ok: true }));

    expect(order).toEqual([
      "ready",
      ...allCommands.flatMap((command) => command.actionId === undefined
        ? []
        : [`associate:${command.actionId}`]),
    ]);
  });

  it("applies the registered Mac shortcut defaults through the Office port", async () => {
    const replacements: Array<Record<string, string>> = [];
    vi.stubGlobal("Office", {
      actions: {
        areShortcutsInUse: async () => [],
        associate: () => undefined,
        getShortcuts: async () => ({ FormatInteger: "Ctrl+Shift+1" }),
        replaceShortcuts: async (shortcuts: Record<string, string>) => {
          replacements.push(shortcuts);
        },
      },
      addin: { showAsTaskpane: async () => undefined },
      onReady: async () => undefined,
    });

    await expect(applyMacShortcutDefaults()).resolves.toMatchObject({ ok: true });
    expect(replacements[0]).toMatchObject({
      FormatInteger: "Ctrl+Shift+1",
      FormatTopBorder: "Ctrl+Alt+Shift+Up",
      OpenQuickTraceDependents: "Ctrl+Alt+D",
      OpenQuickTracePrecedents: "Ctrl+Alt+P",
    });
    expect(replacements[0]).not.toHaveProperty("OpenAnalystLens");
  });
});
