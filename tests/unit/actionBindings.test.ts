import { describe, expect, it } from "vitest";
import {
  bindManifestActions,
  updateShortcuts,
  type OfficeActionPort,
} from "../../src/runtime/actionBindings";
import { allCommands } from "../../src/domain/commands/commandRegistry";

const createActionPort = (conflicting = false) => {
  const handlers = new Map<string, () => Promise<void>>();
  const replacements: Array<Record<string, string>> = [];
  let paneOpens = 0;
  const port: OfficeActionPort = {
    areShortcutsInUse: async (shortcuts) =>
      shortcuts.map((shortcut) => ({ inUse: conflicting, shortcut })),
    associate(actionId, handler) {
      handlers.set(actionId, handler);
    },
    getShortcuts: async () => ({ OpenAnalystLens: "Command+Option+L" }),
    async replaceShortcuts(shortcuts) {
      replacements.push(shortcuts);
    },
    async showTaskpane() {
      paneOpens += 1;
    },
  };
  return {
    handlers,
    get paneOpens() {
      return paneOpens;
    },
    port,
    replacements,
  };
};

describe("shared-runtime action bindings", () => {
  it("associates every manifest action and opens the task pane", async () => {
    const fake = createActionPort();
    const dispatched: string[] = [];

    bindManifestActions(fake.port, async (commandId) => {
      dispatched.push(commandId);
      return { ok: true };
    });

    expect([...fake.handlers.keys()]).toEqual(
      allCommands.flatMap((command) =>
        command.actionId === undefined ? [] : [command.actionId]),
    );
    await fake.handlers.get("OpenAnalystLens")?.();
    await fake.handlers.get("FormatInteger")?.();
    await fake.handlers.get("OpenQuickTracePrecedents")?.();
    await fake.handlers.get("OpenQuickTraceDependents")?.();
    expect(fake.paneOpens).toBe(1);
    expect(dispatched).toEqual([
      "format.integer",
      "audit.quickTrace.precedents",
      "audit.quickTrace.dependents",
    ]);
  });

  it("reports conflicts instead of silently replacing an occupied shortcut", async () => {
    const fake = createActionPort(true);

    const result = await updateShortcuts(fake.port, {
      OpenAnalystLens: "Command+Option+L",
    });

    expect(result).toEqual({
      conflicts: [
        { actionId: "OpenAnalystLens", shortcut: "Command+Option+L" },
      ],
      ok: false,
    });
    expect(fake.replacements).toEqual([]);
  });

  it("replaces occupied shortcuts only after explicit conflict confirmation", async () => {
    const fake = createActionPort(true);
    const requested = { FormatInteger: "Ctrl+Shift+1" };

    await expect(updateShortcuts(fake.port, requested, {
      replaceConflicts: true,
    })).resolves.toEqual({
      current: { OpenAnalystLens: "Command+Option+L" },
      ok: true,
    });
    expect(fake.replacements).toEqual([requested]);
  });

  it("replaces only conflict-free shortcuts and reads them back", async () => {
    const fake = createActionPort();
    const requested = { OpenAnalystLens: "Command+Option+L" };

    await expect(updateShortcuts(fake.port, requested)).resolves.toEqual({
      current: requested,
      ok: true,
    });
    expect(fake.replacements).toEqual([requested]);
  });
});
