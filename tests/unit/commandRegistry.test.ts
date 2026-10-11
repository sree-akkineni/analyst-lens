import { describe, expect, it } from "vitest";
import {
  allCommands,
  getCommand,
  getMacShortcutDefaults,
} from "../../src/domain/commands/commandRegistry";

describe("command registry", () => {
  it("classifies every command with an honest workbook effect", () => {
    expect(
      allCommands.every(
        (command) =>
          command.effect === "read" ||
          command.effect === "verified-write" ||
          command.effect === "structural-write" ||
          command.effect === "view-change",
      ),
    ).toBe(true);
  });

  it("keeps command and action identifiers unique", () => {
    const commandIds = allCommands.map((command) => command.id);
    const actionIds = allCommands.flatMap((command) =>
      command.actionId === undefined ? [] : [command.actionId],
    );

    expect(new Set(commandIds).size).toBe(commandIds.length);
    expect(new Set(actionIds).size).toBe(actionIds.length);
    expect(actionIds).toContain("OpenAnalystLens");
  });

  it("marks revert as a verified workbook write", () => {
    expect(
      allCommands.find((command) => command.id === "system.revertLastAction"),
    ).toMatchObject({ effect: "verified-write" });
  });

  it("registers safety-lock acknowledgement as a local system action", () => {
    expect(getCommand("system.acknowledgeSafetyLock")).toMatchObject({
      category: "system",
      effect: "read",
      id: "system.acknowledgeSafetyLock",
    });
  });

  it("registers every supported speed-layer format as a verified write", () => {
    const expected = [
      "format.input",
      "format.formula",
      "format.link",
      "format.warning",
      "format.integer",
      "format.decimal.one",
      "format.decimal.two",
      "format.percent.oneDecimal",
      "format.currency.integer",
      "format.multiple.oneDecimal",
      "format.date",
      "format.decimals.increase",
      "format.decimals.decrease",
      "format.border.top",
      "format.border.bottom",
      "format.border.doubleBottom",
      "format.border.clear",
      "format.clear",
    ];

    expect(expected.map((id) => getCommand(id))).toEqual(
      expected.map((id) => expect.objectContaining({ effect: "verified-write", id })),
    );
  });

  it("maps speed tools to Macabacus-compatible Mac shortcuts", () => {
    const expected = {
      "format.border.bottom": "Ctrl+Option+Shift+Down",
      "format.border.clear": "Ctrl+Option+Shift+N",
      "format.border.doubleBottom": "Ctrl+Option+Shift+D",
      "format.border.top": "Ctrl+Option+Shift+Up",
      "format.clear": "Ctrl+Option+Shift+C",
      "format.currency.integer": "Ctrl+Shift+4",
      "format.date": "Ctrl+Shift+2",
      "format.decimal.one": "Ctrl+Option+Shift+1",
      "format.decimal.two": "Ctrl+Option+Shift+2",
      "format.decimals.decrease": "Ctrl+Option+Shift+E",
      "format.decimals.increase": "Ctrl+Option+Shift+I",
      "format.formula": "Ctrl+Option+Shift+K",
      "format.input": "Ctrl+Option+Shift+B",
      "format.integer": "Ctrl+Shift+1",
      "format.link": "Ctrl+Option+Shift+G",
      "format.multiple.oneDecimal": "Ctrl+Shift+8",
      "format.percent.oneDecimal": "Ctrl+Shift+5",
      "format.warning": "Ctrl+Option+Shift+R",
    } as const;

    expect(Object.fromEntries(Object.keys(expected).map((id) => [
      id,
      getCommand(id)?.macShortcut,
    ]))).toEqual(expected);
    expect(Object.keys(expected).every((id) => getCommand(id)?.actionId !== undefined)).toBe(true);
    expect(getMacShortcutDefaults()).toMatchObject({
      FormatInteger: "Ctrl+Shift+1",
      FormatTopBorder: "Ctrl+Option+Shift+Up",
      OpenAnalystLens: "Command+Option+L",
      OpenQuickTraceDependents: "Ctrl+Option+D",
      OpenQuickTracePrecedents: "Ctrl+Option+P",
    });
  });

  it("registers Quick Trace actions as audit-only reads", () => {
    expect(getCommand("audit.quickTrace.precedents")).toMatchObject({
      actionId: "OpenQuickTracePrecedents",
      effect: "read",
      macShortcut: "Ctrl+Option+P",
    });
    expect(getCommand("audit.quickTrace.dependents")).toMatchObject({
      actionId: "OpenQuickTraceDependents",
      effect: "read",
      macShortcut: "Ctrl+Option+D",
    });
  });

  it("registers one-layer structural commands with explicit manual inverses", () => {
    const inversePairs = [
      ["structure.groupRows", "structure.ungroupRows"],
      ["structure.ungroupRows", "structure.groupRows"],
      ["structure.groupColumns", "structure.ungroupColumns"],
      ["structure.ungroupColumns", "structure.groupColumns"],
    ] as const;

    expect(inversePairs.map(([id]) => getCommand(id))).toEqual(
      inversePairs.map(([id, manualInverseId]) => expect.objectContaining({
        effect: "structural-write",
        id,
        manualInverseId,
        revertPolicy: "manual-inverse",
      })),
    );
  });

  it("renames detail commands as history-free universal outline visibility", () => {
    const visibilityIds = [
      "structure.showDetails",
      "structure.hideDetails",
    ];

    expect(visibilityIds.map((id) => getCommand(id))).toEqual([
      expect.objectContaining({
        effect: "view-change",
        id: "structure.showDetails",
        label: "Open All Groups",
        revertPolicy: "none",
      }),
      expect.objectContaining({
        effect: "view-change",
        id: "structure.hideDetails",
        label: "Close All Groups",
        revertPolicy: "none",
      }),
    ]);
  });

  it("keeps every structure command available after the approved grouping-acceptance decision", () => {
    const structureIds = [
      "structure.groupRows",
      "structure.ungroupRows",
      "structure.groupColumns",
      "structure.ungroupColumns",
      "structure.showDetails",
      "structure.hideDetails",
    ];

    expect(structureIds.map((id) => getCommand(id))).toEqual(
      structureIds.map((id) => expect.objectContaining({
        availability: undefined,
        id,
      })),
    );
  });
});
