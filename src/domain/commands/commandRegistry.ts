import type { CommandId } from "../../shared/types";

export type CommandDefinition = {
  actionId?: string;
  availability?: { blockedBy: "grouping-acceptance" } | undefined;
  category: "format" | "structure" | "navigate" | "audit" | "system";
  effect: "read" | "verified-write" | "structural-write" | "view-change";
  id: CommandId;
  label: string;
  macShortcut?: string;
  manualInverseId?: CommandId;
  requiredCapabilities: string[];
  revertPolicy: "automatic" | "manual-inverse" | "none";
};

const formatCommand = (
  id: CommandId,
  label: string,
  actionId: string,
  macShortcut: string,
): CommandDefinition => ({
  actionId,
  category: "format",
  effect: "verified-write",
  id,
  label,
  macShortcut,
  requiredCapabilities: ["formatCellProperties"],
  revertPolicy: "automatic",
});

const structureCommand = (
  id: CommandId,
  label: string,
  manualInverseId: CommandId,
): CommandDefinition => ({
  availability: undefined,
  category: "structure",
  effect: "structural-write",
  id,
  label,
  manualInverseId,
  requiredCapabilities: ["outlineGroups"],
  revertPolicy: "manual-inverse",
});

export const allCommands: readonly CommandDefinition[] = [
  {
    actionId: "OpenAnalystLens",
    category: "system",
    effect: "read",
    id: "system.openAnalystLens",
    label: "Open Analyst Lens",
    macShortcut: "Command+Option+L",
    requiredCapabilities: ["sharedRuntime"],
    revertPolicy: "none",
  },
  {
    actionId: "OpenQuickTracePrecedents",
    category: "audit",
    effect: "read",
    id: "audit.quickTrace.precedents",
    label: "Quick Trace In (precedents)",
    macShortcut: "Ctrl+Option+P",
    requiredCapabilities: ["DialogApi 1.2", "directPrecedents"],
    revertPolicy: "none",
  },
  {
    actionId: "OpenQuickTraceDependents",
    category: "audit",
    effect: "read",
    id: "audit.quickTrace.dependents",
    label: "Quick Trace Out (dependents)",
    macShortcut: "Ctrl+Option+D",
    requiredCapabilities: ["DialogApi 1.2", "directDependents"],
    revertPolicy: "none",
  },
  {
    category: "system",
    effect: "verified-write",
    id: "system.revertLastAction",
    label: "Revert last Analyst Lens action",
    requiredCapabilities: [],
    revertPolicy: "none",
  },
  {
    category: "system",
    effect: "read",
    id: "system.acknowledgeSafetyLock",
    label: "Acknowledge safety lock",
    requiredCapabilities: [],
    revertPolicy: "none",
  },
  formatCommand("format.input", "Input blue", "FormatInputBlue", "Ctrl+Option+Shift+B"),
  formatCommand("format.formula", "Formula black", "FormatFormulaBlack", "Ctrl+Option+Shift+K"),
  formatCommand("format.link", "Linked formula green", "FormatLinkedGreen", "Ctrl+Option+Shift+G"),
  formatCommand("format.warning", "Warning red", "FormatWarningRed", "Ctrl+Option+Shift+R"),
  formatCommand("format.integer", "Integer", "FormatInteger", "Ctrl+Shift+1"),
  formatCommand("format.decimal.one", "One decimal", "FormatOneDecimal", "Ctrl+Option+Shift+1"),
  formatCommand("format.decimal.two", "Two decimals", "FormatTwoDecimals", "Ctrl+Option+Shift+2"),
  formatCommand("format.percent.oneDecimal", "Percentage", "FormatPercentage", "Ctrl+Shift+5"),
  formatCommand("format.currency.integer", "Currency", "FormatCurrency", "Ctrl+Shift+4"),
  formatCommand("format.multiple.oneDecimal", "Multiple", "FormatMultiple", "Ctrl+Shift+8"),
  formatCommand("format.date", "Date", "FormatDate", "Ctrl+Shift+2"),
  formatCommand("format.decimals.increase", "Increase decimals", "IncreaseDecimals", "Ctrl+Option+Shift+I"),
  formatCommand("format.decimals.decrease", "Decrease decimals", "DecreaseDecimals", "Ctrl+Option+Shift+E"),
  formatCommand("format.border.top", "Top border", "FormatTopBorder", "Ctrl+Option+Shift+Up"),
  formatCommand("format.border.bottom", "Bottom border", "FormatBottomBorder", "Ctrl+Option+Shift+Down"),
  formatCommand("format.border.doubleBottom", "Double bottom border", "FormatDoubleBottomBorder", "Ctrl+Option+Shift+D"),
  formatCommand("format.border.clear", "Clear borders", "ClearBorders", "Ctrl+Option+Shift+N"),
  formatCommand("format.clear", "Clear formats", "ClearFormats", "Ctrl+Option+Shift+C"),
  structureCommand("structure.groupRows", "Group rows", "structure.ungroupRows"),
  structureCommand("structure.ungroupRows", "Ungroup rows", "structure.groupRows"),
  structureCommand("structure.groupColumns", "Group columns", "structure.ungroupColumns"),
  structureCommand("structure.ungroupColumns", "Ungroup columns", "structure.groupColumns"),
  {
    availability: undefined,
    category: "structure",
    effect: "view-change",
    id: "structure.showDetails",
    label: "Open All Groups",
    requiredCapabilities: ["outlineGroups"],
    revertPolicy: "none",
  },
  {
    availability: undefined,
    category: "structure",
    effect: "view-change",
    id: "structure.hideDetails",
    label: "Close All Groups",
    requiredCapabilities: ["outlineGroups"],
    revertPolicy: "none",
  },
] as const;

export const getCommand = (id: CommandId): CommandDefinition | undefined =>
  allCommands.find((command) => command.id === id);

export const getMacShortcutDefaults = (): Record<string, string> =>
  Object.fromEntries(allCommands.flatMap((command) =>
    command.actionId === undefined || command.macShortcut === undefined
      ? []
      : [[command.actionId, command.macShortcut]],
  ));
