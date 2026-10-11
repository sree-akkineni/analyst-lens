import { getMacShortcutDefaults } from "../../domain/commands/commandRegistry";
import type {
  OfficeActionPort,
  ShortcutUpdateOptions,
  ShortcutUpdateResult,
} from "../../runtime/actionBindings";
import type { CommandId, CommandResult } from "../../shared/types";
import { bindManifestActions, updateShortcuts } from "../../runtime/actionBindings";

export const officeReady = () => Office.onReady();

export const createOfficeActionPort = (): OfficeActionPort => ({
  areShortcutsInUse: (shortcuts) =>
    Office.actions.areShortcutsInUse(shortcuts),
  associate(actionId, handler) {
    Office.actions.associate(actionId, handler);
  },
  getShortcuts: () => Office.actions.getShortcuts(),
  replaceShortcuts: (shortcuts) => Office.actions.replaceShortcuts(shortcuts),
  showTaskpane: () => Office.addin.showAsTaskpane(),
});

export const initializeOfficeActionBindings = async (
  executeCommand: (commandId: CommandId) => Promise<CommandResult>,
) => {
  await officeReady();
  const port = createOfficeActionPort();
  bindManifestActions(port, executeCommand);
  return port;
};

export const applyMacShortcutDefaults = async (
  options: ShortcutUpdateOptions = {},
): Promise<ShortcutUpdateResult> => {
  await officeReady();
  const replacements = Object.fromEntries(
    Object.entries(getMacShortcutDefaults()).flatMap(([actionId, shortcut]) =>
      actionId === "OpenAnalystLens"
        ? []
        : [[actionId, shortcut.replace("Option", "Alt")]],
    ),
  );
  return updateShortcuts(
    createOfficeActionPort(),
    replacements,
    options,
  );
};
