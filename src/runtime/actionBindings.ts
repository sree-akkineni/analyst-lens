import type { CommandId, CommandResult } from "../shared/types";
import { allCommands } from "../domain/commands/commandRegistry";

export type OfficeActionPort = {
  areShortcutsInUse(
    shortcuts: string[],
  ): Promise<Array<{ inUse: boolean; shortcut: string }>>;
  associate(actionId: string, handler: () => Promise<void>): void;
  getShortcuts(): Promise<Record<string, string | null>>;
  replaceShortcuts(shortcuts: Record<string, string>): Promise<void>;
  showTaskpane(): Promise<void>;
};

export type ShortcutConflict = { actionId: string; shortcut: string };
export type ShortcutUpdateOptions = { replaceConflicts?: boolean };
export type ShortcutUpdateResult =
  | { current: Record<string, string | null>; ok: true }
  | { conflicts: ShortcutConflict[]; ok: false };

export const bindManifestActions = (
  port: OfficeActionPort,
  executeCommand: (commandId: CommandId) => Promise<CommandResult>,
) => {
  for (const command of allCommands) {
    if (command.actionId === undefined) continue;
    port.associate(command.actionId, async () => {
      if (command.id === "system.openAnalystLens") {
        await port.showTaskpane();
        return;
      }
      await executeCommand(command.id);
    });
  }
};

export const updateShortcuts = async (
  port: OfficeActionPort,
  requested: Record<string, string>,
  { replaceConflicts = false }: ShortcutUpdateOptions = {},
): Promise<ShortcutUpdateResult> => {
  const entries = Object.entries(requested);
  const inUse = await port.areShortcutsInUse(
    entries.map(([, shortcut]) => shortcut),
  );
  const occupied = new Set(
    inUse.filter((result) => result.inUse).map((result) => result.shortcut),
  );
  const conflicts = entries
    .filter(([, shortcut]) => occupied.has(shortcut))
    .map(([actionId, shortcut]) => ({ actionId, shortcut }));
  if (conflicts.length > 0 && !replaceConflicts) {
    return { conflicts, ok: false };
  }

  await port.replaceShortcuts(requested);
  const current = await port.getShortcuts();
  const unassigned = entries
    .filter(([actionId]) => current[actionId] === null)
    .map(([actionId, shortcut]) => ({ actionId, shortcut }));
  return unassigned.length > 0
    ? { conflicts: unassigned, ok: false }
    : { current, ok: true };
};
