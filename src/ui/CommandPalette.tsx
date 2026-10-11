import { useMemo, useState } from "react";
import { allCommands } from "../domain/commands/commandRegistry";
import type { CommandId, CommandResult } from "../shared/types";
import { ModalPanel } from "./ModalPanel";

export type ContextAction = Readonly<{
  id: string;
  label: string;
  availability: "ready" | "coming-soon";
  kind: "command" | "ai";
  description: string;
}>;

type CommandPaletteProps = {
  onClose(): void;
  open: boolean;
  runCommand(id: CommandId): Promise<CommandResult>;
};

export function CommandPalette({ onClose, open, runCommand }: CommandPaletteProps) {
  const [query, setQuery] = useState("");
  const [result, setResult] = useState<string | null>(null);
  const commands = useMemo(() => {
    const normalized = query.trim().toLocaleLowerCase();
    return allCommands.filter((command) =>
      command.id !== "system.openAnalystLens" &&
      (normalized.length === 0 || command.label.toLocaleLowerCase().includes(normalized)));
  }, [query]);
  if (!open) return null;

  return (
    <ModalPanel labelledBy="command-palette-title" onClose={onClose}>
      <div className="overlay-heading">
        <h2 id="command-palette-title">Command palette</h2>
        <button aria-label="Close command palette" onClick={onClose} type="button">Close</button>
      </div>
      <label>
        Find a command
        <input
          autoFocus
          onChange={(event) => setQuery(event.currentTarget.value)}
          type="search"
          value={query}
        />
      </label>
      <ul className="palette-list">
        {commands.map((command) => (
          <li key={command.id}>
            <button
              aria-label={command.macShortcut === undefined
                ? command.label
                : `${command.label} ${command.macShortcut}`}
              onClick={() => {
                void runCommand(command.id).then((commandResult) => {
                  setResult(commandResult.ok ? `${command.label} completed.` : commandResult.message);
                });
              }}
              type="button"
            >
              <span>{command.label}</span>
              {command.macShortcut === undefined
                ? null
                : <kbd className="command-shortcut">{command.macShortcut}</kbd>}
            </button>
          </li>
        ))}
      </ul>
      {result === null ? null : <p role="status">{result}</p>}
    </ModalPanel>
  );
}
