import { useState } from "react";
import { getCommand } from "../domain/commands/commandRegistry";
import type { CommandId, CommandResult } from "../shared/types";

type SpeedLayerProps = {
  runCommand(id: CommandId): Promise<CommandResult>;
};

const groups = [
  {
    id: "model-colors",
    label: "Model colors",
    commands: ["format.input", "format.formula", "format.link", "format.warning"],
  },
  {
    id: "number-formats",
    label: "Number formats",
    commands: [
      "format.integer",
      "format.decimal.one",
      "format.decimal.two",
      "format.percent.oneDecimal",
      "format.currency.integer",
      "format.multiple.oneDecimal",
      "format.date",
      "format.decimals.increase",
      "format.decimals.decrease",
    ],
  },
  {
    id: "borders-clear",
    label: "Borders & clear",
    commands: [
      "format.border.top",
      "format.border.bottom",
      "format.border.doubleBottom",
      "format.border.clear",
      "format.clear",
    ],
  },
] as const;

const structureCommands = [
  "structure.groupRows",
  "structure.ungroupRows",
  "structure.groupColumns",
  "structure.ungroupColumns",
  "structure.showDetails",
  "structure.hideDetails",
] as const;

const commandLabel = (id: CommandId): string => getCommand(id)?.label ?? id;
const commandShortcut = (id: CommandId): string | undefined => getCommand(id)?.macShortcut;

const successMessage = (id: CommandId): string => {
  if (id === "system.revertLastAction") {
    return "Last Analyst Lens action reverted.";
  }
  if (id === "system.acknowledgeSafetyLock") {
    return "Safety lock acknowledged. Writes are available again.";
  }
  if (id === "structure.showDetails") {
    return "All groups opened. Group definitions were not changed.";
  }
  if (id === "structure.hideDetails") {
    return "All groups closed. Group definitions were not changed.";
  }
  const command = getCommand(id);
  if (command?.effect === "structural-write" && command.manualInverseId !== undefined) {
    return `${command.label} completed one outline layer. To reverse it, select the affected range and run ${commandLabel(command.manualInverseId)} once. Automatic Revert is unavailable for structural changes.`;
  }
  return `${commandLabel(id)} applied. Revert is available.`;
};

export function SpeedLayer({ runCommand }: SpeedLayerProps) {
  const [pending, setPending] = useState<CommandId | null>(null);
  const [result, setResult] = useState<
    { code?: string; kind: "error" | "success"; message: string } | null
  >(null);

  const execute = async (id: CommandId) => {
    setPending(id);
    setResult(null);
    try {
      const commandResult = await runCommand(id);
      setResult(commandResult.ok
        ? {
            kind: "success",
            message: successMessage(id),
          }
        : { code: commandResult.code, kind: "error", message: commandResult.message });
    } catch {
      setResult({
        code: "OFFICE_BOOTSTRAP_FAILED",
        kind: "error",
        message: "Analyst Lens could not connect to Excel. Try the command again.",
      });
    } finally {
      setPending(null);
    }
  };

  return (
    <div className="speed-layer">
      <div className="speed-layer-toolbar">
        <p className="shortcut-limitation">
          Macabacus-compatible Mac shortcuts are active. Press Command+Option+L to reopen Analyst Lens.
        </p>
        <button
          className="revert-button"
          disabled={pending !== null}
          onClick={() => void execute("system.revertLastAction")}
          type="button"
        >
          Revert last Analyst Lens action
        </button>
      </div>

      <div className="command-groups">
        {groups.map((group) => (
          <section aria-labelledby={group.id} className="command-group" key={group.id}>
            <h3 id={group.id}>{group.label}</h3>
            <div className="command-list">
              {group.commands.map((id) => (
                <button
                  aria-label={`${commandLabel(id)} ${commandShortcut(id) ?? ""}`.trim()}
                  className="command-button"
                  disabled={pending !== null}
                  key={id}
                  onClick={() => void execute(id)}
                  type="button"
                >
                  <span>{commandLabel(id)}</span>
                  <kbd className="command-shortcut">{commandShortcut(id)}</kbd>
                </button>
              ))}
            </div>
          </section>
        ))}

        <section aria-labelledby="structure-tools" className="command-group">
          <h3 id="structure-tools">Structure</h3>
          <p className="group-note">
            Each Group/Ungroup command adds or removes exactly one outline layer on the rows or columns intersecting the selection. Repeated commands handle nested layers. Automatic Revert is unavailable for structural changes; use the opposite command once. Open/Close All changes visibility only.
          </p>
          <div className="command-list">
            {structureCommands.map((id) => (
              <button
                className="command-button"
                disabled={pending !== null}
                key={id}
                onClick={() => void execute(id)}
                type="button"
              >
                <span>{commandLabel(id)}</span>
              </button>
            ))}
          </div>
        </section>
      </div>

      {result?.kind === "success" ? (
        <p className="command-result success-result" role="status">{result.message}</p>
      ) : null}
      {result?.kind === "error" ? (
        <div className="command-result error-result" role="alert">
          <p>{result.message}</p>
          {result.code === "SAFETY_LOCKED" ||
          result.code === "REVERT_VERIFY_FAILED" ||
          result.code?.startsWith("WRITE_VERIFY_FAILED") === true ? (
            <button
              className="safety-lock-action"
              disabled={pending !== null}
              onClick={() => void execute("system.acknowledgeSafetyLock")}
              type="button"
            >
              Acknowledge safety lock
            </button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
