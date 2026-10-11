import { useState } from "react";
import { allCommands, type CommandDefinition } from "../domain/commands/commandRegistry";
import type {
  ShortcutUpdateOptions,
  ShortcutUpdateResult,
} from "../runtime/actionBindings";
import { ModalPanel } from "./ModalPanel";

type SettingsProps = {
  onClose(): void;
  open: boolean;
  updateShortcuts?: ((options?: ShortcutUpdateOptions) => Promise<ShortcutUpdateResult>) | undefined;
};

type ShortcutRow = Readonly<{ label: string; shortcut: string }>;

const toRow = (command: CommandDefinition): ShortcutRow | null =>
  command.macShortcut === undefined
    ? null
    : { label: command.label, shortcut: command.macShortcut };

const shortcutRows = (predicate: (command: CommandDefinition) => boolean): ShortcutRow[] =>
  allCommands.flatMap((command) =>
    predicate(command) ? (toRow(command) === null ? [] : [toRow(command) as ShortcutRow]) : []);

/** Audit + pane access — the shortcuts a daily user touches most. */
const coreIds = new Set([
  "system.openAnalystLens",
  "audit.quickTrace.precedents",
  "audit.quickTrace.dependents",
]);

/** High-frequency number formats worth committing to muscle memory first. */
const everydayFormatIds = new Set([
  "format.integer",
  "format.percent.oneDecimal",
  "format.currency.integer",
  "format.multiple.oneDecimal",
  "format.date",
]);

const ShortcutTable = ({ rows }: { rows: ShortcutRow[] }) => (
  <dl className="shortcut-map">
    {rows.map((row) => (
      <div key={row.label}>
        <dt>{row.label}</dt>
        <dd><kbd className="command-shortcut">{row.shortcut}</kbd></dd>
      </div>
    ))}
  </dl>
);

export function Settings({ onClose, open, updateShortcuts }: SettingsProps) {
  const [shortcutResult, setShortcutResult] = useState<ShortcutUpdateResult | null>(null);
  const [shortcutError, setShortcutError] = useState<string | null>(null);
  const [updating, setUpdating] = useState(false);
  const applyShortcuts = async (replaceConflicts = false) => {
    if (updateShortcuts === undefined) return;
    setUpdating(true);
    setShortcutError(null);
    try {
      setShortcutResult(await updateShortcuts({ replaceConflicts }));
    } catch (error) {
      setShortcutError(error instanceof Error ? error.message : String(error));
    } finally {
      setUpdating(false);
    }
  };
  if (!open) return null;
  return (
    <ModalPanel labelledBy="settings-title" onClose={onClose}>
      <div className="overlay-heading">
        <h2 id="settings-title">Inspector settings</h2>
        <button aria-label="Close settings" onClick={onClose} type="button">Close</button>
      </div>
      <p>Formula analysis is local. Workbook values, formulas, graph state, and scan results are not persisted.</p>
      <fieldset>
        <legend>Default safe constants</legend>
        <label><input defaultChecked type="checkbox" /> 0</label>
        <label><input defaultChecked type="checkbox" /> 1</label>
        <label><input defaultChecked type="checkbox" /> −1</label>
      </fieldset>
      <section aria-labelledby="shortcut-settings-title" className="shortcut-settings">
        <h3 id="shortcut-settings-title">Macabacus-compatible Mac defaults</h3>
        <p>
          Macabacus Control chords stay on Control, while Alt becomes Option. Commands that
          Macabacus exposes only as a cycle or punctuation use an explicit mnemonic fallback.
        </p>
        <p>
          Macabacus used Control+Shift+[ and Control+Shift+]. Excel add-in shortcuts on this Mac
          reject terminal punctuation, so Analyst Lens uses Control+Option+P for precedents and
          Control+Option+D for dependents.
        </p>
        {updateShortcuts === undefined ? null : (
          <button
            disabled={updating}
            onClick={() => void applyShortcuts()}
            type="button"
          >Apply Macabacus defaults</button>
        )}
        {shortcutResult?.ok === false ? (
          <div className="shortcut-conflicts" role="alert">
            <p>These shortcuts are already in use:</p>
            <ul>
              {shortcutResult.conflicts.map((conflict) => (
                <li key={`${conflict.actionId}-${conflict.shortcut}`}>{conflict.shortcut}</li>
              ))}
            </ul>
            <button
              disabled={updating}
              onClick={() => void applyShortcuts(true)}
              type="button"
            >Use Analyst Lens anyway</button>
          </div>
        ) : null}
        {shortcutResult?.ok === true ? (
          <p role="status">
            {Object.keys(shortcutResult.current).length} shortcut
            {Object.keys(shortcutResult.current).length === 1 ? " is" : "s are"} active.
          </p>
        ) : null}
        {shortcutError === null ? null : <p role="alert">{shortcutError}</p>}

        <div className="shortcut-tier">
          <h4>Core — pane &amp; trace</h4>
          <ShortcutTable rows={shortcutRows((command) => coreIds.has(command.id))} />
        </div>
        <div className="shortcut-tier">
          <h4>Everyday number formats</h4>
          <ShortcutTable rows={shortcutRows((command) => everydayFormatIds.has(command.id))} />
        </div>
        <div className="shortcut-tier">
          <h4>Colors, decimals &amp; borders</h4>
          <ShortcutTable rows={shortcutRows((command) =>
            command.macShortcut !== undefined &&
            command.category === "format" &&
            !everydayFormatIds.has(command.id))}
          />
        </div>
      </section>
    </ModalPanel>
  );
}
