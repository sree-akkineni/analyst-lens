import { useEffect, useRef, useState } from "react";
import type {
  InspectorIntent,
  InspectorOverlay,
  InspectorState,
} from "../runtime/controller";
import type {
  ShortcutUpdateOptions,
  ShortcutUpdateResult,
} from "../runtime/actionBindings";
import {
  canonicalCellId,
  type CommandId,
  type CommandResult,
  type DependencyNode,
} from "../shared/types";
import { CommandPalette } from "./CommandPalette";
import { DependencyGraph } from "./DependencyGraph";
import { ModalPanel } from "./ModalPanel";
import { Settings } from "./Settings";

export type OpenQuickTrace = (direction: "precedent" | "dependent") => Promise<void>;

type InspectorProps = {
  dispatch(intent: InspectorIntent): void;
  openQuickTrace?: OpenQuickTrace | undefined;
  quickTraceSupported?: boolean | undefined;
  runCommand(id: CommandId): Promise<CommandResult>;
  state: InspectorState;
  updateShortcuts?: ((options?: ShortcutUpdateOptions) => Promise<ShortcutUpdateResult>) | undefined;
};

const cellKindLabel = {
  blank: "Blank cell",
  error: "Excel error",
  formula: "Formula",
  hardcode: "Hardcoded value",
  multi: "Multiple selection",
} as const;

const romanNumerals = ["i.", "ii.", "iii.", "iv.", "v.", "vi.", "vii.", "viii.", "ix.", "x."] as const;

const glossFor = (node: DependencyNode): string => {
  if (node.unresolved === "external") return "external workbook";
  if (node.unresolved === "dynamic") return "dynamic reference, unresolved";
  return "";
};

const AnnotationRow = ({
  dispatch,
  index,
  node,
}: {
  dispatch(intent: InspectorIntent): void;
  index: number;
  node: DependencyNode;
}) => {
  const gloss = glossFor(node);
  const numeral = romanNumerals[index] ?? `${index + 1}.`;
  const badges: string[] = [];
  if (node.unresolved === "external") badges.push("ext");
  if (node.unresolved === "dynamic") badges.push("dyn");
  return (
    <li className="annotation">
      <button
        aria-label={`${node.ref.sheetName} ${node.ref.address}`}
        className="annotation-button"
        onClick={() => dispatch({ ref: node.ref, type: "navigate" })}
        type="button"
      >
        <span aria-hidden="true" className="annotation-num">{numeral}</span>
        <span className="annotation-body">
          <span className="annotation-ref">{node.ref.sheetName} {node.ref.address}</span>
          {gloss === "" ? null : <span className="annotation-gloss">{gloss}</span>}
        </span>
        <span className="annotation-val">
          {node.value === undefined ? "—" : String(node.value)}
          {badges.map((badge) => (
            <span className={`badge ${badge}`} key={badge}>{badge}</span>
          ))}
        </span>
      </button>
    </li>
  );
};

export function Inspector({
  dispatch,
  openQuickTrace,
  quickTraceSupported,
  runCommand,
  state,
  updateShortcuts,
}: InspectorProps) {
  const graphInvoker = useRef<HTMLButtonElement>(null);
  const paletteInvoker = useRef<HTMLButtonElement>(null);
  const settingsInvoker = useRef<HTMLButtonElement>(null);
  const [quickTracePending, setQuickTracePending] = useState(false);
  const [quickTraceError, setQuickTraceError] = useState<string | null>(null);
  const canOpenQuickTrace = quickTraceSupported === true && openQuickTrace !== undefined;
  const overlayActive = Object.values(state.overlays).some(Boolean);
  useEffect(() => {
    const openPalette = (event: KeyboardEvent) => {
      const target = event.target;
      const editing = target instanceof HTMLElement && (
        target.isContentEditable || ["INPUT", "SELECT", "TEXTAREA"].includes(target.tagName)
      );
      if (
        overlayActive ||
        editing ||
        !event.metaKey ||
        event.altKey ||
        event.ctrlKey ||
        event.key.toLowerCase() !== "k"
      ) return;
      event.preventDefault();
      dispatch({ overlay: "palette", type: "openOverlay" });
    };
    window.addEventListener("keydown", openPalette);
    return () => window.removeEventListener("keydown", openPalette);
  }, [dispatch, overlayActive]);
  const launchQuickTrace = async (direction: "precedent" | "dependent") => {
    if (!canOpenQuickTrace || quickTracePending) return;
    setQuickTracePending(true);
    setQuickTraceError(null);
    try {
      await openQuickTrace(direction);
    } catch (error) {
      setQuickTraceError(error instanceof Error ? error.message : String(error));
    } finally {
      setQuickTracePending(false);
    }
  };
  const closeOverlay = (overlay: InspectorOverlay, invoker: React.RefObject<HTMLButtonElement | null>) => {
    dispatch({ overlay, type: "closeOverlay" });
    invoker.current?.focus();
  };

  if (state.view.kind === "loading") {
    return <section className="inspector-state" role="status">Loading selection</section>;
  }
  if (state.view.kind === "unsupported" || state.view.kind === "failure") {
    return <section className="inspector-state"><p role="alert">{state.view.message}</p></section>;
  }

  const { cellKind, facts } = state.view;
  if (cellKind === "multi") {
    return (
      <section className="inspector-state">
        <p className="section-label">Selection</p>
        <h2>{facts.rowCount} × {facts.columnCount} selection</h2>
        <p>Choose one cell to inspect formula context and direct dependencies.</p>
      </section>
    );
  }
  const breadcrumbPast = state.history?.past.filter((ref) =>
    canonicalCellId(ref) !== canonicalCellId(state.history?.origin ?? ref)) ?? [];
  const currentIsOrigin = state.history !== null &&
    canonicalCellId(state.history.current) === canonicalCellId(state.history.origin);

  return (
    <section aria-labelledby="inspector-title" className="inspector">
      <div className="inspector-hero">
        <div className="hero-eyebrow">
          <span aria-hidden="true" className="hero-live-dot" />
          <span className="hero-live-label">Live</span>
          <span className="hero-kind">· {cellKindLabel[cellKind]}</span>
        </div>
        <h2 className="hero-addr" id="inspector-title">
          <span className="sheet">{facts.sheetName}</span> {facts.address}
        </h2>
        <p className="hero-gloss">selected cell</p>
        <div className="hero-value-row">
          <span className="hero-value">{facts.value === null ? "Blank" : String(facts.value)}</span>
          <span className="hero-format">{facts.numberFormat}</span>
          <button
            aria-label="Pin key output"
            className="hero-pin"
            onClick={() => dispatch({ type: "pinCurrent" })}
            type="button"
          >Pin output</button>
        </div>
      </div>

      <div className="inspector-toolbar">
        <nav aria-label="Navigation history" className="history-controls">
          <button
            disabled={(state.history?.past.length ?? 0) === 0}
            onClick={() => dispatch({ type: "back" })}
            type="button"
          >Back</button>
          <button
            disabled={(state.history?.future.length ?? 0) === 0}
            onClick={() => dispatch({ type: "forward" })}
            type="button"
          >Forward</button>
          <button
            disabled={state.history === null}
            onClick={() => dispatch({ type: "returnToOrigin" })}
            type="button"
          >Return to origin</button>
        </nav>

        <div className="quick-trace-openers">
          <button
            aria-label="Open Quick Trace In (Ctrl+Option+P)"
            disabled={!canOpenQuickTrace || quickTracePending}
            onClick={() => void launchQuickTrace("precedent")}
            type="button"
          >Quick Trace In <kbd>Ctrl+Option+P</kbd></button>
          <button
            aria-label="Open Quick Trace Out (Ctrl+Option+D)"
            disabled={!canOpenQuickTrace || quickTracePending}
            onClick={() => void launchQuickTrace("dependent")}
            type="button"
          >Quick Trace Out <kbd>Ctrl+Option+D</kbd></button>
        </div>
        {canOpenQuickTrace ? null : (
          <p className="quick-trace-availability">Requires DialogApi 1.2 for a floating Office dialog.</p>
        )}
        {quickTraceError === null ? null : <p className="quick-trace-error" role="alert">{quickTraceError}</p>}

        {state.history === null ? null : (
          <ol aria-label="Navigation breadcrumb" className="navigation-breadcrumb">
            <li>
              <button
                aria-label={`Go to origin ${state.history.origin.sheetName} ${state.history.origin.address}`}
                onClick={() => dispatch({ ref: state.history?.origin ?? facts, type: "navigate" })}
                type="button"
              >Origin · {state.history.origin.sheetName} {state.history.origin.address}</button>
            </li>
            {breadcrumbPast.slice(-3).map((ref) => (
              <li key={`${ref.sheetId}-${ref.address}`}>
                <button onClick={() => dispatch({ ref, type: "navigate" })} type="button">
                  {ref.sheetName} {ref.address}
                </button>
              </li>
            ))}
            {currentIsOrigin ? null : (
              <li aria-current="page">{state.history.current.sheetName} {state.history.current.address}</li>
            )}
          </ol>
        )}
      </div>

      <div className="inspector-cockpit">
        <div className="inspector-detail">
      {facts.formula === null ? null : (
        <section aria-labelledby="formula-title" className="formula-panel">
          <h3 id="formula-title">Formula</h3>
          <code>
            =
            {state.tokens.map((token) =>
              token.kind === "reference" ? (
                <button
                  className="formula-reference"
                  key={`${token.start}-${token.end}`}
                  onClick={() => dispatch({ text: token.text, type: "navigateReference" })}
                  type="button"
                >{token.text}</button>
              ) : <span className={`formula-${token.kind}`} key={`${token.start}-${token.end}`}>{token.text}</span>)}
          </code>
        </section>
      )}

      <section aria-labelledby="risk-title" className="risk-section">
        <h3 id="risk-title">Noted</h3>
        <ul className="risk-list">
          {state.risks.length === 0 ? <li className="risk-none">None detected lexically</li> : null}
          {state.risks.map((risk) => (
            <li key={`${risk.kind}-${risk.start}`} title={risk.message}>{risk.kind}</li>
          ))}
        </ul>
      </section>

      <section aria-labelledby="sources-title" className="sources-section">
        <div className="sources-head">
          <h3 id="sources-title">Sources</h3>
          <span className="sources-count">
            {state.precedents.length === 0 ? "none" : `${state.precedents.length}`}
          </span>
        </div>
        {state.precedents.length === 0 ? (
          <p className="annotation-empty">No direct precedents reported by Excel.</p>
        ) : (
          <ul aria-label="Direct precedents" className="annotation-list">
            {state.precedents.map((node, index) => (
              <AnnotationRow dispatch={dispatch} index={index} key={node.id} node={node} />
            ))}
          </ul>
        )}
      </section>

      <section aria-labelledby="neighbor-title">
        <h3 id="neighbor-title">Neighbor comparison</h3>
        {state.neighbors.length === 0 ? <p className="muted-row">No adjacent comparison loaded.</p> : null}
        <ul className="neighbor-list">
          {state.neighbors.map((neighbor) => (
            <li key={neighbor.label}>
              <strong>{neighbor.label}</strong>
              <span>{neighbor.status === "same"
                ? "Matches selected formula"
                : neighbor.status === "different"
                  ? "Different pattern"
                  : "No formula"}</span>
              {neighbor.normalizedFormula === null ? null : <code>{neighbor.normalizedFormula}</code>}
            </li>
          ))}
        </ul>
      </section>

      <section aria-labelledby="dependents-title" className="sources-section">
        <div className="sources-head">
          <h3 id="dependents-title">Dependents</h3>
          <span className="sources-count">
            {state.dependents.length === 0 ? "none" : `${state.dependents.length}`}
          </span>
        </div>
        {state.dependents.length === 0 ? (
          <p className="annotation-empty">No direct dependents reported by Excel.</p>
        ) : (
          <ul aria-label="Direct dependents" className="annotation-list">
            {state.dependents.map((node, index) => (
              <AnnotationRow dispatch={dispatch} index={index} key={node.id} node={node} />
            ))}
          </ul>
        )}
      </section>
        </div>

      {state.graph === null ? null : (
        <div className="compact-dependency-map">
          <h3>Compact map</h3>
          <DependencyGraph capped={state.capped} dispatch={dispatch} state={state.graph} />
        </div>
      )}
      </div>

      <div className="command-bar">
        <button
          aria-label="Open commands for this cell, Command-K"
          className="command-launcher"
          onClick={() => dispatch({ overlay: "palette", type: "openOverlay" })}
          ref={paletteInvoker}
          type="button"
        >
          <span>Open commands for this cell</span>
          <kbd>⌘K</kbd>
        </button>
        <button
          aria-label="Open dependency map"
          className="cmd-icon-btn"
          onClick={() => dispatch({ overlay: "graph", type: "openOverlay" })}
          ref={graphInvoker}
          type="button"
        >◎</button>
        <button
          aria-label="Settings"
          className="cmd-icon-btn"
          onClick={() => dispatch({ overlay: "settings", type: "openOverlay" })}
          ref={settingsInvoker}
          type="button"
        >⋯</button>
      </div>

      {state.notice === null ? null : <p className="inspector-notice" role="status">{state.notice}</p>}

      {state.overlays.graph && state.graph !== null ? (
        <ModalPanel
          labelledBy="full-graph-title"
          onClose={() => closeOverlay("graph", graphInvoker)}
        >
          <div className="overlay-heading">
            <h2 id="full-graph-title">Dependency map</h2>
            <div className="overlay-actions">
              <button
                aria-label="Return to origin from dependency map"
                onClick={() => dispatch({ type: "returnToOrigin" })}
                type="button"
              >Return to origin</button>
              <button
                aria-label="Close dependency map"
                onClick={() => closeOverlay("graph", graphInvoker)}
                type="button"
              >Close</button>
            </div>
          </div>
          <DependencyGraph capped={state.capped} dispatch={dispatch} state={state.graph} />
        </ModalPanel>
      ) : null}
      <CommandPalette
        onClose={() => closeOverlay("palette", paletteInvoker)}
        open={state.overlays.palette}
        runCommand={runCommand}
      />
      <Settings
        onClose={() => closeOverlay("settings", settingsInvoker)}
        open={state.overlays.settings}
        updateShortcuts={updateShortcuts}
      />
    </section>
  );
}
