import { useEffect, useMemo, useRef, useState } from "react";
import type { QuickTraceUserIntent } from "../../runtime/quickTraceController";
import type { QuickTraceSnapshot } from "../../shared/quickTraceProtocol";

export type QuickTraceAppProps = Readonly<{
  send(intent: QuickTraceUserIntent): void;
  state: QuickTraceSnapshot;
}>;

const WINDOW_SIZE = 72;
const ROW_HEIGHT = 30;
const OVERSCAN = 6;

const fullReference = (ref: QuickTraceSnapshot["root"]): string =>
  `${ref.sheetName} ${ref.address}`;

const rowDomId = (rowId: string): string =>
  `quick-trace-row-${encodeURIComponent(rowId)}`;

const issueLabel = (issue: QuickTraceSnapshot["rows"][number]["issue"]): string | null => {
  if (issue === null) return null;
  return issue[0]?.toUpperCase() + issue.slice(1);
};

const initialWindowStart = (state: QuickTraceSnapshot): number => {
  const selected = state.rows.findIndex((row) => row.rowId === state.selectedRowId);
  if (state.rows.length <= WINDOW_SIZE || selected < 0) return 0;
  return Math.max(0, Math.min(
    state.rows.length - WINDOW_SIZE,
    selected - Math.floor(WINDOW_SIZE / 2),
  ));
};

const isEditable = (target: EventTarget | null): boolean => {
  if (!(target instanceof HTMLElement)) return false;
  return target.isContentEditable || ["INPUT", "SELECT", "TEXTAREA"].includes(target.tagName);
};

export function QuickTraceApp({ send, state }: QuickTraceAppProps) {
  const treeRef = useRef<HTMLDivElement>(null);
  const focused = useRef(false);
  const [windowStart, setWindowStart] = useState(() => initialWindowStart(state));

  useEffect(() => {
    if (state.status === "ready" && !focused.current) {
      treeRef.current?.focus();
      focused.current = true;
    }
  }, [state.status]);

  useEffect(() => {
    setWindowStart(initialWindowStart(state));
  }, [state.rows, state.selectedRowId]);

  const selectedRow = state.rows.find((row) => row.rowId === state.selectedRowId);
  const rootRow = state.rows.find((row) => row.depth === 0);
  const directionLabel = state.direction === "precedent" ? "Trace In" : "Trace Out";
  const emptyLabel = state.direction === "precedent"
    ? "No direct precedents"
    : "No direct dependents";

  const siblingMetadata = useMemo(() => {
    const groups = new Map<string, QuickTraceSnapshot["rows"]>();
    state.rows.forEach((row) => {
      const key = row.parentRowId ?? "__root__";
      groups.set(key, [...(groups.get(key) ?? []), row]);
    });
    const metadata = new Map<string, Readonly<{ position: number; size: number }>>();
    groups.forEach((rows) => rows.forEach((row, index) => metadata.set(row.rowId, {
      position: index + 1,
      size: rows.length,
    })));
    return metadata;
  }, [state.rows]);

  const renderedRows = useMemo(() => {
    if (state.rows.length <= WINDOW_SIZE) return state.rows;
    const end = Math.min(state.rows.length, windowStart + WINDOW_SIZE);
    const indices = new Set<number>();
    for (let index = windowStart; index < end; index += 1) indices.add(index);
    const selected = state.rows.findIndex((row) => row.rowId === state.selectedRowId);
    if (selected >= 0) indices.add(selected);
    return [...indices]
      .sort((left, right) => left - right)
      .flatMap((index) => state.rows[index] === undefined ? [] : [state.rows[index]]);
  }, [state.rows, state.selectedRowId, windowStart]);

  const onKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (event.target !== event.currentTarget || isEditable(event.target)) return;
    if (event.metaKey || event.ctrlKey || event.altKey) return;
    let intent: QuickTraceUserIntent | null = null;
    let blockedByWorkbook = false;
    switch (event.key) {
      case "ArrowUp":
        intent = { movement: "previous", type: "move" };
        blockedByWorkbook = true;
        break;
      case "ArrowDown":
        intent = { movement: "next", type: "move" };
        blockedByWorkbook = true;
        break;
      case "ArrowLeft":
        intent = { type: "left" };
        blockedByWorkbook = true;
        break;
      case "ArrowRight":
        intent = { type: "right" };
        blockedByWorkbook = true;
        break;
      case "Enter":
        if (selectedRow !== undefined) {
          intent = { rowId: selectedRow.rowId, type: "reroot" };
          blockedByWorkbook = true;
        }
        break;
      case "Escape":
        intent = { type: "close" };
        break;
      default: {
        const key = event.key.toLocaleLowerCase();
        if (key === "t") {
          intent = {
            direction: state.direction === "precedent" ? "dependent" : "precedent",
            type: "setDirection",
          };
        } else if (key === "o") {
          intent = { type: "returnToOrigin" };
          blockedByWorkbook = true;
        }
      }
    }
    if (intent === null) return;
    event.preventDefault();
    if (blockedByWorkbook && state.navigationBlocked) return;
    send(intent);
  };

  const onScroll = (event: React.UIEvent<HTMLDivElement>) => {
    if (state.rows.length <= WINDOW_SIZE) return;
    const next = Math.max(0, Math.min(
      state.rows.length - WINDOW_SIZE,
      Math.floor(event.currentTarget.scrollTop / ROW_HEIGHT) - OVERSCAN,
    ));
    setWindowStart(next);
  };

  const topSpacer = state.rows.length <= WINDOW_SIZE ? 0 : windowStart * ROW_HEIGHT;
  const bottomSpacer = state.rows.length <= WINDOW_SIZE
    ? 0
    : Math.max(0, (state.rows.length - windowStart - WINDOW_SIZE) * ROW_HEIGHT);

  return (
    <main className="quick-trace-shell">
      <header className="quick-trace-header">
        <div className="quick-trace-header-top">
          <div aria-label="Trace direction" className="quick-trace-direction" role="group">
            <button
              aria-pressed={state.direction === "precedent"}
              onClick={() => send({ direction: "precedent", type: "setDirection" })}
              type="button"
            >Trace In</button>
            <button
              aria-pressed={state.direction === "dependent"}
              onClick={() => send({ direction: "dependent", type: "setDirection" })}
              type="button"
            >Trace Out</button>
          </div>
          <p className="quick-trace-origin">Origin: {fullReference(state.origin)}</p>
        </div>
        <section
          aria-label={`Current root ${fullReference(state.root)}`}
          className="quick-trace-root quick-trace-rootline"
        >
          <span className="root-label">Root:</span>
          <span className="root-ref">{fullReference(state.root)}</span>
          {rootRow === undefined ? null : (
            <span className="root-value">{rootRow.displayValue}</span>
          )}
        </section>
        <code className="quick-trace-formula">{state.formula ?? "No formula"}</code>
      </header>

      <div aria-hidden="true" className="quick-trace-columns">
        <span>Reference</span>
        <span>Value</span>
      </div>
      <div
        aria-activedescendant={selectedRow === undefined ? undefined : rowDomId(selectedRow.rowId)}
        aria-busy={state.status === "loading"}
        aria-disabled={state.navigationBlocked || undefined}
        aria-label={`${directionLabel} dependencies`}
        className="quick-trace-tree"
        onKeyDown={onKeyDown}
        onScroll={onScroll}
        ref={treeRef}
        role="tree"
        tabIndex={0}
      >
        {topSpacer === 0 ? null : <div aria-hidden="true" style={{ height: topSpacer }} />}
        {renderedRows.map((row) => {
          const metadata = siblingMetadata.get(row.rowId);
          const badge = issueLabel(row.issue);
          const reference = fullReference(row.node.ref);
          return (
            <div
              aria-expanded={row.disclosure === "leaf"
                ? undefined
                : row.disclosure === "expanded"}
              aria-label={`${reference} ${row.displayValue}${badge === null ? "" : ` ${badge}`}`}
              aria-level={row.depth + 1}
              aria-posinset={metadata?.position}
              aria-selected={row.rowId === state.selectedRowId}
              aria-setsize={metadata?.size}
              className="quick-trace-row"
              id={rowDomId(row.rowId)}
              key={row.rowId}
              onClick={() => {
                if (!state.navigationBlocked) send({ rowId: row.rowId, type: "navigate" });
              }}
              role="treeitem"
              style={{ paddingInlineStart: `${row.depth * 16 + 8}px` }}
            >
              <span className="quick-trace-reference">
                {row.disclosure === "leaf" ? (
                  <span aria-hidden="true" className="quick-trace-chevron-placeholder" />
                ) : (
                  <button
                    aria-label={`${row.disclosure === "expanded" ? "Collapse" : "Expand"} ${reference}`}
                    className="quick-trace-chevron"
                    onClick={(event) => {
                      event.stopPropagation();
                      if (!state.navigationBlocked) {
                        send({ rowId: row.rowId, type: "toggleBranch" });
                      }
                    }}
                    type="button"
                  >{row.disclosure === "expanded" ? "▾" : "▸"}</button>
                )}
                <span>{reference}</span>
                {badge === null ? null : <span className={`quick-trace-badge is-${row.issue}`}>{badge}</span>}
              </span>
              <span className="quick-trace-value">{row.displayValue}</span>
            </div>
          );
        })}
        {bottomSpacer === 0 ? null : <div aria-hidden="true" style={{ height: bottomSpacer }} />}
      </div>

      <div className="quick-trace-secondary">
        <div className="quick-trace-live-regions">
          {state.rows.length === 1 && state.status === "ready"
            ? <p role="status">{emptyLabel}</p>
            : null}
          {state.status === "loading" ? <p role="status">Loading dependencies…</p> : null}
          {state.capped ? <p role="status">Results stopped at the 500-node limit.</p> : null}
          {state.notice === null ? null : state.status === "error" ? (
            <div className="quick-trace-error" role="alert">
              <span>{state.notice}</span>
              {selectedRow === undefined ? null : (
                <button
                  disabled={state.navigationBlocked}
                  onClick={() => {
                    if (!state.navigationBlocked) {
                      send({ rowId: selectedRow.rowId, type: "navigate" });
                    }
                  }}
                  type="button"
                >Retry selected reference</button>
              )}
            </div>
          ) : <p>{state.notice}</p>}
          {state.pendingExcelSelection === null ? null : (
            <section className="quick-trace-notice" role="status">
              <p>Excel selection changed</p>
              <button
                onClick={() => send({ type: "useExcelSelection" })}
                type="button"
              >Use {fullReference(state.pendingExcelSelection)} as new root</button>
            </section>
          )}
        </div>

        <section aria-label="Selected reference detail" className="quick-trace-detail">
          {selectedRow === undefined ? <p>No selected reference</p> : (
            <>
              <p className="detail-ref">{fullReference(selectedRow.node.ref)}</p>
              <code>{state.selectedFormula ?? "No formula"}</code>
              <button
                disabled={state.navigationBlocked}
                onClick={() => send({ rowId: selectedRow.rowId, type: "reroot" })}
                type="button"
              >Make selected reference the root</button>
            </>
          )}
        </section>
      </div>

      <footer className="quick-trace-footer">
        <span className="quick-trace-keys" aria-label="Keyboard shortcuts">
          <span className="key-combo"><kbd>↑</kbd><kbd>↓</kbd> navigate</span>
          <span className="key-combo"><kbd>←</kbd><kbd>→</kbd> collapse · expand</span>
          <span className="key-combo"><kbd>⏎</kbd> root</span>
          <span className="key-combo optional"><kbd>T</kbd> switch</span>
          <span className="key-combo optional"><kbd>O</kbd> origin</span>
          <span className="key-combo optional"><kbd>esc</kbd> close</span>
        </span>
        <span className="quick-trace-count">
          {state.nodeCount} nodes{state.capped ? " · capped" : ""}
        </span>
        <div className="quick-trace-footer-actions">
          <button
            disabled={state.navigationBlocked}
            onClick={() => send({ type: "returnToOrigin" })}
            type="button"
          >Return to origin</button>
          <button onClick={() => send({ type: "close" })} type="button">Close</button>
        </div>
      </footer>
    </main>
  );
}
