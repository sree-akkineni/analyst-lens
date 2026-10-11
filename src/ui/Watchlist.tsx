import { useEffect, useState } from "react";
import {
  watchDisplayRef,
  type WatchItem,
  type WatchlistState,
} from "../domain/watchlist/watchlist";
import type { CellRef } from "../shared/types";

export type WatchlistIntent =
  | Readonly<{ id: string; type: "unpinWatch" }>
  | Readonly<{ category: string; id: string; label: string; type: "updateWatchMetadata" }>
  | Readonly<{ ref: CellRef; type: "navigate" }>
  | Readonly<{ type: "pinCurrent" | "refreshWatchlist" | "saveWatchlistWithWorkbook" }>;

type WatchlistProps = {
  dispatch(intent: WatchlistIntent): void;
  state: WatchlistState;
};

const displayValue = (value: WatchItem["currentValue"]): string =>
  value === null ? "—" : String(value);

const WatchRow = ({
  affected,
  dispatch,
  item,
}: {
  affected: boolean;
  dispatch(intent: WatchlistIntent): void;
  item: WatchItem;
}) => {
  const [category, setCategory] = useState(item.category);
  const [label, setLabel] = useState(item.label);
  useEffect(() => setCategory(item.category), [item.category]);
  useEffect(() => setLabel(item.label), [item.label]);
  const refLabel = watchDisplayRef(item.ref).replace("!", " ");
  const displayName = item.label.trim() || refLabel;
  const saveMetadata = () => dispatch({
    category,
    id: item.id,
    label,
    type: "updateWatchMetadata",
  });

  return (
    <li className={`watch-item watch-item-${item.status}`}>
      <div className="watch-heading">
        <div>
          <p className="watch-ref">{watchDisplayRef(item.ref)}</p>
          <h4 className="watch-name">{displayName}</h4>
        </div>
        <button
          aria-label={`Unpin ${displayName}`}
          onClick={() => dispatch({ id: item.id, type: "unpinWatch" })}
          type="button"
        >Unpin</button>
      </div>
      {item.status === "missing" ? (
        <p role="alert">{item.message}</p>
      ) : (
        <div className="watch-values">
          <span>Current: {displayValue(item.currentValue)}</span>
          <span>Last: {displayValue(item.lastValue)}</span>
        </div>
      )}
      <div className="watch-badges">
        {item.formulaChanged ? <span>Formula changed</span> : null}
        {item.valueChanged ? <span>Value changed</span> : null}
        {affected ? <span>Affected by current selection</span> : null}
      </div>
      <div className="watch-metadata">
        <label>
          Label
          <input
            aria-label={`Label for ${refLabel}`}
            onBlur={saveMetadata}
            onChange={(event) => setLabel(event.currentTarget.value)}
            value={label}
          />
        </label>
        <label>
          Category
          <input
            aria-label={`Category for ${refLabel}`}
            onBlur={saveMetadata}
            onChange={(event) => setCategory(event.currentTarget.value)}
            value={category}
          />
        </label>
      </div>
      <button
        aria-label={`Go to ${displayName}`}
        disabled={item.status === "missing"}
        onClick={() => dispatch({ ref: item.ref, type: "navigate" })}
        type="button"
      >Go to cell</button>
    </li>
  );
};

export function Watchlist({ dispatch, state }: WatchlistProps) {
  const [persistenceConfirmed, setPersistenceConfirmed] = useState(false);
  return (
    <div className="watchlist">
      <h3 className="watch-section-title">Key outputs</h3>
      <div className="watch-actions">
        <button
          disabled={state.refreshing}
          onClick={() => dispatch({ type: "refreshWatchlist" })}
          type="button"
        >{state.refreshing ? "Refreshing…" : "Refresh key outputs"}</button>
        <button onClick={() => dispatch({ type: "pinCurrent" })} type="button">
          Pin current selection
        </button>
        <span>{state.affected.length} affected by current selection</span>
      </div>
      {state.items.length === 0 ? (
        <p>No key outputs pinned for this session.</p>
      ) : (
        <ul className="watch-items">
          {state.items.map((item) => (
            <WatchRow
              affected={state.affected.includes(watchDisplayRef(item.ref))}
              dispatch={dispatch}
              item={item}
              key={item.id}
            />
          ))}
        </ul>
      )}
      <div className="watch-persistence">
        <label>
          <input
            checked={persistenceConfirmed}
            onChange={(event) => setPersistenceConfirmed(event.currentTarget.checked)}
            type="checkbox"
          />
          Store only addresses, labels, and categories in this workbook
        </label>
        <button
          disabled={!persistenceConfirmed || state.persistence === "saving" || state.persistence === "unavailable"}
          onClick={() => dispatch({ type: "saveWatchlistWithWorkbook" })}
          type="button"
        >Save watchlist with workbook</button>
        <p role="status">
          {state.persistence === "saved"
            ? "Watchlist metadata saved with this workbook."
            : state.persistence === "failed"
              ? "Watchlist metadata could not be saved. Session pins are unchanged."
              : state.persistence === "unavailable"
                ? "Workbook persistence is unavailable in this Excel host."
                : "Values and formulas always remain session-only."}
        </p>
      </div>
    </div>
  );
}
