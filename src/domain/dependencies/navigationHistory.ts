import { canonicalCellId, type CellRef } from "../../shared/types";

export type NavigationHistory = Readonly<{
  current: CellRef;
  future: readonly CellRef[];
  origin: CellRef;
  past: readonly CellRef[];
}>;

const MAX_VISITS = 200;

const sameCell = (left: CellRef, right: CellRef): boolean =>
  canonicalCellId(left) === canonicalCellId(right);

const boundedPast = (past: readonly CellRef[]): CellRef[] =>
  past.slice(-(MAX_VISITS - 1));

export const startHistory = (origin: CellRef): NavigationHistory => ({
  current: origin,
  future: [],
  origin,
  past: [],
});

export const visit = (
  history: NavigationHistory,
  target: CellRef,
): NavigationHistory => {
  if (target.workbookId !== history.current.workbookId) {
    return startHistory(target);
  }
  if (sameCell(history.current, target)) return history;
  return {
    current: target,
    future: [],
    origin: history.origin,
    past: boundedPast([...history.past, history.current]),
  };
};

export const back = (history: NavigationHistory): NavigationHistory => {
  const current = history.past.at(-1);
  if (current === undefined) return history;
  return {
    current,
    future: [history.current, ...history.future],
    origin: history.origin,
    past: history.past.slice(0, -1),
  };
};

export const forward = (history: NavigationHistory): NavigationHistory => {
  const [current, ...future] = history.future;
  if (current === undefined) return history;
  return {
    current,
    future,
    origin: history.origin,
    past: boundedPast([...history.past, history.current]),
  };
};

export const returnToOrigin = (
  history: NavigationHistory,
): NavigationHistory => visit(history, history.origin);
