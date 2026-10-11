import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createWatchItem, refreshWatchItem } from "../../src/domain/watchlist/watchlist";
import type { CellFacts } from "../../src/shared/types";
import { Watchlist } from "../../src/ui/Watchlist";

afterEach(cleanup);

const facts = (
  address = "$F$42",
  value: CellFacts["value"] = 18.4,
  formula = "=IRR(B2:B8)",
): CellFacts => ({
  address,
  columnCount: 1,
  formula,
  numberFormat: "0.0%",
  rowCount: 1,
  sheetId: "returns",
  sheetName: "Returns",
  value,
  workbookId: "workbook-1",
});

describe("Watchlist", () => {
  it("shows current/last values, distinct changes, and selected-assumption impact", () => {
    const item = refreshWatchItem(
      { ...createWatchItem(facts()), category: "Returns", label: "IRR" },
      facts("$F$42", 19.1, "=XIRR(B2:B8,A2:A8)"),
    );

    render(<Watchlist
      dispatch={() => undefined}
      state={{
        affected: ["Returns!$F$42"],
        items: [item],
        persistence: "session",
        refreshing: false,
      }}
    />);

    expect(screen.getByRole("heading", { name: "Key outputs" })).toBeTruthy();
    expect(screen.getByRole("heading", { name: "IRR" })).toBeTruthy();
    expect(screen.getByText("Current: 19.1")).toBeTruthy();
    expect(screen.getByText("Last: 18.4")).toBeTruthy();
    expect(screen.getByText("Formula changed")).toBeTruthy();
    expect(screen.getByText("Value changed")).toBeTruthy();
    expect(screen.getByText("Affected by current selection")).toBeTruthy();
  });

  it("exposes refresh, navigation, metadata editing, unpin, and explicit persistence opt-in", () => {
    const dispatch = vi.fn();
    const item = { ...createWatchItem(facts()), category: "Returns", label: "IRR" };
    render(<Watchlist
      dispatch={dispatch}
      state={{ affected: [], items: [item], persistence: "session", refreshing: false }}
    />);

    fireEvent.click(screen.getByRole("button", { name: "Refresh key outputs" }));
    fireEvent.click(screen.getByRole("button", { name: "Go to IRR" }));
    fireEvent.change(screen.getByRole("textbox", { name: "Label for Returns $F$42" }), {
      target: { value: "Levered IRR" },
    });
    fireEvent.blur(screen.getByRole("textbox", { name: "Label for Returns $F$42" }));
    fireEvent.click(screen.getByRole("button", { name: "Unpin IRR" }));

    const save = screen.getByRole("button", { name: "Save watchlist with workbook" });
    expect(save.hasAttribute("disabled")).toBe(true);
    fireEvent.click(screen.getByRole("checkbox", {
      name: "Store only addresses, labels, and categories in this workbook",
    }));
    expect(save.hasAttribute("disabled")).toBe(false);
    fireEvent.click(save);

    expect(dispatch).toHaveBeenCalledWith({ type: "refreshWatchlist" });
    expect(dispatch).toHaveBeenCalledWith(expect.objectContaining({
      ref: expect.objectContaining({ address: "$F$42" }),
      type: "navigate",
    }));
    expect(dispatch).toHaveBeenCalledWith(expect.objectContaining({
      id: item.id,
      label: "Levered IRR",
      type: "updateWatchMetadata",
    }));
    expect(dispatch).toHaveBeenCalledWith({ id: item.id, type: "unpinWatch" });
    expect(dispatch).toHaveBeenCalledWith({ type: "saveWatchlistWithWorkbook" });
  });

  it("keeps a missing or renamed sheet pin visible and non-destructive", () => {
    const item = {
      ...createWatchItem(facts()),
      message: "The pinned sheet is missing or was renamed.",
      status: "missing" as const,
    };
    render(<Watchlist
      dispatch={() => undefined}
      state={{ affected: [], items: [item], persistence: "session", refreshing: false }}
    />);

    expect(screen.getByText("The pinned sheet is missing or was renamed.")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Unpin Returns $F$42" })).toBeTruthy();
  });
});
