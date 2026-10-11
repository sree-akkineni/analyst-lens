import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import {
  CommandPalette,
  type ContextAction,
} from "../../src/ui/CommandPalette";

afterEach(cleanup);

describe("CommandPalette", () => {
  it("preserves the future context-action boundary without shipping AI actions", () => {
    const action: ContextAction = {
      availability: "coming-soon",
      description: "Reserved for a future provider-backed action.",
      id: "future.action",
      kind: "ai",
      label: "Future action",
    };

    expect(action.kind).toBe("ai");
  });

  it("shows registered Mac shortcuts with their commands", () => {
    render(<CommandPalette
      onClose={() => undefined}
      open
      runCommand={async () => ({ ok: true })}
    />);

    expect(screen.getByRole("button", { name: /Percentage Ctrl\+Shift\+5/ })).toBeTruthy();
    expect(screen.getByRole("button", { name: /Bottom border Ctrl\+Option\+Shift\+Down/ })).toBeTruthy();
    expect(screen.getByRole("button", {
      name: /Quick Trace In.*Ctrl\+Option\+P/u,
    })).toBeTruthy();
  });

  it("exposes the approved structural and universal visibility commands", () => {
    render(<CommandPalette
      onClose={() => undefined}
      open
      runCommand={async () => ({ ok: true })}
    />);

    expect(screen.getByRole("button", { name: "Group rows" }).hasAttribute("disabled")).toBe(false);
    expect(screen.getByRole("button", { name: "Ungroup columns" }).hasAttribute("disabled")).toBe(false);
    expect(screen.getByRole("button", { name: "Open All Groups" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Close All Groups" })).toBeTruthy();
  });
});
