import { describe, expect, it } from "vitest";
import {
  parseQuickTraceDialogIntent,
  parseQuickTraceHostEvent,
  serializeQuickTraceMessage,
} from "../../src/shared/quickTraceProtocol";

const session = { nonce: "nonce-1", sessionId: "session-1" } as const;
const origin = {
  address: "$F$42",
  sheetId: "dcf",
  sheetName: "DCF",
  workbookId: "fixture://sample-model",
} as const;
const node = { id: "fixture://sample-model::dcf::$F$42", ref: origin } as const;
const state = {
  capped: false,
  direction: "precedent" as const,
  formula: "=Inputs!$B$7*2",
  navigationBlocked: false,
  nodeCount: 1,
  notice: null,
  origin,
  pendingExcelSelection: null,
  root: origin,
  rows: [{
    depth: 0,
    disclosure: "leaf" as const,
    displayValue: "12",
    issue: null,
    node,
    parentRowId: null,
    rowId: "root",
  }],
  selectedFormula: "=Inputs!$B$7*2",
  selectedRowId: "root",
  status: "ready" as const,
};

describe("Quick Trace protocol", () => {
  it.each([
    {
      ...session,
      protocol: 1 as const,
      sequence: 1,
      type: "quickTrace.ready" as const,
    },
    {
      ...session,
      payload: { movement: "next" as const },
      protocol: 1 as const,
      sequence: 2,
      type: "quickTrace.move" as const,
    },
    {
      ...session,
      protocol: 1 as const,
      sequence: 3,
      type: "quickTrace.left" as const,
    },
    {
      ...session,
      protocol: 1 as const,
      sequence: 4,
      type: "quickTrace.right" as const,
    },
    {
      ...session,
      payload: { direction: "dependent" as const },
      protocol: 1 as const,
      sequence: 5,
      type: "quickTrace.setDirection" as const,
    },
    {
      ...session,
      payload: { rowId: "root/child" },
      protocol: 1 as const,
      sequence: 6,
      type: "quickTrace.navigate" as const,
    },
    {
      ...session,
      payload: { rowId: "root/child" },
      protocol: 1 as const,
      sequence: 7,
      type: "quickTrace.toggleBranch" as const,
    },
    {
      ...session,
      payload: { rowId: "root/child" },
      protocol: 1 as const,
      sequence: 8,
      type: "quickTrace.reroot" as const,
    },
    {
      ...session,
      protocol: 1 as const,
      sequence: 9,
      type: "quickTrace.returnToOrigin" as const,
    },
    {
      ...session,
      protocol: 1 as const,
      sequence: 10,
      type: "quickTrace.useExcelSelection" as const,
    },
    {
      ...session,
      protocol: 1 as const,
      sequence: 11,
      type: "quickTrace.close" as const,
    },
  ])("round-trips final dialog intent $type", (message) => {
    expect(parseQuickTraceDialogIntent(
      serializeQuickTraceMessage(message),
      session,
    )).toEqual(message);
  });

  it.each([
    {
      ...session,
      payload: { state },
      protocol: 1 as const,
      sequence: 1,
      type: "quickTrace.initialize" as const,
    },
    {
      ...session,
      payload: { state },
      protocol: 1 as const,
      sequence: 2,
      type: "quickTrace.state" as const,
    },
  ])("round-trips complete host snapshot $type", (message) => {
    expect(parseQuickTraceHostEvent(
      serializeQuickTraceMessage(message),
      session,
    )).toEqual(message);
  });

  it.each([
    ["wrong version", { ...session, protocol: 2, sequence: 1, type: "quickTrace.ready" }],
    ["wrong nonce", { ...session, nonce: "other", protocol: 1, sequence: 1, type: "quickTrace.ready" }],
    ["unknown action", { ...session, protocol: 1, sequence: 1, type: "quickTrace.writeFormula" }],
    ["invalid movement", {
      ...session,
      payload: { movement: "sideways" },
      protocol: 1,
      sequence: 1,
      type: "quickTrace.move",
    }],
  ])("rejects %s", (_label, message) => {
    expect(() => parseQuickTraceDialogIntent(JSON.stringify(message), session)).toThrow();
  });

  it("rejects a host event from a stale session", () => {
    const event = JSON.stringify({
      ...session,
      payload: { code: "STALE", message: "stale" },
      protocol: 1,
      sequence: 3,
      type: "quickTrace.error",
    });

    expect(() => parseQuickTraceHostEvent(event, {
      nonce: "nonce-2",
      sessionId: "session-2",
    })).toThrow();
  });
});
