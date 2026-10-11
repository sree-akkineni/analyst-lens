import type {
  QuickTraceDialogHost,
  QuickTraceDialogSession,
} from "../../src/platform/office/quickTraceDialogHost";
import {
  createQuickTraceController,
  type QuickTraceUserIntent,
} from "../../src/runtime/quickTraceController";
import {
  parseQuickTraceDialogIntent,
  parseQuickTraceHostEvent,
  QUICK_TRACE_PROTOCOL_VERSION,
  serializeQuickTraceMessage,
  type QuickTraceDialogIntent,
  type QuickTraceSessionIdentity,
  type QuickTraceSnapshot,
} from "../../src/shared/quickTraceProtocol";
import type { DependencyDirection } from "../../src/shared/types";
import { createSimulatedWorkbookFixture } from "./simulatedWorkbook";

export type QuickTraceFixtureClient = Readonly<{
  getState(): QuickTraceSnapshot;
  send(intent: QuickTraceUserIntent): Promise<void>;
  subscribe(listener: () => void): () => void;
}>;

const identity: QuickTraceSessionIdentity = {
  nonce: "quick-trace-fixture-nonce",
  sessionId: "quick-trace-fixture-session",
};

const directionFromLocation = (): DependencyDirection =>
  new URLSearchParams(window.location.search).get("direction") === "dependent"
    ? "dependent"
    : "precedent";

export async function createQuickTraceFixtureClient(): Promise<QuickTraceFixtureClient> {
  const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
  const clientListeners = new Set<() => void>();
  const closedListeners = new Set<(error: number) => void>();
  const dialogListeners = new Set<(message: string, origin?: string) => void>();
  let current: QuickTraceSnapshot | null = null;
  let incomingSequence = -1;
  let outgoingSequence = 0;
  let closed = false;

  const dialog: QuickTraceDialogSession = {
    close() {
      if (closed) return;
      closed = true;
      [...closedListeners].forEach((listener) => listener(0));
      closedListeners.clear();
      dialogListeners.clear();
    },
    post(raw) {
      const event = parseQuickTraceHostEvent(raw, identity);
      if (event.sequence <= incomingSequence) return;
      incomingSequence = event.sequence;
      if (event.type === "quickTrace.error") {
        throw new Error(event.payload.message);
      }
      current = event.payload.state;
      clientListeners.forEach((listener) => listener());
    },
    subscribeClosed(listener) {
      closedListeners.add(listener);
      return () => closedListeners.delete(listener);
    },
    subscribeMessage(listener) {
      dialogListeners.add(listener);
      return () => dialogListeners.delete(listener);
    },
  };
  const dialogHost: QuickTraceDialogHost = {
    isSupported: () => true,
    open: async () => dialog,
  };
  const controller = createQuickTraceController({
    dialogHost,
    dialogUrl: new URL("/quick-trace.html", window.location.origin).href,
    gateway: fixture.gateway,
    identity: () => identity,
  });
  await controller.open(directionFromLocation());

  const envelope = () => ({
    ...identity,
    protocol: QUICK_TRACE_PROTOCOL_VERSION,
    sequence: ++outgoingSequence,
  });
  const postIntent = (message: QuickTraceDialogIntent) => {
    const raw = serializeQuickTraceMessage(message);
    parseQuickTraceDialogIntent(raw, identity);
    dialogListeners.forEach((listener) => listener(raw, window.location.origin));
  };
  postIntent({ ...envelope(), type: "quickTrace.ready" });

  if (current === null) {
    throw new Error("Quick Trace fixture did not receive its initial state");
  }

  return {
    getState() {
      if (current === null) throw new Error("Quick Trace fixture is closed");
      return current;
    },
    async send(intent) {
      let message: QuickTraceDialogIntent;
      switch (intent.type) {
        case "move":
          message = {
            ...envelope(),
            payload: { movement: intent.movement },
            type: "quickTrace.move",
          };
          break;
        case "navigate":
          message = {
            ...envelope(),
            payload: { rowId: intent.rowId },
            type: "quickTrace.navigate",
          };
          break;
        case "reroot":
          message = {
            ...envelope(),
            payload: { rowId: intent.rowId },
            type: "quickTrace.reroot",
          };
          break;
        case "toggleBranch":
          message = {
            ...envelope(),
            payload: { rowId: intent.rowId },
            type: "quickTrace.toggleBranch",
          };
          break;
        case "setDirection":
          message = {
            ...envelope(),
            payload: { direction: intent.direction },
            type: "quickTrace.setDirection",
          };
          break;
        case "left":
          message = { ...envelope(), type: "quickTrace.left" };
          break;
        case "right":
          message = { ...envelope(), type: "quickTrace.right" };
          break;
        case "returnToOrigin":
          message = { ...envelope(), type: "quickTrace.returnToOrigin" };
          break;
        case "useExcelSelection":
          message = { ...envelope(), type: "quickTrace.useExcelSelection" };
          break;
        case "close":
          message = { ...envelope(), type: "quickTrace.close" };
      }
      postIntent(message);
      await Promise.resolve();
    },
    subscribe(listener) {
      clientListeners.add(listener);
      return () => clientListeners.delete(listener);
    },
  };
}
