import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { createOfficeQuickTraceDialogChild } from "./platform/office/quickTraceDialogChild";
import type { QuickTraceUserIntent } from "./runtime/quickTraceController";
import {
  parseQuickTraceHostEvent,
  QUICK_TRACE_PROTOCOL_VERSION,
  serializeQuickTraceMessage,
  type QuickTraceDialogIntent,
  type QuickTraceSessionIdentity,
  type QuickTraceSnapshot,
} from "./shared/quickTraceProtocol";
import { QuickTraceApp } from "./ui/quick-trace/QuickTraceApp";
import "./ui/quick-trace/quickTrace.css";

declare const __ANALYST_LENS_FIXTURES__: boolean;

type QuickTraceClient = Readonly<{
  getState(): QuickTraceSnapshot;
  send(intent: QuickTraceUserIntent): Promise<void>;
  subscribe(listener: () => void): () => void;
}>;

const rootElement = document.getElementById("quick-trace-root");

if (rootElement === null) {
  throw new Error("Analyst Lens Quick Trace root was not found");
}

const render = async () => {
  const parameters = new URLSearchParams(window.location.search);
  if (
    __ANALYST_LENS_FIXTURES__ &&
    parameters.get("fixture") === "sample-model"
  ) {
    const { createQuickTraceFixtureClient } = await import(
      "../tests/fixtures/quickTraceFixture"
    );
    const fixtureClient: QuickTraceClient = await createQuickTraceFixtureClient();
    const root = createRoot(rootElement);
    const renderFixture = () => root.render(
      <StrictMode>
        <QuickTraceApp
          send={(intent) => void fixtureClient.send(intent)}
          state={fixtureClient.getState()}
        />
      </StrictMode>,
    );
    fixtureClient.subscribe(renderFixture);
    renderFixture();
    return;
  }
  const identity: QuickTraceSessionIdentity = {
    nonce: parameters.get("nonce") ?? "",
    sessionId: parameters.get("sessionId") ?? "",
  };
  if (identity.nonce.length === 0 || identity.sessionId.length === 0) {
    throw new Error("Quick Trace requires an active session identity");
  }

  const root = createRoot(rootElement);
  const child = await createOfficeQuickTraceDialogChild();
  let outgoingSequence = 0;
  let incomingSequence = -1;

  const envelope = () => ({
    ...identity,
    protocol: QUICK_TRACE_PROTOCOL_VERSION,
    sequence: ++outgoingSequence,
  });

  const send = (intent: QuickTraceUserIntent) => {
    let message: QuickTraceDialogIntent;
    switch (intent.type) {
      case "move":
        message = { ...envelope(), payload: { movement: intent.movement }, type: "quickTrace.move" };
        break;
      case "navigate":
        message = { ...envelope(), payload: { rowId: intent.rowId }, type: "quickTrace.navigate" };
        break;
      case "reroot":
        message = { ...envelope(), payload: { rowId: intent.rowId }, type: "quickTrace.reroot" };
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
    child.post(serializeQuickTraceMessage(message));
  };

  child.subscribe((raw) => {
    let event;
    try {
      event = parseQuickTraceHostEvent(raw, identity);
    } catch {
      return;
    }
    if (event.sequence <= incomingSequence) return;
    incomingSequence = event.sequence;
    if (event.type === "quickTrace.error") {
      root.render(<p role="alert">{event.payload.message}</p>);
      return;
    }
    root.render(
      <StrictMode>
        <QuickTraceApp send={send} state={event.payload.state} />
      </StrictMode>,
    );
  });
  const ready: QuickTraceDialogIntent = {
    ...envelope(),
    type: "quickTrace.ready",
  };
  child.post(serializeQuickTraceMessage(ready));
};

void render();
