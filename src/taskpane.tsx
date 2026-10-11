import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import {
  applyMacShortcutDefaults,
  initializeOfficeActionBindings,
  officeReady,
} from "./platform/office/officeActionPort";
import { createExcelOfficeGateway } from "./platform/office/excelOfficeGateway";
import { createQuickTraceDialogHost } from "./platform/office/quickTraceDialogHost";
import { createOfficeGatewayProvider } from "./runtime/officeGatewayBootstrap";
import {
  createQuickTraceController,
  type QuickTraceController,
} from "./runtime/quickTraceController";
import type { CommandId } from "./shared/types";
import { App } from "./ui/App";
import "./ui/app.css";

declare const __ANALYST_LENS_FIXTURES__: boolean;

const root = document.getElementById("root");

if (root === null) {
  throw new Error("Analyst Lens task pane root was not found");
}

const render = async () => {
  const parameters = new URLSearchParams(window.location.search);
  if (
    __ANALYST_LENS_FIXTURES__ &&
    parameters.get("fixture") === "sample-model"
  ) {
    const { createSimulatedWorkbookFixture } = await import(
      "../tests/fixtures/simulatedWorkbook"
    );
    const fixture = createSimulatedWorkbookFixture({
      initialState: parameters.get("selection") ?? "formula",
    });
    createRoot(root).render(
      <StrictMode>
        <App getGateway={async () => fixture.gateway} runCommand={fixture.runCommand} />
      </StrictMode>,
    );
    return;
  }

  await officeReady();
  const getOfficeGateway = createOfficeGatewayProvider({
    createGateway: createExcelOfficeGateway,
    ready: officeReady,
  });
  const quickTraceDialogHost = createQuickTraceDialogHost();
  let quickTrace: QuickTraceController | null = null;
  const getQuickTrace = async () => {
    quickTrace ??= createQuickTraceController({
      dialogHost: quickTraceDialogHost,
      dialogUrl: new URL("/quick-trace.html", window.location.origin).href,
      gateway: await getOfficeGateway(),
    });
    return quickTrace;
  };
  const executeCommand = async (id: CommandId) => {
    if (id === "audit.quickTrace.precedents") {
      await (await getQuickTrace()).open("precedent");
      return { ok: true } as const;
    }
    if (id === "audit.quickTrace.dependents") {
      await (await getQuickTrace()).open("dependent");
      return { ok: true } as const;
    }
    return (await getOfficeGateway()).runWrite({ id });
  };
  createRoot(root).render(
    <StrictMode>
      <App
        getGateway={getOfficeGateway}
        openQuickTrace={async (direction) => (await getQuickTrace()).open(direction)}
        quickTraceSupported={quickTraceDialogHost.isSupported()}
        runCommand={executeCommand}
        updateShortcuts={applyMacShortcutDefaults}
      />
    </StrictMode>,
  );
  void initializeOfficeActionBindings(executeCommand).catch(() => undefined);
};

void render();
