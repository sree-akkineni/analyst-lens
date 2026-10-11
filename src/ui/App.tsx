import { useEffect, useState, useSyncExternalStore } from "react";
import type { OfficeGateway } from "../platform/office/officeGateway";
import {
  createController,
  type InspectorController,
} from "../runtime/controller";
import type { CommandId, CommandResult } from "../shared/types";
import type {
  ShortcutUpdateOptions,
  ShortcutUpdateResult,
} from "../runtime/actionBindings";
import { Inspector, type OpenQuickTrace } from "./Inspector";
import { HealthScan } from "./HealthScan";
import { SpeedLayer } from "./SpeedLayer";
import { scanWorkbook } from "../domain/health/scan";
import { Watchlist } from "./Watchlist";
import { Disclosure } from "./Disclosure";

export type AppProps = {
  controller?: InspectorController;
  getGateway?(): Promise<OfficeGateway>;
  openQuickTrace?: OpenQuickTrace | undefined;
  quickTraceSupported?: boolean | undefined;
  runCommand?(id: CommandId): Promise<CommandResult>;
  updateShortcuts?: ((options?: ShortcutUpdateOptions) => Promise<ShortcutUpdateResult>) | undefined;
};

const unavailableRunner = async (id: CommandId): Promise<CommandResult> => ({
  code: "COMMAND_NOT_READY",
  message: `${id} is not available yet`,
  ok: false,
});

const InspectorPanel = ({
  controller,
  openQuickTrace,
  quickTraceSupported,
  runCommand,
  updateShortcuts,
}: {
  controller: InspectorController;
  openQuickTrace?: OpenQuickTrace | undefined;
  quickTraceSupported?: boolean | undefined;
  runCommand(id: CommandId): Promise<CommandResult>;
  updateShortcuts?: ((options?: ShortcutUpdateOptions) => Promise<ShortcutUpdateResult>) | undefined;
}) => {
  const state = useSyncExternalStore(
    controller.subscribe,
    controller.getState,
    controller.getState,
  );
  return (
    <>
      <Inspector
        dispatch={(intent) => void controller.dispatch(intent)}
        openQuickTrace={openQuickTrace}
        quickTraceSupported={quickTraceSupported}
        runCommand={runCommand}
        state={state}
        updateShortcuts={updateShortcuts}
      />
      <Disclosure
        defaultOpen={state.watchlist.items.length > 0}
        eyebrow="Session-local monitoring"
        meta={`${state.watchlist.items.length} pinned`}
        openWhen={state.watchlist.items.length > 0}
        title="Key outputs"
        titleId="watchlist-title"
      >
        <Watchlist
          dispatch={(intent) => void controller.dispatch(intent)}
          state={state.watchlist}
        />
      </Disclosure>
    </>
  );
};

const exportLocalMarkdown = (markdown: string) => {
  const url = URL.createObjectURL(new Blob([markdown], { type: "text/markdown" }));
  const link = document.createElement("a");
  link.download = "analyst-lens-health.md";
  link.href = url;
  link.click();
  URL.revokeObjectURL(url);
};

const HealthPanel = ({ gateway }: { gateway: OfficeGateway }) => {
  const getHealthWorkbookFacts = gateway.getHealthWorkbookFacts;
  if (getHealthWorkbookFacts === undefined) return null;
  return (
    <Disclosure
      eyebrow="Evidence-first audit"
      meta="On demand"
      title="Model health"
      titleId="health-scan-title"
    >
      <HealthScan
        exportLocal={exportLocalMarkdown}
        navigate={(ref) => void gateway.selectCell(ref)}
        runScan={async ({ onProgress, signal }) => scanWorkbook(
          await getHealthWorkbookFacts(signal),
          { onProgress, signal },
        )}
      />
    </Disclosure>
  );
};

export function App({
  controller: injectedController,
  getGateway,
  openQuickTrace,
  quickTraceSupported,
  runCommand = unavailableRunner,
  updateShortcuts,
}: AppProps = {}) {
  const [controller, setController] = useState<InspectorController | null>(
    injectedController ?? null,
  );
  const [gateway, setGateway] = useState<OfficeGateway | null>(null);
  const [bootstrapError, setBootstrapError] = useState<string | null>(null);

  useEffect(() => {
    let disposed = false;
    let active = injectedController ?? null;
    const start = async () => {
      try {
        if (active === null && getGateway !== undefined) {
          const resolvedGateway = await getGateway();
          active = createController(resolvedGateway);
          if (disposed) {
            active.stop();
            return;
          }
          setGateway(resolvedGateway);
          setController(active);
        }
        await active?.start();
      } catch (error) {
        if (!disposed) {
          setBootstrapError(error instanceof Error ? error.message : String(error));
        }
      }
    };
    void start();
    return () => {
      disposed = true;
      active?.stop();
    };
  }, [getGateway, injectedController]);

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Excel for Mac</p>
          <h1>Analyst Lens</h1>
        </div>
        <p className="ready-status" role="status">
          <span aria-hidden="true" className="ready-dot" />
          Local and ready
        </p>
      </header>

      <div className="workspace-grid">
        {controller !== null ? (
          <InspectorPanel
            controller={controller}
            openQuickTrace={openQuickTrace}
            quickTraceSupported={quickTraceSupported}
            runCommand={runCommand}
            updateShortcuts={updateShortcuts}
          />
        ) : bootstrapError !== null ? (
          <section className="inspector-state"><p role="alert">{bootstrapError}</p></section>
        ) : (
          <section aria-labelledby="getting-started" className="empty-state">
            <p className="section-label">Workbook cockpit</p>
            <h2 id="getting-started">Select a cell to begin</h2>
            <p>Inspector, dependency navigation, model health, and watchlist context will appear here.</p>
          </section>
        )}
        <Disclosure
          defaultOpen
          eyebrow="Selected range"
          meta="Formats · structure"
          title="Speed tools"
          titleId="speed-tools"
        >
          <SpeedLayer runCommand={runCommand} />
        </Disclosure>
        {gateway === null ? null : <HealthPanel gateway={gateway} />}
      </div>
    </main>
  );
}
