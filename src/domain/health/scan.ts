import type { HealthFinding } from "../../shared/types";
import { runRules, type HealthWorkbookFacts } from "./rules";

export type ScanProgress = Readonly<{ completed: number; total: number }>;
export type ScanStore = { current: readonly HealthFinding[] | null };

export type ScanOptions = Readonly<{
  chunkSize?: number;
  now?: () => number;
  onProgress?(progress: ScanProgress): void;
  progressDelayMs?: number;
  signal?: AbortSignal;
  store?: ScanStore;
  yieldControl?: () => Promise<void>;
}>;

export const createScanStore = (): ScanStore => ({ current: null });
export const scanStore = createScanStore();
const activeStoreRequests = new WeakMap<ScanStore, symbol>();

const abortError = (): DOMException =>
  new DOMException("The health scan was cancelled", "AbortError");

const throwIfAborted = (signal: AbortSignal | undefined) => {
  if (signal?.aborted === true) throw abortError();
};

const defaultYield = (): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, 0));

const cellRuleIds = [
  "ERR_VALUE",
  "EMBEDDED_CONSTANT",
  "EXTERNAL_LINK",
  "VOLATILE_FUNCTION",
] as const;

const workbookRuleIds = [
  "FORMULA_DISCONTINUITY",
  "HARDCODE_IN_RUN",
  "MANUAL_CALCULATION",
  "INCONSISTENT_FORMAT",
] as const;

const orderFindings = (findings: readonly HealthFinding[]): HealthFinding[] =>
  [...findings].sort((left, right) =>
    left.ref.sheetName.localeCompare(right.ref.sheetName) ||
    left.ref.address.localeCompare(right.ref.address, undefined, { numeric: true }) ||
    left.ruleId.localeCompare(right.ruleId));

export const scanWorkbook = async (
  facts: HealthWorkbookFacts,
  options: ScanOptions = {},
): Promise<readonly HealthFinding[]> => {
  const chunkSize = Math.max(1, options.chunkSize ?? 1_000);
  const now = options.now ?? performance.now.bind(performance);
  const progressDelayMs = options.progressDelayMs ?? 500;
  const store = options.store ?? scanStore;
  const yieldControl = options.yieldControl ?? defaultYield;
  const startedAt = now();
  const findings: HealthFinding[] = [];
  const request = Symbol("health-scan");
  activeStoreRequests.set(store, request);
  store.current = null;
  throwIfAborted(options.signal);

  for (let completed = 0; completed < facts.cells.length; completed += chunkSize) {
    throwIfAborted(options.signal);
    const nextCompleted = Math.min(completed + chunkSize, facts.cells.length);
    findings.push(...runRules(
      { ...facts, cells: facts.cells.slice(completed, nextCompleted) },
      { ruleIds: cellRuleIds },
    ));
    await yieldControl();
    throwIfAborted(options.signal);
    if (
      nextCompleted < facts.cells.length &&
      now() - startedAt >= progressDelayMs
    ) {
      options.onProgress?.({ completed: nextCompleted, total: facts.cells.length });
    }
  }

  for (const ruleId of workbookRuleIds) {
    throwIfAborted(options.signal);
    findings.push(...runRules(facts, { ruleIds: [ruleId] }));
    await yieldControl();
    throwIfAborted(options.signal);
  }
  throwIfAborted(options.signal);
  const ordered = orderFindings(findings);
  if (activeStoreRequests.get(store) === request) store.current = ordered;
  if (now() - startedAt >= progressDelayMs) {
    options.onProgress?.({ completed: facts.cells.length, total: facts.cells.length });
  }
  return ordered;
};

const escapeCell = (value: string): string => value.replaceAll("|", "\\|");

export const exportHealthMarkdown = (
  findings: readonly HealthFinding[],
  title = "Analyst Lens model health report",
): string => {
  const lines = [
    `# ${title}`,
    "",
    `Findings: ${findings.length}`,
    "",
  ];
  for (const current of findings) {
    lines.push(
      `## ${current.ruleId} · ${current.severity}`,
      "",
      `- Cell: ${escapeCell(`${current.ref.sheetName}!${current.ref.address}`)}`,
      `- Evidence: ${escapeCell(current.evidence)}`,
      `- Remediation: ${escapeCell(current.remediation)}`,
      "",
    );
  }
  return lines.join("\n");
};
