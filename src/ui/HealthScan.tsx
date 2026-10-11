import { useEffect, useMemo, useRef, useState } from "react";
import { exportHealthMarkdown, type ScanProgress } from "../domain/health/scan";
import type { CellRef, HealthFinding } from "../shared/types";

export type HealthScanRunner = (options: Readonly<{
  onProgress(progress: ScanProgress): void;
  signal: AbortSignal;
}>) => Promise<readonly HealthFinding[]>;

type HealthScanProps = {
  exportLocal(markdown: string): void;
  navigate(ref: CellRef): void;
  runScan: HealthScanRunner;
};

const findingsPageSize = 100;

export function HealthScan({ exportLocal, navigate, runScan }: HealthScanProps) {
  const active = useRef<AbortController | null>(null);
  const mounted = useRef(true);
  const [findings, setFindings] = useState<readonly HealthFinding[] | null>(null);
  const [progress, setProgress] = useState<ScanProgress | null>(null);
  const [page, setPage] = useState(0);
  const [rule, setRule] = useState("all");
  const [severity, setSeverity] = useState("all");
  const [sheet, setSheet] = useState("all");
  const [status, setStatus] = useState("Ready to scan on explicit command.");
  const [running, setRunning] = useState(false);

  const rules = useMemo(() => [...new Set(findings?.map((item) => item.ruleId) ?? [])], [findings]);
  const sheets = useMemo(() => [...new Set(findings?.map((item) => item.ref.sheetName) ?? [])], [findings]);
  const visible = findings?.filter((item) =>
    (severity === "all" || item.severity === severity) &&
    (rule === "all" || item.ruleId === rule) &&
    (sheet === "all" || item.ref.sheetName === sheet)) ?? [];
  const pageStart = page * findingsPageSize;
  const pagedFindings = visible.slice(pageStart, pageStart + findingsPageSize);
  const firstVisible = visible.length === 0 ? 0 : pageStart + 1;
  const lastVisible = Math.min(pageStart + findingsPageSize, visible.length);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      active.current?.abort();
    };
  }, []);

  const start = () => {
    const controller = new AbortController();
    active.current = controller;
    setFindings(null);
    setProgress(null);
    setPage(0);
    setRule("all");
    setSeverity("all");
    setSheet("all");
    setRunning(true);
    setStatus("Scanning workbook…");
    void runScan({
      onProgress: (nextProgress) => {
        if (mounted.current && active.current === controller) setProgress(nextProgress);
      },
      signal: controller.signal,
    }).then((results) => {
      if (controller.signal.aborted || !mounted.current) return;
      setFindings(results);
      setStatus(`Scan complete. ${results.length} findings.`);
    }).catch((error: unknown) => {
      if (!mounted.current) return;
      const normalizedAbort = typeof error === "object" && error !== null &&
        "code" in error && error.code === "ABORTED";
      if (
        controller.signal.aborted ||
        (error instanceof DOMException && error.name === "AbortError") ||
        normalizedAbort
      ) {
        setFindings(null);
        setStatus("Scan cancelled. No partial findings were kept.");
        return;
      }
      setFindings(null);
      setStatus(error instanceof Error ? error.message : "The health scan failed.");
    }).finally(() => {
      if (active.current === controller) active.current = null;
      if (mounted.current) setRunning(false);
    });
  };

  return (
    <div className="health-scan">
      <div className="section-heading">
        <p className="health-scan-hint" role="status">
          {status}
          {progress === null ? "" : ` ${progress.completed} of ${progress.total} cells.`}
        </p>
        {running ? (
          <button onClick={() => active.current?.abort()} type="button">Cancel health scan</button>
        ) : (
          <button onClick={start} type="button">Run model health scan</button>
        )}
      </div>

      {findings === null ? null : (
        <>
          <div className="health-filters">
            <label>Severity<select aria-label="Severity" onChange={(event) => {
              setSeverity(event.currentTarget.value); setPage(0);
            }} value={severity}>
              <option value="all">All</option><option value="error">Error</option><option value="warning">Warning</option><option value="info">Info</option>
            </select></label>
            <label>Rule<select aria-label="Rule" onChange={(event) => {
              setRule(event.currentTarget.value); setPage(0);
            }} value={rule}>
              <option value="all">All</option>{rules.map((item) => <option key={item}>{item}</option>)}
            </select></label>
            <label>Sheet<select aria-label="Sheet" onChange={(event) => {
              setSheet(event.currentTarget.value); setPage(0);
            }} value={sheet}>
              <option value="all">All</option>{sheets.map((item) => <option key={item}>{item}</option>)}
            </select></label>
          </div>
          <p className="health-result-count">
            Showing {firstVisible}–{lastVisible} of {visible.length} filtered findings ({findings.length} total).
          </p>
          <ul aria-label="Health findings" className="health-findings">
            {pagedFindings.map((item) => (
              <li key={`${item.ruleId}-${item.ref.sheetId}-${item.ref.address}`}>
                <button
                  aria-label={`${item.ref.sheetName} ${item.ref.address} ${item.ruleId}`}
                  onClick={() => navigate(item.ref)}
                  type="button"
                >
                  <strong>{item.ruleId}</strong>
                  <span>{item.ref.sheetName} {item.ref.address} · {item.severity}</span>
                  <span>{item.evidence}</span>
                  <span>{item.remediation}</span>
                </button>
              </li>
            ))}
          </ul>
          <div className="health-pagination">
            <button
              aria-label="Previous findings page"
              disabled={page === 0}
              onClick={() => setPage((current) => Math.max(0, current - 1))}
              type="button"
            >Previous</button>
            <button
              aria-label="Next findings page"
              disabled={lastVisible >= visible.length}
              onClick={() => setPage((current) => current + 1)}
              type="button"
            >Next</button>
          </div>
          <button
            onClick={() => exportLocal(exportHealthMarkdown(visible))}
            type="button"
          >Export local Markdown</button>
        </>
      )}
    </div>
  );
}
