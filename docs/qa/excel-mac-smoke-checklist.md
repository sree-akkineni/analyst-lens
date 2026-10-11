# Excel for Mac smoke checklist

> **Public-release redaction notice.** This checklist preserves the procedures
> and completion state of the original local qualification run. All
> workbook-identifying details from the private qualification workbook —
> source path, sheet names, formulas, cell values, and SHA-256 digests — are
> replaced below with synthetic-fixture equivalents or omitted. Use this
> checklist only with a timestamped QA copy. Never open the source workbook
> for write testing.

## Environment and copy gate

- [x] Record macOS and Excel versions: macOS 26.3.1 (25D2128), Excel 16.111 (16.111.26071325).
- [x] Run `npm ci && npm run verify` under Node 22.23.1/npm 10.9.8: 44 files/420 tests and 17/17 browser flows passed.
- [x] Run the uninstall/absent/install/status replacement sequence; final status `healthy`.
- [x] Create and record a fresh copy: `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260717-215319.xlsx`.
- [x] Source and copy SHA-256 matched before Excel.
- [x] Record plan-baseline comparison: the historical baseline remains an unreconciled historical mismatch.
- [x] Confirm the Excel title bar showed `Sample_Model_qa_20260717-215319` before any write.

## Smoke workflows

For each row, record the starting sheet/cell, button or shortcut, elapsed time where budgeted, pass/fail, and a screenshot path.

1. **Formula inspection**
   - Select one formula output.
   - Verify sheet/address, formula, displayed value, number format, direct precedents, and direct dependents.
   - Warm selection-to-inspector target: p95 at or below 600 ms.

2. **Three-level trace**
   - Navigate at least three precedent levels across sheets.
   - Choose a branch and return to origin.
   - Direct navigation target: p95 at or below 500 ms after load; learned workflow target: under 10 seconds.

3. **Expanded graph**
   - Open the full graph.
   - Expand upstream and downstream.
   - Verify cycles and unresolved external/dynamic leaves remain bounded.
   - One-level expansion target: at or below one second for up to 50 new nodes.

4. **Verified formats and revert**
   - Apply and read back input blue, formula black, linked formula green, integer, percentage, multiple, top border, and bottom border.
   - Run **Revert last Analyst Lens action** after every category and verify the original property is restored.
   - Stop immediately on a write/readback or revert mismatch.
   - Exercise the Macabacus-style shortcuts for integer (`Ctrl+Shift+1`), percentage
     (`Ctrl+Shift+5`), currency (`Ctrl+Shift+4`), multiple (`Ctrl+Shift+8`), top border
     (`Ctrl+Option+Shift+Up`), bottom border (`Ctrl+Option+Shift+Down`), and clear borders
     (`Ctrl+Option+Shift+N`).
   - In Settings, use **Apply Macabacus defaults**. Verify occupied shortcuts are listed and
     not replaced until **Use Analyst Lens anyway** is explicitly selected.

5. **Structure commands**
   - Record the exact starting selection. Group its intersecting rows once, repeat on the same or nested rows, ungroup once, and confirm one layer remains; ungroup again and confirm the next layer is removed.
   - Repeat the same one-layer nested workflow for intersecting columns.
   - Create both row and column groups. Run **Close All Groups**, then **Open All Groups**, and confirm visibility changed while group definitions remained available.
   - Confirm the opposite structural command is the documented one-layer manual inverse.
   - Apply and verify a formatting command, then perform a structural mutation. Confirm **Revert last Analyst Lens action** refuses to revert the stale formatting snapshot and directs the user to the corresponding inverse.
   - Exercise multi-area/mixed/protected/outline-limit failure behavior without bypassing workbook safety.
   - Do not claim outline-state readback or arbitrary prior nested-outline reconstruction; Office.js exposes neither.

   Final result: **PASS under the approved structural exception.** Rows `68:70` and columns `AH:AI` each passed repeated one-layer grouping and one-layer inverse ungrouping. Both dimensions survived Close All/Open All. Protected-sheet mutation failed closed. A structural write blocked stale formatting Revert and directed the opposite command. Multi-area/mixed/unsupported and typed Office-error cases passed automated contracts. Excel 16.111 capped the real row outline at level 8 with success/no-op instead of returning an Office error; this is documented as a host limitation, not as verified outline readback.

6. **Model health**
   - Start a scan, cancel once, and verify no partial findings remain.
   - Rerun to completion and inspect exact findings on at least three sheets.
   - Target: at most 20 seconds for up to 100,000 used cells, with cancellation available after 500 ms.

7. **Key-output watchlist**
   - Pin at least three transaction/return outputs.
   - Change one assumption in the QA copy, recalculate, and verify affected-output highlighting.
   - Record whether persistence is session-only or explicitly saved with the workbook.

8. **Excel and host lifecycle**
   - Close and reopen Excel with the QA copy.
   - Verify the task pane loads the current bundle and the loopback host remains healthy.
   - Run `./scripts/analyst-lens repair && ./scripts/analyst-lens status`.

## Floating Quick Trace final matrix — 2026-07-17

Run only against `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260717-093812.xlsx`. The completed results are recorded in `fixture-qa-report.md`.

| # | Start | Action | Required result | Final result |
| --- | --- | --- | --- | --- |
| 1 | `Staffing Schedule!H12`, pane closed | `Control+Option+P` | Floating Trace In opens with formula and direct precedents | **PASS** |
| 2 | H12 Trace In | Up/Down | Dialog focus retained; Excel follows references | **PASS**, 20/20 acknowledgements |
| 3 | H12 Trace In | Right twice on H5 | Three levels, one lazy level per expansion | **PASS**, H12 → H5 → H6 |
| 4 | Deep branch | Left twice | Select parent, then collapse without losing origin | **PASS** |
| 5 | Selected precedent | Return | Selected precedent becomes root; origin remains H12 | **PASS** |
| 6 | Re-rooted branch | `O` and visible **Return to origin** | Both restore H12 and keep dialog open | **PASS** |
| 7 | H12 session | `T` and segmented control | Same dialog switches direction without changing root/origin | **PASS** |
| 8 | `Assumptions!C22`, pane closed | `Control+Option+D` | Floating Trace Out opens with downstream dependents | **PASS** |
| 9 | Open dialog | Drag and resize | Native frame moves/resizes; footer remains reachable | **PASS**, 622×612 → 782×722 |
| 10 | Open dialog | Repeat shortcut | No second dialog | **PASS** |
| 11 | Open dialog | Click `Assumptions!AG57` manually | Current tree remains; explicit new-root action appears | **PASS** |
| 12 | Close/reopen | Shortcut | Fresh session; no stale rows/errors/nonce messages | **PASS** |
| 13 | Settings | Apply defaults with conflicts | List conflicts; require explicit replacement confirmation | **PASS**, confirmation not selected |
| 14 | Entire audit | Close without save and re-hash | No workbook mutation | **PASS**, source and copy unchanged |

- [x] Three-level upstream branch verified in real Excel.
- [x] Three-level downstream branch verified in real Excel.
- [x] Reference and Value are the only columns; no Argument or inferred meaning appears.
- [x] Cycle, external, dynamic, cap, empty, loading, error, and workbook-changed states pass deterministic production-shaped tests.
- [x] Automated budgets pass: 100 loaded moves, 100 cached branch toggles, 20 uncached 50-node expansions, and 100 cached direction toggles.
- [x] Screenshot evidence captured for Trace In, three-level depth, and Trace Out (captures removed for public release; procedures retained).
- [x] `grouping-acceptance` qualification passed in real Excel under the owner-approved manual-inverse exception.

## Integrity and release gate

- [x] Compare audit-only behavior and close the final QA workbook without saving.
- [x] Re-hash the source after QA: it remained unchanged; the historical baseline mismatch remains documented.
- [x] Re-hash the final QA copy: it remained byte-identical.
- [x] Run the uninstall/absence/reinstall/status sequence from the README; final status `healthy`.
- [x] Run `npm run verify` after reinstall under the pinned runtime.
- [x] Attach exact selections/results in `fixture-qa-report.md`.
- [x] Record quick-trace-shortcut-activation as resolved only after a closed task pane reopened from the real Excel shortcut.
- [x] Keep grouping-acceptance open until the approved one-layer grouping contract passes complete real-Excel qualification; qualification passed on 2026-07-17.
- [x] Call the usable local MVP complete only after Office runtime, write safety, grouping acceptance, and source integrity all pass.

Final Quick Trace decision: **PASS**. Final Analyst Lens local MVP decision: **PASS** under the approved one-layer structural manual-inverse exception.
