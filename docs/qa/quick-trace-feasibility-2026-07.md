# Floating Quick Trace feasibility — July 2026

> **Public-release redaction notice.** This report preserves the procedures,
> checklists, and outcomes of the original local feasibility run. All
> workbook-identifying details from the private qualification workbook —
> source path, sheet names, formulas, cell values, and SHA-256 digests — are
> replaced below with synthetic-fixture equivalents or omitted.

> **Post-closeout status — 2026-07-17:** The `grouping-acceptance` references below record the earlier Quick Trace checkpoint. The subsequent one-layer grouping qualification passed in real Excel and closed the usable local MVP; see `fixture-qa-report.md` for current status and evidence.

## Outcome

**Task 1 feasibility PASSED.** A bounded dialog-host lifecycle correction now clears the live Quick Trace session when the host closes the dialog, even when Excel for Mac does not emit `DialogEventReceived`. After a clean production build and reinstall, both pane-closed shortcuts opened a native floating dialog, the dialog moved and resized, and it retained keyboard focus while Excel navigated.

All ten mandatory real-Excel checks passed on a timestamped byte-identical qualification-workbook copy. No docked fallback was used. This accepts only the Floating Quick Trace feasibility gate and authorizes Tasks 2–7; it does not complete the Analyst Lens MVP. Grouping acceptance remains unresolved separately.

## Test identity

| Field | Evidence |
| --- | --- |
| Date/time | 2026-07-17, approximately 08:37–08:50 MDT |
| macOS | 26.3.1 (25D2128) |
| Excel for Mac | 16.111 (16.111.26071325) |
| Required Node / npm | 22.23.1 / 10.9.8 |
| Installed payload | `./scripts/analyst-lens status` reported `healthy` after clean build and reinstall |
| Installed XML identity | `1.0.1.0`; source and installed manifest hashes matched |
| Installed task-pane bundle | Source and installed bundle hashes matched |
| Deterministic starting cell | `Staffing Schedule!H12` |
| Visible formula | `=$F12*(H$5>=$C$3)*$G12*(1+Assumptions!$C$22)/12` |
| Shortcuts exercised | `Control+Option+P` and `Control+Option+D` |
| Tracking-issue feasibility blocker | quick-trace-shortcut-activation, resolved by this evidence |
| Overall MVP blocker | grouping-acceptance, still unresolved |

## Safe-copy and hash integrity

Only the timestamped QA copy below was opened in Excel. The source workbook was never opened or write-tested.

| Artifact | Path | Before SHA-256 | After SHA-256 | Result |
| --- | --- | --- | --- | --- |
| Source workbook | `$ANALYST_LENS_QA_SOURCE_FILE` | `redacted` | `redacted` | Unchanged |
| Timestamped QA copy | `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260717-083729.xlsx` | `redacted` | `redacted` | Byte-identical and unchanged after **Don't Save** |

The current source and QA-copy hashes match each other. The historical planning baseline still does not match this run's source hash. That historical mismatch remains unreconciled evidence; this run does not explain or supersede it.

## Root cause and correction

The earlier `1.0.1.0` manifest-identity correction repaired Excel's stale ActionService catalog and made both Quick Trace actions available. A narrower lifecycle bug remained: `createQuickTraceDialogHost` retained its `live` session until Excel raised `DialogEventReceived`. Excel for Mac does not reliably raise that event after the host calls `dialog.close()`. The next shortcut therefore reused a dead session and completed without displaying a window.

Temporary visible lifecycle telemetry proved that both the original and subsequent shortcut handlers resolved; the missing window was downstream of action activation. An official Microsoft keyboard-shortcut sample comparison also confirmed that the handlers should return promises and do not require `event.completed()`.

The RED contract test closes a session without emitting a native dialog event and requires the next `open()` to create a distinct session. The production correction adds one idempotent finalizer shared by programmatic close and native close. It clears the host's live/opening state, listeners, and buffered messages before calling native `dialog.close()`. The focused suite and full verification then passed.

Real Excel falsified the earlier shared-runtime hypothesis: pane-closed actions were executing. Instrumented proof showed fresh sessions after both HTML close and native window close. Session identities were opaque UUIDs; all three transitions produced a distinct fresh session.

The temporary telemetry was removed before the clean production build and final acceptance run.

## Final real-Excel acceptance run

1. Built and reinstalled the clean production payload; installed manifest and task-pane bundle hashes matched their source artifacts.
2. Created the timestamped byte-identical QA copy and opened only that copy in Excel.
3. With the task pane closed, pressed `Control+Option+P`. A native floating Trace In dialog opened at the current workbook selection without a docked fallback.
4. Selected `Staffing Schedule!H12`, confirmed the formula above, and pressed P. The dialog opened a fresh session with root `H12` and direct precedents `H5`, `C3`, `F12:G12`, and `Assumptions!C22`.
5. Repeated P. Excel retained exactly one dialog and the same session identity.
6. Moved the native frame by its title bar and resized it from approximately 622×612 to 1294×768 pixels using its native zoom control.
7. Pressed Down and Up. The selected dialog row changed from `H12` to `H5` and back while `Analyst Lens` remained the active window and Excel followed the published selection.
8. Pressed Left and Right. The tree collapsed to the root and expanded back to five rows without an accidental Excel cursor move.
9. Pressed `T` to switch Trace In to Trace Out and used the segmented control to switch back. Both preserved origin `H12`.
10. Selected dialog row `H5`; origin `H12` and the five-row expanded tree remained intact.
11. Switched to Excel and manually selected unrelated cell `H20`, then returned to the dialog. The session, origin `H12`, selected row `H5`, and expanded state were retained, and **Use selection as new root** appeared.
12. Closed the dialog and pressed `Control+Option+D` with the task pane closed. A fresh session opened at new root `H20` in Trace Out with dependent `H31`; no stale H12 data remained.
13. Opened Settings and chose **Apply Macabacus defaults** only far enough to run conflict detection. It explicitly listed `Ctrl+Alt+P` and `Ctrl+Alt+D` as already in use and required **Use Analyst Lens anyway**. That confirmation was not selected, so no silent replacement occurred.
14. Closed the workbook, chose **Don't Save**, and re-hashed source and copy. Both retained identical content and SHA-256.

## Ten-check feasibility record

| # | Required check | Result | Evidence |
| --- | --- | --- | --- |
| 1 | With the task pane closed, both chords are accepted or show an explicit conflict, and accepted `Control+Option+P` opens Quick Trace | **PASS** | P cold-started the native floating dialog with the pane closed. D subsequently opened after the prior dialog was closed, also with the pane closed. |
| 2 | Exactly one floating dialog; repeated shortcut reuses it | **PASS** | Repeated P retained one window and session. |
| 3 | Native floating frame moves and resizes over Excel | **PASS** | The macOS title-bar drag moved it; native zoom resized it from approximately 622×612 to 1294×768. No docked fallback was used. |
| 4 | Up/Down retains dialog keyboard focus while Excel selects published precedents | **PASS** | Down selected `H5`; Up returned to `H12`; the active window remained `Analyst Lens`. |
| 5 | Left/Right collapses and expands the direct tree without accidental Excel cursor movement | **PASS** | Left reduced the tree to root-only; Right restored all five rows while selection stayed coherent. |
| 6 | `T` and the segmented control switch Trace In/Out while preserving origin | **PASS** | `T` produced Trace Out with dependent `H31`; the segmented control restored Trace In; origin remained `H12`. |
| 7 | Dialog-driven selection events preserve origin, row, and expanded state | **PASS** | Selecting row `H5` retained origin `H12` and the five-row tree. |
| 8 | Unrelated manual Excel selection retains the tree and shows **Use selection as new root** | **PASS** | Manual Excel selection `H20` retained session/origin/row/expanded state and displayed the required action. |
| 9 | Close and reopen via `Control+Option+D` creates a fresh identity without stale data | **PASS** | D opened a fresh session at `H20`, Trace Out, with `H31`. |
| 10 | Settings reports both chord conflicts without silent replacement | **PASS** | Conflict detection explicitly listed `Ctrl+Alt+P` and `Ctrl+Alt+D` and gated replacement behind **Use Analyst Lens anyway**, which was not selected. |

## Screenshots

The original native captures (pane-closed Trace In dialog rooted at `H12` with only **Reference** and **Value** columns, plus the same dialog after move/resize) embedded workbook content and were removed for public release.

The final clean run was additionally verified through Excel's live accessibility state, including the exact session identities, retained active-window focus, tree state, unrelated-selection notice, and explicit shortcut-conflict list recorded above.

## Automated evidence

- RED: `tests/contract/quickTraceDialogHost.test.ts` failed because a programmatically closed session without a native event was reused by the next `open()`.
- GREEN: the focused dialog-host suite passed (1 file / 5 tests) after adding the idempotent local finalizer.
- Full `npm run verify` under Node 22.23.1/npm 10.9.8 passed: 42 test files / 335 tests and 11/11 browser flows, followed by lint, typecheck, production build, fixture exclusion, XML and unified manifest validation, Office-boundary verification, and network verification.
- Lifecycle checks and installed-payload health passed with Excel fully quit.
- Production output retained the separate same-origin `quick-trace.html`, JavaScript, and CSS entry assets.
- The UI/protocol slice remains audit-only and exposes only **Reference** and **Value** columns.

## Gate decision

The mandatory Task 1 feasibility gate is accepted. Pane-closed shortcut activation, floating move/resize, retained dialog focus, all ten checks, and source/copy integrity passed in real Excel. Tasks 2–7 may proceed task-by-task under strict RED-GREEN-REFACTOR with one focused verified commit per numbered task.

This is not an Analyst Lens MVP completion claim. Grouping acceptance remains the unresolved overall-MVP blocker, and the historical workbook-hash mismatch remains unreconciled evidence.

## Task 7 continuation qualification

The completed Tasks 2–6 build was reinstalled and qualified again on Excel 16.111 using the fresh copy `Sample_Model_qa_20260717-093812.xlsx`. The final run preserved every Task 1 result and added:

- three-level Trace In traversal H12 → H5 → H6, keyboard collapse, re-root, `O`, and the visible **Return to origin** control;
- three-level Trace Out traversal `Assumptions!C22 → H10:AE30 → H31:AE31 → AF31:AG31 / H42:AE42 / H57:AE57`;
- native move/resize from 622×612 to 782×722 with footer controls retained;
- same-session direction switching and one-dialog reuse;
- unrelated manual selection notice at `Assumptions!AG57`;
- fresh close/reopen session and nonce with no stale rows or errors;
- Settings conflict output for both final letter-based chords, without selecting the replacement confirmation;
- 44 test files / 400 tests and 16/16 browser flows, plus production build, manifest, fixture-exclusion, Office-boundary, and network gates;
- final installed Quick Trace HTML/JS/CSS and shortcut hashes recorded in `fixture-qa-report.md`;
- unchanged source and QA-copy hashes after closing without saving.

Checkpoint decision: **Quick Trace PASS / Analyst Lens MVP PARTIAL at the time of this gate**. The historical planning-baseline mismatch remains unreconciled. The subsequent grouping closeout resolved grouping acceptance for the usable local MVP under the approved one-layer manual-inverse exception.
