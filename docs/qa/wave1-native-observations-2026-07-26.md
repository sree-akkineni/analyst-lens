# Wave 1 native observation log — 2026-07-26

> **Public-release redaction notice.** This log preserves the procedures and
> observations of the original native acceptance run. All workbook-identifying
> details from the private qualification workbook — source path, sheet names,
> formulas, cell values, and SHA-256 digests — are replaced below with
> synthetic-fixture equivalents or omitted. Aggregate outcome counts are
> preserved. Native screenshot captures from the original run embedded
> workbook content and were removed.

## Scope and custody

- Exact candidate: `79cbcb546ba0c6a6b9762b94993cc37abe35e476`
- Excel: `16.111.1` (`16.111.26071913`)
- Manifest registration ID: `5db4c51e-9c38-4d8b-976b-2d0f707e27f5`
- Protected source, never opened:
  `$ANALYST_LENS_QA_SOURCE_FILE` (see [SETUP.md](SETUP.md))
- Timestamped audit copy:
  `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260726-175842-audit.xlsx`
- Timestamped retention copy:
  `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260726-175842-retention.xlsx`
- Source and final audit-copy SHA-256 matched before and after the run.
- Final cleaned retention SHA-256 differed only by the intended saved changes.

The audit copy was closed with **Don't Save** after the final evidence
recapture. Excel was then quit. The protected source was not selected in
Excel's recent-workbook list or opened at any point.

## Exact-candidate automated qualification

The canonical gate was rerun on the exact candidate HEAD with
`/opt/homebrew/opt/node@22/bin` first on `PATH`.

- Node `v22.23.1`; npm `10.9.8`.
- `npm ci`: PASS, 379 packages installed and 380 audited in 12 seconds.
- `npm run lint`: PASS, zero warnings or errors.
- `npm run typecheck`: PASS, zero TypeScript errors.
- `npm test`: PASS, 46/46 files and 431/431 tests.
- `npm run test:e2e`: PASS, 34/34 Playwright tests in 25.3s.
- `npm run build`: PASS, 132 modules transformed; Vite 63ms; standalone
  fixture code absent.
- XML and unified manifest validation: PASS.
- Office boundary: PASS.
- Network allowlist: PASS.
- Post-build lifecycle status: `healthy`.
- Post-build `diff -qr` between `dist` and the installed release assets:
  empty.

The same `npm ci` reported four high-severity package-audit entries:
`adm-zip`, `brace-expansion`, `fast-uri`, and the direct
`office-addin-manifest` entry inherited from `adm-zip`. This is recorded as a
distribution risk, not hidden inside the canonical gate result.

## Native AX and visible observations

These entries preserve the exact observations used for Real Excel matrix rows
whose interactive accessibility snapshots were otherwise transient.

| ID | Retained observation |
| --- | --- |
| EXCEL-02 | With Key outputs and Model health closed, the AX tree retained the disclosure buttons with `collapsed` state and omitted the disclosure-body containers and controls. The corresponding task-pane light/dark captures show the closed visual state. |
| EXCEL-03 | Clicking `Open commands for this cell` and pressing Command-K each opened one command-palette modal with its search field focused. Escape closed the modal, and AX focus returned to `Open commands for this cell, Command-K`. |
| EXCEL-04 | At `Staffing Schedule!H12`, Excel's formula bar and Analyst Lens reported value `0`, format `#,##0;(#,##0);-`, formula `=$F12*(H$5>=$C$3)*$G12*(1+Assumptions!$C$22)/12`, four precedent groups (`H5`, `C3`, `F12:G12`, `Assumptions!C22`), and dependent `H31`. |
| EXCEL-05 | A fresh Analyst Lens session was initialized with `Staffing Schedule!G16` selected. The breadcrumb then advanced `G16 → H12 → Assumptions!C22`; Return to origin restored `Staffing Schedule!G16`, confirmed by the pane heading and Excel name box. |
| EXCEL-06 | Floating Trace In at H12 showed H5, C3, F12:G12, and `Assumptions!C22`. Floating Trace Out at `Assumptions!C22` showed `Staffing Schedule!H10:AE30`, `H34:AE34`, and `H38:AE39` in one native dialog. |
| EXCEL-07 | Quick Trace moved and resized from 622×612 to 852×702. The same dialog session expanded H12→H5→H6, rerooted to H5 while retaining H12 as origin, changed direction, and reused the single dialog. Closing and reopening produced a different opaque session ID, used the current selection as root, and contained no stale expansion. |
| EXCEL-08 | Fast cancel returned `Scan cancelled. No partial findings were kept.` The rerun returned `Scan complete. 1976 findings.` AX results included Validation, Deferred Sales Schedule, and Assumptions; the final wide capture retains the completed result. |
| EXCEL-09 | Session-local pins retained `AE16 = 17455.8`, `AE31 = 298700.4`, and `AE42 = 401555.2`. Selecting `Assumptions!C22` through the evidence link reported `3 affected by current selection`. Workbook persistence stayed disabled. |
| EXCEL-10 | Input blue returned `Input blue applied. Revert is available.` Revert returned `Last Analyst Lens action reverted.` The exact H12 formula remained unchanged. |
| EXCEL-11 | Rows 68:70 and columns AH:AI each gained one outline layer; each matching inverse removed one layer. Open/Close All changed visibility without deleting definitions. Structural Revert showed the opposite-command instruction. Protect Sheet returned `Structure changes are unavailable on a protected sheet`; the sheet was immediately unprotected. The final wide grouping capture retains the one-layer message and fully legible command labels. |
| EXCEL-12 | After Analyst Lens was initialized from Excel's Developer Add-ins tile, pane-closed Command+Option+L reopened it and selection changes refreshed the pane. On the first shortcut attempt in each cold workbook context, Excel instead showed `Add-in Error This add-in is no longer available`; no shortcut-registration or lifecycle workaround was added because those paths are frozen outside Wave 1. |

## Retention semantic reopen

Excel reopened the cleaned retention copy without a repair or recovery prompt.
The native accessibility tree reported:

- workbook title
  `Sample_Model_qa_20260726-175842-retention`;
- `Staffing Schedule`, sheet 6 of 10;
- a frozen-pane workbook view;
- selected cell H29 and formula
  `=$F29*(H$5>=$C$3)*$G29*(1+Assumptions!$C$22)/12`.

The package was then compared read-only with the byte-exact audit copy using a
normalized semantic signature covering:

- ordered sheet names, visibility, and dimensions;
- defined names;
- every formula and every non-formula constant;
- merged ranges;
- non-default row and column hidden/outline/collapsed state;
- freeze panes and autofilters.

Both files produced identical semantic SHA-256 digests
with exact equality:

- 10 sheets;
- 13,343 formulas;
- 8,576 constants;
- 379 outlined/hidden row records;
- 6 outlined/hidden column records;
- `Staffing Schedule` freeze boundary H29.

Representative formula/value readback after Excel Calculate Now:

| Sheet/cell | Formula or constant | Cached value |
| --- | --- | --- |
| `Staffing Schedule!H12` | `=$F12*(H$5>=$C$3)*$G12*(1+Assumptions!$C$22)/12` | `0` |
| `Staffing Schedule!G16` | `195020` | `195020` |
| `Staffing Schedule!H31` | `=SUM(H10:H30)` | `0` |
| `Staffing Schedule!AE16` | `=$F16*(AE$5>=$C$3)*$G16*(1+Assumptions!$C$22)/12` | `17455.8` |
| `Staffing Schedule!AE31` | `=SUM(AE10:AE30)` | `298700.4` |
| `Staffing Schedule!AE42` | `=AE31+AE35+AE40` | `401555.2` |
| `Assumptions!C22` | `0.15` | `0.15` |
| `Validation!H21` | `=IF(ABS(F21)<=G21,"PASS","FAIL")` | `PASS` |

The only intended persisted cell-style change was the controlled H12
font-color cycle. Saved-blue package readback was `FF0000FF`; final cleanup is
explicit black `FF000000`. No watchlist persistence was enabled.

## Retained native screenshots

The original native captures embedded workbook content and were removed for
public release. The capture procedure — adaptive task pane and Quick Trace in
light/dark, command-launcher focus, health/watchlist, format revert, and
grouping retention — is preserved in the acceptance matrix below and in
[wave1-production-acceptance-2026-07-21.md](wave1-production-acceptance-2026-07-21.md).
