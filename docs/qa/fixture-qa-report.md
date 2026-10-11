# Sample Model fixture qualification report — 2026-07-15/17

> **Public-release redaction notice.** This report preserves the procedures,
> checklists, and outcomes of the original local qualification run. All
> workbook-identifying details from the private qualification workbook —
> source path, sheet names, formulas, cell values, and SHA-256 digests — are
> replaced below with synthetic-fixture equivalents or omitted. Aggregate
> outcome counts and timing budgets are preserved. Native screenshot captures
> from the original run embedded workbook content and were removed; the
> capture procedures that produced them are retained in
> [excel-mac-smoke-checklist.md](excel-mac-smoke-checklist.md).

## Outcome

**Analyst Lens local MVP PASS.** The owner-approved one-layer grouping contract passed real Excel 16.111 qualification for rows, columns, nested layers, one-layer manual inverses, non-destructive Open/Close visibility, protected-sheet failure, and the structural automatic-Revert barrier. The earlier inspector, verified formatting, scan, watchlist, shortcut, and floating Quick Trace evidence remains valid, and the pane-closed Quick Trace shortcut was rechecked on the final grouping build. `grouping-acceptance` and `quick-trace-shortcut-activation` are accepted for this local build; no publication or deployment is claimed.

## Environment

| Item | Observed |
| --- | --- |
| macOS | 26.3.1 (25D2128) |
| Microsoft Excel for Mac | 16.111 |
| Node.js / npm | 22.23.1 / 10.9.8 |
| Automated gate | PASS — 44 test files, 420 tests, 17/17 browser flows, production build, manifest validation, fixture exclusion, Office boundary, and network checks all green |
| Final grouping QA workbook | `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260717-215319.xlsx` |
| Final audit-only QA workbook | `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260717-093812.xlsx` |
| Earlier write/revert QA workbook | `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260715-220420.xlsx` |
| Source workbook | `$ANALYST_LENS_QA_SOURCE_FILE` (see [SETUP.md](SETUP.md)) |

## Integrity evidence

| Check | Result |
| --- | --- |
| Source before QA | `redacted` |
| Final grouping QA copy before/after | `redacted` — PASS, byte-identical, closed without saving |
| Source after continued QA | `redacted` — PASS, unchanged |
| Final audit-only QA copy before/after | `redacted` — PASS, byte-identical and unchanged after closing without saving |
| Required plan baseline | `redacted` — mismatch existed before Excel opened |
| Earlier on-disk write/revert QA copy after save | `redacted` |

The source path is not present in the committed tree, so there is no committed blob to reconcile against the planning-time baseline. The source file was read and copied but never opened in Excel.

Excel changed the QA-copy package when it was saved. A read-only structural comparison found zero entered-value/formula changes, zero sheet-name changes, and zero worksheet-protection changes, while cached values and format serializations changed through Excel recalculation/normalization. All continued write testing targeted only this QA copy, and every completed Analyst Lens write was reverted before the final save.

## Installed-build evidence

The final production and installed assets are byte-identical:

- final task-pane HTML SHA-256: `27cb10219fdb3245a7e27394634ada45cfd3c33711614ff4dce7811ef5716e8c`
- final task-pane JavaScript SHA-256: `9a59cf31ac99edd924bb5c6a22ccfc747577695c0fd5fbe91b5beb4ec73d4822`
- final task-pane CSS SHA-256: `da3ca0e1522937936f50f31b89dc294bf6162113af55eea1f5fa6fc7d65a039d`
- final Quick Trace HTML SHA-256: `43359b21b46e3a4486cd6886048aeff73b4c4fec1d14fcde89870c112b918558`
- final Quick Trace JavaScript SHA-256: `d7d00d75d8d666075daf31794b6f02c9a07d7bcfb170d8ba2d72d05854481921`
- final Quick Trace CSS SHA-256: `3b2a9811ff4231c9e989769b8071a12e6d693f56f6d8a262f87683d5729e0dda`
- final 21-action shortcut catalog SHA-256: `f58e663f1b0230f2f5ff1b8fab1ed90071923a857e8da6756faafa0f3c13ab14`

The final lifecycle sequence reported `healthy`, `uninstalled`, `installed`, and `healthy`. Excel then loaded the replacement payload in a fresh process. The readable production shortcut catalog contains 21 actions and is 4,107 bytes.

## Resolved Office-runtime blocker

Web Inspector showed that the previous content-security policy allowed only the literal `office.js` URL. Office.js then tried to load sibling runtime resources such as `excel-mac-16.00.js`, `en-us/office_strings.js`, telemetry under the hosted runtime root, and the legacy `MicrosoftAjax.js` bootstrap; the policy blocked them before `Office.context` and `Excel` could initialize.

The local host now permits only the hosted Office runtime root and exact MicrosoftAjax bootstrap in addition to local assets. The static network verifier permits those exact literals but still rejects unapproved paths on the same hosts. After reload, `Office.context` and `Excel` were present, selection inspection initialized, and the shortcut action reopened the pane.

## Workflow matrix

| # | Workflow | Starting selection | Result | Evidence |
| --- | --- | --- | --- | --- |
| 1 | Formula output inspection | `Staffing Schedule!H12` | **PASS** | Inspector showed value `0`, format `#,##0;(#,##0);-`, formula `=$F12*(H$5>=$C$3)*$G12*(1+Assumptions!$C$22)/12`, cross-sheet and embedded-number risks, four direct precedents, direct dependent `H31`, and matching neighbors. |
| 2 | Three-level cross-sheet precedent trace and return | Origin `Staffing Schedule!G16` | **PASS** | Breadcrumb reached `G16 → H12 → Assumptions!C22`; **Return to origin** restored `G16`. |
| 3 | Lazy upstream/downstream graph | `Staffing Schedule!G16` | **PASS for real branch expansion** | Expanded `G16 → H16:AE16 → H31:AE31 → AF31:AG31 / H42:AE42 / H57:AE57`. No cycle, external leaf, or dynamic reference was present in the chosen real branch; standalone E2E covers those states. |
| 4 | Visible formatting with verified readback and revert | `Staffing Schedule!G16`, `Assumptions!C22`, and `Staffing Schedule!H12` | **PASS for exercised categories** | Formula black, linked green, warning red, integer, percentage, currency, multiple, date, top border, bottom border, and Input blue each passed independent apply/readback/revert checks. For the final Input blue replay, the saved H12 font changed from black to `FF0000FF`; revert preserved the formula and restored displayed/settable black (`FF000000`). |
| 5 | Row/column group, ungroup, and revert | Rows `68:70`; columns `AH:AI` | **PASS under the approved structural exception** | Repeated Group created nested layers; each opposite command removed one layer; Open/Close preserved definitions; protected sheets failed closed; structural history blocked stale format Revert and directed the one-layer inverse. |
| 6 | Health scan cancel/rerun and three-sheet findings | QA workbook | **PASS** | Fast cancel produced `Scan cancelled. No partial findings were kept.` Rerun completed with 1,976 findings across `Validation`, `Deferred Sales Schedule`, and `Assumptions`, including inconsistent formats, embedded constants, and formula discontinuities. |
| 7 | Three-output watchlist and assumption impact | `AE16`, `AE31`, `AE42`; assumption `Assumptions!C22` | **PASS after range-aware fix** | Values `17455.8`, `298700.4`, and `401555.2` were pinned session-locally. Selecting `Assumptions!C22` marked all three outputs affected. Workbook persistence stayed opt-in and disabled. |
| 8 | Excel/task-pane/host lifecycle and shortcut | QA copy | **PASS** | Healthy uninstall/reinstall/status, add-in reload, pane close, and `Command+Option+L` reopening all succeeded. Inspector returned fully initialized after the shortcut. |
| 9 | Macabacus-style Mac shortcut catalog | Settings, visible SpeedLayer controls, pane closed, and `H12` selected | **PASS for catalog load, pane activation, and representative format dispatch** | Excel loaded all 19 actions with private shortcut telemetry `FValid:true`. `Command+Option+L` reopened the closed pane, and `Control+Option+Shift+B` dispatched Input blue to H12. Automated tests cover every mapping and action association; the QA run did not press all 19 chords individually. |

## One-layer structure qualification — 2026-07-17

The final grouping run used only `Sample_Model_qa_20260717-215319.xlsx`. The source and copy were both hashed before Excel opened. Excel was closed without saving after the structural, failure, format/revert, and Quick Trace regression cases; the source and copy hashes still matched.

| Case | Exact selection | Result |
| --- | --- | --- |
| Row nesting | `Staffing Schedule!68:70` | **PASS** — first Group added one bracket; second Group exposed outline levels `1 2 3`. |
| Row inverse | `Staffing Schedule!68:70` | **PASS** — first Ungroup returned the sheet to levels `1 2` with a bracket remaining; second removed the added grouping. |
| Column nesting | `Staffing Schedule!AH:AI` | **PASS** — first and second Group added one layer each and exposed `1 2 3`. |
| Column inverse | `Staffing Schedule!AH:AI` | **PASS** — the two Ungroup calls removed one layer each and restored the baseline. |
| Close/Open visibility | Row `68:70` plus columns `AH:AI` | **PASS** — Close All hid the details while plus controls remained; Open All restored details and the same definitions. Neither operation entered write history. |
| Structural Revert barrier | Input blue on `H12`, then Group rows on `68:70` | **PASS** — Revert refused to restore the stale formatting snapshot and displayed `Automatic Revert is unavailable. Select the affected rows and run Ungroup rows once.` |
| Protected sheet | Protected active sheet, then Group rows on `68:70` | **PASS** — no mutation; the pane displayed `Structure changes are unavailable on a protected sheet.` The sheet was immediately unprotected. |
| Maximum depth | Repeated Group rows on `68:70` | **HOST LIMIT OBSERVED** — Excel exposed outline buttons through level 8, then returned success/no-op rather than an Office error. The app preserves Office errors as typed fail-closed results when supplied, but cannot distinguish this host no-op without the outline getter that Office.js does not expose. |
| Formatting regression | Input blue then Revert on `H12` | **PASS** — verified readback and automatic Revert remained operational. |
| Quick Trace regression | Pane closed, `H12`, `Control+Option+P` | **PASS** — a fresh floating Trace In dialog showed the exact formula and five reference/value nodes. |

The raw capture sequence was recorded outside the repository and is not committed. The compressed repository captures from the original run embedded workbook content and were removed for public release.

## Floating Quick Trace final qualification — 2026-07-17

The final run used only the timestamped, initially byte-identical QA copy. The source workbook was never opened. The dialog remained an Office floating window throughout; no docked fallback was used. The UI exposed only **Reference** and **Value** columns and performed no workbook writes.

| # | Start | Action | Elapsed / observation | Result | Evidence |
| --- | --- | --- | --- | --- | --- |
| 1 | `Staffing Schedule!H12`, pane closed | `Control+Option+P` | Dialog visible in the next accessibility snapshot | **PASS** | Trace In opened with the exact H12 formula and direct precedents H5, C3, F12:G12, and `Assumptions!C22`. |
| 2 | H12 Trace In | Up/Down | 20 alternating visible row acknowledgements; raw Computer Use plus macOS accessibility p95 507 ms | **PASS** | Every sample alternated H12/H5, Excel followed the published reference, and the focused window stayed `Analyst Lens`. The 507 ms includes the external accessibility-tree query; deterministic controller timing is recorded below. |
| 3 | H12 Trace In | Right on H5, then Right to H6 | Uncached visible expansion acknowledgements 468–515 ms | **PASS** | H12 → H5 → H6 appeared lazily; each Right expanded at most one level. |
| 4 | H6 deep branch | Left twice | Immediate keyboard transitions | **PASS** | First Left selected H5; second collapsed H5 without changing origin. |
| 5 | Selected H5 | Return | Immediate re-root followed by Office navigation | **PASS** | Root became H5 while origin remained H12. |
| 6 | Re-rooted H5 | `O`, then visible **Return to origin** after an H6 re-root | Both controls kept the dialog open | **PASS** | Each restored the original root; the explicit button was exercised separately from the mnemonic. |
| 7 | H12 session | `T`, then segmented control | Same session identity | **PASS** | Trace In/Out switched without changing root or origin and without opening another dialog. |
| 8 | `Assumptions!C22`, pane closed | `Control+Option+D` | Dialog visible in the next accessibility snapshot | **PASS** | Trace Out showed H10:AE30, H34:AE34, and H38:AE39. |
| 9 | Open Trace Out | Drag title bar; resize lower-right corner | 622×612 → 782×722 in the sampled run | **PASS** | Native frame moved and resized; footer actions remained present and reachable. |
| 10 | Open dialog | Press `Control+Option+D` again | Same URL session and one Excel dialog | **PASS** | No duplicate window opened. |
| 11 | C22 tree open | Manually click `Assumptions!AG57` | Current tree retained | **PASS** | Explicit **Use Assumptions $AG$57 as new root** appeared; no implicit re-root occurred. |
| 12 | Close C22 dialog | Reopen with `Control+Option+D` | New opaque session identity | **PASS** | Fresh AG57 root, new nonce, one node, and no stale rows, errors, or prior nonce messages. |
| 13 | Inspector Settings | **Apply Macabacus defaults** | Conflict result returned without replacement | **PASS** | Both `Ctrl+Option+P` and `Ctrl+Option+D` were listed and gated by **Use Analyst Lens anyway**. The confirmation was not selected. |
| 14 | Entire audit run | Close without saving; SHA-256 compare | Source and QA copy identical before/after | **PASS** | Source and copy hashes matched before and after. |

Downstream depth also passed in the same real workbook: `Assumptions!C22 → Staffing Schedule!H10:AE30 → H31:AE31 → AF31:AG31 / H42:AE42 / H57:AE57`. The selected branch remained bounded and every expansion added only one dependency level.

### Performance qualification

The production performance suite ran in the final verified build and used more samples than the manual minimum:

| Budget | Samples | Result |
| --- | --- | --- |
| Loaded Quick Trace row movement p95 ≤100 ms | 100 | **PASS** |
| Cached one-level expand/collapse p95 ≤250 ms with no new Office reads | 100 | **PASS** |
| Uncached one-level expansion of 50 nodes p95 ≤1 second | 20 | **PASS**; every state contained 52 rows including root and parent |
| Cached direction toggle p95 ≤250 ms with no new Office reads | 100 | **PASS** |

The real-Excel keyboard sampling confirmed 20/20 correct row acknowledgements and no focus loss. Its raw 507 ms p95 is not treated as product execution time because each sample includes a separate out-of-process macOS accessibility-tree round trip. The in-product controller budget above is the reproducible application measurement; real Excel supplied functional and focus evidence.

### Native captures

The integrated Excel captures from the original run isolated the active Office dialog, but the uncropped live-desktop verification still exposed workbook content, so all JPEG evidence was removed for public release. The capture procedure — Trace In, three-level depth, Trace Out, move/resize, and one-dialog identity — is preserved in the smoke checklist.

One intentionally rapid automation sequence sent another keyboard command before Excel finished the prior selection request and briefly displayed **A newer Office request superseded this result** with **Retry selected reference**. The notice was bounded, the next stable command cleared it, and the fresh close/reopen session contained no stale notice, rows, or nonce messages. No Office error code, focus loss, workbook write, or dialog crash occurred.

## Fixed defects

### AL-QA-001 — Office gateway never becomes ready in Excel 16.111

**Resolved.** Root cause was an over-restrictive CSP that allowed `office.js` but blocked the sibling scripts that it loads. The corrected allowlist is regression-tested and passed real Excel verification.

### AL-QA-003 — Selected assumption misses pins contained by coalesced dependent ranges

**Resolved.** Excel reports the direct dependents of `Assumptions!C22` as ranges such as `H10:AE30`. Exact-ID matching could not recognize an individual pin such as `AE16` inside that range. The controller now performs same-workbook/same-sheet A1 containment matching; a failing regression test was added before the fix. Real Excel then changed the count from zero to three affected outputs.

### AL-QA-004 — Border revert succeeds in mocks but fails through Excel Mac proxies

**Resolved.** Excel Mac ignored border restoration through `setCellProperties`, leaving the
workbook changed and correctly safety-locking later Analyst Lens writes. The gateway now
restores every captured cell edge through Excel's direct border proxies. A focused failing
regression preceded the fix; top-border and bottom-border apply/revert then passed in the QA
copy without engaging the safety lock.

### AL-QA-005 — Punctuation chord invalidates the Excel shortcut catalog

**Resolved.** The original Clear Borders mapping ended in `-`. Excel rejected an 18-action
catalog containing that chord even when the JSON was compacted to 2,978 bytes, while a
17-action catalog without it parsed successfully. Replacing only that terminal key with a
letter made 18 actions valid, and the final readable 19-action, 3,678-byte catalog parses with
`FValid:true`. Clear Borders now uses the Mac mnemonic `Control+Option+Shift+N`; Clear Formats
uses `Control+Option+Shift+C`.

The final Excel diagnostic log is:
`~/Library/Containers/com.microsoft.Excel/Data/Library/Logs/Diagnostics/EXCEL/Primary1784233273268553000_D5547FBA-B9EF-4798-A596-0BF25694AAF8.log`.
It records the 3,678-byte fetch, valid parse, pane shortcut dispatch, Input blue dispatch, and
successful action execution.

## Remaining limits

- Structural Group/Ungroup intentionally has no generic automatic Revert. The opposite one-layer command is the supported inverse because Office.js exposes no outline membership/level getter and Analyst Lens does not claim arbitrary nested reconstruction.
- At maximum row-outline depth, Excel 16.111 returned success/no-op instead of an Office error. The typed fail-closed error path remains covered for hosts that report an error, but this host no-op cannot be identified as a failure without unsupported readback.
- The planning-time source hash does not match the source file observed before QA. This run proves the observed source stayed unchanged, but cannot reconcile the older baseline.
- The selected real dependency branches did not contain a cycle, external leaf, dynamic reference, or cap event. Those states, plus empty/loading/error/workbook-changed handling, passed in the deterministic production-shaped fixture and contract suites.
- An already-healthy `install` or `repair` is not an upgrade operation. Local qualification upgrades require uninstall/install and an add-in reload or Excel restart.
- Excel's displayed/settable font-color readback reports automatic black as black. The Input blue revert restored that property and the rendered color, but saving normalized H12 from an implicit/automatic black encoding to explicit `FF000000`. The formula was unchanged. Exact package encoding is therefore not promised for a verified formatting write/revert on the disposable QA copy.
- During file selection, Excel's recent list briefly opened an unrelated newer workbook revision; it was immediately closed without editing or saving. The protected source workbook was never opened. This event did not involve either the source artifact or the final QA copy, but is retained here as honest run evidence.

## Completion decision

The Analyst Lens usable local MVP is accepted. `grouping-acceptance` and `quick-trace-shortcut-activation` have real-Excel evidence, and the inspector, graph, scan, watchlist, verified formatting writes/reverts, 21-action shortcut catalog, pane shortcut, floating Quick Trace, one-layer grouping, manual structural inverses, outline visibility, and local lifecycle are qualified. This decision is local-only and does not claim a push, PR, publication, deployment, or generic structural rollback.
