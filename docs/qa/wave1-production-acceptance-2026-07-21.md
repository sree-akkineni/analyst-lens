# Analyst Lens Wave 1 Production Acceptance — 2026-07-21

> **Public-release redaction notice.** This report preserves the procedures,
> gates, and decision of the original Wave 1 acceptance run. All
> workbook-identifying details from the private qualification workbook —
> source path, sheet names, formulas, cell values, and SHA-256 digests — are
> replaced below with synthetic-fixture equivalents or omitted. Aggregate
> outcome counts are preserved. Native screenshot captures from the original
> run embedded workbook content and were removed; the browser-fixture PNG
> evidence under `docs/qa/ui-modernization/` is synthetic and retained.

## Candidate identity

- Branch: `codex/analyst-lens-ui-modernization`
- Received clean source base: `7e6d7aba7520667258ac4c2e1b4a437c7a901bff` — the initial complete gate reached the final network check and was BLOCKED by the capture-provenance punctuation defect described below
- First integration-review candidate: `920ab02855759afeeb1e4a058eed31595f47b57c`
- Clean fix source SHA: `950e54d5cfd1c0dcd0cd4fe377053097f8bc93cd`
- Task 7 qualified clean automated candidate SHA: `5f513a659e4defb7ea50def5d8ef36a6324c056d`
- Exact native-acceptance and current-head requalification candidate SHA: `79cbcb546ba0c6a6b9762b94993cc37abe35e476`
- Task 8 evidence commit and final-review HEAD: `d7cac1aa614dd93075a2b45eab0cef763655ce88`
- Source status at final review: `PASS — clean worktree; no product-file drift after native acceptance`
- Qualification timestamp: `2026-07-24T02:12:10+0200` (`2026-07-24T00:12:10Z`)
- Native acceptance completed: `2026-07-26` in Excel for Mac `16.111.1`
- Exact native candidate canonical gate rerun: `2026-07-26`
- macOS: `26.3.1` (`25D2128`)
- Installed Microsoft Excel bundle: `16.111.1` (`16.111.26071913`), read from the application bundle without launching Excel
- Node: `v22.23.1`
- npm: `10.9.8`
- Evidence-reporting identity: the later commit containing this report is intentionally not self-referenced. This report qualifies the clean source SHA above.

## Automated gate

| Gate ID | Exact command | Result | Test count | Duration | Log or artifact hash |
| --- | --- | --- | --- | --- | --- |
| ENV-01 | `test "$(node --version)" = "v22.23.1"` | PASS — exact runtime matched | 1 assertion | <1s | Environment log SHA-256 `238e3df5451fc8eb72ed26cbf8b867407067c45bb54628ba0c99c1bd3d6a0b1c` |
| ENV-02 | `test "$(npm --version)" = "10.9.8"` | PASS — exact runtime matched | 1 assertion | <1s | Environment log SHA-256 `238e3df5451fc8eb72ed26cbf8b867407067c45bb54628ba0c99c1bd3d6a0b1c` |
| DEPS-01 | `npm ci` | PASS — 379 packages installed and 380 audited; npm also reported 3 high-severity audit findings | 379 installed packages | 10.63s | Log SHA-256 `2d2dd9ab97538100cb2debb5ae9e75f0e6baa36df394eca32c18020c2c4de636` |
| PREP-01 | `npm run build` | PASS — 132 modules transformed and the production bundle excluded standalone fixture code | 1 production build | 1.69s total; Vite 261ms | Log SHA-256 `a1773c814945a2aa5eb990f09a86bd9f35e7dc5737d8e9256c93aac731422716` |
| FIX-01 | `npm test -- tests/component/Inspector.test.tsx` | PASS — overlay-safe Command-K and focus-return component gate | 1/1 file; 21/21 tests | 1.21s total; Vitest 913ms | Log SHA-256 `b243c97e4d215b9a672c29e6729b3c4b3b3e87f475301956b5326a07d973e5d8` |
| FIX-02 | `npm run test:e2e -- tests/e2e/adaptive-ui.spec.ts tests/e2e/inspector.spec.ts tests/e2e/health-watchlist.spec.ts tests/e2e/quick-trace.spec.ts` | PASS — modal safety plus exact Task 5 binding cross-product | 31/31 Playwright tests; 36 task-pane and 8 Quick Trace axe/overflow matrix cases | 26.98s total; Playwright 26.7s | Log SHA-256 `c7d4f4cd3795d971d7c4c78f5fd18c525f08f8c2b2793aef090df41385edb666` |
| FIX-03 | `npm run typecheck` | PASS — zero TypeScript errors | Not applicable | 1.52s | Log SHA-256 `470fcfdef44e000c2d3875ac8cbf632006370fe8ebe395aa27956dc2c611c4e4` |
| FIX-04 | `npm run build` | PASS — 132 modules transformed and production bundle excluded standalone fixture code | 1 production build | 0.68s total; Vite 90ms | Log SHA-256 `86957526813613f58c6e7f4c9183fef86d1b8d2b40bd9d6b059516ffc5a6075f` |
| FIX-05 | `npm test` | PASS | 45/45 files; 430/430 tests | 33.71s total; Vitest 33.44s | Log SHA-256 `259ec8be1738fd3693fc9d8c6b692d85b009311fc62f33180b2f6246c4712557` |
| CAPTURE-01 | `npm run build:e2e` | PASS — fixture bundle built from clean fix commit | 134 transformed modules | 0.56s total; Vite 61ms | Log SHA-256 `731aec56553fa21a9ab876d595f38f25b1a276e06766c9564554b1305acd0d3a` |
| CAPTURE-02 | `npm run qa:capture` | PASS — 12/12 images regenerated and inspected | 12 captures | 6.31s | Log SHA-256 `9bce1fffca2102160d48927b06f389ffd6c76330a1a19d9c97f0be86a78fcc89` |
| VERIFY-00 | `npm run verify` | PASS — complete canonical gate from clean `5f513a659e4defb7ea50def5d8ef36a6324c056d` | 45/45 Vitest files, 430/430 Vitest tests, 34/34 Playwright tests; no skipped test reported | 57.21s | Log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-01 | `npm run lint` via `npm run verify` | PASS — zero warnings/errors | Not applicable | Included in 57.21s; not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-02 | `npm run typecheck` via `npm run verify` | PASS — zero TypeScript errors | Not applicable | Included in 57.21s; not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-03 | `npm test` via `npm run verify` | PASS | 45/45 files; 430/430 tests | 28.15s | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-04 | `npm run test:e2e` via `npm run verify` | PASS | 34/34 Playwright tests | 23.8s | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-05 | `npm run build` via `npm run verify` | PASS — production bundle excluded standalone fixture code | 132 transformed modules | Vite 56ms; full subgate not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-06 | `office-addin-manifest validate manifest.xml` via `npm run verify:manifest` | PASS — XML manifest valid | 1 manifest | Included in 57.21s; not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-07 | `office-addin-manifest validate manifest.json` via `npm run verify:manifest` | PASS — unified manifest valid | 1 manifest | Included in 57.21s; not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-08 | `npm run verify:office-boundary` via `npm run verify` | PASS — `src/platform/office` remains the only Office.js caller | 1 boundary scan | Included in 57.21s; not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-09 | `npm run verify:network` via `npm run verify` | PASS — only loopback plus the approved Office.js and MicrosoftAjax runtime URLs were allowed | 1 network scan | Included in 57.21s; not separately timed | Verify log SHA-256 `d918d4d780825dede8c937142b1af4cdf3e4ab88d425cc83105a1a096cbff0b0` |
| VERIFY-CURRENT-00 | `npm ci` at exact HEAD `79cbcb546ba0c6a6b9762b94993cc37abe35e476` | PASS — 379 packages installed and 380 audited; npm reported four high-severity audit entries, retained below as a distribution risk | 379 installed packages | Not separately timed | `docs/qa/wave1-native-observations-2026-07-26.md` |
| VERIFY-CURRENT-01 | `npm run verify` at exact HEAD `79cbcb546ba0c6a6b9762b94993cc37abe35e476` | PASS — lint, typecheck, unit, E2E, production build, both manifests, Office boundary, network allowlist, and canonical screenshot checksums | 46/46 Vitest files; 431/431 Vitest tests; 34/34 Playwright tests; no skipped test reported | Not separately timed | `docs/qa/wave1-native-observations-2026-07-26.md` |
| VERIFY-CURRENT-02 | `./scripts/analyst-lens status` plus post-build source/installed asset comparison | PASS — status `healthy`; exact normalized 12-file hash lists matched | 12 assets | <1s | Normalized SHA-256 `58c5a7d3165895f167434e3c96b690eedf581975f29cb358a1407d0e89a7e6e3` |
| VERIFY-FINAL-00 | `npm ci` at exact final-review HEAD `d7cac1aa614dd93075a2b45eab0cef763655ce88` | PASS — 379 packages installed and 380 audited; four high-severity audit entries remain disclosed below | 379 installed packages | 11s | Task 9 clean-verification record |
| VERIFY-FINAL-01 | `npm run verify` at exact final-review HEAD `d7cac1aa614dd93075a2b45eab0cef763655ce88` | PASS — lint, typecheck, unit, E2E, production build, both manifests, Office boundary, network allowlist, and canonical screenshot checksums | 46/46 Vitest files; 431/431 Vitest tests; 34/34 Playwright tests; no skipped test reported | Vitest 27.11s; Playwright 23.4s; Vite 62ms | Task 9 clean-verification record |
| VERIFY-FINAL-02 | `./scripts/analyst-lens status` and `shasum -a 256 -c docs/qa/ui-modernization/current/manifest.sha256` | PASS — lifecycle status `healthy`; all 12 canonical capture hashes matched | 1 lifecycle status; 12 captures | Not separately timed | Task 9 clean-verification record |
| VERIFY-FINAL-03 | `git diff --check` and `git status --short --branch` | PASS — no whitespace errors and clean branch before closeout documentation | 2 repository checks | <1s | Task 9 clean-verification record |

The first complete run exposed a trailing-period defect in capture-script provenance text and ended BLOCKED at `verify:network`. That focused repair was proven RED then GREEN without weakening the verifier, committed separately, and followed by a clean-source recapture. The later integration review of the first candidate identified the two Important findings below. The fix, capture-provenance, and final Task 7 complete PASS run superseded the earlier automated counts. The later native shortcut/lifecycle changes exceeded the frozen Wave 1 boundary and were functionally removed by a reviewed rollback; the accepted candidate retains only the awaited Office-ready bootstrap and its focused regression test. The complete canonical gate was rerun against the exact native candidate before Task 8 evidence closure.

## Integration review and bounded fix wave

- **BLOCKED — first integration review:** the first candidate had two open Important findings.
  1. Inspector's pane-scoped Command-K handler could open the command palette while graph or settings was already modal, producing two simultaneous `aria-modal` dialogs and moving focus away from the original modal.
  2. Task 5's axe coverage did not bind every required state to both themes and both required viewports/directions; the required cross-product was therefore not proven.
- **RED — modal component:** `npm test -- tests/component/Inspector.test.tsx` failed the three graph/palette/settings suppression cases (3 failed, 18 passed) because `openOverlay("palette")` was dispatched.
- **RED — modal E2E:** `npm run test:e2e -- tests/e2e/inspector.spec.ts` failed graph and settings with two dialogs (2 failed, 8 passed).
- **GREEN — modal safety:** Inspector now suppresses only its pane-scoped Command-K handler whenever any Inspector overlay is active. Editable-field suppression and normal Command-K opening remain covered. Component coverage passes 21/21 and graph/palette/settings E2E requires exactly one dialog, focus retained in the original modal, and focus returned to its invoker after close.
- **GREEN — exact axe matrix:** formula, hardcode, blank, error, multi, command palette open, dependency map open, Model health open/completed, and Key outputs with a pin each pass light/dark at 360×900 and 760×900. Quick Trace In and Out each pass light/dark at 420×390 and 900×700. Every case retains the unchanged WCAG 2 A/AA serious-or-critical filter, checks zero horizontal overflow, and checks state-required controls/reachability.
- **PASS — independent integration re-review:** a fresh reviewer approved the fixed head with no Critical or Important findings and returned `Proceed to Task 8: Yes`.

## Final whole-branch review

- Review range: Task 0 parent through Task 8 evidence (historical SHAs in Candidate identity above).
- **Critical:** none.
- **Important:** none.
- **Minor:** stale “primary input” CSS comment; ambiguous visible `◎` / `⋯` controls; duplicate inner Key outputs heading; four high-severity dependency audit entries.
- The fresh reviewer checked the approved plan and both design specifications, the full diff, acceptance evidence, accessibility, privacy/network boundaries, Office architecture, automated tests, captures, installed identity, native Excel behavior, workbook integrity, and claim scope.
- Explicit reviewer decision: **READY FOR TASK 9 CLOSEOUT**.

## Browser matrix

These are browser fixture captures, not Excel evidence. All 12 PNGs were regenerated from the clean fix source, freshly inspected, and remained byte-identical to the prior set.

| ID | Surface | Fixture/state | Viewport | Theme | Expected | Observed | PASS/PARTIAL/BLOCKED | Screenshot path | SHA-256 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| B-01 | Task pane | `sample-model`, formula selection | 360×900 | Light | Reference/gloss text remains in its track; values and badges remain separate | Byte-identical recapture; no overlap or clipping in the inspected state | PASS | `docs/qa/ui-modernization/current/light/taskpane-narrow-formula.png` | `d21bf5dc48118bb8736034b2aed93f0572fac27fe0bf992738102aa3c996662b` |
| B-02 | Task pane | `sample-model`, hardcode selection | 360×900 | Light | No clipping, dead content, or unreachable control | Byte-identical recapture; inspected layout and controls remain intact | PASS | `docs/qa/ui-modernization/current/light/taskpane-narrow-hardcode.png` | `37ba271b7e52f511c3afa81ec90fc358d19ab6367b67eb9bf399c1f8f9c2443c` |
| B-03 | Task pane | `sample-model`, formula selection | 760×900 | Light | Source/reference/value alignment, compact map, and controls remain intact | Byte-identical recapture; inspected wide layout remains intact | PASS | `docs/qa/ui-modernization/current/light/taskpane-wide-formula.png` | `ff478b7a9d5ae2d3e757b3658a4f7f639a41a879c0d4d834bdbcdedde0927e31` |
| B-04 | Quick Trace In | `sample-model`, precedent direction | 420×390 | Light | Selected-reference action, tree, and footer remain reachable | Byte-identical recapture; inspected compact layout remains intact | PASS | `docs/qa/ui-modernization/current/light/quick-trace-in-420x390.png` | `63cd7f7f3d18c7b3bd18a5d4593e3066f32b7bd1246f6c4331bc16479bd36a2a` |
| B-05 | Quick Trace In | `sample-model`, precedent direction | 900×700 | Light | Tree, detail action, footer, badges, and contrast remain intact | Byte-identical recapture; inspected expanded layout remains intact | PASS | `docs/qa/ui-modernization/current/light/quick-trace-in-900x700.png` | `0d0f446e283974d856fa092b9fcb47e5293af5f49ac8b7ef73588281f70c20f4` |
| B-06 | Quick Trace Out | `sample-model`, dependent direction | 420×390 | Light | Selected-reference action remains fully above the footer | Byte-identical recapture; inspected compact action remains fully visible | PASS | `docs/qa/ui-modernization/current/light/quick-trace-out-420x390.png` | `ca023f2db0b2020284377e711c9ad1ba93ff5a7fc321c441df5615e31daa8b5e` |
| B-07 | Task pane | `sample-model`, formula selection | 360×900 | Dark | Reference/gloss text remains in its track; values and badges remain separate | Byte-identical recapture; no overlap or clipping in the inspected state | PASS | `docs/qa/ui-modernization/current/dark/taskpane-narrow-formula.png` | `c4aca9537406f81136eae8a98c674d98f1368168577163137e976b9195ad6972` |
| B-08 | Task pane | `sample-model`, hardcode selection | 360×900 | Dark | No clipping, dead content, or unreachable control | Byte-identical recapture; inspected layout and controls remain intact | PASS | `docs/qa/ui-modernization/current/dark/taskpane-narrow-hardcode.png` | `3b6eb60b3765e8c1e563a622efd56f782154f7b41a78654ca1c5cccd97670fc0` |
| B-09 | Task pane | `sample-model`, formula selection | 760×900 | Dark | Source/reference/value alignment, compact map, and controls remain intact | Byte-identical recapture; inspected wide layout remains intact | PASS | `docs/qa/ui-modernization/current/dark/taskpane-wide-formula.png` | `83ca1134016aef77c72402f8f2c29043156fb0319e7c4c2aba969f5423b0d2d2` |
| B-10 | Quick Trace In | `sample-model`, precedent direction | 420×390 | Dark | Selected-reference action, tree, and footer remain reachable | Byte-identical recapture; inspected compact layout remains intact | PASS | `docs/qa/ui-modernization/current/dark/quick-trace-in-420x390.png` | `4dcc8334fef99f906090f2cbee2f62746067142dd4e4437fbd10a056b477b4a8` |
| B-11 | Quick Trace In | `sample-model`, precedent direction | 900×700 | Dark | Tree, detail action, footer, badges, and contrast remain intact | Byte-identical recapture; inspected expanded layout remains intact | PASS | `docs/qa/ui-modernization/current/dark/quick-trace-in-900x700.png` | `ae2cc05b2da0eb42020ce639cd69146bfb4731210eed459b62497d69add40e54` |
| B-12 | Quick Trace Out | `sample-model`, dependent direction | 420×390 | Dark | Selected-reference action remains fully above the footer | Byte-identical recapture; inspected compact action remains fully visible | PASS | `docs/qa/ui-modernization/current/dark/quick-trace-out-420x390.png` | `fe98ac64570b964c79b0e78829b3eac6b5d4d7169c2320d5be2dcf70983bd859` |

Capture manifest source: the clean fix commit in Candidate identity above. Capture manifest SHA-256: `6d0967ef489e8f76998ea70672b3f7dfc150d17dcbeeae35b9ca0cd93f2d764e`. Capture checksum-list SHA-256: `38e295ceaeed83df21f05062f0da4b4770f2a119a79ca111126f2f40d0b4d993`.

## Installed-build identity

PASS — uninstall, absent-status, install, status, repair, and final status
completed in the required order. The
final status is `healthy`. Manifest ID
`5db4c51e-9c38-4d8b-976b-2d0f707e27f5`, version `1.0.1.0`, and
manifest SHA-256
`301468cfd09180c42899d964df2f8227ac5cdd1404d8fb1a949833701158fbbf`
were installed with the exact candidate.

`diff -qr dist "$HOME/Library/Application Support/Analyst Lens/release/assets"`
and the normalized sorted hash diff were empty. Both 12-file checksum lists
have normalized SHA-256
`58c5a7d3165895f167434e3c96b690eedf581975f29cb358a1407d0e89a7e6e3`.

| Asset | Source SHA-256 | Installed SHA-256 | Result |
| --- | --- | --- | --- |
| `assets/jsx-runtime-5Zz-bvlC.js` | `fbd4d6dc0bfe21802769969439efe736711386c1bfb10780d3cff14d99f1bf19` | `fbd4d6dc0bfe21802769969439efe736711386c1bfb10780d3cff14d99f1bf19` | PASS |
| `assets/quick-trace.css` | `021ead8b40d50c0c875281e9d9456ba5e6b383823aa8bea40363a2d27d3e6071` | `021ead8b40d50c0c875281e9d9456ba5e6b383823aa8bea40363a2d27d3e6071` | PASS |
| `assets/quick-trace.js` | `3786c78addc0750be2879f587dea8ce80057a3705392812bca8f401d2dbbfc4b` | `3786c78addc0750be2879f587dea8ce80057a3705392812bca8f401d2dbbfc4b` | PASS |
| `assets/taskpane.css` | `afcc522e71dc0d3fed9281eae408888b2a2dc78c3b802d3607e523231e15ab84` | `afcc522e71dc0d3fed9281eae408888b2a2dc78c3b802d3607e523231e15ab84` | PASS |
| `assets/taskpane.js` | `17adcadd7c9f9cdcb86fc9ab30952a5a1ec61fc4d55d787449a8b1232d37d2f7` | `17adcadd7c9f9cdcb86fc9ab30952a5a1ec61fc4d55d787449a8b1232d37d2f7` | PASS |
| `icons/icon-16.png` | `d3c8da63b8adffb7d67c301e18b65b8e6cb5775ea05ae5b2bc18be378eb3de1a` | `d3c8da63b8adffb7d67c301e18b65b8e6cb5775ea05ae5b2bc18be378eb3de1a` | PASS |
| `icons/icon-192.png` | `8c6e27cc9d4f0bb46ee4d839a837b65a920e821dc37e2760c25c9bac7619cc06` | `8c6e27cc9d4f0bb46ee4d839a837b65a920e821dc37e2760c25c9bac7619cc06` | PASS |
| `icons/icon-32.png` | `8308007a0eed7ac61969d20747fd747dfe6be3172e95a2e704b7b5325c2fdf09` | `8308007a0eed7ac61969d20747fd747dfe6be3172e95a2e704b7b5325c2fdf09` | PASS |
| `icons/icon-80.png` | `5c26c1fa86660a664e2708c8b5c322f5cddc07ecba78e10d95fb15701edf2289` | `5c26c1fa86660a664e2708c8b5c322f5cddc07ecba78e10d95fb15701edf2289` | PASS |
| `quick-trace.html` | `43359b21b46e3a4486cd6886048aeff73b4c4fec1d14fcde89870c112b918558` | `43359b21b46e3a4486cd6886048aeff73b4c4fec1d14fcde89870c112b918558` | PASS |
| `shortcuts.json` | `f58e663f1b0230f2f5ff1b8fab1ed90071923a857e8da6756faafa0f3c13ab14` | `f58e663f1b0230f2f5ff1b8fab1ed90071923a857e8da6756faafa0f3c13ab14` | PASS |
| `taskpane.html` | `27cb10219fdb3245a7e27394634ada45cfd3c33711614ff4dce7811ef5716e8c` | `27cb10219fdb3245a7e27394634ada45cfd3c33711614ff4dce7811ef5716e8c` | PASS |

## Real Excel matrix

The protected source workbook was never opened. All native actions used the
timestamped audit or retention copy named in the Integrity section.
Transient AX observations are retained in
`docs/qa/wave1-native-observations-2026-07-26.md` at SHA-256
`2c1dcacdfb4ded23e0188777922637635ce7622ad2723c22355cca8d2791dcee`;
matrix references to the observation log bind to that hash.
Native captures from the original run embedded workbook content and were
removed for public release; capture procedures are preserved per row.

| ID | QA copy | Exact sheet/cell | Action/shortcut | Expected | Observed | PASS/PARTIAL/BLOCKED | Elapsed | Evidence |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| EXCEL-01 | Audit | Task pane and Quick Trace | Switch macOS Appearance between Light and Dark | Both native surfaces follow the host appearance without an app-specific toggle. | Both surfaces adapted automatically; dark appearance was restored after capture. | PASS | Not separately timed | Observation log; capture procedure: adaptive task pane + Quick Trace in light/dark |
| EXCEL-02 | Audit | Key outputs and Model health disclosures | Inspect closed visual and accessibility states | Closed bodies are absent visually and from AX while their disclosure buttons remain operable. | Closed bodies were absent visually and from AX; the buttons remained present with `collapsed` state. | PASS | Not separately timed | Observation log |
| EXCEL-03 | Audit | `Staffing Schedule!H12` | Open command launcher by click and Command-K; close with Escape | Search receives focus, one modal remains active, Escape closes it, and focus returns to the invoker. | Search received focus; Escape closed the palette and focus returned to `Open commands for this cell, Command-K`. | PASS | Not separately timed | Observation log; capture procedure: command-launcher focus |
| EXCEL-04 | Audit | `Staffing Schedule!H12` | Formula inspection | Exact value, format, formula, precedents, and dependents are reported without mutating the cell. | Value `0`, format `#,##0;(#,##0);-`, exact formula, four precedent groups, and dependent `H31` were reported. | PASS | Not separately timed | Observation log; capture procedure: format/revert state |
| EXCEL-05 | Audit | `G16 → H12 → Assumptions!C22` | Three-cell navigation and Return to origin | Breadcrumb preserves all three cells and Return to origin restores G16. | A fresh session was initialized at G16; the exact chain was present; Return to origin restored `Staffing Schedule!G16`, confirmed by the pane heading and Excel name box. | PASS | Not separately timed | Observation log |
| EXCEL-06 | Audit | `H12`; `Assumptions!C22` | Floating Trace In and Trace Out | Native floating traces report the expected cross-sheet precedent and dependent ranges. | Trace In showed H5, C3, F12:G12, and `Assumptions!C22`; Trace Out from `Assumptions!C22` showed H10:AE30, H34:AE34, and H38:AE39. | PASS | Not separately timed | Observation log; capture procedure: Quick Trace light/dark |
| EXCEL-07 | Audit | Floating Quick Trace | Move, resize, focus, expand, reroot, switch direction, reuse, close/reopen | One dialog supports all in-session actions and reopening starts a fresh current-selection session without stale expansion. | Dialog moved/resized 622×612→852×702; H12→H5→H6 expanded; reroot and direction reuse stayed in one dialog; reopen used a new session and no stale expansion. | PASS | Not separately timed | Observation log |
| EXCEL-08 | Audit | Workbook health scan | Cancel, then rerun | Cancellation keeps no partial findings; rerun completes across at least three sheets. | Cancel returned the exact no-partial-results message; rerun completed with 1,976 findings including Validation, Deferred Sales Schedule, and Assumptions. | PASS | Not separately timed | Observation log; capture procedure: health/watchlist state |
| EXCEL-09 | Audit | `AE16`, `AE31`, `AE42`; `Assumptions!C22` | Pin three outputs and inspect impact | Three session-local pins retain exact values and the shared assumption reports three affected outputs without workbook persistence. | Values 17455.8, 298700.4, and 401555.2 were pinned; `Assumptions!C22` reported `3 affected by current selection`; persistence stayed disabled. | PASS | Not separately timed | Observation log; capture procedure: health/watchlist state |
| EXCEL-10 | Audit | `H12` | Apply Input blue, verify command readback, Revert | Verified format applies, Revert becomes available, Revert restores the prior format, and the formula is unchanged. | Exact apply and Revert messages appeared; the formula remained unchanged. | PASS | Not separately timed | Observation log; capture procedure: format/revert state |
| EXCEL-11 | Audit | Rows `68:70`; columns `AH:AI` | Group, one-layer inverse, Close/Open All, protected failure, structural Revert | Every group command changes one layer; inverse removes one layer; visibility commands preserve definitions; protection fails closed; Revert explains the structural boundary. | Row and column behavior matched; protection returned the exact fail-closed message and was immediately removed. | PASS | Not separately timed | Observation log; capture procedure: grouping retention |
| EXCEL-12 | Audit/Retention | Pane closed; representative same-sheet and cross-sheet selections | Command+Option+L and selection changes | Once the add-in is initialized for the workbook, the pane-closed shortcut reopens it and selections refresh selected-cell evidence. | After Developer Add-ins initialization, pane-closed Command+Option+L reopened Analyst Lens and the named same-sheet/cross-sheet selections refreshed the pane. The separate cold-workbook pre-initialization error is recorded below as a distribution risk. | PASS | Not separately timed | Observation log |
| RET-01 | Retention | `Staffing Schedule!H12`; representative cells and workbook structure | Format, save, close, reopen, calculate, clean up, save, reopen; compare formulas/constants/structure | No repair prompt; calculation succeeds; exact formulas, constants, sheets, outlines, filters, and freeze boundary remain semantically intact; only the controlled style cycle persists. | Reopened without repair; 10 sheets, 13,343 formulas, 8,576 constants, 379 row and 6 column outline records, defined names, filters, and H29 freeze boundary matched the audit copy exactly under a shared semantic digest; representative formulas matched and H12 was cleaned to explicit black. | PASS | Not separately timed; three reopen cycles | Observation log |

## Integrity

| Artifact | Pre hash and size | Post hash and size | Semantic reopen result |
| --- | --- | --- | --- |
| Source workbook — `$ANALYST_LENS_QA_SOURCE_FILE` — never opened | `redacted` | Same hash | PASS — custody only; never opened or write-tested |
| Audit copy — `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260726-175842-audit.xlsx` | `redacted` | Same hash after the final **Don't Save** close | PASS — byte-exact |
| Retention copy — `$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260726-175842-retention.xlsx` | `redacted` | Differed only by intentional save, cleanup, and restoration of the original H29 freeze boundary | PASS — no repair prompt; Calculate Now completed; normalized semantic digest exactly matched the audit copy across 10 sheets, 13,343 formulas, 8,576 constants, defined names, filters, and non-default row/column outlines; final H12 font color is explicit black; no watchlist persistence was enabled |

## Exceptions and risks

| Status | Exact unsupported behavior or risk | Scope | Mitigation | Owner-approved status |
| --- | --- | --- | --- | --- |
| PARTIAL | The visible `◎` / `⋯` command glyphs remain ambiguous without labels | Browser presentation only | Retain for final triage; do not treat the glyphs as native Excel qualification | PARTIAL — deferred during Task 6 review; no broader owner sign-off recorded |
| PARTIAL | Current `npm audit` reports four high-severity entries: three unique advisories (`adm-zip`, `brace-expansion`, `fast-uri`) plus the direct `office-addin-manifest` entry inherited from `adm-zip` | Installed dependency graph; not part of the canonical `npm run verify` contract | Triage dependency paths before any distribution or broader launch decision; the `office-addin-manifest` suggested fix is a semver-major downgrade and must not be applied blindly | PARTIAL — untriaged; no owner waiver recorded |
| PARTIAL | The first Command+Option+L attempt in each cold workbook context returned Excel's stale `Add-in Error`; after one Developer Add-ins launch, pane-closed Command+Option+L passed | Native activation/distribution readiness | Keep lifecycle and shortcut-registration paths frozen in Wave 1; require explicit manual initialization in the current candidate and resolve in a separately approved distribution-hardening wave | PARTIAL — no public-beta or unattended-distribution claim |
| PASS | The one-layer grouping limitation remains explicit: automatic Revert does not roll structure back, and each opposite command removes one layer | Native grouping semantics | Retain the visible inverse instruction and do not claim arbitrary structural rollback | PASS — limitation accepted in the approved plan |
| PASS | Installed payload identity, post-initialization real Excel behavior, source custody, audit integrity, and retention semantic reopen are qualified for the exact candidate above | Native Task 8 gate | Preserve the evidence and run Task 9 final verification/review | PASS |
| PASS | The first integration review's two Important findings were fixed, covered, and independently approved | Task 7 Step 5 | Proceed to Task 8 against the preserved candidate identities | PASS — no Critical or Important findings; `Proceed to Task 8: Yes` |
| PASS | The final whole-branch review found no Critical or Important issue across the exact Task 0-parent-to-Task 8-evidence range | Task 9 review gate | Retain the four Minor items in the deferred hardening queue | PASS — `READY FOR TASK 9 CLOSEOUT` |

## Decision

**PARTIAL — Wave 1 adaptive production hardening is closed, but distribution readiness is not approved.** Tasks 0–9 are complete. All automated checks, capture checksums, installed-build identity, post-initialization real-Excel audit and retention matrices, source/audit custody, semantic reopen integrity, and the fresh whole-branch review pass. The protected source workbook was never opened. The decision remains PARTIAL because cold-workbook Command+Option+L requires one manual Developer Add-ins initialization and the remaining dependency and UI-clarity limitations have no owner waiver. The documented next gate is public-beta packaging and distribution-hardening design; this report does not claim public-beta launch or unattended-distribution readiness.
