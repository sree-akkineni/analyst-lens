# Analyst Lens

Analyst Lens is a local-only Excel for Mac add-in for formula inspection, dependency navigation, verified formatting commands, model-health review, and a key-output watchlist. Workbook data stays on the Mac. The only non-loopback runtime dependencies are Microsoft's hosted Office.js runtime resources and its exact legacy MicrosoftAjax bootstrap.

## The problem

Excel for Mac has the modeling engine but not the analyst control surface: no keyboard-first formatting layer, no navigable precedent/dependent history through branching cross-sheet chains, and no fast way to triage an inherited model for hardcodes, broken runs, or hidden dependencies. Analyst Lens is that missing layer — built for people who learned modeling on Windows with Macabacus-style muscle memory and now work on Mac.

## What got built

- **Hybrid Inspector** — selected-cell formula, value, format, direct precedents/dependents with values, risk badges, neighbor comparison, and a lazy 500-node dependency graph with origin breadcrumb and Back/Forward history.
- **Quick Trace** — a floating, movable, resizable Office dialog (not a pane overlay) for keyboard-only precedent/dependent branch walks: Up/Down navigate, Right expands exactly one level, Return re-roots, `T` switches direction, `O` returns to origin. Audit-only; two columns (Reference, Value), no invented semantics.
- **Speed layer** — one-keystroke formatting (input blue, formula black, linked green, warning red, borders, decimals), plus one-layer Group/Ungroup with the opposite command as the documented manual inverse (Office.js exposes no outline-state getter, so generic automatic Revert is honestly refused for structure).
- **Verified writes** — every formatting write does preflight → snapshot → write → readback, with one **Revert last Analyst Lens action**. A structural mutation installs a Revert barrier that fails closed instead of restoring a stale snapshot.
- **Model health + watchlist** — cancellable structural scan (errors, discontinuities, hardcodes, embedded constants, external links, volatiles, …) with exact cell evidence, and session-local key-output pins with assumption-impact highlighting.
- **Local install lifecycle** — loopback-only HTTPS host, user `launchd` agent, owned XML manifest registration, and idempotent `install` / `status` / `repair` / `uninstall` with byte-identical installed-payload verification.

## What's interesting

- **Honest boundaries instead of fake coverage.** Where Office.js can't read state back (outline membership, max-depth no-ops), the product says so in the UI and the QA reports record host limits as observations, not passes.
- **Evidence-first QA.** Every qualification run mints a timestamped byte-identical workbook copy, hashes source and copy before/after, and records exact selections, on-screen copy, and screenshot hashes. See [docs/qa/SETUP.md](docs/qa/SETUP.md) and the reports linked below.
- **A fixture-shaped automated gate.** 420 unit/component/contract tests plus 17 browser flows run against a synthetic workbook fixture; production builds are statically verified to exclude fixture code, to keep all Office.js calls behind `src/platform/office`, and to allow no unexpected outbound destinations.
- **The floating-dialog lifecycle bug.** Excel for Mac doesn't reliably raise `DialogEventReceived` after the host closes a dialog, so the next shortcut reused a dead session and showed nothing. Fixed with one idempotent finalizer shared by programmatic and native close — see [docs/qa/quick-trace-feasibility-2026-07.md](docs/qa/quick-trace-feasibility-2026-07.md).

## Supported qualification environment

- Microsoft Excel for Mac 16.111
- Node.js 22.23.1 and npm 10.9.8
- Local HTTPS host at `https://127.0.0.1:43119`
- Add-in-only XML manifest registered under the Analyst Lens manifest ID

## Build and verify

```bash
npm ci
npm run verify
```

The verification command runs lint, type checking, unit/component/contract tests, browser flows, a production build, both manifest validators, the Office API boundary check, fixture-bundle exclusion, and the outbound-network allowlist.

## Macabacus-style Mac shortcuts

Analyst Lens keeps Macabacus's physical `Control` chords on Mac and translates `Alt` to
`Option`. This avoids macOS intercepting `Command+Shift+4` and `Command+Shift+5` for
screenshots. Commands that Macabacus exposes only as a cycle or with unsupported punctuation
use explicit `Control+Option+Shift` mnemonics.

| Command | Mac shortcut |
| --- | --- |
| Integer / Date / Currency / Percentage / Multiple | `Ctrl+Shift+1` / `2` / `4` / `5` / `8` |
| Top / Bottom border | `Ctrl+Option+Shift+Up` / `Down` |
| Clear borders | `Ctrl+Option+Shift+N` |
| Input blue / Formula black / Linked green / Warning red | `Ctrl+Option+Shift+B` / `K` / `G` / `R` |
| One / Two decimals | `Ctrl+Option+Shift+1` / `2` |
| Increase / Decrease decimals | `Ctrl+Option+Shift+I` / `E` |
| Double bottom border / Clear formats | `Ctrl+Option+Shift+D` / `C` |
| Quick Trace In / Quick Trace Out | `Ctrl+Option+P` / `Ctrl+Option+D` |
| Open Analyst Lens | `Command+Option+L` |
| Group/Ungroup rows or columns; Open/Close All Groups | Task-pane controls only; no shortcut |

The same mapping is visible beside every speed tool, in the Inspector, in the command palette,
and in Settings. Quick Trace opens as a floating, movable, resizable Office dialog—not as a
task-pane overlay—and remains audit-only.

## Install lifecycle

```bash
./scripts/analyst-lens install
./scripts/analyst-lens status
./scripts/analyst-lens repair
./scripts/analyst-lens uninstall
```

`status` exits zero and prints `healthy` only when the owned manifest registration, launch agent and PID, certificate, tooling, and loopback health response all agree. `uninstall` removes only artifacts proven to be owned by Analyst Lens.

An already-healthy `install` or `repair` validates the installed payload but does not replace it with a newer source build. For a local qualification upgrade, use the verified replacement sequence:

```bash
./scripts/analyst-lens uninstall
if ./scripts/analyst-lens status; then exit 1; fi
./scripts/analyst-lens install
./scripts/analyst-lens status
```

After replacement, use **Add-in options > Reload** in an open task pane or restart Excel so the web view does not retain the previous JavaScript bundle.

## Safe qualification-workbook QA

Never write-test the source workbook. Create a byte-identical timestamped copy first:

```bash
qa_copy="$(bash scripts/qa-copy-fixture.sh)"
printf '%s\n' "$qa_copy"
```

The helper takes the source path as an argument or from `ANALYST_LENS_QA_SOURCE_FILE` (see [docs/qa/SETUP.md](docs/qa/SETUP.md)). It refuses an in-place target, refuses an existing target, requires the copy to live directly under the QA root, and verifies the initial SHA-256 before printing the path. The full manual matrix is in [docs/qa/excel-mac-smoke-checklist.md](docs/qa/excel-mac-smoke-checklist.md).

## Current qualification status

The continued 2026-07-15/17 Excel Mac run is **Analyst Lens local MVP PASS**. The Office gateway initializes; inspection, navigation, graph, scan, watchlist, verified formatting write/revert, the 21-action shortcut catalog, pane reopening, and floating Quick Trace are qualified. The final Excel 16.111 run also passed one-layer row/column grouping, repeated nested layers, one-layer inverse ungrouping, non-destructive Open/Close All visibility, protected-sheet failure, and the structural automatic-Revert barrier. Structural commands intentionally use the opposite one-layer command instead of generic automatic Revert because Office.js exposes no outline-state getter. Wave 1 adaptive production hardening is closed as **PARTIAL** because cold-workbook shortcut initialization and pre-distribution risks remain explicit; see the [Wave 1 production acceptance report](docs/qa/wave1-production-acceptance-2026-07-21.md). The qualification is local-only and claims no push, PR, publication, public-beta launch, or unattended-distribution readiness. See [docs/qa/fixture-qa-report.md](docs/qa/fixture-qa-report.md) and [docs/qa/quick-trace-feasibility-2026-07.md](docs/qa/quick-trace-feasibility-2026-07.md) for earlier exact evidence.

## License

MIT — see [LICENSE](LICENSE).
