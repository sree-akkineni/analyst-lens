# Changelog

All notable changes to Analyst Lens are documented here.

## [0.1.0] — 2026-10-12 — Initial public release

Local-only Excel for Mac add-in for formula inspection, dependency navigation,
verified formatting commands, model-health review, and a key-output watchlist.
Workbook data stays on the Mac; the only non-loopback runtime dependencies are
Microsoft's hosted Office.js runtime resources and its exact legacy
MicrosoftAjax bootstrap.

- **Hybrid Inspector** — selected-cell formula, value, format, direct
  precedents/dependents with values, risk badges, neighbor comparison, and a
  lazy 500-node dependency graph with origin breadcrumb and Back/Forward
  history.
- **Quick Trace** — floating, movable, resizable Office dialog for keyboard-only
  precedent/dependent branch walks (audit-only; two columns: Reference, Value).
- **Speed layer** — one-keystroke formatting (input blue, formula black, linked
  green, warning red, borders, decimals) plus one-layer Group/Ungroup.
- **Verified writes** — every formatting write does preflight → snapshot →
  write → readback, with one "Revert last Analyst Lens action". Structural
  mutations install a Revert barrier that fails closed.
- **Model health + watchlist** — cancellable structural scan (errors,
  discontinuities, hardcodes, embedded constants, external links, volatiles)
  with exact cell evidence; session-local key-output pins with
  assumption-impact highlighting.
- **Local install lifecycle** — loopback-only HTTPS host, user `launchd`
  agent, owned XML manifest registration, idempotent install / status /
  repair / uninstall with byte-identical installed-payload verification.

### Public-import notes

This is the initial public import: one squashed commit, no prior history. The
following were scrubbed or replaced before publication:

- All workbook fixtures, screenshots, sheet names, formulas, and values are
  synthetic (`Sample_Model.xlsx`); production builds statically exclude
  fixture code.
- 27 QA screenshots containing real client workbook data were dropped.
- QA reports, checklists, and specs were redacted with notices; internal
  tracking references were generalized.
- Internal execution runbooks in `docs/superpowers/plans/` were removed.

### Verification

`npm run verify` green: lint, type checking, 431 unit/component/contract tests,
browser flows, production build, both manifest validators, the Office API
boundary check, fixture-bundle exclusion, and the outbound-network allowlist.
See [docs/qa/SETUP.md](docs/qa/SETUP.md).
