# Grouping acceptance RED evidence — 2026-07-17

## Environment

- Starting implementation: one-layer grouping contract approved for local qualification
- Node: `v22.23.1`
- npm: `10.9.8`
- Pre-RED baseline: 44 Vitest files / 400 tests and 16/16 Playwright flows passed, followed by build, both manifest validators, Office-boundary, fixture-exclusion, and network gates.

## Focused RED command

```bash
PATH=/opt/homebrew/opt/node@22/bin:$PATH npm test -- \
  tests/unit/capabilityMap.test.ts \
  tests/unit/commandRegistry.test.ts \
  tests/contract/excelOfficeGateway.test.ts \
  tests/contract/writeExecutor.test.ts \
  tests/component/SpeedLayer.test.tsx \
  tests/component/CommandPalette.test.tsx
```

Observed result: **6 test files failed; 25 tests failed and 43 passed**.

Expected missing-behavior failures:

- 2 capability failures: ExcelApi 1.10 outline support is not yet registered.
- 3 command-registry failures: structure commands still use `verified-write`, retain the grouping-acceptance availability block, have no manual inverse metadata, and retain misleading Show/Hide labels.
- 1 write-executor failure: no structural Revert barrier exists, so Revert consumes the older formatting snapshot.
- 15 Excel-gateway failures: structure commands still return `UNVERIFIABLE_GROUPING_STATE` before Office instead of normalizing selected rows/columns, applying one layer, failing closed with typed selection/protection/table/outline errors, and running visibility-only Open/Close.
- 1 command-palette failure: the approved structure command remains disabled.
- 3 SpeedLayer failures: controls remain disabled, Open/Close labels/copy are absent, and structural success cannot direct the user to the one-layer inverse.

No failure was caused by a malformed test import, missing test fixture, or unrelated regression. The failures directly demonstrate the old blocked implementation.

## Playwright RED command

```bash
PATH=/opt/homebrew/opt/node@22/bin:$PATH npm run test:e2e -- tests/e2e/speed-layer.spec.ts
```

Observed result: **2 tests failed and 1 passed**.

- Protected structure failure could not execute because **Group rows** remained disabled.
- The one-layer/history flow could not find **Open All Groups** because the UI retained the legacy Show/Hide detail contract.
- The existing formatting apply/revert flow remained green.

## RED decision

The tests fail for the intended missing grouping-acceptance behavior. Production implementation can now begin under strict GREEN scope; acceptance must not be weakened to preserve the old block.
