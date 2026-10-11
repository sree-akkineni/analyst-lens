# Qualification-workbook setup

Real-Excel qualification never opens or write-tests a source workbook. Every
run starts by minting a timestamped byte-identical QA copy with
`scripts/qa-copy-fixture.sh`, then opens only the copy.

## Configure the source

Point the helper at any workbook you own. Pass the path as the first argument
or export it once per shell:

```bash
export ANALYST_LENS_QA_SOURCE_FILE="$HOME/qa-workbooks/Sample_Model.xlsx"
```

The helper accepts a workbook anywhere on disk; nothing in the repository
assumes a particular directory layout.

## Mint a QA copy

```bash
qa_copy="$(bash scripts/qa-copy-fixture.sh)"
printf '%s\n' "$qa_copy"
```

or with an explicit path argument:

```bash
qa_copy="$(bash scripts/qa-copy-fixture.sh "$HOME/qa-workbooks/Sample_Model.xlsx")"
```

The helper prints the copy path (for example
`$ANALYST_LENS_QA_ROOT/Sample_Model_qa_20260717-215319.xlsx`) and guarantees:

- the copy is byte-identical (SHA-256 verified before the path is printed);
- the target is never the source (in-place QA is refused);
- an existing target is never overwritten;
- the copy lands directly beneath the QA root.

## Optional overrides

| Variable | Default | Meaning |
| --- | --- | --- |
| `ANALYST_LENS_QA_ROOT` | `$TMPDIR/analyst-lens-qa` | Directory receiving QA copies |
| `ANALYST_LENS_QA_TARGET` | Derived from the source name | Exact copy path |
| `ANALYST_LENS_QA_TIMESTAMP` | Current `%Y%m%d-%H%M%S` | Timestamp embedded in the copy name |

## Run the matrix

Follow [excel-mac-smoke-checklist.md](excel-mac-smoke-checklist.md) against the
minted copy, then re-hash the source to prove it stayed untouched. Prior
completed runs are recorded in
[fixture-qa-report.md](fixture-qa-report.md),
[quick-trace-feasibility-2026-07.md](quick-trace-feasibility-2026-07.md),
[wave1-production-acceptance-2026-07-21.md](wave1-production-acceptance-2026-07-21.md),
and
[wave1-native-observations-2026-07-26.md](wave1-native-observations-2026-07-26.md).
