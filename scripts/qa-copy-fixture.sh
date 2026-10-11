#!/usr/bin/env bash
# Create a timestamped byte-identical QA copy of a qualification workbook.
#
# The source workbook is never opened or write-tested directly. Point this
# helper at any workbook you own via the first argument or the
# ANALYST_LENS_QA_SOURCE_FILE environment variable:
#
#   qa_copy="$(bash scripts/qa-copy-fixture.sh "$ANALYST_LENS_QA_SOURCE_FILE")"
#
# Optional environment overrides:
#   ANALYST_LENS_QA_ROOT       directory receiving QA copies (default: $TMPDIR/analyst-lens-qa)
#   ANALYST_LENS_QA_TARGET      exact copy path (default: derived from the source name)
#   ANALYST_LENS_QA_TIMESTAMP   copy timestamp (default: current date +%Y%m%d-%H%M%S)
#
# See docs/qa/SETUP.md for the full qualification-workbook setup.
set -euo pipefail

source_file="${1:-${ANALYST_LENS_QA_SOURCE_FILE:-}}"
if [[ -z "$source_file" ]]; then
  printf 'QA source workbook is required: pass a path or set ANALYST_LENS_QA_SOURCE_FILE\n' >&2
  exit 1
fi
tmp_root="${TMPDIR:-/tmp}"
qa_root="${ANALYST_LENS_QA_ROOT:-${tmp_root%/}/analyst-lens-qa}"
timestamp="${ANALYST_LENS_QA_TIMESTAMP:-$(date +%Y%m%d-%H%M%S)}"
source_base="$(basename -- "$source_file")"
source_stem="${source_base%.xlsx}"
target="${ANALYST_LENS_QA_TARGET:-${qa_root}/${source_stem}_qa_${timestamp}.xlsx}"

[[ -f "$source_file" ]] || {
  printf 'QA source workbook is missing: %s\n' "$source_file" >&2
  exit 1
}

if [[ "$target" == "$source_file" ]] || [[ -e "$target" && "$source_file" -ef "$target" ]]; then
  printf 'Analyst Lens refuses in-place QA against the source workbook\n' >&2
  exit 1
fi

mkdir -p "$qa_root"
qa_root_canonical="$(CDPATH= cd -- "$qa_root" && pwd -P)"
target_parent="$(dirname -- "$target")"
target_parent_canonical="$(CDPATH= cd -- "$target_parent" && pwd -P)"

[[ "$target_parent_canonical" == "$qa_root_canonical" ]] || {
  printf 'Analyst Lens QA copies must be created directly beneath: %s\n' "$qa_root_canonical" >&2
  exit 1
}

[[ ! -e "$target" ]] || {
  printf 'Analyst Lens refuses to overwrite an existing QA copy: %s\n' "$target" >&2
  exit 1
}

source_hash="$(shasum -a 256 "$source_file" | awk '{ print $1 }')"
cp -p "$source_file" "$target"
copy_hash="$(shasum -a 256 "$target" | awk '{ print $1 }')"

if [[ "$source_hash" != "$copy_hash" ]]; then
  rm -f "$target"
  printf 'Analyst Lens QA copy checksum mismatch\n' >&2
  exit 1
fi

printf '%s\n' "$target"
