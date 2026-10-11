# Excel Analyst Add-in — Product Requirements and Architecture

**Working name:** Analyst Lens  
**Status:** Approved for MVP implementation
**Owner:** Sree Akkineni  
**Primary user:** One buy-side / corp-dev analyst using Microsoft Excel for Mac  
**Target host:** Microsoft Excel for Mac 16.103 or later; first qualification build is Excel 16.111  
**Last updated:** 2026-07-17

**Approval:** Product direction, Hybrid Inspector option C, and the P0 architecture were approved by the owner on 2026-07-15. The implementation branch must use the scoped MVP plan before expanding into P1 or P2.

## 1. Executive summary

Analyst Lens is a locally installed Excel add-in for finance professionals who learned modeling on Windows with Macabacus-style shortcuts and now work on Mac. It combines a keyboard-first formatting and structure layer with a docked formula-auditing cockpit that lets the user move through complex precedent and dependent chains without losing context.

The initial product is one Office.js task-pane add-in using a shared runtime. It runs its application files from a local HTTPS service, keeps workbook data on the Mac, and uses Microsoft-hosted Office.js only to communicate with Excel. It does not require a VBA workbook, an Analyst Lens account, telemetry, a backend, or a Microsoft 365 work/school tenant. Local Mac installation uses an add-in-only XML manifest and requires no cloud title acquisition.

The approved interaction model is the **Hybrid Inspector**: a compact breadcrumb trail and keyboard navigator for routine work, paired with an expandable dependency graph for inherited or unfamiliar models.

This document is the umbrella product contract. Implementation is deliberately decomposed into independently shippable subprojects; each subproject receives its own reviewed implementation plan before code is written.

## 2. Problem

Excel for Mac provides the underlying modeling functions but not the integrated speed layer familiar to investment-banking and buy-side analysts on Windows:

- Repetitive formatting requires too many ribbon and dialog interactions.
- Grouping, tracing, and navigation shortcuts do not match established analyst muscle memory.
- Native tracer arrows do not preserve a navigable history through branching, cross-sheet dependency chains.
- Reviewing inherited models requires manually finding hardcodes, inconsistent formulas, external links, hidden dependencies, and fragile functions.
- Comparing revised models and watching key outputs are separate manual workflows.

The result is not a missing calculation engine. It is a missing analyst control surface.

## 3. Product vision

The add-in should feel like a flight deck for a financial model:

1. Every frequent command is reachable without leaving the keyboard.
2. The selected cell is always explainable in context.
3. The user can move upstream, downstream, and back to the origin without getting lost.
4. Read-only audit actions never alter workbook values or formulas.
5. Formatting writes are narrow, visible, verified, reversible through an add-in-specific revert command, and limited to the selected range. Structural outline writes use the explicitly approved one-layer manual-inverse exception in section 8.6.
6. Large or unfamiliar models can be scanned and triaged before detailed review.
7. No workbook content is transmitted to a vendor or remote service.

## 4. Product principles

- **Keyboard first, pointer available.** Every core action has a shortcut and a visible control.
- **Context before complexity.** Show direct relationships immediately; expand deeper levels only on demand.
- **Read-only by default.** Opening the inspector or running an audit must not modify cells, formulas, workbook structure, or calculation settings.
- **Progressive disclosure.** Routine navigation stays compact; graphs and detailed diagnostics expand when requested.
- **Evidence over opaque scoring.** Every health finding links to cells and states the exact rule that fired.
- **Local privacy.** No analytics, telemetry, formula logging, or workbook upload.
- **Finance conventions, configurable.** Ship sensible analyst defaults while keeping shortcut and formatting maps editable.
- **Performance is a feature.** Batch Office.js calls, lazily expand graphs, and make long scans cancellable.

## 5. Non-goals for the initial release

- Market-data ingestion or vendor integrations.
- AI-generated formula explanations or external LLM calls.
- PowerPoint linking or presentation-book automation.
- Full parity with every Macabacus feature.
- Editing or evaluating formulas in a custom calculation engine.
- Inspecting formulas inside a closed external workbook.
- Team administration, cloud synchronization, billing, or multi-user collaboration.
- Windows, Excel web, or iPad qualification beyond avoiding unnecessary platform lock-in.

## 6. User and core jobs

### Primary user

A technically proficient buy-side, investment-banking, private-equity, hedge-fund, or corporate-development analyst who spends several hours per day in financial models and values speed, auditability, and keyboard control.

### Core jobs

1. Format selected model ranges to house conventions in one keystroke.
2. Group or ungroup rows and columns without navigating the ribbon.
3. Understand what drives a selected output and what the selected assumption drives.
4. Traverse branching dependency chains across worksheets and return to the starting cell.
5. Identify formula inconsistencies, hardcodes, errors, fragile functions, and hidden dependencies.
6. Keep critical transaction and return outputs visible while changing assumptions.
7. Understand what changed between two versions of a model.

## 7. Success criteria

### Workflow outcomes

- A user can trace three levels upstream, inspect a branch, and return to the origin in under 10 seconds after learning the shortcuts.
- A user can apply any P0 format or grouping operation with one simultaneous shortcut or one ribbon command.
- A health-scan result always identifies the exact cells, rule, severity, and remediation guidance.
- A pinned output visibly updates after workbook recalculation without requiring the user to navigate back to it.
- Audit-only commands make zero changes to workbook values, formulas, formats, names, worksheets, protection, or calculation settings.

### Performance budgets on Excel 16.111

- Warm selection-to-inspector update: p95 at or below 600 ms for a single formula cell.
- Direct precedent/dependent navigation: p95 at or below 500 ms after data is loaded.
- One-level graph expansion: p95 at or below 1 second for up to 50 newly discovered nodes.
- Health scan: completes within 20 seconds for a workbook with up to 100,000 used cells, with progress and cancellation available after 500 ms.
- Task-pane interactions remain responsive during graph layout and scan result rendering.

### Reliability and privacy

- No uncaught exceptions in a 60-minute dogfood session on three representative financial models.
- No runtime non-loopback outbound requests other than loading Microsoft Office.js resources; automated tests fail on any unexpected first-party destination. Installation, repair, status, and uninstall are local-only and receive no workbook content.
- Install, start-at-login, repair, and uninstall instructions are verified on a clean macOS user profile.
- Every formatting write, and every other write whose state Excel exposes, has an add-in-specific **Revert last Analyst Lens action** path.
- Group/Ungroup is the sole approved P0 exception: Excel exposes the mutation but not outline-level or group-membership readback, so each command is one layer and its user-facing inverse is the opposite command on the applicable selection.

## 8. Product scope

### 8.1 P0 — Hybrid Formula Inspector

The docked task pane updates when the Excel selection changes.

It displays:

- Workbook, worksheet, address, value, number format, and formula.
- A syntax-highlighted formula with clickable A1 references, ranges, sheet references, names, functions, operators, strings, and numeric literals.
- Direct precedents and direct dependents, including current values and sheet names.
- Back and forward history plus an origin breadcrumb.
- Keyboard commands for next/previous precedent, next/previous dependent, back, forward, pin, and open graph.
- Native Excel trace-arrow controls when `ExcelApiDesktop 1.1` is available.
- Risk badges for hardcoded numeric literals, cross-sheet references, external links, hidden-sheet references, named ranges, errors, volatile functions, dynamic references, and inconsistent neighboring formulas.
- A neighbor comparison showing normalized formulas to the left, right, above, and below.

Behavior:

- Clicking or keyboard-selecting a reference activates that worksheet and range in Excel.
- Every navigation action appends a node to history; branching never discards the path back to the origin.
- History is session-local and capped at 200 entries per workbook window.
- A non-formula selection shows value type, format, and relevant audit/formatting actions instead of an empty error state.
- Multi-cell selection shows aggregate information and disables single-cell dependency navigation until one cell is chosen.

### 8.2 P0 — Dependency Graph

The compact orientation map expands into a full graph within the task pane.

- Nodes represent workbook ranges, not individual cells when Excel reports a contiguous range dependency.
- Directed edges point from precedent to dependent.
- The starting cell is visually distinct and remains centered on first open.
- Users expand one level upstream or downstream from any node.
- Users can collapse branches, fit the graph, search by sheet/address/name, and return to origin.
- Cycles are detected and rendered as explicit cycle edges rather than expanded repeatedly.
- Duplicate nodes are de-duplicated by workbook, worksheet ID, and canonical address.
- Graph expansion is lazy and capped at 500 visible nodes. At the cap, the UI explains how to narrow the graph or continue a specific branch.
- Dynamic references from `INDIRECT`, text-built addresses, and similar constructs appear as **dynamic unresolved** nodes when Excel cannot resolve them.
- References into closed external workbooks appear as leaf nodes with the source path or link name when available.

The graph uses Excel dependency APIs as the source of truth. The formula tokenizer improves display and click targets but does not invent graph edges.

### 8.3 P0 — Model Health Scan

The scan operates on explicit user command and never runs continuously in the background.

Initial rules:

- Excel error values and broken references.
- Formula discontinuities within repeated horizontal or vertical runs.
- Hardcoded values interrupting formula runs.
- Numeric constants embedded inside formulas, excluding `0`, `1`, `-1`, and configured safe constants.
- External workbook links.
- References to hidden or very-hidden sheets, hidden rows, or hidden columns.
- Volatile functions: `OFFSET`, `INDIRECT`, `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `CELL`, and `INFO`.
- Circular references reported by Excel or found during graph traversal.
- Formulas longer than 500 characters or nested more than 12 function levels.
- Mixed formulas within a selected comparison band after relative-reference normalization.
- Calculation mode set to manual.
- Inconsistent number format, font color, or sign convention inside a repeated model row.
- Merged cells inside used financial-model ranges.

Results include severity, rule ID, workbook location, observed evidence, and suggested remediation. Clicking a result navigates to the cell. Users can filter by severity, sheet, or rule, rerun the scan, export a local Markdown report, and mark a finding ignored for the current session.

The initial scan does not assert that a financial model is economically correct. It reports structural and formula evidence only.

### 8.4 P0 — Analyst Speed Layer

Formatting commands operate on the current selection:

- Input font blue, formula font black, linked formula font green, warning font red, automatic font color.
- Integer, one-decimal, two-decimal, percentage, currency, multiple, date, and text formats.
- Zero as dash, negatives in parentheses, and configurable scale markers.
- Increase/decrease decimals.
- Top, bottom, and double-bottom borders; clear borders.
- Bold, italic, underline, indentation, alignment, wrap, and center across selection.
- Copy values only, formulas only, formats only, or number formats only.
- Clear formats without clearing values or formulas.

Structure and navigation commands:

- Group or ungroup the rows intersecting the current selection, exactly one outline layer per command.
- Group or ungroup the columns intersecting the current selection, exactly one outline layer per command.
- Repeated Group commands may add nested levels up to Excel's supported maximum; repeated Ungroup commands remove one level at a time.
- Open All Groups and Close All Groups change outline visibility with `Worksheet.showOutlineLevels(8, 8)` and `Worksheet.showOutlineLevels(1, 1)` respectively. They never add/remove group definitions and never enter write/revert history.
- Jump to next/previous formula, hardcode, error, blank, visible cell, or bookmarked cell.
- Open the Hybrid Inspector, return to origin, and toggle trace arrows.

Shortcut behavior:

- Defaults use conflict-resistant Mac chords.
- Commands and shortcuts are visible in Settings and in a searchable command palette.
- The user can replace supported shortcuts through Office's shortcut API.
- Conflicts are detected and explained; the add-in never silently steals an established shortcut.

### 8.5 P0 — Key Output Watchlist

- Pin any single cell with an optional label and category.
- Show address, current value, formula-change state, and last observed value.
- Refresh after recalculation or on explicit refresh when recalculation events are unavailable.
- Highlight which pinned outputs are downstream of the current selection.
- Show the count and names of pinned outputs affected by a selected assumption.
- Session-local by default.
- **Save with workbook** is an explicit opt-in that stores only addresses, labels, and categories in add-in document settings; it never stores values or formulas.

### 8.6 P0 — Add-in Revert

Office add-in actions do not promise integration with Excel's native multi-level Undo stack.

Before a write command, Analyst Lens captures the exact properties it intends to change. It then exposes one command: **Revert last Analyst Lens action**.

- Formatting snapshots use displayed/settable cell properties for the selected range and retain mandatory post-write/readback/revert verification.
- The snapshot is in memory only and is replaced by the next verified Analyst Lens write.
- Group Rows, Group Columns, Ungroup Rows, and Ungroup Columns are explicitly excluded from generic automatic Revert because Office.js exposes no outline-level or group-membership getter. A successful `context.sync()` is not verified readback and must never be described as such.
- The documented manual inverse is exactly one opposite structural command on the applicable selection: Group ↔ Ungroup. Analyst Lens does not claim arbitrary prior nested-outline reconstruction.
- After a successful structural mutation, **Revert last Analyst Lens action** must fail clearly and direct the user to the corresponding one-layer inverse. It must not silently revert an older formatting snapshot.
- Open All Groups and Close All Groups are non-destructive visibility actions. They do not change group definitions and do not replace, clear, or enter mutation/revert history.
- Commands other than the approved structural exception remain excluded from P0 when their required state cannot be safely read back and reverted.
- Audit, navigation, graph, scan, and watchlist refresh actions do not require revert because they do not modify cells.

### 8.7 P1 — Model Version Comparison

The user chooses a baseline and revised workbook that are both open in Excel.

- Compare worksheet presence/order, names, formulas, values, number formats, styles, names, hidden state, and row/column structure.
- Normalize formulas so copied formulas can be compared by relative structure.
- Let the user ignore value-only recalculation changes, formatting-only changes, or selected sheets.
- Connect changed cells to affected pinned outputs.
- Produce a local Markdown change report with counts, material paths, and exact cell references.

P1 is a separate subproject because it needs its own comparison model, performance budgets, fixtures, and UX review.

### 8.8 P2 — Scenario and Sensitivity Tools

- Base, upside, and downside assumption sets.
- Two-way data-table builder.
- IRR, MOIC, leverage, CAGR, LTM/NTM, and accretion/dilution helper commands.
- Source/provenance annotations for hardcoded assumptions.

P2 remains outside initial ticket creation until P0 has passed dogfood acceptance.

## 9. UI contract

The approved UI family is documented in [Interactive UI mockups](assets/excel-analyst-addin-ui-mockups.html).

Primary surfaces:

1. **Inspector:** compact formula, breadcrumbs, risk badges, precedent/dependent lists, neighbor comparison, compact graph, and navigation controls.
2. **Full graph:** search, upstream/downstream expansion, branch collapse, graph limits, and origin return.
3. **Health scan:** progress, cancellation, severity/rule/sheet filters, exact evidence, and navigation.
4. **Watchlist:** pinned output values, change indicators, downstream impact, and persistence choice.
5. **Commands and settings:** searchable commands, editable shortcuts, formatting conventions, safe constants, privacy state, repair link, and version information.
6. **Model diff:** baseline/revised selectors, change filters, summary, cell-level evidence, and pinned-output impact.

The task pane is resizable. At narrow widths it uses a single-column inspector and hides the compact graph behind an **Open map** action. At wider widths it shows the inspector and compact graph side by side.

## 10. Architecture decision record

### Decision

Build one locally served Office.js add-in with an add-in-only XML manifest for Excel on Mac, a shared runtime, React/TypeScript task pane, and pure domain modules behind an Excel adapter. Retain the unified Microsoft 365 manifest only as a future organizational-deployment artifact.

### Why

- The selected Hybrid Inspector needs a docked, resizable, interactive panel that a VBA `.xlam` cannot provide cleanly on Mac.
- Office.js supports custom keyboard shortcuts, ribbon commands, grouping, cross-sheet precedent/dependent retrieval, selection events, and native tracer arrows on the target Excel version.
- One add-in avoids the operational and testing complexity of a VBA plus JavaScript hybrid.
- Domain logic can be isolated from Office.js and tested without Excel.

### Alternatives considered

1. **Pure VBA `.xlam`.** Fastest for formatting and grouping, but the audit UI would be a floating UserForm or workbook sheet, graph rendering would be brittle, and cross-sheet dependency retrieval is weaker.
2. **VBA plus Office.js hybrid.** Strongest native macro compatibility, but creates two runtimes, two installers, duplicated command routing, and more failure modes.
3. **Native macOS helper plus Excel automation.** Maximum shortcut control, but requires Accessibility permissions, signing, and brittle UI automation. Deferred unless Office shortcut conflicts block an essential workflow.

### Consequences

- A local HTTPS process must be available while Excel uses the add-in.
- The installer must configure a trusted development certificate, register an exact owned XML manifest in Excel's local `wef` folder, and register a `launchd` user agent. It must not require a work/school tenant or cloud title acquisition.
- The add-in loads Microsoft Office.js from Microsoft's supported CDN, but application code makes no outbound requests and sends no workbook content off-device.
- Excel 16.103 is the minimum supported Mac version for the chosen shared-runtime and desktop API surface. Excel 16.111 is the first release qualification target.

## 11. Component boundaries

### `platform/office`

The only layer allowed to call `Excel.run`, `Office.context`, or shared-runtime APIs.

- Selection and recalculation event adapters.
- Range metadata and dependency retrieval.
- Formatting, grouping, navigation, and trace-arrow commands.
- Capability detection and normalized errors.
- Batched `context.sync()` calls.

### `domain/dependencies`

Pure TypeScript graph behavior.

- Canonical node IDs.
- Precedent/dependent edge construction.
- Cycle detection, de-duplication, depth/node caps, expansion, collapse, and search.
- Navigation history and origin behavior.

### `domain/formulas`

Pure TypeScript lexical analysis for display and diagnostics.

- Tokenizes references, functions, operators, strings, and numeric literals.
- Normalizes relative formulas for neighbor/version comparison.
- Does not calculate formulas or invent dependencies.

### `domain/health`

Pure rule engine.

- Consumes normalized workbook facts.
- Produces deterministic findings with rule ID, severity, evidence, and remediation.
- Keeps rules independently testable and selectively runnable.

### `domain/commands`

- Command registry, labels, categories, capability requirements, shortcuts, and write/read classification.
- Revert snapshot contracts for every verified write command plus explicit manual-inverse metadata for structural Group/Ungroup.

### `ui/taskpane`

- Inspector, graph, scan, watchlist, command palette, settings, and diff views.
- Contains no direct Office.js calls.
- Receives domain state and sends typed intents to the application controller.

### `runtime/controller`

- Coordinates Office events, domain services, caches, cancellation, and UI state.
- Debounces selection events and invalidates workbook-scoped caches on workbook changes.

### `local-host`

- Serves built static assets over HTTPS on loopback only.
- Exposes a health endpoint with version only; it never accepts workbook content.
- Starts through a user `launchd` agent.

### `installer`

- Installs certificate, manifest, local host, and `launchd` agent.
- Provides idempotent install, repair, status, and uninstall commands.
- Owns the local add-in-only XML registration and certificate lifecycle. It never receives or transmits workbook content, and lifecycle commands have no permitted non-loopback destination.
- Never changes Excel macro-security settings.

## 12. Data flow

1. Excel raises a selection-changed event.
2. The controller debounces for 100 ms and requests the active range metadata through `platform/office`.
3. The Office adapter performs the minimum batched load and returns plain TypeScript data.
4. Formula and dependency domains normalize the data and update inspector state.
5. Direct dependencies are cached by workbook, sheet ID, address, formula text, and calculation version.
6. Graph expansion requests additional dependencies only for the chosen node and direction.
7. The UI renders state and emits intents; it never accesses Excel directly.
8. A verified write command captures its revert snapshot, applies the selected operation, verifies the resulting properties, and publishes success or a normalized error. A structural Group/Ungroup command instead performs fail-closed preflight, applies one layer, records an automatic-Revert barrier, and publishes the exact manual inverse without claiming readback.

## 13. Persistence and privacy

- Settings and shortcut preferences live in local browser storage within the add-in origin.
- Navigation history, scan ignores, graph state, and watchlists are session-local by default.
- Explicit **Save with workbook** stores watchlist addresses/labels/categories in Office document settings. It stores no values, formulas, scan results, or graph snapshots.
- No telemetry or crash-upload service.
- Logs are local, redact formulas and values, and include only timestamps, command IDs, capability flags, durations, and error codes.
- Debug formula/value logging exists only in development builds and is off by default.
- The local host binds to `127.0.0.1`, rejects non-loopback Host headers, and serves immutable built assets.

## 14. Error handling and degraded states

- Unsupported requirement set: disable only the affected command and show the required Excel version/API.
- Protected worksheet or locked range: show a non-destructive explanation and do not partially apply a command.
- No formula/no dependencies: show an informative empty state, not an exception.
- Dynamic unresolved reference: render a leaf node and explain why it cannot be expanded.
- Closed external workbook: render an external leaf and offer to refresh after the workbook is opened.
- Selection changed during an async request: cancel or discard the stale result by request ID.
- Workbook closed or context invalidated: clear workbook-scoped state and return to the landing screen.
- Graph cap reached: stop automatic expansion and allow explicit branch continuation.
- Scan exceeds 500 ms: show progress and cancellation; cancellation leaves no partial persisted result.
- Verified-write readback fails: attempt the captured revert once, report both outcomes, and stop further writes until the user dismisses the safety lock.
- Structural Group/Ungroup fails or reaches Excel's outline limit: normalize the Office error, make no automatic inverse claim, and leave the prior verified-write history untouched unless Excel accepted the structural mutation.
- Local host unavailable: Excel shows a repair message with the exact local status command.

## 15. Testing strategy

### Unit tests

Mandatory for every domain function and rule:

- Formula tokenization and relative normalization.
- Graph IDs, de-duplication, expansion, collapse, cycles, search, limits, and history.
- Health rules and false-positive exclusions.
- Command capability gates, shortcut mapping, and revert descriptors.
- State reducers and cancellation semantics.

Tests use Vitest with deterministic fixtures and no Office globals.

### UI component tests

- Inspector state variants: formula, hardcode, blank, error, multi-selection, loading, and unsupported capability.
- Keyboard navigation and focus order.
- Graph node selection, expansion intent, cap state, and origin return.
- Health filters, result selection, progress, and cancellation.
- Watchlist refresh and formula-change indicators.
- Settings conflict and repair states.

Tests use Testing Library semantics and assert accessible roles/names rather than CSS structure.

### Office adapter contract tests

- A narrow fake Office context records loads, queued writes, and `context.sync()` calls.
- Every adapter test proves the returned plain data and expected batching behavior.
- Formatting-write tests verify preflight, snapshot, mutation, post-write readback, and failure normalization.
- Structural tests verify selection-to-row/column normalization, one mutation call per invocation, nesting without de-duplication, one-layer ungroup, fail-closed selection/protection handling, explicit Revert barriers, and visibility-only Open/Close behavior. They do not fabricate outline-state readback.
- Contract tests never claim real Excel compatibility; they protect adapter intent.

### Excel fixture integration tests

A committed fixture corpus covers:

- Horizontal and vertical operating-model formula runs.
- Cross-sheet precedents and dependents.
- Named ranges and sheet-scoped names.
- Hidden sheets, rows, and columns.
- Grouped/nested rows and columns.
- External workbook links with both workbooks open and with the source closed.
- `INDIRECT`, `OFFSET`, volatile functions, long formulas, errors, and circular references.
- Protected sheets and mixed locked/unlocked cells.
- Large used ranges and sparse models.
- Multiple number formats, colors, borders, merged cells, and localized display formats.

Each integration case states the exact workbook, starting selection, command, expected cell/formula/format outcome, and expected inspector/scan result.

### End-to-end tests

- Standalone task-pane shell tests run in Playwright with a simulated workbook adapter.
- Real Excel for Mac smoke tests run against the sideloaded manifest and fixture workbooks.
- A human-visible QA checklist captures screenshots and exact on-screen copy for install, inspector, graph, scan, shortcuts, grouping, revert, watchlist, host repair, and uninstall.
- Real Excel evidence is required before an issue that changes Office integration can move to Done.

### Performance tests

- Generated graph fixtures at 50, 100, 250, and 500 nodes.
- Formula-run scans at 10,000, 50,000, and 100,000 used cells.
- Selection-event burst tests prove only the latest request renders.
- Office adapter tests record sync counts and fail when a workflow exceeds its budget.

## 16. Engineering and autonomous-agent operating contract

### Source-of-truth order

1. This approved PRD and architecture decision.
2. The active subproject implementation plan.
3. The tracking issue and its acceptance tests.
4. Current Microsoft API behavior observed on Excel 16.111.
5. Existing code and comments.

If two sources conflict, the worker stops and reports the exact conflict instead of choosing silently.

### Ticket execution loop

Every implementation issue follows this loop:

1. Read the issue, linked subproject plan, and applicable architecture sections.
2. Confirm lane ownership and inspect the current branch/worktree.
3. Write one failing test for the next behavior.
4. Run it and record the expected failure.
5. Implement the minimum behavior.
6. Run focused tests, then the relevant package suite.
7. Run fixture or real-Excel verification proportional to the change.
8. Review the diff for scope, privacy, mutation safety, and API batching.
9. Attach evidence and update the issue with what changed, what passed, and remaining limits.
10. Stop at the issue boundary; do not opportunistically implement neighboring tickets.

### Verification loop

Every subproject has three gates:

- **Domain gate:** pure logic tests and static checks pass.
- **Adapter/integration gate:** fixture and Office adapter contracts pass.
- **Product gate:** Excel 16.111 smoke test and visible evidence pass.

A subproject is not complete when only mocked tests pass.

### QA loop

1. Start from a clean Excel launch and known fixture checksum.
2. Execute the exact workflow from the issue acceptance test.
3. Record Excel version, macOS version, workbook, starting selection, shortcut/button used, outcome, and screenshot.
4. Reopen the workbook where persistence is relevant.
5. Repeat verified formatting writes with **Revert last Analyst Lens action**. Repeat structural writes with the documented one-layer opposite command and verify that generic Revert refuses to consume a stale formatting snapshot.
6. Run an adjacent regression workflow.
7. File failures as separate issues with reproducible steps and fixture references.

### Stop conditions

An autonomous worker must stop and request review when:

- Required Office API behavior differs from the PRD on Excel 16.111.
- A change would transmit workbook data or add a new outbound dependency.
- A write command other than the owner-approved one-layer Group/Ungroup exception cannot be verified or safely reverted.
- A ticket requires changing another subproject's published interface.
- Real Excel verification is unavailable for an Office-integration change.
- Performance exceeds a stated budget after one bounded optimization pass.
- The correct handling of financial-model semantics is ambiguous.

## 17. Delivery decomposition

The umbrella product is split into independently planned subprojects. Tracking issues should target one independently testable one-to-two-day deliverable; work estimated above three days must be split.

### E0 — Platform and local installation

Manifest, shared runtime, task-pane shell, capability checks, local HTTPS host, `launchd`, install/repair/status/uninstall, and a clean Excel 16.111 smoke test.

### E1 — Office adapter and command spine

Typed Office boundary, selection events, command registry, shortcut routing, capability errors, write verification, and revert infrastructure.

### E2 — Analyst speed layer

Formatting, copy/paste modes, one-layer grouping/ungrouping with manual inverses, non-destructive Open/Close visibility, and workbook navigation with verified fixture outcomes.

### E3 — Dependency engine

Direct and multi-level dependency retrieval, canonical graph, history, cycles, caps, caching, and formula tokenization.

### E4 — Hybrid Inspector UI

Inspector, breadcrumbs, reference lists, neighbor comparison, compact graph, full graph, keyboard focus, and native tracer-arrow controls.

### E5 — Model Health Scan

Fact extraction, rule engine, initial rule set, progress/cancellation, result navigation, filters, and Markdown export.

### E6 — Key Output Watchlist

Pinning, refresh/recalculation behavior, formula-change detection, current-selection impact, and opt-in workbook persistence.

### E7 — Release qualification and dogfood

Fixture corpus, automated gates, Excel Mac QA matrix, privacy/network validation, performance budgets, documentation, and signed release checklist.

### E8 — Model Version Comparison (P1)

Separate PRD-to-plan cycle after P0 dogfood acceptance.

### E9 — Scenarios and Sensitivities (P2)

Separate PRD-to-plan cycle after P1 prioritization.

## 18. Tracking-issue contract

Each future issue must include:

- **Goal:** one user-visible or independently testable outcome.
- **Why:** link to the PRD job and requirement.
- **Scope:** exact components/files or interfaces owned.
- **Out of scope:** neighboring behavior the worker must not add.
- **Inputs/interfaces:** exact types or upstream contracts consumed.
- **Outputs/interfaces:** exact types or downstream contracts produced.
- **Acceptance tests:** concrete test names and Excel fixture steps.
- **Verification:** focused command, suite command, real Excel requirement, and evidence format.
- **QA risks:** protected sheets, dynamic references, large selections, shortcut conflicts, or privacy implications as applicable.
- **Dependencies:** blocking issue identifiers once created.
- **Stop conditions:** issue-specific reasons to pause.
- **Completion evidence:** test output summary, fixture/version, screenshots for visible changes, and known limitations.

## 19. Release gates

P0 cannot be called complete until all are true:

- Lint, typecheck, unit, component, adapter, integration, and standalone E2E suites pass.
- Add-in-only XML and retained unified manifest validation pass.
- Install, repair, status, and uninstall pass on a clean macOS user profile.
- All P0 commands have documented shortcuts and visible controls.
- All formatting and otherwise readable P0 writes pass preflight, readback verification, and revert tests. Structural Group/Ungroup passes the approved fail-closed one-layer/manual-inverse contract and generic Revert barrier tests.
- Audit-only workflows show no workbook mutations in before/after inspection.
- Unexpected network request test passes.
- Performance budgets pass on the defined fixture sizes.
- Accessibility keyboard/focus pass is complete.
- Excel 16.111 smoke matrix is complete with screenshots.
- Three representative real-model dogfood sessions totaling at least three hours have no P0/P1 unresolved defects.
- User signs off on the Inspector, graph, scan, speed layer, and watchlist workflows.

## 20. References

- Microsoft Office Add-ins overview: https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins
- Shared runtime: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime
- Custom keyboard shortcuts: https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts
- Excel Range API: https://learn.microsoft.com/en-us/javascript/api/excel/excel.range
- Excel requirement sets: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets
- Task pane API: https://learn.microsoft.com/en-us/javascript/api/office/office.taskpane
- Sideload Office Add-ins on Mac: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac
