# Analyst Lens Floating Quick Trace — Design Specification

**Status:** Approved for implementation
**Owner:** Sree Akkineni
**Target host:** Microsoft Excel for Mac 16.111
**Date:** 2026-07-16
**Parent contract:** `2026-07-15-excel-analyst-addin-design.md`
**Visual review artifact:** `assets/analyst-lens-quick-trace-mockup.html`

## 1. Decision summary

Add a compact **Quick Trace** window for keyboard-first precedent and dependent auditing. It must be a genuinely floating, movable, resizable, nonmodal window over Excel. It is not a task-pane view and has no docked fallback.

Quick Trace is inspired by Macabacus Pro Precedents / Pro Dependents, now called Trace In / Trace Out, but it is adapted to the Office.js and Excel for Mac execution model:

- **Trace In** walks precedents.
- **Trace Out** walks dependents.
- Up and Down select visible tree rows and navigate Excel to the selected range.
- Right lazily expands one level in the current direction.
- Left collapses the selected branch or moves to its parent.
- Return re-roots the audit on the selected range.
- Return to Origin restores the cell where the session began.
- All workbook interaction remains audit-only.

The implementation approach is the Office Dialog API. The existing shared runtime remains the only owner of Office.js workbook operations. The dialog renders state and sends typed user intents to the shared runtime.

This design is owner-approved. Product implementation remains gated on the real-Excel feasibility spike described below.

## 2. Relationship to the existing Analyst Lens product

Quick Trace complements rather than replaces the Hybrid Inspector:

| Surface | Primary job | Presentation |
| --- | --- | --- |
| Hybrid Inspector | Explain the selected cell, compare neighbors, show risks, and expose the full dependency map | Docked task pane |
| Quick Trace | Traverse one dependency branch rapidly without losing keyboard position | Floating dialog |
| Full dependency map | Explore and search a larger branching model | Expanded task-pane view |

All three surfaces use the existing canonical dependency graph, navigation history, caching, and Office gateway. There must not be a second dependency engine or a dialog-specific Office adapter.

The task pane and floating dialog have independent visibility. Closing one does not close the other, and closing the task pane must not prevent a shortcut from opening Quick Trace through the shared runtime.

## 3. Goals and non-goals

### P0 goals

1. Open a floating Trace In or Trace Out window from the current Excel selection.
2. Traverse at least three dependency levels across worksheets using only the keyboard.
3. Keep the session origin and provide an explicit return path.
4. Show exact references and trustworthy values without inventing semantic labels.
5. Preserve dialog keyboard focus while Excel activates the selected range.
6. Reuse the existing lazy graph, cycle handling, unresolved leaves, and 500-node cap.
7. Keep the workflow read-only and local.

### P0 non-goals

- Editing a formula from the dialog or entering Excel Point/Edit mode.
- Inferring business meanings such as `salary`, `period`, or `overhead` from formula references.
- Evaluating Excel functions or arbitrary expression groups in a custom calculation engine.
- Automatically opening a closed external workbook.
- Applying row/column shading or any other workbook formatting as navigation highlighting.
- Displaying an entire 500-node graph at once; the dialog is a branch-oriented tree.
- Replacing the existing inspector or full dependency map.

### Deferred enhancements

- **Evaluate expressions:** add expression nodes and formal Excel function arguments such as `logical_test` or `value_if_true` after implementing a tested formula AST and function-signature catalog.
- **Formula editing:** investigate safe edit-mode handoff only if Office.js exposes a deterministic Mac path.
- **External workbook assistance:** offer an explicit open/refresh workflow if Excel exposes a safe supported API.
- **Persistent placement:** remember a user-selected dialog position only if the Office Dialog API exposes a supported mechanism. The host controls initial placement today.

## 4. Information architecture

### 4.1 Window frame

The initial target is approximately 480 by 420 logical pixels, subject to the percentage-based sizing and minimums imposed by the Office Dialog API. Excel initially centers the dialog. The user can move and resize it using the native window frame.

Only one Quick Trace dialog may exist per Excel host window. A second open command updates the existing session or reports that Quick Trace is already open; it never creates a duplicate.

### 4.2 Header

The header contains:

- A segmented **Trace In / Trace Out** mode control.
- The current root workbook, worksheet, and address.
- A compact **Origin** address.
- The current root value.
- The current root formula, wrapped and selectable for reading.

The formula is evidence, not a custom editor.

### 4.3 Tree

P0 uses two columns only:

| Column | Contract |
| --- | --- |
| Reference | Exact worksheet and address or an explicit unresolved/cycle label |
| Value | Current single-cell value when available; a range-size label for multi-cell ranges |

There is no **Argument** or business-role column in P0.

Macabacus uses its Argument column for formal function parameters when its optional function/group evaluation is enabled. For example, a child expression of `IF` may be identified as `logical_test` or `value_if_true`. Ordinary references and arithmetic groups do not receive business-semantic labels. Analyst Lens must not turn a positional formula reference into an invented model meaning.

The tree contains dependency ranges reported by Excel:

- The root is visually distinct.
- A chevron indicates whether a node can be expanded or is already expanded.
- Indentation communicates depth.
- The selected row uses a high-contrast selection treatment.
- Cross-sheet references include the sheet name.
- A same-sheet reference may abbreviate the sheet visually, but its accessible label contains the full reference.
- A multi-cell dependency displays its canonical range and dimensions rather than pretending the top-left value represents the range.
- Dynamic, external, cycle, unavailable, and capped nodes use explicit badges.

Sibling order must be deterministic within a session. Prefer Excel's returned order; if the API does not guarantee it, normalize by workbook, worksheet, and address before publishing state.

### 4.4 Selected-row detail

When facts for the selected node are available, a compact detail line beneath the tree may show its full address and formula. Loading that detail must not block row navigation. A pending detail read shows a subtle loading state, and a stale result is discarded by request ID.

### 4.5 Footer

The footer contains:

- Concise keyboard hints.
- Current node count and cap state.
- Loading or normalized error status.
- Visible **Return to origin** and **Close** controls.

The UI does not rely on undocumented gestures or color alone.

## 5. Interaction contract

### 5.1 Opening a session

1. The user selects one cell or contiguous range in Excel.
2. The user invokes Trace In or Trace Out from a registered shortcut or visible Analyst Lens control.
3. The shared runtime snapshots the active reference as both `origin` and `root`.
4. The shared runtime opens the same-origin Quick Trace dialog if needed.
5. Direct dependencies for the selected direction load lazily.
6. The dialog takes keyboard focus after its initial state is ready.

A multi-area selection, invalid context, or unsupported workbook state produces a clear non-destructive error before the dialog is opened.

### 5.2 Tree keyboard behavior

| Input | Behavior |
| --- | --- |
| Up / Down | Move to the previous or next visible row and request Excel navigation to that range |
| Right | Expand the selected node one dependency level; if already expanded, move to its first child |
| Left | Collapse an expanded node; otherwise move to its parent |
| Return | Make the selected node the new root while preserving the original session origin |
| T | Toggle Trace In / Trace Out when focus is not inside a text control |
| O | Return to the session origin when focus is not inside a text control |
| Escape | Close Quick Trace |
| Command+W | Allow the native dialog window close behavior |

Pointer controls mirror every core action. Clicking a row selects and navigates to it; clicking a chevron expands or collapses without requiring a re-root.

### 5.3 Excel navigation and focus

The critical experience is that Up and Down continue to operate the dialog while Excel visibly activates the corresponding worksheet and range.

The dialog sends a typed `navigate` intent to the shared runtime. The runtime performs the Excel navigation and acknowledges the result. The dialog does not call Excel directly and does not intentionally relinquish keyboard focus.

The selected Excel range is the navigation highlight. P0 does not shade rows, columns, or cells because a formatting-based highlight would mutate the workbook and can clear Excel's Undo stack.

### 5.4 Selection-change correlation

Programmatic Quick Trace navigation triggers Excel selection-change events. Those events must not silently reset the tree.

Before navigating, the controller records a short-lived correlation containing the session ID, intent sequence, and target canonical reference. When the matching selection event arrives, it is acknowledged as dialog-initiated and preserves `origin`, `root`, expansions, and selected row.

If Excel raises an unrelated selection event while Quick Trace is open:

- Keep the current tree and origin.
- Show **Excel selection changed** with a **Use selection as new root** action.
- Do not silently replace the audit session.

If the active workbook changes, disable navigation against stale nodes and require an explicit new root or dialog close.

### 5.5 Re-root and origin behavior

Re-rooting starts a fresh branch view at the selected node but does not overwrite the session origin. The prior roots are retained in session navigation history so the existing Back/Forward behavior can be exposed later without changing the state model.

Return to Origin:

1. Activates the origin workbook, sheet, and range when still available.
2. Restores the origin as the tree root.
3. Restores the original session direction unless the user explicitly chooses to keep the current direction in a future setting.
4. Leaves the dialog open and focused.

## 6. Shortcut mapping

Historical Macabacus defaults use `Ctrl+Shift+[` for Pro Precedents and `Ctrl+Shift+]` for Pro Dependents. Excel's Office add-in shortcut catalog rejected terminal punctuation during Analyst Lens qualification, so those chords cannot be copied literally.

The proposed Mac mnemonic fallbacks are:

| Action | Historical Macabacus | Proposed Analyst Lens Mac default |
| --- | --- | --- |
| Open Trace In / precedents | `Ctrl+Shift+[` | `Control+Option+P` |
| Open Trace Out / dependents | `Ctrl+Shift+]` | `Control+Option+D` |

`P` means precedents and `D` means dependents. The reduced modifier set distinguishes the new audit commands from existing `Control+Option+Shift` formatting mnemonics, including the existing Double Bottom Border command.

These become shipping defaults only after the feasibility spike proves all of the following in real Excel:

- The catalog accepts both chords.
- `Office.actions.areShortcutsInUse` reports their conflict state.
- An unoccupied chord activates its registered action with the task pane closed.
- A conflict is displayed and requires explicit user confirmation before replacement.
- Settings and the command palette display the actual assigned chord.

If either default conflicts, Analyst Lens does not silently steal it. The visible ribbon/control path remains available while the user selects another supported letter-based chord.

## 7. Architecture

### 7.1 Process boundary

The Office dialog is a separate web runtime. It cannot own the Excel object model or assume access to shared-runtime memory.

```text
Excel workbook
      ↕ Office.js
Shared runtime / controller
      ↕ typed, validated dialog messages
Floating Quick Trace dialog
```

Responsibilities:

| Component | Owns |
| --- | --- |
| `platform/office` | `Excel.run`, selection activation, dependency and cell-fact reads, capability normalization; isolated Office Dialog API parent/child bridges and lifecycle events |
| `runtime/controller` | Quick Trace session, dialog handle, request IDs, cache use, selection correlation, message routing |
| `domain/dependencies` | Canonical nodes, expansions, cycles, caps, deterministic tree projection |
| `ui/quick-trace` | Rendering, focus, keyboard interpretation, accessible status, typed intents |

The dialog page is served from the same local HTTPS origin as the task pane. It loads Office.js only to use the supported `Office.context.ui.messageParent` transport through a dedicated module under `platform/office`. It never calls `Excel.run` or reads the workbook object model. Workbook access remains in the host shared runtime. The dialog has no network destinations beyond the already approved local assets and Microsoft Office.js runtime policy.

### 7.2 Message protocol

Every message is a discriminated JSON object with a protocol version, session nonce, session ID, and monotonic sequence number. Both sides reject unknown versions, message types, malformed references, and stale session nonces.

Dialog-to-host intents include:

- `quickTrace.ready`
- `quickTrace.navigate`
- `quickTrace.expand`
- `quickTrace.collapse`
- `quickTrace.setDirection`
- `quickTrace.reroot`
- `quickTrace.returnToOrigin`
- `quickTrace.useExcelSelection`
- `quickTrace.close`

Host-to-dialog events include:

- `quickTrace.initialize`
- `quickTrace.statePatch`
- `quickTrace.navigationResult`
- `quickTrace.externalSelection`
- `quickTrace.error`
- `quickTrace.close`

Message handlers publish plain serializable data only. No Office proxy objects cross the boundary.

### 7.3 Session state

The controller owns at least:

```ts
type QuickTraceSession = Readonly<{
  sessionId: string;
  nonce: string;
  direction: "precedent" | "dependent";
  origin: CellRef;
  root: CellRef;
  selectedNodeId: string;
  graph: DependencyGraphState;
  expandedNodeIds: readonly string[];
  visibleNodeIds: readonly string[];
  pendingExternalSelection?: CellRef;
  requestSequence: number;
  status: "opening" | "ready" | "loading" | "error";
}>;
```

The exact implementation type may differ, but ownership and invariants must remain explicit and testable.

### 7.4 State invariants

- `origin` changes only when the user begins a new session or explicitly accepts a new Excel selection as the root.
- `selectedNodeId` always identifies a visible node when status is `ready`.
- A direction switch preserves the origin and root but may use a separate cached expansion projection.
- An aborted or stale request cannot mutate the visible tree.
- The graph never exceeds 500 canonical nodes.
- A closing or closed dialog cannot receive a state patch.
- Only one live dialog handle exists per Excel host window.

## 8. Dependency and value truthfulness

Excel dependency APIs remain the source of truth for edges. The formula tokenizer may improve display and risk badges but cannot add an edge merely because text resembles a reference.

Values are loaded through the existing cell-facts gateway:

- A single cell shows its current value using a safe display representation.
- A contiguous multi-cell range shows dimensions and does not claim a scalar value.
- A value that cannot be loaded shows `Unavailable`, never a guessed or stale value without a stale indicator.
- Formula/value data is never written to production logs.

No P0 field claims why a reference exists. Finance-specific names can be shown only when they come from explicit workbook evidence such as a defined name or label feature already covered by a separate approved design.

## 9. Error and degraded states

| State | User-visible behavior |
| --- | --- |
| No direct dependencies | Empty tree message naming the current direction |
| Dynamic unresolved reference | Non-expandable `Dynamic` leaf with explanation |
| Closed external workbook | Non-expandable `External` leaf; no automatic open |
| Cycle | Cycle badge and link back to the existing canonical node; no recursive expansion |
| 500-node cap | Preserve loaded nodes, mark the capped branch, and advise re-rooting narrower |
| Stale async response | Discard silently; keep the latest loading state authoritative |
| Navigation failure | Keep dialog selection, show normalized error, and offer retry |
| Manual Excel selection | Preserve tree; offer explicit new-root action |
| Workbook changed/closed | Freeze stale navigation and require restart or close |
| Dialog messaging failure | Close the dialog session cleanly and keep the task pane/runtime usable |
| Unsupported Dialog API | Disable Quick Trace controls and report the required host capability; do not open a docked substitute |

## 10. Accessibility and visual behavior

- The tree uses standard `tree` / `treeitem` semantics with `aria-expanded`, level, set size, and position where supported.
- Direction controls expose pressed/selected state and do not rely on color.
- Loading, navigation failures, and external-selection notices use appropriate polite or assertive live regions.
- Focus starts on the root row after initialization.
- Focus remains visible at all times.
- Text and controls meet WCAG AA contrast in light and dark macOS appearance.
- The formula wraps without forcing horizontal window growth.
- Resize down to the Office-supported minimum without hiding Close or Return to Origin.
- Values align for scanning but remain accessible as text.

## 11. Performance budgets

After direct dependency data is loaded:

- Up/Down selection and visible dialog update: p95 at or below 100 ms.
- Excel range activation acknowledgement: p95 at or below 500 ms.
- Cached one-level expansion: p95 at or below 250 ms.
- Uncached one-level expansion of up to 50 nodes: p95 at or below 1 second.
- Direction toggle with cached direct dependencies: p95 at or below 250 ms.

The dialog renders only visible tree rows when the flattened tree becomes large. Office reads remain batched, cancellable where possible, and guarded by request IDs.

## 12. Privacy and mutation contract

Quick Trace is audit-only:

- It does not change values, formulas, formats, names, worksheet visibility, protection, calculation settings, or workbook structure.
- It does not capture a revert snapshot because it performs no workbook writes.
- Excel selection and worksheet activation are navigation, not workbook mutation.
- It does not send workbook data off-device.
- Local production logs may record command IDs, durations, capability flags, and normalized error codes, but not workbook names, sheet names, addresses, formulas, or values.

## 13. Mandatory real-Excel feasibility spike

No product implementation proceeds past a minimal spike until a timestamped Excel 16.111 run proves the Office Dialog API can deliver the core experience.

The spike must prove:

1. `Control+Option+P` and `Control+Option+D` are accepted or produce a visible conflict result.
2. An accepted shortcut opens Quick Trace while the task pane is closed.
3. Exactly one dialog is open per Excel host window.
4. The window is visibly floating over Excel and can be moved and resized.
5. Up/Down remains captured by the dialog while the shared runtime activates the corresponding Excel range.
6. Right/Left expands and collapses a test tree without moving Excel's worksheet cursor accidentally.
7. Trace In / Trace Out can switch within the same dialog.
8. A dialog-initiated selection event preserves the origin and tree.
9. A manual Excel selection is distinguishable or at minimum does not silently destroy the session.
10. Close and reopen produces a fresh, correctly initialized session without stale messages.

Evidence must record Excel version, macOS version, add-in commit, starting workbook copy, starting cell, shortcut/control used, exact result, and screenshots or video frames.

If items 2, 4, or 5 fail, stop and report a product blocker. A task-pane fallback is explicitly out of scope unless the owner later rescopes the floating-window requirement.

## 14. Testing strategy

Implementation follows strict TDD with one focused verified commit per task.

### Pure unit tests

- Tree flattening and deterministic order.
- Up/Down/Right/Left state transitions.
- Re-root and Return to Origin invariants.
- Direction-specific cached projections.
- Cycle, unresolved, multi-cell, and cap presentation.
- Stale request and selection-correlation behavior.
- Protocol validation and stale nonce rejection.

### Component tests

- Two-column Reference / Value contract with no Argument field.
- Accessible tree semantics and visible focus.
- Keyboard and pointer parity.
- Loading, empty, external-selection, error, and cap states.
- Formula wrapping and narrow-window behavior.

### Contract tests

- Host/dialog message serialization and validation.
- One-dialog lifecycle and cleanup.
- Office action registration and shortcut conflict handling.
- Gateway navigation, dependency reads, batching, cancellation, and normalized errors.
- No Excel object-model calls from the dialog; Office.js use is isolated to the dialog message bridge.

### Browser tests

- Open, navigate, expand, collapse, switch direction, re-root, and return to origin using the keyboard.
- Reopen after close without stale state.
- Resize at supported minimum and large sizes.

### Real Excel QA

After the spike passes, run the completed workflow against a timestamped copy of the qualification workbook. Never write-test the source workbook. Re-hash the source before and after QA and retain the existing MVP acceptance blockers, including the grouping-acceptance blocker, until resolved or explicitly rescoped.

## 15. Delivery boundaries

The implementation plan should split this feature into small verified commits, beginning with the feasibility spike and ending with real Excel QA. No implementation task may claim the overall Analyst Lens MVP complete while the grouping-acceptance blocker or a Quick Trace feasibility blocker remains unresolved.

Suggested delivery slices, to be refined in the implementation plan:

1. Real-Excel dialog/focus/shortcut feasibility spike.
2. Typed dialog protocol and one-dialog lifecycle.
3. Pure Quick Trace tree projection and keyboard state machine.
4. Shared-runtime session orchestration and selection correlation.
5. Accessible floating UI and visible controls.
6. Shortcut/settings integration and conflict handling.
7. Automated verification, packaging, and real Excel qualification-copy QA.

## 16. Source references

- Macabacus Precedents & Dependents: https://help.macabacus.com/article/280-precedents-dependents
- Microsoft Office Dialog API: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
- Microsoft custom keyboard shortcuts: https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts
- Analyst Lens umbrella design: `2026-07-15-excel-analyst-addin-design.md`
- Analyst Lens usable MVP plan and qualification notes: internal execution plan, not shipped in the public tree
