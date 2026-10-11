# Analyst Lens Adaptive Production Hardening — Design Specification

**Status:** Approved for planning
**Owner:** Sree Akkineni
**Target host:** Microsoft Excel for Mac 16.111
**Date:** 2026-07-18
**Parent contracts:** `2026-07-15-excel-analyst-addin-design.md`, `2026-07-16-analyst-lens-floating-quick-trace-design.md`

## 1. Decision summary

Refine the current Analyst Lens modernization into a production-ready local Excel for Mac add-in without changing its proven workbook-operation architecture.

The product will automatically follow the host macOS/Excel light or dark appearance. Both modes are first-class presentations of one semantic design system. The current instrument-like hierarchy remains: live and computed state uses teal, audit context uses amber, dangerous or unresolved state uses red, and routine chrome stays neutral.

Deterministic analyst tools remain the product core. The new bottom command surface becomes a real, keyboard-accessible launcher for existing commands and a small set of future contextual AI actions. This pass does not connect an AI provider, recreate a general chat product, or send workbook data off-device.

## 2. Goals and non-goals

### P0 goals

1. Ship a coherent adaptive light and dark appearance that follows `prefers-color-scheme` without a manual theme setting.
2. Repair the disclosure regression so closed content is not visible, focusable, clickable, or exposed to assistive technology.
3. Restore accessible dependency-list names and preserve keyboard navigation, focus return, and status semantics.
4. Make all existing automated verification green under Node 22.23.1 and npm 10.9.8.
5. Keep the task pane usable at 360 by 900 and Quick Trace usable at 420 by 390 and 900 by 700.
6. Replace the inert AI-looking field with an honest command launcher that opens the existing command palette today.
7. Preserve the local-only Office boundary, network allowlist, write/revert safety, grouping contract, and Quick Trace lifecycle.
8. Qualify the installed build in real Excel for Mac against a timestamped QA copy without opening or write-testing the qualification source workbook.

### P0 non-goals

- Connecting OpenAI, Anthropic, or another model provider.
- Building a generic chat transcript, model picker, prompt library, or autonomous workbook editor.
- Changing formula parsing, dependency discovery, formatting behavior, grouping behavior, or persistence policy.
- Publishing, pushing, creating a pull request, or deploying beyond the local Mac.
- Supporting a manual theme override in this pass.

## 3. Product and visual system

### 3.1 Visual thesis

Analyst Lens is a calm financial instrument that feels native beside Excel: dense numerical clarity, restrained chrome, and one obvious working context.

The selected cell remains the visual anchor. Formula evidence, risk notes, precedents, dependents, and neighbor comparison form the primary audit trail. Quick Trace and deterministic commands are secondary actions. Watchlist, formatting, structure, and model health remain progressively disclosed.

### 3.2 Adaptive appearance

One semantic token file owns both modes. The default token values are light. A `@media (prefers-color-scheme: dark)` block supplies dark values. `color-scheme: light dark` allows native controls to follow the same appearance.

Tokens describe purpose rather than color names:

- background, surface, raised surface, and sunken surface;
- primary, secondary, muted, and decorative text;
- routine, strong, and subtle borders;
- live/computed, annotation/audit, warning, danger, and focus states;
- shadows and selection treatments appropriate to each appearance.

Small informational text must meet WCAG AA contrast. Decorative text may use a fainter token only when it does not carry required information.

### 3.3 Information hierarchy

The task pane order remains:

1. host status and selected-cell inspector;
2. contextual command launcher;
3. key-output watchlist;
4. speed tools;
5. model health.

At narrow width, the inspector is a single column and the compact dependency map is omitted from the first reading path. At wide width, the map may sit alongside the audit trail. The command launcher is visually stable and keyboard reachable without pretending that AI is already connected.

## 4. Interaction contract

### 4.1 Progressive disclosure

Each disclosure header is a real button with `aria-expanded` and `aria-controls`. Closed content remains mounted only if internal state preservation requires it, but it must be removed from layout and the accessibility tree using the native `hidden` contract. Opening restores the content without resetting internal state.

Default state:

- Key outputs opens automatically only when at least one output is pinned.
- Speed tools opens by default.
- Model health remains closed until requested.

### 4.2 Contextual command launcher

The current inert “Ask about this cell…” treatment becomes a real button named **Open commands for this cell**. It displays the current keyboard shortcut and opens the existing command palette while returning focus correctly when the palette closes.

The palette remains deterministic in this pass. Its architecture allows later contextual actions to be registered beside commands without changing the inspector:

```ts
type ContextAction = Readonly<{
  id: string;
  label: string;
  availability: "ready" | "coming-soon";
  kind: "command" | "ai";
  description: string;
}>;
```

Coming-soon AI actions may be shown only in an explicitly labeled exploration section and must never appear executable when no provider exists.

### 4.3 Motion and focus

Motion is limited to disclosure, overlay, focus, and status transitions. It must be fast, restrained, and disabled under `prefers-reduced-motion`. Every interactive element has a visible focus treatment in both themes. Closing an overlay returns focus to its invoker.

## 5. Complementary AI direction

Analyst Lens should use model providers for reasoning while owning the Excel-specific context, evidence, and action boundary. It should not compete with OpenAI or Anthropic on general conversation.

The first candidate actions are:

1. **Explain this formula** — package the selected formula, exact references, number format, and lexical risks into a concise provider-ready context bundle.
2. **Check nearby consistency** — compare the selected formula with left/right or row/column peers and ask a model to explain only the material differences already detected locally.
3. **Summarize the trace** — turn the deterministic precedent/dependent tree into a short audit narrative with citations back to exact cell references.
4. **Draft an audit note** — produce a local note from selected evidence without changing the workbook.
5. **Suggest next checks** — recommend deterministic Analyst Lens actions, such as tracing an external link or running model health, instead of directly mutating cells.

These actions share four constraints:

- evidence is assembled locally and displayed before transmission;
- the user explicitly invokes every request;
- the provider adapter is replaceable and does not leak into the Office gateway;
- model output is advisory until a separate, approved write contract exists.

Provider connection, consent UX, redaction, data retention, and cost controls require a separate approved design.

## 6. Architecture and data flow

No workbook-operation boundary changes are required.

```text
Excel workbook
      ↕ existing Office gateway
Inspector controller and deterministic domain state
      ↕ typed UI intents
Adaptive task pane / Quick Trace / command palette
      ↕ future explicit context bundle
Replaceable model-provider adapter (not implemented in this pass)
```

Responsibilities:

| Unit | Responsibility |
| --- | --- |
| `tokens.css` | Semantic adaptive appearance and contrast-safe state colors |
| `Disclosure` | Native expanded/collapsed behavior and state preservation |
| `Inspector` | Selected-cell evidence, accessible dependency lists, and launcher invocation |
| `CommandPalette` | Existing deterministic command discovery; future context-action host |
| `QuickTraceApp` | Compact keyboard-first dependency traversal with stable textual semantics |
| Tests and capture script | Behavioral contracts and current-source visual evidence |

No AI code may call Office.js. A future provider adapter receives plain, validated, serializable evidence assembled from domain state.

## 7. Error handling and safety

- A collapsed section cannot expose visible dead controls.
- Unavailable Quick Trace remains disabled with a visible reason.
- A failed command or workbook operation retains the existing normalized error path and never claims success.
- Theme selection has no JavaScript dependency and cannot block the task pane.
- Missing AI connectivity is represented as unavailable capability, not a failed chat request.
- Production captures must be regenerated from the exact source being qualified and stored under one clearly named `current` directory; stale explorations remain separate.

## 8. Testing and acceptance

Implementation follows red-green-refactor. Each repaired behavior receives a failing regression test before production code changes.

Required automated gates:

1. Disclosure component tests prove closed bodies have the native hidden contract, retain internal state, and reopen accessibly.
2. Inspector tests prove named precedent/dependent lists, all selection states, launcher invocation, and focus return.
3. Theme tests inspect both computed light and dark token behavior and contrast-critical token use.
4. Browser tests exercise health scan and watchlist flows through their disclosure headers.
5. Quick Trace tests use stable accessible references rather than CSS-generated punctuation while preserving reroot and lifecycle behavior.
6. Responsive browser checks cover 360 by 900, 760 by 900, 420 by 390, and 900 by 700 without horizontal overflow or unreachable primary controls.
7. `npm run verify` passes under the pinned runtime.

Real Excel acceptance begins only after the automated gate is green:

- build, uninstall, verify absent, install, status healthy, and reload;
- open only a timestamped qualification QA copy;
- verify automatic light/dark appearance against the current macOS setting;
- exercise inspector navigation, Quick Trace, watchlist, health scan, formatting/revert, and one-layer grouping;
- close without saving and confirm source and QA-copy integrity as required by the existing checklist.

## 9. Completion boundary

This pass is complete only when the adaptive UI, full automated suite, current-source screenshots, and real Excel acceptance all pass. The AI lane is complete for this pass when the command surface and context-action boundary are honest and extensible; no provider integration is implied.
