# Analyst Lens Product PRD

**Status:** Product direction approved; written specification pending final review

**Date:** 2026-07-21

**Initial platform:** Microsoft Excel for Mac

**Initial audience:** Solo finance professionals

**Product posture:** Local-first deterministic utility with provider-independent AI access

## Executive Summary

Analyst Lens is the missing finance power layer for Excel on Mac. It helps finance professionals understand, audit, navigate, change, and defend consequential financial models without giving up the Mac workflows they prefer.

The product is not another AI chat pane and does not attempt to compete with ChatGPT, Claude, or Copilot on general model reasoning, workbook generation, research, or content creation. Those providers are rapidly absorbing natural-language spreadsheet creation, financial-model updates, formula explanation, scenario analysis, data connectors, and cross-document workflows.

Analyst Lens instead owns the deterministic operating layer around the workbook:

- Mac-native hotkeys and fast navigation;
- selected-cell inspection and dependency tracing;
- reproducible finance-model checks;
- persistent workbook maps and fingerprints;
- version comparison and change provenance;
- previewed, validated, reversible edits; and
- structured tools that Claude, Codex, and future agents can call.

The first public release is a production-grade Mac Excel utility that is valuable without AI. Later releases add a local evidence engine and MCP integrations so users can bring the Claude or Codex account they already use. AI remains replaceable reasoning; Analyst Lens remains the trusted workbook control plane.

## Product Thesis

> Analyst Lens turns an unfamiliar or changing Excel model into an inspectable, verifiable system, then lets the user bring whichever AI model they prefer to reason over trusted evidence.

The product succeeds when it becomes part of a finance professional's Excel muscle memory. It should feel less like opening a separate application and more like gaining the auditing, navigation, and safety capabilities that Excel for Mac should already have.

## Why Now

Three trends create the opening.

### 1. Finance professionals increasingly want Mac-native workflows

The target user has chosen a Mac partly to remain close to leading-edge desktop AI experiences, but still depends on Excel for serious modeling. This is a target-market hypothesis grounded in the founder's own workflow; the public beta will validate its breadth rather than presenting it as an established market statistic.

### 2. The strongest traditional finance add-ins remain Windows-centered

Macabacus states that Office for Mac does not support the sophisticated COM add-ins on which its full product depends and recommends running Windows in a virtual machine. Newer cross-platform products validate demand for formula auditing and finance productivity, but none yet owns the combination of Mac muscle memory, deterministic workbook intelligence, and provider-independent agent control.

### 3. General AI spreadsheet work is becoming a platform capability

ChatGPT for Excel, Claude for Excel, Microsoft Copilot, and role-specific finance agents already build, update, explain, audit, and scenario-test workbooks. OpenAI's Public Equity Investing and Investment Banking plugins cover model updates, diligence, valuation, and deliverable creation. Anthropic offers finance-agent templates including model builder, valuation reviewer, earnings reviewer, and statement auditor.

Generic workbook chat is therefore not a durable independent product thesis. Analyst Lens must own the evidence, controls, interaction speed, and change safety beneath those agents.

## Market Boundary

### Capabilities we will not claim as differentiation

- Explain a formula.
- Audit a workbook through a prompt.
- Build a DCF, LBO, merger model, or three-statement model from scratch.
- Update a model for a new quarter.
- Run a scenario or sensitivity table.
- Draft a memo, pitchbook, or board presentation.
- Connect to filings, market data, research, email, or data rooms.
- Produce a narrative audit report.
- Ask before making a large AI-generated edit.

These may appear in workflows powered by Claude or Codex, but they are not the reason Analyst Lens exists.

### Capabilities Analyst Lens must own

1. **Instant Mac workflow:** hotkeys, command palette, tracing, navigation, and formatting that work without prompting.
2. **Canonical workbook evidence:** a reusable dependency graph, workbook fingerprint, snapshots, findings, and change ledger.
3. **Reproducible controls:** versioned rules that return the same findings for the same workbook state regardless of model provider.
4. **Version intelligence:** exact formula, value, structure, and output changes with impact propagation.
5. **Transactional edits:** cell-level preview, approval, validation, rollback, and revert.
6. **Provider independence:** the same evidence and controls exposed to Claude, Codex, or a future compatible client.
7. **Value without AI:** a complete deterministic utility even when no model client is connected.

## Target User

### Primary persona

A solo finance professional who:

- works in investing, banking, corporate development, or strategic/corporate finance;
- spends significant time in multi-sheet financial models;
- uses Excel on a Mac despite weaker professional-modeling add-in support;
- frequently inherits third-party models or returns to older models;
- already uses Claude or Codex separately; and
- is personally accountable for understanding and defending model outputs.

The unifying characteristic is responsibility for a consequential workbook, not a specific title or employer.

### Triggering moment

> I need to understand, validate, or change this model quickly, and I cannot afford to break it.

Examples include:

- receiving a diligence, valuation, operating, or transaction model;
- reviewing another analyst's work;
- revisiting a model after weeks or months;
- investigating why an output changed;
- preparing a workbook for investment committee, CFO, senior banker, auditor, or client review; and
- comparing a new model version with the one previously reviewed.

### Core job to be done

> Help me establish what drives this model, identify what may be wrong, make controlled changes, and prove what changed.

## Goals

1. Deliver a useful model insight within the first minute of use.
2. Make selected-cell inspection feel instantaneous.
3. Let users explain why an output changed using exact workbook evidence.
4. Reduce the time and cognitive load required to orient within an unfamiliar model.
5. Catch material formula, structure, and assumption issues before the workbook is shared or relied upon.
6. Ensure no AI agent can silently change the workbook.
7. Create repeated weekly usage through hotkeys, tracing, and deterministic checks.
8. Let users use their existing Claude or Codex account without buying AI credits from Analyst Lens.

## Non-Goals

- Replacing Excel.
- Replacing ChatGPT, Claude, or Copilot.
- Building complete models from natural-language prompts.
- Becoming a market-data, research, data-room, or document-generation platform.
- Offering a general-purpose chat experience inside the add-in.
- Generating pitchbooks, investment memos, or board decks as first-party features.
- Building enterprise administration, team workspaces, or sales-led onboarding for the initial product.
- Achieving feature parity with every Windows COM add-in.
- Supporting general spreadsheet use cases outside finance.
- Hosting inference or reselling model usage in the initial product.

## Product Principles

1. **Excel is the product surface.** Users should remain in the workbook.
2. **Determinism before intelligence.** Checks and evidence come from code; interpretation may come from a model.
3. **Contextual actions, not generic chat.** AI entry points appear when a cell, finding, snapshot, or change creates a useful next step.
4. **Keyboard first, discoverable second.** Experts gain speed without making the product inaccessible to new users.
5. **Local by default.** Workbook-derived state remains on the Mac unless the user explicitly enables an agent session.
6. **No silent writes.** All agent-proposed edits are previewed and approved inside Excel.
7. **Fail closed.** Uncertain or stale evidence is labeled and cannot authorize a write.
8. **Provider independent.** Analyst Lens behavior must not depend on one model provider's output style.
9. **Honest coverage.** Unsupported constructs produce PARTIAL coverage, not a false pass.

## Core User Journey

1. The user opens a workbook and Analyst Lens maps its available structure.
2. The add-in presents a compact orientation and model-health summary.
3. The user selects an important cell and invokes Inspector or Quick Trace by shortcut.
4. Analyst Lens shows direct drivers, downstream consumers, checks, and relevant model context.
5. The user navigates, watches, formats, groups, or audits through deterministic actions.
6. When AI connectivity exists, the user chooses a contextual action such as **Investigate why this changed** or **Pressure-test key drivers**.
7. Claude or Codex calls bounded Analyst Lens tools and receives structured workbook evidence.
8. If the agent recommends a change, Analyst Lens renders a cell-level proposed change set.
9. The user approves or rejects the proposal in Excel.
10. Analyst Lens applies an approved transaction, recalculates, validates, records the result, and preserves a revert point.

## Functional Architecture

### 1. Excel for Mac add-in

The add-in is the primary product and owns:

- adaptive Excel/macOS appearance;
- selected-cell Inspector;
- Quick Trace;
- direct precedent and dependent navigation;
- model-health findings;
- watchlist;
- speed-layer formatting and finance actions;
- keyboard shortcuts and command palette;
- change preview, approval, status, validation, and revert UI; and
- install, status, repair, and uninstall workflows.

It must continue providing deterministic value when the local intelligence service or an AI client is unavailable.

### 2. Local Analyst Lens engine

The local engine is the long-term differentiated core. It owns:

- canonical workbook dependency graph;
- workbook identity and fingerprint;
- local model snapshots;
- formula, value, formatting, link, name, visibility, and structure metadata;
- versioned deterministic rule execution;
- model-version comparison;
- evidence bundles;
- transaction validation and rollback; and
- local change and review ledger.

The engine exposes stable, provider-neutral contracts. Its outputs must be understandable and testable without an LLM.

### 3. MCP bridge

The MCP layer makes Analyst Lens callable from Claude and Codex. It is an interface, not the product moat.

Initial conceptual tools include:

- `get_workbook_summary`
- `inspect_selected_cell`
- `trace_dependencies`
- `list_model_findings`
- `get_finding_evidence`
- `compare_snapshots`
- `explain_change_path`
- `create_proposed_change_set`
- `get_proposed_change_status`

The model client may create a proposal but may not directly execute a workbook change. Application remains an explicit Excel-side action.

### 4. Optional future cloud layer

A remote service may later support sign-in, licensing, shared rule packs, team conventions, and cross-device history. It is excluded from the public beta and is justified only by observed user demand.

## Data Flow and Permissions

### Deterministic mode

- Requires no Analyst Lens account.
- Uses no model provider.
- Runs locally and works offline.
- Sends no workbook content to Analyst Lens infrastructure.

### AI session mode

1. The user explicitly enables Analyst Lens tools from Claude or Codex.
2. Access is scoped to the currently open workbook and current session.
3. The client receives structured tool responses, not unrestricted filesystem access.
4. Analyst Lens returns only evidence relevant to the requested tool call.
5. The client may request a proposed change set.
6. The proposal appears in Excel and remains inert until the user approves it.
7. The Excel add-in applies, validates, records, or rolls back the transaction.

The user's Claude or Codex account supplies inference. Analyst Lens does not obtain a reusable token that lets its own Excel pane consume the user's model subscription.

## AI Interaction Design

### Primary pattern: contextual actions

AI actions appear next to concrete workbook context, for example:

- Investigate why this changed.
- Explain this output's driver chain.
- Pressure-test the key assumptions.
- Review this finding against nearby formula patterns.
- Compare this output with the prior snapshot.
- Propose a safe correction.

The result opens as an evidence view, not as a conversation transcript.

### Secondary pattern: command palette

All contextual actions are also available through the keyboard-first command palette. The palette is the expert accelerator; contextual placement remains the discovery path.

### Explicitly excluded pattern: persistent chat pane

Analyst Lens will not include a general-purpose prompt box or duplicate the Claude, ChatGPT, or Copilot conversation experience.

## Launch Scope and Sequencing

The north-star product is too large for one implementation plan. Work is divided into independent waves with explicit acceptance gates.

### Wave 1: Production-grade Mac utility

**Goal:** Public LinkedIn beta with immediate value and no AI dependency.

In scope:

- production-hardening the existing adaptive UI;
- automatic Excel/macOS light and dark appearance;
- reliable disclosure behavior and accessibility;
- selected-cell Inspector and direct precedents/dependents;
- Quick Trace;
- model-health scan;
- watchlist;
- speed-layer actions;
- existing grouping, formatting, navigation, and revert behavior;
- keyboard shortcuts and command palette;
- clear empty, loading, unavailable, and error states;
- simple install, status, repair, and uninstall;
- sample finance workbook and two-minute onboarding; and
- real Excel for Mac acceptance.

Out of scope:

- local model snapshots and full fingerprints;
- model-version comparison;
- MCP tools;
- AI contextual actions that pretend to operate before an agent bridge exists;
- user accounts, licensing, teams, or cloud synchronization.

Wave 1 implementation is governed by the [approved adaptive production-hardening design](./2026-07-18-analyst-lens-adaptive-production-hardening-design.md). The inert placeholder AI prompt is removed rather than shipped as decorative UI.

### Wave 2: Evidence engine and read-oriented MCP

**Goal:** Give Claude and Codex trustworthy, structured access to the current model.

In scope:

- workbook graph and fingerprint;
- local snapshots;
- formula and structure comparison;
- stable finding identifiers and evidence payloads;
- read-oriented MCP tools for Claude and Codex;
- opt-in session permission UX;
- contextual investigation actions; and
- provider-parity contract tests.

Wave 2 begins only after Wave 1 demonstrates repeated use of tracing, inspection, and model-health workflows.

### Wave 3: Controlled action and personal memory

**Goal:** Let agents propose safe changes while Analyst Lens owns approval and execution.

In scope:

- proposed change-set schema;
- cell-level impact preview;
- approval and rejection UI;
- atomic apply, recalculation, validation, rollback, and revert;
- personal conventions and custom checks;
- model-type review packs; and
- optional community-contributed rule packs.

Wave 3 begins only after users trust the evidence engine and repeatedly ask agents to recommend workbook changes.

## Deterministic Finding Contract

Every finding must include:

- stable rule identifier;
- rule version;
- severity tied to decision impact;
- workbook location;
- concise finding;
- supporting evidence;
- coverage or confidence state;
- recommended human review action; and
- links or commands to inspect the affected cells.

The same workbook state and rule version must produce the same finding regardless of whether Claude, Codex, or no model is connected.

## Change-Set Contract

Every proposed change set must include:

- proposal identifier;
- originating client and session;
- affected workbook fingerprint;
- exact target cells;
- before and proposed after values or formulas;
- stated rationale;
- expected downstream impact;
- required validations;
- approval status;
- application result; and
- revert reference.

If the current workbook fingerprint no longer matches the proposal's fingerprint, the proposal becomes stale and cannot be applied without regeneration.

## Failure and Recovery Model

- **Stale graph:** refresh affected evidence before important findings or writes; block the write if freshness cannot be established.
- **Local engine unavailable:** retain core add-in features and mark evidence/MCP actions unavailable.
- **Agent unavailable or timed out:** leave the workbook unchanged and preserve the deterministic context.
- **Malformed agent output:** reject it at the schema boundary and show a concise explanation.
- **Unsupported workbook construct:** mark coverage PARTIAL and identify the unsupported area.
- **Write conflict:** invalidate the proposal when workbook state differs from the proposal fingerprint.
- **Partial application:** roll back the transaction automatically; surface recovery status and the revert point.
- **Recalculation or validation failure:** roll back unless the user explicitly accepts a reviewed partial outcome in a future workflow; the initial action release remains fail-closed.
- **Crash or restart:** reopen with the last durable transaction state and offer recovery when an incomplete operation is detected.

## Privacy and Telemetry

Workbook-derived state remains local unless the user explicitly invokes an external model client.

Product analytics must not collect:

- workbook files or filenames;
- sheet names;
- cell addresses tied to workbook identity;
- cell values;
- formulas;
- labels;
- comments or notes; or
- agent evidence payloads.

Optional anonymous analytics may collect coarse events such as feature invocation, success/failure category, duration bucket, and explicit useful/noise feedback. Analytics are disabled until the public-beta disclosure and opt-out experience are implemented.

## Production Acceptance

A release is not production-ready until it passes all applicable gates.

### Automated gates

- lint;
- TypeScript/build validation;
- unit and integration tests;
- manifest validation;
- Office API boundary tests;
- network allowlist/no-network checks;
- automated UI flows;
- accessibility checks;
- light and dark appearance regression; and
- performance checks on representative multi-sheet workbooks.

### Real Excel for Mac gates

- install, launch, repair, and uninstall;
- selected-cell updates;
- shortcuts and command palette;
- Inspector and Quick Trace;
- precedents and dependents;
- model-health scan and watchlist;
- formatting/grouping actions;
- recalculation;
- add-in-specific revert;
- save, close, reopen, and retained workbook integrity; and
- timestamped test-copy verification so the original workbook remains untouched.

Coverage must be reported as PASS, PARTIAL, or BLOCKED. Browser-only success cannot qualify a release as production-ready.

### Later MCP/action gates

- provider-parity tests for structured responses;
- explicit current-workbook/session scoping;
- no unrestricted file access;
- no agent-direct writes;
- stale-fingerprint rejection;
- apply/rollback/revert recovery;
- timeout and malformed-response behavior; and
- verified absence of workbook content in product telemetry.

## Launch and Distribution

### Founder-led story

The initial public narrative is:

> I moved my finance workflow to Mac to stay closer to the leading edge of AI tools, but Excel for Mac still lacked the auditing, navigation, and hotkey workflows I relied on. I built Analyst Lens for myself, it has materially improved how I work with models, and I am sharing the beta while I keep improving it.

The launch promise is:

> **The missing finance power layer for Excel on Mac.**

Claude and Codex connectivity is presented as the roadmap, not as functionality required to justify the first release.

### Initial distribution

- free public beta;
- personal LinkedIn post and short workflow demonstrations;
- one frustration and one outcome per demonstration;
- simple self-service install and uninstall;
- sample model for a two-minute first experience;
- public changelog;
- visible near-term roadmap; and
- lightweight feedback from the exact cell, finding, or action in use.

The initial product has no sales calls, team deployment, enterprise onboarding, or AI-credit purchase flow.

## Feedback and Discovery Loop

Useful prompts include:

- Was this finding useful or noise?
- What were you trying to understand?
- Which Excel for Mac shortcut do you miss most?
- What do you repeatedly ask Claude or Codex about this workbook?
- Suggest a check.

Feedback remains optional and must not include workbook content unless the user deliberately attaches or pastes an example.

## Monetization Posture

The public beta is free. Monetization begins only after repeated weekly usage demonstrates a habit and identifies capabilities users would meaningfully miss.

The likely eventual structure is:

- **Free:** core hotkeys, formatting, basic tracing, and limited checks.
- **Individual Pro:** advanced auditing, snapshots, version comparison, custom rules, MCP connectivity, and personal workflow memory.
- **Default billing:** simple annual individual subscription, with final pricing based on observed willingness to pay.
- **Model usage:** no Analyst Lens AI credits; users bring their Claude or Codex account.
- **Enterprise:** no enterprise tier until team demand appears organically.

## Success Metrics

### Primary

1. **Activation:** the user reaches a useful trace or finding in the first session.
2. **Retention:** the user returns weekly during real modeling work.
3. **Habit:** repeated shortcut, trace, navigation, and check usage.
4. **Trust:** low explicit noise rate, no silent changes, successful recovery, and reliable revert.
5. **Pull:** retained users request version comparison, custom checks, or agent connectivity without being prompted.
6. **Conversion readiness:** retained users say they would be meaningfully inconvenienced if Analyst Lens disappeared.

### Secondary

- install completion;
- time to first useful action;
- crash-free sessions;
- feature success/failure rate;
- model-health finding usefulness;
- investigation actions per retained user;
- proposal acceptance/rejection rates after Wave 3; and
- opt-in referrals or public recommendations.

Downloads, LinkedIn impressions, and one-time audit reports are supporting indicators, not the core success measure.

## Investment Gates

### Build Wave 2 when

- users repeatedly use deterministic inspection, tracing, and model health;
- users describe a recurring set of questions they take from Excel to Claude or Codex;
- users request model comparison or persistent review context; and
- the Wave 1 utility meets its real-Excel acceptance bar.

### Build Wave 3 when

- users trust Analyst Lens findings and snapshots;
- users repeatedly ask agents to recommend specific model changes;
- proposed-change review is more valuable than simple explanation; and
- atomic apply, rollback, and revert can be proven in real Excel.

### Consider a cloud/team layer when

- users ask to share standards, checks, or review history;
- individual local state becomes a collaboration bottleneck;
- licensing or distribution requires an account; and
- the privacy cost is justified by demonstrated demand.

## Key Risks and Mitigations

### First-party AI products absorb more auditing UX

**Risk:** ChatGPT, Claude, or Microsoft ships deeper tracing, checks, or model diffs.

**Mitigation:** prioritize instantaneous Mac workflow, reproducible rules, cross-provider evidence, and reversible local control rather than model-specific reasoning.

### Office.js limits professional add-in behavior

**Risk:** Some COM-era interactions remain unavailable or inconsistent on Mac.

**Mitigation:** use runtime capability detection, document PARTIAL coverage, design around supported APIs, and qualify every critical workflow in real Excel.

### False positives destroy trust

**Risk:** Noisy health checks make users ignore the product.

**Mitigation:** stable rule IDs, decision-impact severity, evidence, inspectability, and explicit useful/noise feedback.

### The product becomes another crowded sidebar

**Risk:** Feature accumulation reduces workbook visibility and speed.

**Mitigation:** compact disclosures, contextual actions, keyboard access, progressive detail, and no persistent chat.

### MCP integration leaks too much workbook context

**Risk:** Tool responses expose unnecessary sensitive model information.

**Mitigation:** current-workbook/session scope, bounded tools, minimum relevant evidence, explicit enablement, and telemetry exclusion.

### North-star scope delays launch

**Risk:** Building the local engine and AI bridge postpones real user learning.

**Mitigation:** ship Wave 1 first and require observed pull before each later investment gate.

## Execution Decomposition

This PRD governs product direction but is not a single implementation specification. Execution proceeds through separate designs and plans:

1. **Wave 1 adaptive production hardening** — existing approved design; next implementation plan.
2. **Wave 1 onboarding, installer, and public-beta packaging** — separate design after the hardened utility passes Excel QA.
3. **Wave 2 local evidence engine** — separate architecture and data-contract design.
4. **Wave 2 Claude/Codex MCP integrations** — separate client, permission, and tool-contract design.
5. **Wave 3 controlled change transactions** — separate safety-critical design.
6. **Wave 3 personal rules and workflow memory** — separate design based on beta evidence.

The immediate next implementation plan covers item 1 only.

## Decision Record

The following product choices were explicitly approved during design review:

- model review and inheritance is the beachhead;
- solo Mac finance users are the initial audience;
- Analyst Lens is the deterministic Excel control plane, not a finance agent;
- the architecture is Excel add-in + local intelligence engine + Claude/Codex MCP bridge;
- the public launch is utility-first, followed by the evidence engine and MCP;
- AI appears through contextual actions plus the command palette;
- there is no general-purpose chat pane;
- deterministic use is local, account-free, and offline-capable;
- agents may propose but never directly apply workbook changes;
- production readiness requires real Excel for Mac verification; and
- the initial beta is free and founder-led through LinkedIn.

## Research Sources

- [ChatGPT for Excel and Google Sheets](https://help.openai.com/en/articles/20001063-chatgpt-for-excel-and-google-sheets)
- [OpenAI Public Equity Investing plugin](https://openai.com/business/plugins/public-equity-investing/)
- [OpenAI Investment Banking plugin](https://openai.com/business/plugins/investment-banking/)
- [OpenAI finance solutions](https://openai.com/business/solutions/finance/)
- [Anthropic agents for financial services](https://www.anthropic.com/news/finance-agents)
- [Microsoft Copilot in Excel](https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel)
- [Macabacus system requirements](https://help.macabacus.com/article/263-system-requirements)
- [Excelerate FP&A productivity toolkit](https://www.exceleratefinancialmodelling.com/)
- [Formula Audit XL](https://formulaudit.com/about/)
- [BankerToolBench](https://arxiv.org/abs/2604.11304)
- [AuditFlow](https://arxiv.org/abs/2606.03031)
