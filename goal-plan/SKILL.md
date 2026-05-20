---
name: goal-plan
description: Turn a rough idea into a tight, one-shottable design spec optimised for Claude Code's /goal command. Use whenever the user wants to plan, prep, scope, or design a piece of work they intend to hand off to /goal — including phrases like "plan a goal for X", "help me write a /goal", "scope this for /goal", "I want to one-shot X with /goal", or any time the user mentions /goal alongside a task they haven't fully scoped yet. Also trigger when the user describes ambitious autonomous work (migrations, refactors, backlog clearing, dashboards, feature builds) and indicates they want to run it unattended. Interviews the user adaptively, translates fuzzy intent into mechanically verifiable exit criteria that force Claude to surface proof in the transcript, and outputs both the /goal command string and a lightweight design spec.
---

# /goal-plan

A planning skill for Claude Code's `/goal` command. Bridges the gap between "I want to build X" and a `/goal` condition tight enough to one-shot.

## Why this exists

`/goal` is only as good as its completion condition. Vague conditions like "make it production-ready" burn tokens forever because the evaluator (Haiku by default) reads only the transcript — it can't judge anything Claude's own output doesn't surface. Most users write bad conditions on their first try.

This skill runs the front-half: an adaptive interview that captures intent, translates it into transcript-verifiable exit criteria, and outputs a lightweight spec the user can hand off.

## Principle 1 — transcript-only

**The evaluator sees only what Claude has typed. Nothing else.**

It does not run commands. It does not read files. It does not visit URLs. It does not inspect the spec file. It reads the conversation transcript and answers yes/no.

And it reads shallow. Bias to literal patterns (string match, exit code) over analysis (counts, comparisons, judgement).

Every translation, every criterion, every condition must pass this test:

> *"Reading only what Claude has typed across N turns, can Haiku conclude done?"*

If the answer requires the evaluator to do anything else — broken condition. Rewrite.

This is the load-bearing principle. Everything below is a consequence of it.

## Principle 2 — force the worker to surface proof

Because the evaluator can't run anything, **Claude (the worker) has to type the evidence into the transcript**. The condition must direct Claude's output behaviour, not just the work.

Wrong: `all tests in /test/auth pass`
Right: `run \`npm test test/auth\`, quote the full output in the same turn, exit code must be 0`

Wrong: `migration is scoped to src/api/v2/`
Right: `final turn must contain \`git diff --stat\` output showing only paths under src/api/v2/`

Wrong: `no console errors in the dashboard`
Right: `run headless playwright against /dashboard, paste the full console log into the turn, log must contain zero \`error\` level entries`

**"Surfacing proof" is part of the condition, not a hint.** This is where most conditions silently fail — the work happens but the evaluator can't see it.

## Principle 3 — the condition has two jobs

The condition string is read twice with different jobs:
1. **Turn 1**: a directive — what Claude should do
2. **Every turn after**: a checklist — has it happened yet?

A good condition reads cleanly as both. Most failed conditions are good at one and bad at the other.

**Good directive, bad checklist:**
> `Migrate all v1 API calls to v2 and ensure tests pass.`
>
> Reads fine as instructions. But "migrate" is not transcript-verifiable — Haiku can't tick that box without seeing proof.

**Good checklist, bad directive:**
> `final turn must contain git diff with only src/api/v2/ paths AND npm test output with exit 0`
>
> Reads fine as a checklist. But Claude on turn 1 doesn't know *what* to migrate, *from* what, or *what tests*.

**Good as both:**
> `Migrate all v1 API calls in src/ to v2 per docs/specs/api-v2-migration.md. After each module is migrated, run its test file and quote output. Final turn must contain: (a) git diff --stat showing only src/api/v2/ and src/api/v1/ paths modified, (b) npm test full output with exit 0, (c) grep -r "v1/" src/ output showing zero matches. Stop after 40 turns and report status if not complete.`

Write every condition to pass both reads.

## Principle 4 — turn cap is mandatory

`/goal` has no native budget cap. Without one, a slightly-wrong condition burns indefinitely.

**Every condition includes a turn cap, every time.** Even surgical jobs.

Template: `Stop after N turns and output a status report if condition not met.`

Rough N by tier: S=5, M=15, L=40, XL=don't use /goal.

This isn't a harness, it's a circuit breaker.

## Operating style

**Light by default.** The spec is the minimum needed for `/goal` to verify done — not a PRD. Scale only when work demands.

**Adaptive depth.** Function rename ≠ analytics pipeline. Match interview depth to ambition.

**Propose, then confirm.** Users know what "done" feels like, not what command proves it. Translate intent into mechanical criteria, propose back, confirm. Misalignment hides in this step.

**One question at a time when fuzzy.** Batch only when the user has signalled clarity.

**Dry-run the evaluator before handoff.** Mandatory final step (see Step 6).

## How to use

### Step 1 — Capture the seed

Ask one open question: **"What are you trying to ship?"**

Listen. From the answer, infer:
- **Work type**: greenfield, migration, bug fix, refactor, backlog clearing, dashboard, infra, content
- **Ambition tier** (next step)
- **Whether they know what "done" looks like** or are still figuring it out

Don't jump to framework. Just listen.

### Step 2 — Triage ambition

Pick a tier silently. Let it shape interview depth and turn cap.

| Tier | Signal | Depth | Turn cap |
|------|--------|-------|----------|
| **S — Surgical** | Single function/file, clear scope, known fix | 2–3 questions | ~5 |
| **M — Medium** | Multi-file, one module, deterministic | 4–6 questions | ~15 |
| **L — Large** | Multi-module, design judgment, ambiguous done | Full interview + sanity check | ~40 |
| **XL — Not /goal** | Taste-dependent, layout, UX feel, no finish line | — | Don't ship |

**If XL: say so.** Don't force a bad fit. Suggest two-phase: interactive design first, then `/goal` for the build.

### Step 3 — Adaptive interview

Cover these dimensions only as deep as the tier demands. Reorder by where the user is weakest.

**Always cover (even S):**
1. **End state** — What exists when done? (the thing, not the feeling)
2. **Verify check + surfacing** — What command proves it AND how does its output land in the transcript?

**M and above:**
3. **Scope boundary** — Which files/modules in scope? Out of scope?
4. **Constraints / invariants** — What must not change?

**L and above:**
5. **Work shape** — Greenfield, migration, fix, refactor — affects condition phrasing
6. **Failure modes** — Most likely way this goes wrong. Bake a guard in.

**The high-leverage move: translation craft.**

When the user says something fuzzy, translate it back and confirm. Examples:

| User says | Translate to |
|-----------|--------------|
| *"I want the dashboard to feel polished"* | `headless playwright loads /dashboard, full console log pasted into turn, log contains zero error-level entries, every chart in docs/specs/dashboard.md has a matching [data-testid] that returns non-empty content via document.querySelectorAll output quoted in turn` |
| *"Make the migration safe"* | `final turn must contain (a) git diff --name-only output showing only src/api/v2/ paths modified, (b) npm test full output with exit 0, (c) git log --oneline of test files since branch start showing zero commits` |
| *"Clear the bug backlog"* | `For each Jira issue with label:bug status:todo as of run start: quote the issue key, state action taken (closed/PR linked), and quote the resulting status. Final turn lists all keys processed and any skipped with reason.` |

Make the translation explicit, every time. That's the value.

### Step 4 — Anti-pattern check

Run the condition against these. If two or more fire, push back before producing the spec.

| Anti-pattern | What it looks like | Why it fails |
|--------------|---------------------|--------------|
| Unfalsifiable | "production-ready", "clean", "polished" | Evaluator can't verify from transcript |
| Self-judging | "Claude is confident it's done" | Claude judging Claude |
| No measurable end | "Refactor for clarity" | Nothing to check |
| Evaluator can't see proof | "The deployed app works" | Evaluator doesn't visit URLs |
| No surfacing instruction | "Tests pass" (without "quote output") | Work happens, evaluator can't see it |
| Scope creep | "Fix all bugs" with no list/cutoff | Goalposts move forever |
| Missing constraint | "Migrate to v2" without "without breaking tests" | Claude may break things to satisfy condition |
| Design judgment | "UI looks good" | Needs a human |
| No turn cap | Condition with no stop clause | Cost blowup |

### Step 5 — Produce the artifacts

**Output two things.**

**1. The `/goal` command string** — primary deliverable.

```
/goal <directive>. <surfacing instructions>. <final-turn checklist>. <constraints>. Spec: <path>. Stop after N turns and output status if not met.
```

**2. The design spec file** — saved to `docs/specs/<kebab-name>.md`.

**Important: the evaluator never reads this file.** The spec is for the worker (Claude), to give it context the condition string can't carry. The condition must instruct Claude to surface compliance with the spec *into the transcript* — e.g., "after completing each section in docs/specs/X.md, state which section was just completed and quote the verification output."

If the condition doesn't reference the spec, the spec is irrelevant. If it references the spec without telling Claude to surface compliance, the evaluator still can't see it.

**Spec template (light):**

```markdown
# <Name>

## Goal
<one sentence — what done looks like>

## Acceptance criteria
- [ ] <criterion 1, mechanically verifiable>
- [ ] <criterion 2>

## Scope
**In:** <files, modules, behaviours>
**Out:** <explicitly>

## Constraints
- <invariant 1>

## Verification
Run: `<command>`
Expect in transcript: <output pattern, exit code, file state>

## Surfacing instructions for Claude
- After <step>, quote <output> in the same turn
- Final turn must contain: <list>

## Notes
<context, gotchas, links>
```

Only include sections that earn their place. S-tier may need only Goal + Acceptance + Verification + Surfacing.

### Step 6 — Dry-run the evaluator (mandatory)

Before handing off, run this exercise with the user:

> *"Imagine Claude has just finished. Mentally scan its last 3–5 turns. If you were Haiku reading only those turns, would you say done? If not — what specific phrase would Claude have had to type for you to say yes?"*

Whatever's missing goes into the surfacing instructions.

This catches what the anti-pattern table doesn't: the *specific gaps* where Claude did the work but didn't type the evidence.

One round. Then hand off.

## Worked examples

### Example A — Surgical (S)

**User:** *"There's a flaky test in test/auth/login.test.ts that fails maybe 1 in 10 runs. Need to fix it."*

**Interview:**
- End state? *"Test passes reliably."*
- How do we prove it? *"Run it 20 times in a row, all pass."*
- Anything that mustn't change? *"Don't change the actual auth logic, just the test."*

**Final `/goal`:**
```
/goal Fix the flaky test in test/auth/login.test.ts. Modifications limited to that file only. After fixing, run `for i in {1..20}; do npm test test/auth/login.test.ts; done` and quote full output. Final turn must contain (a) git diff showing only test/auth/login.test.ts modified, (b) the full 20-run output, (c) confirmation all 20 runs exited 0. Stop after 5 turns if not met.
```

No spec file needed — condition carries all of it.

### Example B — Medium (M)

**User:** *"I want to migrate our PostHog event tracking calls in src/analytics/ to the new SDK. There's about 30 of them."*

**Interview:**
- End state? *"All calls use the new SDK, old import removed."*
- Verify? *"Tests pass, no references to old SDK anywhere."*
- Scope? *"Just src/analytics/, don't touch the components that call into it."*
- Constraints? *"Don't change the event names or payloads, just the SDK calls. Existing analytics tests must keep passing."*

**Spec at `docs/specs/posthog-sdk-migration.md`:**
```markdown
# PostHog SDK migration

## Goal
Migrate all src/analytics/ calls from posthog-js v1 to posthog-js v2.

## Acceptance criteria
- [ ] All imports of `posthog-js` in src/analytics/ use v2 syntax
- [ ] Zero references to v1-only methods (.capture, .identify legacy form)
- [ ] Event names and payloads unchanged
- [ ] All tests in test/analytics/ pass

## Scope
**In:** src/analytics/**
**Out:** Any file outside src/analytics/, including callers

## Constraints
- Event names unchanged
- Event payloads unchanged
- No edits outside src/analytics/

## Surfacing instructions for Claude
- After each file migration, quote git diff for that file
- Final turn must contain: full output of `grep -rn "posthog" src/analytics/`, `npm test test/analytics/`, `git diff --stat`
```

**Final `/goal`:**
```
/goal Execute the migration in docs/specs/posthog-sdk-migration.md. After each file, quote its git diff. Final turn must contain: (a) `grep -rn "posthog" src/analytics/` output showing only v2 patterns, (b) `npm test test/analytics/` full output with exit 0, (c) `git diff --stat` showing only src/analytics/ paths. Stop after 15 turns and report status if not complete.
```

### Example C — Large (L)

**User:** *"Build a Tech Diagnostics dashboard page that shows the 5 tribe-level metrics we agreed on. Pulls from the existing API endpoints."*

**Interview:**
- End state? *"Page renders, all 5 charts show data, no console errors."*
- Verify? *"Playwright test loads the page and checks each chart."*
- Scope? *"New page at src/pages/diagnostics.tsx and any new components in src/components/diagnostics/. API endpoints already exist."*
- Constraints? *"Don't modify the API. Don't change other pages."*
- Failure mode I'm worried about? *"Claude inventing endpoints that don't exist or skipping charts."*
- What charts exactly? *"Listed in docs/specs/tech-diagnostics-dashboard.md."*

**Spec lists all 5 charts with their endpoint, chart type, and acceptance criterion each.**

**Final `/goal`:**
```
/goal Build the diagnostics dashboard per docs/specs/tech-diagnostics-dashboard.md. Implement only the 5 charts listed there; do not invent endpoints. After each chart, run the playwright assertion for that chart and quote its output. Final turn must contain: (a) `git diff --stat` showing only paths under src/pages/diagnostics.tsx and src/components/diagnostics/, (b) `npx playwright test diagnostics` full output with exit 0, (c) confirmation each of the 5 chart [data-testid] selectors returned non-empty content (quote the assertion lines), (d) `grep -rn "src/api" src/pages/diagnostics.tsx` output showing only references to endpoints listed in the spec. Stop after 40 turns and output a status report if not complete.
```

## When to refuse or redirect

- **Design-dependent work** (taste, layout, copy, UX feel) — say so. Propose two-phase: interactive design to lock the spec, then `/goal` for the build.
- **User can't articulate a verify check after two attempts** — `/goal` isn't right for this. Don't manufacture one.
- **Tiny work** — point out it'll cost more in evaluator turns than running it interactively.

## Output style

Match user energy. Terse when they're moving fast, thorough when they're working through complexity. Default concise. Don't lecture on `/goal` mechanics — assume they know.

The deliverable is the spec + command. Everything else is supporting work.
