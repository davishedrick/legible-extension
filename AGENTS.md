# AGENTS.md

Rules for this Chrome extension repo:
# Agents.md — Scriptor Extension

## Purpose

This document exists to maximize extension reliability while minimizing token usage.

The objective is not documentation.

The objective is:

1. Accurate session tracking.
2. Accurate word count tracking.
3. Accurate document binding.
4. Reliable extension behavior.
5. Prevention of regressions.

Every bug fixed should make the extension more reliable.

Every discovered bug should improve scenario coverage.

---

# Core Principle

The extension exists for one purpose:

Track writing activity accurately inside Google Docs.

When evaluating a bug, feature, or refactor:

Prioritize:

- Session accuracy
- Word count accuracy
- Document identity accuracy
- Project synchronization
- User data safety

Avoid:

- Unnecessary abstractions
- Large refactors
- Duplicate systems
- Process-heavy documentation

---

# Extension Architecture Philosophy

The extension should always know:

1. Which Google Doc it is attached to.
2. Which Scriptor project it is attached to.
3. Whether a writing session is active.
4. The current document word count.
5. The last synchronized word count.

If any of these become ambiguous:

Treat it as a bug.

---

# Bug Resolution Workflow

## Step 1 — Reproduce

Before modifying code:

Identify:

- Exact user workflow
- Starting state
- Expected behavior
- Actual behavior

Never guess.

Always reproduce if possible.

---

## Step 2 — Find Root Cause

Identify:

- Exact file(s)
- Exact function(s)
- Exact state transition(s)
- Exact event sequence

Root cause must be explainable clearly.

Example:

Bad:

"Session handling appears broken."

Good:

"The session stores the active tab ID but not the originating document ID. When ending from another document, the wrong word count is used."

---

## Step 3 — Prove Root Cause

Verify the identified cause through:

- Logging
- State inspection
- Reproduction
- Existing tests

Do not fix assumptions.

---

## Step 4 — Implement Fix

Fix the smallest surface area possible.

Prefer:

- Existing state structures
- Existing messaging systems
- Existing storage systems

Avoid:

- Rewrites
- New architectures
- New persistence layers

Unless absolutely necessary.

---

## Step 5 — Verify

Required:

### Direct Test

Original bug is fixed.

### Neighbor Test

Related workflows still function.

### Regression Test

Bug cannot silently return.

---

# Scriptor Extension Invariants

These rules must always be true.

If any become false, a bug exists.

---

## Session Invariants

A session belongs to exactly one document.

A session belongs to exactly one project.

A session has exactly one starting word count.

A session has exactly one ending word count.

Word deltas must always be:

Ending Count - Starting Count

Negative sessions are valid.

---

## Document Invariants

A Google Doc may only be bound to one project.

A project may have multiple bound documents.

Deleting a document must remove any associated binding.

A deleted document must never appear as bound.

A document identity must remain stable across browser refreshes.

---

## Binding Invariants

When binding:

Extension must immediately:

1. Identify document.
2. Read current word count.
3. Sync project.
4. Store starting word count.

The first bind establishes baseline counts.

The first bind must never generate a writing session.

The first bind must never generate catch-up words.

---

## Synchronization Invariants

Current document count must never unexpectedly become zero.

If a count becomes zero:

Validate before saving.

Never overwrite valid counts with invalid counts.

Document synchronization must always compare:

Current Document Count

against

Last Known Document Count

Never against unrelated projects.

Never against unrelated documents.

---

## Catch-Up Session Invariants

Catch-up sessions only exist to account for writing performed outside tracked sessions.

Catch-up sessions should trigger:

### Binding Mismatch

Project count differs from document count.

### Pre-Session Mismatch

Last known document count differs from current document count.

Catch-up sessions must never appear:

- During active writing
- While typing
- Mid-session

---

## Multi-Tab Invariants

The most fragile area of the extension.

Every session must store:

- Project ID
- Document ID
- Tab ID (if applicable)

Document identity is always more important than tab identity.

Switching tabs must not:

- Reset session
- Reassign session
- Recalculate from another document

Ending a session must always reference:

The document that started the session.

Not the currently viewed tab.

---

## Minimized Extension Invariants

When minimized:

Session remains active.

Timer remains active.

Tracking remains active.

Word counts continue updating.

Only UI visibility changes.

Expanding restores full interface.

Closing browser behavior remains unchanged.

---

# Scenario Driven Development

Every bug reveals a missing scenario.

Whenever a bug is fixed:

Ask:

"What user workflow exposed this issue?"

If missing:

Add it to this extension repo's `SCENARIOS.md`.

---

# Required Scenario Categories

Maintain coverage for these categories.

---

## Binding

### New Document Binding

Create project.

Bind document.

Verify starting count.

---

### Existing Document Binding

Create project.

Document already contains words.

Bind.

Verify baseline count.

---

### Deleted Document

Bind document.

Delete document.

Verify binding removed.

---

## Sessions

### Normal Session

Start session.

Write words.

End session.

Verify delta.

---

### Negative Session

Start session.

Delete words.

End session.

Verify negative delta.

---

### Long Session

Start session.

Leave running.

Return later.

End session.

Verify counts.

---

## Tab Switching

### Same Document Different Tabs

Switch tabs.

Return.

End session.

Verify counts.

---

### Different Documents

Start session in Document A.

Open Document B.

Attempt interaction.

Verify session remains tied to Document A.

---

### Wrong Tab End

Start session in Document A.

Move to Document B.

End session.

Verify Document A remains source of truth.

---

## Synchronization

### Sync After Writing Outside Session

Write without session.

Start session.

Verify catch-up detection.

---

### Sync After Binding

Bind existing document.

Verify immediate synchronization.

---

### Zero Count Protection

Valid count exists.

Temporary read failure returns zero.

Verify valid count preserved.

---

## Minimized Mode

### Minimize Active Session

Start session.

Minimize.

Continue writing.

Restore.

Verify counts.

---

### Minimize Then Switch Tabs

Start session.

Minimize.

Switch tabs.

Return.

Verify session survives.

---

# Testing Priorities

Highest priority:

1. Word count accuracy
2. Session accuracy
3. Binding accuracy
4. Multi-tab behavior
5. Synchronization
6. Catch-up sessions

Lower priority:

- Visual issues
- Styling
- Layout
- Copy changes

Never risk a tracking system to improve UI.

---

# BUGS.md Rules

Use this extension repo's `BUGS.md` for extension bugs only.

Use this extension repo's `SCENARIOS.md` for extension workflow coverage only.

Keep bug entries short.

Format:

## Open

### EXT-001

Status:
Open

Summary:

Reproduction:

---

## Fixed

### EXT-001

Root Cause:

Fix:

Regression:

Scenario Added:

---

# Token Efficiency Rules

Prefer:

- Regression tests
- Scenario coverage
- Small fixes

Avoid:

- Long reports
- Excessive summaries
- Rewriting documentation
- Duplicate explanations

The extension becomes more stable through tests and scenarios, not documentation volume.

---

# Definition of Done

A bug is complete only when:

1. Root cause identified.
2. Root cause proven.
3. Fix implemented.
4. Original bug reproduced before fix.
5. Original bug cannot reproduce after fix.
6. Regression test exists.
7. Scenario exists or is updated.

Only then may the issue be closed.

---

# Success Definition

The extension succeeds when:

- Sessions are accurate.
- Word counts are accurate.
- Document bindings are accurate.
- Synchronization is reliable.
- Multi-tab behavior is predictable.

Every fix should move the extension closer to those goals.
