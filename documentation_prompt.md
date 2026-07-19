# Documentation Generation Prompt — AI Lead Generation Engine

Paste this into your code editor (Claude Code CLI) pointed at the full `f:\Lead_Generation` repo. The goal is a single, complete reference document that anyone — including someone who has never touched this codebase — could use to understand, operate, and maintain the entire system. Do not summarize or skip files; read every `.gs` file, every sheet/tab referenced in code, and every config value actually in use before writing a single line of documentation. Where the code and any prior spec disagree, document what the code actually does, and flag the discrepancy separately.

---

## Required Output

Produce a single markdown file named `PROJECT_DOCUMENTATION.md` with the sections below, in this order. Do not skip a section for being "obvious" — a reader with zero context should be able to go from "I've never seen this project" to "I can operate and modify it" using only this document.

---

### 1. System Overview

- One paragraph: what this project does, in plain language (it's an automated lead-generation and outreach pipeline running on Google Sheets + Apps Script + Gmail, scoring leads with Gemini, validating emails with ZeroBounce, and managing a full send/follow-up lifecycle).
- A simple end-to-end diagram in text/ASCII or a numbered flow showing how a lead moves from raw import → scored → validated → queued → sent → classified → followed up.
- List every external service this depends on (Google Sheets, Gmail, Gemini API, ZeroBounce API) and what each one is responsible for.

### 2. Full File Inventory — What Every File Does

For **every single `.gs` file in the project** (not just the ones built in the last upgrade — the original files too), produce an entry with:
- **File name**
- **One-line role** (what this file is responsible for, in the system as a whole)
- **Every function in the file**, each with: what it does, what triggers it (manual run, time-based trigger, called by another function), what it reads from the sheet, what it writes back, and what external APIs (if any) it calls
- **Dependencies**: what other files' functions does this file call, and what calls into this file from elsewhere

Do this for every file — do not group multiple files into one entry. If a file has more than 10 functions, still document all of them; don't stop at "the important ones."

### 3. Full Sheet/Tab Inventory — What Every Tab and Column Does

For every tab in the spreadsheet (Leads, Config, Templates, Log, Ready to Send, Account Config, and any others that exist in the actual sheet, not just the ones mentioned in past specs):
- **Tab name and purpose**
- **Every column**, in order, with: column name, what it stores, valid values (if it's a dropdown/enum), which script writes to it, which script reads from it, and whether it's meant to ever be edited manually vs. system-only
- Note any column that's currently unused by any script (dead column) or referenced by a script but missing from the sheet (broken reference) — check this directly rather than assuming consistency

### 4. Every Feature, Explained With Its Full Logic ("Clauses")

For each of the following features (and any others found in the code that aren't listed here — add them), write a dedicated subsection with:
- **What it does** (plain language)
- **The exact rule/threshold/condition** it enforces — quote the actual constant names and values from `Config.gs` or wherever they live, not paraphrased numbers
- **Edge cases it handles** (and any it doesn't, if found during the read-through)
- **What happens on failure** (API timeout, missing data, etc.)

Cover at minimum:
- Lead ingestion and AI scoring (Gemini)
- Email validation (ZeroBounce) and send-priority assignment
- The catch-all cap (state clearly: current cap value, and whether it's per-account or global, as actually implemented — not as originally speced)
- Daily quota system and the bootstrap-quota hard-stop
- The daily send window / scheduling logic (prep window vs. send window timing)
- Spillover of skipped leads to the next day
- Response classification (replied / bounced / out-of-office) including exactly what signals each detector looks for
- Bounce-rate tracking (state the actual window used, e.g. rolling 7-day, as implemented)
- The follow-up state machine (3-day and 10-day rules, OOO carve-out, threading behavior)
- Any additional features present in the code not covered above

### 5. Configuration Reference

A single table of **every configurable value in the system** — every constant in `Config.gs`, every value in the Config sheet tab, every per-account setting in Account Config:

| Setting | Current Value | What it controls | Where it lives | Confirmed by manager? |
|---|---|---|---|---|

Mark clearly (using whatever markers exist in code comments, e.g. `TODO(confirm-with-manager)`) which values are provisional defaults versus confirmed business decisions.

### 6. Setup Guide (First-Time Install)

Step-by-step instructions to get this running from scratch on a brand-new Google account, written for someone who has never opened this project before:
1. Copying/creating the spreadsheet and running initial sheet setup
2. Where to get and enter the Gemini and ZeroBounce API keys (Script Properties)
3. Setting up Account Config rows for each sending mailbox, including explaining the `Quota Confirmed` checkbox and why it starts unticked
4. Setting `DRY_RUN` / `Outreach Mode` correctly for safe first-time testing
5. Setting up the time-based triggers (which functions need triggers, at what times, and how to add them in the Apps Script trigger UI)
6. What "you're ready to go live" looks like — the minimum checklist before flipping to real sending

### 7. Day-to-Day Usage Guide

Written for a non-engineer who will operate this day to day:
- How to add new leads (where to paste them, what format is expected)
- How to check whether today's batch ran correctly (where to look — Execution Log, which sheet columns to check)
- How to read lead status at a glance (what each `Pipeline Stage`, `Response Status`, `Send Priority` value means for that lead)
- How to manually override something (e.g., force a lead to be excluded from sending, manually mark something as replied)
- What to do if something looks wrong (a practical troubleshooting flow, not just "check the logs")

### 8. Known Limitations & Open Questions

- List every `TODO(confirm-with-manager)` or similar flagged-but-unresolved item found in the code, in plain language, with what decision is pending.
- List anything found during this documentation pass that looks like a bug, inconsistency, or gap between what the spec intended and what the code does — do not fix these, just document them clearly so they can be triaged separately.

---

## Formatting Rules

- Use clear headers matching the numbered sections above.
- Every code-level claim (a threshold, a formula, a condition) must be traceable to an actual file/line — write it as prose but make sure it's verifiably accurate against the real code, not remembered from an old spec.
- Do not pad this document with generic boilerplate about "the importance of email marketing" or similar filler — every sentence should be specific to this actual codebase.
- Where you are uncertain whether something is intentional behavior or a bug, say so explicitly rather than guessing silently.
- This document should be usable on its own — someone should not need to read the original build spec, the audit, or any past chat to fully understand and operate the system after reading this file.
