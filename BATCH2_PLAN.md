# Batch 2 — Implementation Plan (8 Features)

**Principle:** Additive only. Do not remove working logic (hourly batch trigger, DRY_RUN, JSON stripping, thread-based reply detection, per-row account override). New behavior is gated so existing flows keep working.

**Execution order:** Build **Features 1, 2, 4 first → CHECKPOINT** for review → then 3, 6, 5, 8, 7.

---

## New shared infrastructure (built once, in Setup.gs)

- **Canonical field dictionary** `CANONICAL_FIELDS` — map of canonical key → array of accepted header aliases (email, company, first_name, last_name, industry, revenue, funding_date, linkedin_url, title, employee_count, source).
- **New columns added to `CUSTOM_COLUMNS`** (auto-appended to existing sheets by existing missing-column logic):
  - `Raw Name Backup` (hidden) — Feature 2
  - `Score Source` (hidden) — Feature 3
  - `Enrichment Status` — Feature 7
- **New hidden config sheet `_FieldMap`** — stores resolved column mapping per sheet so it's not recomputed per row (Feature 1).
- **New config keys** (DEFAULT_CONFIG): recruitment keyword list, account band pools (`Score Band High Accounts` = "Account A,Account B", `Score Band Mid Accounts` = "Account C,Account D,Account E"), `Draft Quality Threshold` = 7, `Sender Source` = "MainSheet" (flag for Feature 8), Account C/D/E Email/Label/Signature placeholders.

---

## Feature 1 — Generic source-agnostic ingestion  [BUILD FIRST]
- New `Mapping.gs` file: `resolveColumnMapping(sheet)` scans header row, fuzzy-matches (lowercase, strip non-alnum, alias table) against `CANONICAL_FIELDS`, returns `{canonicalKey: colIndex}`. Cache result in `_FieldMap` sheet keyed by sheet name.
- `getCanonical(row, mapping, key)` accessor — returns "" if unmapped. Downstream prompt builders (scoring, outreach) use this so **only non-empty canonical fields are included** in prompts — no empty placeholders.
- If `email` cannot be mapped for a row → set `Pipeline Stage = "Missing: email"`, skip that row, don't throw.
- Wire into ingestion: existing `mapApolloContactToRow` stays; the generic layer is the fallback for manual/other imports.

## Feature 2 — AI name splitting in setup node  [BUILD FIRST]
- New `splitNameWithAI(fullName, config)` in `Ingestion.gs` — LLM prompt, returns `{first, last}`, handles single names / initials / South Indian multi-part names. Non-Gemini failover reuse.
- Trigger: during ingest, per new row, **only if** `last_name` empty AND `first_name` has >1 token.
- Before overwrite: copy original into hidden `Raw Name Backup`. Then write cleaned first/last back.
- Skip entirely if last name already populated (no credit spend). Replaces the naive `.split(" ")`.

## Feature 4 — Recruitment-industry hard filter  [BUILD FIRST]
- New `isRecruitmentIndustry(industryText, config)` — keyword match against config list (recruitment, staffing, talent acquisition, executive search, hr consulting, headhunting).
- Insert as **pre-scoring gate** in `runFullPipelineMenu` (Main.gs) and the scoring-only menu: if matched → `Pipeline Stage = "Disqualified — Recruitment Industry"`, skip scoring/validation/drafting/sending.
- Applies regardless of source (F1) and regardless of manual score (F3) — even a manual "9" is blocked; reason shown in row.

--- CHECKPOINT: review 1/2/4 before continuing ---

## Feature 3 — Manual score override
- Remove dropdown lock on `Score` (change `applyLeadDropdownValidations` to skip Score / apply numeric 1–10 validation with allowInvalid).
- In pipeline: read `Score` before AI. If blank → AI as normal. If numeric 1–10 → skip AI, set `Score Source = "Manual"`. If non-numeric non-blank → treat as invalid, fall back to AI (no crash).
- Downstream categorization treats manual score identically. Recruitment filter (F4) still overrides.

## Feature 6 — Conditional ZeroBounce
- In `validateEmailQuality`: before second-pass, if Apollo `Email Status === "verified"` → set `Validation Status = "Verified (Apollo)"`, skip ZeroBounce (already ~half-present at line 84; formalize the status string + skip).
- Else call ZeroBounce (existing path) → `Verified (ZeroBounce)` / `Invalid (ZeroBounce)`. This status gates sending.

## Feature 5 — Score-band account routing
- New `selectAccountByScore(scoreNum, config)` — score ≥ 8 → random of High pool; 6 ≤ score < 8 → random of Mid pool; < 6 → default/hold. Pools from config (scaffolded C/D/E).
- Insert as tier 2 in precedence: `forcedAccount || rowOverride || **scoreBand** || default`.
- Respect per-account send caps: if chosen account at cap, pick next available in same band. (Cap tracking: add per-account daily counter in config/properties.)
- ⚠️ Band overlap (8,9) flagged to manager — using highest-first mutually-exclusive bands pending confirmation.

## Feature 8 — "Ready to Send" tab + secondary quality score
- New tab `Ready to Send` created in Setup: `Recipient Name | Company | Email | Subject | Body | Sender Account | Lead Score | Draft Quality Score | Ready for Send`.
- On draft generation, append/update a row in this tab.
- `scoreDraftQuality(subject, body, context, config)` — **non-Gemini** failover call, input is only tab content, returns 1–10.
- `Ready for Send` checkbox: default TRUE if quality > threshold, else FALSE. User can toggle.
- Sender: add `Sender Source` config flag. When = "ReadyTab", hourly sender pulls rows where `Ready for Send = TRUE`, respects existing 10/hr cap + assigned account (no re-selection). When = "MainSheet", unchanged. Default MainSheet (safe rollout).
- After send: mark main sheet `Pipeline Stage = Sent` + lock the tab row (reuse existing protection pattern).

## Feature 7 — Apollo enrichment (manual trigger)
- New `enrichSelectedRows()` menu item — for selected rows where email blank but name+company/domain present.
- Batch ≤10 → Apollo `/v1/people/bulk_match` (endpoint helper already exists at Ingestion.gs:164).
- On match: write email/phone back, tag `Source = "Apollo Enrichment"`. On no match: `Enrichment Status = "No Match"`, don't block.
- Manual-only (credit control).

---

## Files touched
- `Setup.gs` — columns, config keys, `_FieldMap` + `Ready to Send` tabs, Score dropdown change
- `Mapping.gs` (new) — Feature 1 mapping layer
- `Ingestion.gs` — generic mapping use, AI name split, enrichment
- `Main.gs` — recruitment gate, manual-score branch, menu items
- `Scoring.gs` — manual-score skip, canonical-field prompt building
- `Validation.gs` — Apollo-verified skip
- `Outreach.gs` — score-band routing, Ready-to-Send tab population, sender source flag, quality score

## Verification (per feature, DRY_RUN on)
Manual menu run against Sample Data on a handful of rows; confirm sheet state transitions match spec. No live sends until you flip DRY_RUN/Send mode.
