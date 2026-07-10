# AI Lead Generation Engine — Complete Project Reference

> **Platform:** Google Sheets + Google Apps Script  
> **Owner:** Butter Search  
> **Status:** Production-Ready (DRY_RUN=true by default)

---

## Project Overview

An automated B2B outreach engine built entirely in Google Apps Script on top of Google Sheets. It ingests leads from multiple sources, scores them with AI, validates emails, routes them to the correct sender account by score, generates personalized cold emails, and manages follow-ups — all without leaving the spreadsheet.

**Daily target:** 200 leads scored + validated + emailed across 5 accounts.

---

## File Architecture

| File | Purpose |
|------|---------|
| `Setup.gs` | Sheet creation, column definitions, default config, dropdown setup |
| `Main.gs` | Scoring pipeline entry points, recruitment filter, name-split guard |
| `Mapping.gs` | Generic source-agnostic column mapping layer (canonical field resolver) |
| `Ingestion.gs` | Apollo, LinkedIn X-Ray, GitHub, Google Maps ingestion + AI name splitting + Manual Apollo enrichment |
| `Scoring.gs` | AI scoring via Gemini / Groq / Cerebras with key rotation and failover |
| `Validation.gs` | Email quality validation (ZeroBounce bypass for Apollo-verified, company validation) |
| `Outreach.gs` | Email generation, score-band routing, per-account cap enforcement, Ready to Send tab, follow-up engine |
| `Hiring.gs` | Hiring signal detection via Google Jobs / Careers pages |
| `Tracking.gs` | Utilities for reply tracking |

---

## Sheet Structure

### Sheets Created by "Setup Sheets"

| Sheet | Purpose |
|-------|---------|
| `Leads` | Primary data table — all leads, scores, statuses |
| `Config` | All tunable parameters |
| `Templates` | Named email templates with Subject + Body columns |
| `Log` | Per-run audit trail (timestamp, rows processed, errors) |
| `Ready to Send` | Staged drafts with quality scores and send checkbox |
| `_FieldMap` | Hidden — caches canonical column mapping per sheet name |

### Leads Sheet — Custom Tracking Columns

`Score`, `Score Reason`, `Score Source`, `Validation Status`, `Validation Reason`, `Outreach Status`, `Hiring Status`, `Company Validation Status`, `Company Validation Reason`, `Research Status`, `Pipeline Stage`, `Last Sent At`, `Replied`, `Follow-up Status`, `Thread Id`, `Send From Account`, `Sent From Account`, `Raw Name Backup`, `Enrichment Status`

### Ready to Send Tab Columns
`Draft ID` | `Timestamp` | `Company` | `First Name` | `Last Name` | `Email` | `Selected Account` | `Lead Score` | `Draft Quality Score` | `Subject` | `Body` | `Ready for Send` (checkbox)

---

## Config Parameters

| Parameter | Default | Notes |
|-----------|---------|-------|
| Revenue Cutoff (INR) | 50,000,000 | ₹5 Crore minimum for "Qualified" |
| Funding Recency Window (Months) | 12 | Recent funding signal window |
| Email Confidence Threshold | 0.7 | Below this → ZeroBounce triggered |
| Gemini Model | deepseek-ai/deepseek-r1 | Any Gemini/Groq/Cerebras/DeepSeek model |
| Gemini Temperature | 0.2 | Lower = more deterministic |
| Outreach Mode | Draft | `Draft` = Gmail drafts; `Send` = real send |
| Emails Per Run | 10 | Per hourly batch cap |
| Follow-up Delay (Days) | 3 | Days before auto follow-up |
| Account A/B/C/D/E Email + Label + Signature | — | Sender identities |
| Default Send Account | Account A | Fallback routing |
| Recruitment Industry Keywords | recruitment, staffing, talent acquisition... | Disqualification keywords |
| Score Band High Accounts | Account A, Account B | Pool for leads scoring ≥ 8 |
| Score Band Mid Accounts | Account C, Account D, Account E | Pool for leads scoring 6–7 |
| Score Band Low Behavior | Default | `Hold` or `Default` |
| Per Account Daily Cap | 40 | Max sends/account/day before rerouting |
| Draft Quality Threshold | 7 | Auto-check Ready for Send if quality ≥ this |
| Sender Source | MainSheet | `MainSheet` or `ReadyTab` |

---

## Script Properties (API Keys)

Set in Apps Script → Project Settings → Script Properties.

| Key | Required | Purpose |
|-----|----------|---------|
| `GEMINI_API_KEY` | ✅ Yes | AI scoring, name splitting, draft quality scoring |
| `APOLLO_API_KEY` | Recommended | Lead ingestion + manual enrichment |
| `ZEROBOUNCE_API_KEY` | Optional | Email verification |
| `HUNTER_API_KEY` | Optional | Fallback email finder/verifier |
| `GROQ_API_KEY` | Optional | Alternative free scoring provider (14,400 req/day) |
| `CEREBRAS_API_KEY` | Optional | Alternative free scoring provider |

> Auto-managed daily counter keys: `account_sent_YYYY-MM-DD_AccountName`

---

## Menu Actions

| Menu Item | Description |
|-----------|-------------|
| Setup Sheets | Creates all tabs, headers, dropdowns, default config |
| Apollo Ingest | Searches Apollo API, appends new leads |
| LinkedIn X-Ray Ingest | Google CSE search → parses LinkedIn results |
| GitHub Ingest | GitHub user search → developer leads |
| Google Maps Ingest | Places API → local business leads |
| Manual Apollo Enrichment (Selected) | Batch Apollo bulk_match on selected rows (10/call) |
| Score new leads | AI scores all unscored rows (80/run cap) |
| Process Leads (End-to-End) | Score + validate in one pass |
| Generate outreach drafts (All Ready) | Drafts for all "Ready for outreach" rows |
| Generate outreach drafts (Selected) | Drafts for currently selected rows |
| Preview outreach samples | Shows 5 sample emails in popup (no draft created) |
| Set Account A / B for selected | Sets Send From Account on selected rows |
| Scan replies & send follow-ups | Checks Gmail threads for replies, sends follow-ups |

---

## All Implemented Features

### ── Batch 1 ──

#### B1-1 — Multi-Account Outreach (Account A / Account B)
- Two Gmail identities with separate email, label, and signature.
- Force an account via popup (Yes = A, No = B, Cancel = auto).
- Per-row override via `Send From Account` column.

#### B1-2 — AI Lead Scoring (`Scoring.gs`)
- Scores 1–10 with `Score Reason` explanation.
- Supports Gemini, Groq, Cerebras, DeepSeek.
- Key rotation across comma-separated API keys; provider-level failover.
- `Score Source = "AI"` stamped after scoring.

#### B1-3 — Email Quality Validation (`Validation.gs`)
- ZeroBounce / Hunter.io second-pass verification.
- Results: `Validation Status` (Ready/Risky/Flagged) + `Validation Reason`.

#### B1-4 — Company Validation
- Pre-scoring check on company name, website, and industry.
- `Company Validation Status` gates row from proceeding if invalid.

#### B1-5 — Hiring Signal Detection (`Hiring.gs`)
- Scans company careers pages/Google Jobs for active tech hiring.
- Writes `Hiring Status` for use in scoring prompt.

#### B1-6 — Reply Detection & Follow-Up
- Daily trigger scans Gmail by `Thread Id`.
- `Replied = Yes` stops the follow-up sequence.
- Auto follow-up after configurable delay using Templates tab.
- `Follow-up Status`: Awaiting Reply → Follow-up Sent → Replied.

#### B1-7 — Pipeline Stage Tracking
- `Pipeline Stage` updated at every step:  
  `Company Verified` → `Scored` → `Ready for outreach` → `Draft Created` → `Sent` → `Replied`

#### B1-8 — DRY_RUN Safety Mode
- `DRY_RUN = true` at top of `Outreach.gs` (line 8).
- All sends become Gmail drafts prefixed with `[TEST DRAFT]`.
- Must be manually set to `false` for live sending.

---

### ── Batch 2 ──

#### B2-F1 — Generic Source-Agnostic Ingestion (`Mapping.gs`)
- Canonical field dictionary with fuzzy alias matching.
- Non-Apollo headers (`Organization`, `Sector`, `Given Name`) resolve automatically.
- Mapping computed **once per run**, cached in `_FieldMap` (not per-row).
- AI prompts include **only populated fields** — empty lines stripped before API call.
- `flagIfMissingEmail` writes `"Missing: email"` to Pipeline Stage and skips row before any API call.
- `getFieldValue` tries literal header first, then canonical fallback — Apollo sheets unchanged.

> **`Keywords` is NOT aliased as `industry`** — Apollo product tags (e.g. "it staffing") caused false disqualifications on IT companies. `Keywords` has its own `"keywords"` canonical key for scoring use only.

#### B2-F2 — AI Name Splitting
- Triggers at ingestion (Apollo, LinkedIn, GitHub, Google Maps) AND at processing time.
- If `First Name` = multi-word + `Last Name` = blank → calls Gemini with South Indian name handling.
- Writes back to both `First Name` and `Last Name` columns on the sheet.
- Saves original full name to `Raw Name Backup` before overwriting.
- AI failure → graceful fallback to naive space-split.
- **Guard:** if Last Name already filled → AI call skipped entirely (no cost on re-runs).

#### B2-F3 — Manual Score Override
- Enter 1–10 in `Score` column → AI scoring skipped for that row.
- Stamps `Score Source = "Manual"`, `Score Reason = "Manual Override"`.
- Row still proceeds through validation and outreach normally.
- Invalid scores (text, out-of-range, "Error") → falls back to AI scoring.

#### B2-F4 — Recruitment Industry Hard Filter
- Runs **before** any AI call in the row loop.
- Checks `Industry`/`Sector` column only (not Keywords column).
- Substring match against `Recruitment Industry Keywords` config (case-insensitive).
- On match: `Pipeline Stage = "Disqualified — Recruitment Industry"`, `Outreach Status = "Disqualified"`.
- High manual score does NOT override — filter runs first.

#### B2-F5 — Score-Band Account Routing with Daily Cap
- Priority chain: Forced (popup) → Per-row override → Score band → Default account.
- Bands: `≥ 8` = High pool, `6–7` = Mid pool, `< 6` = Low/Hold.
- Random selection within a band's available accounts.
- **Per-account daily cap:** Script Properties key `account_sent_YYYY-MM-DD_AccountName` tracks sends. Capped accounts skipped; system tries next in band.
- If all High band accounts capped → falls to Mid band.
- If all accounts capped → returns `HOLD` rather than exceeding cap.
- `recordAccountSend` wired into both `processSingleOutreach` and `processReadyTabHourly`.
- Cap configurable: `Per Account Daily Cap` in Config (default: 40/account/day).

> ⚠️ **Manager sign-off pending** on band boundaries (≥8 High, 6–7 Mid, <6 Low).

#### B2-F6 — Conditional ZeroBounce Bypass
- `Email Status = "verified"` (Apollo) → immediately `Validation Status = "Ready"`. No ZeroBounce call made.
- All other statuses → ZeroBounce or Hunter.io called.
- ZeroBounce failure/timeout → `"Risky"` (never silently passes as verified).
- Validation Status gates the send step downstream.

#### B2-F7 — Manual Apollo Enrichment (Batched)
- User selects rows → Apollo `/api/v1/people/bulk_match` called in **batches of 10**.
- 1.5 second pause between batches to avoid rate limiting.
- On match: writes `Email`, `Email Status`, `Mobile Phone`, sets `Enrichment Status = "Apollo Enrichment"`.
- No match → `"No Match"`. API error → `"API Error — Retry"`.
- Batch errors don't stop remaining batches.
- Max 50 rows per selection. **Never runs automatically** — menu-only.

#### B2-F8 — "Ready to Send" Drafts Tab
- Drafts route to Ready tab instead of Gmail (when in Draft mode).
- **Draft Quality Score**: separate Gemini call on raw email body (1–10).
- `Ready for Send` auto-set TRUE if quality ≥ threshold (default 7).
- User can manually flip checkbox — sender reads live sheet value.
- Hourly sender (`processReadyTabHourly`): reads `Ready for Send = TRUE` rows, sends up to `Emails Per Run` per hour.
- After send: Leads sheet updated (`Pipeline Stage = "Sent"`, `Last Sent At`, `Outreach Status = "Email Sent"`). Ready tab row deleted.
- DRY_RUN: creates `[TEST DRAFT]` Gmail draft even from Ready tab.

---

## Pipeline Execution Flow

```
[1] INGESTION
    Apollo API / LinkedIn X-Ray / GitHub / Google Maps
        → AI name split (splitNameWithAI) if Last Name blank
        → Raw Name Backup saved
        → Row appended to Leads sheet

[2] SCORE NEW LEADS / PROCESS LEADS (End-to-End)
    For each row (skip if Pipeline Stage contains "Disqualified"):
        → Recruitment filter  →  Disqualify if keyword match (no AI call)
        → ensureNameSplit     →  Split multi-word First Name if Last Name blank
        → flagIfMissingEmail  →  Mark "Missing: email" and skip
        → Company validation  →  Skip unverified companies
        → Manual score check  →  If 1–10 exists, stamp Manual and skip AI
        → AI scoring          →  Gemini/Groq/Cerebras → Score + Reason
        → Email validation    →  Apollo verified → Ready; else ZeroBounce
        → Outreach Status     →  "Ready for outreach" if all checks pass

[3] GENERATE DRAFTS
    For each "Ready for outreach" row:
        → Account selection (priority: popup → row override → score band → default)
        → Per-account cap check → skip capped, try next in band
        → generatePersonalizedEmail → Gemini prompt, account-specific template
        → Flag if body contains hard numbers (qualitative rule)
        → scoreDraftQuality → separate Gemini call (1–10)
        → Append to Ready to Send tab
        → Pipeline Stage → "Draft Created"

[4] SEND (processReadyTabHourly)
    Ready to Send tab, rows where Ready for Send = TRUE:
        → Send via GmailApp (or [TEST DRAFT] in DRY_RUN)
        → recordAccountSend → increment daily counter
        → Delete row from Ready tab
        → Update Leads: Pipeline Stage = "Sent", Last Sent At, Outreach Status = "Email Sent"

[5] FOLLOW-UP (runFollowUpDaily — daily trigger)
        → Scan Gmail threads by Thread Id
        → Reply found → Replied = Yes, stop sequence
        → No reply after delay → send follow-up from Templates tab
```

---

## API Provider Reference & Capacity

| Provider | Free Limit | Use In System |
|----------|-----------|---------------|
| Gemini Flash | 1,500 req/day | Scoring, name split, draft quality |
| Groq (Llama 3) | 14,400 req/day, 30 RPM | Alternative scorer — best free option |
| Cerebras | ~14,400 req/day | Alternative scorer |
| DeepSeek (NVIDIA NIM) | 1,000 credits one-time | Alternative scorer |
| Apollo.io | Plan-dependent | Ingestion + bulk_match |
| ZeroBounce | 100/month free | Email verification |
| Hunter.io | 25/month free | Email find + verify |
| Google CSE | 100 queries/day free | LinkedIn X-Ray |
| GitHub API | 5,000/hr (with key) | Developer leads |

> **For 200 leads/day:** Use Groq as primary model. Add 2+ Groq keys comma-separated in Script Properties to load-balance (doubles RPM to 60 and daily limit to 28,800).

---

## Production Go-Live Checklist

- [ ] Set `DRY_RUN = false` in `Outreach.gs` line 8
- [ ] Fill Account A / B (and C/D/E) email + label + signature in Config
- [ ] Set all API keys in Script Properties
- [ ] Set `Outreach Mode = Send` in Config (or keep Draft for review)
- [ ] Confirm `Score Band High/Mid Accounts` match your real account names in Config
- [ ] Get **manager sign-off** on score band thresholds
- [ ] Set `Sender Source = ReadyTab` to enable quality-gated workflow
- [ ] Create time-based trigger: `processReadyTabHourly` → every 1 hour
- [ ] Create time-based trigger: `runFollowUpDaily` → once daily
- [ ] Run `Setup Sheets` once on fresh spreadsheet
- [ ] Do one test run with `DRY_RUN = true` and verify Ready to Send tab populates
- [ ] Monitor `Log` sheet after first real run for errors

---

## Known Limitations / Open Items

| Item | Status | Notes |
|------|--------|-------|
| Score band boundaries | ⚠️ Pending manager sign-off | ≥8 High, 6–7 Mid, <6 Low |
| ZeroBounce free tier | ⚠️ Insufficient for 200/day | 100/month free — needs paid plan |
| Apollo bulk_match credits | ⚠️ Monitor usage | Each match costs 1 Apollo credit |
| Dashboard tab | ❌ Not built | Score Source column data is present; no visual dashboard |
| Ready to Send row locking | ❌ Not built | Rows are deleted after send (same functional outcome) |
| Google Maps ingestion | ⚠️ Requires paid API key | Google Places API is pay-per-use |
