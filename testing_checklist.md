# AI Lead Generation Engine — Feature Testing Checklist

> **Before starting:** Open your Google Sheet → Apps Script editor → confirm `DRY_RUN = true` at the top of `Outreach.gs`.  
> **Golden rule:** Never test with `DRY_RUN = false` until you're satisfied with every check below.

---

## PRE-FLIGHT: One-Time Setup

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Open the spreadsheet → **Lead Engine → Setup Sheets** | All tabs created: Leads, Config, Templates, Log, Ready to Send, _FieldMap (hidden). No error popup. |
| 2 | Open Apps Script → Project Settings → Script Properties. Confirm these keys exist: `GEMINI_API_KEY`, `APOLLO_API_KEY` (optional), `ZEROBOUNCE_API_KEY` (optional) | Keys are present. |
| 3 | Open the **Config** tab. Confirm these rows exist with values: `Recruitment Industry Keywords`, `Score Band High Accounts`, `Score Band Mid Accounts`, `Draft Quality Threshold` (7), `Sender Source` (MainSheet) | All rows visible with default values. |
| 4 | Open the **Templates** tab. Confirm rows for "Style Reference" (initial email) and "Follow-Up" exist with Subject and Body columns filled. | Templates visible. |

---

## SECTION 1 — Setup & Column Structure (Batch 1 Foundation)

### Test 1.1 — Sheets are correctly structured

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Click **Lead Engine → Setup Sheets** on a fresh sheet | Leads tab created with all required headers in row 1. |
| 2 | Scroll right in Leads tab | Confirm these columns exist: `Score`, `Score Reason`, `Score Source`, `Validation Status`, `Validation Reason`, `Outreach Status`, `Pipeline Stage`, `Send From Account`, `Sent From Account`, `Last Sent At`, `Thread Id`, `Raw Name Backup`, `Enrichment Status` |
| 3 | Click **Lead Engine → Apply Column Dropdowns** | Pipeline Stage, Outreach Status, etc. now show dropdown options in each cell. |
| 4 | Click **Lead Engine → Highlight Required Columns** | Key columns are highlighted/coloured in the header row. |

---

## SECTION 2 — Generic Source-Agnostic Ingestion (Feature 1)

### Test 2.1 — Non-Apollo headers are resolved correctly

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Manually paste a test row into the Leads sheet using these **non-Apollo headers** by temporarily renaming 3 columns: rename `Company` → `Organization`, rename `Industry` → `Sector`, rename `First Name` → `Given Name` | Columns renamed. |
| 2 | Add a test row under those renamed headers with data: Organization=`Acme Corp`, Sector=`Fintech`, Given Name=`Rahul`, Last Name=`Sharma`, Email=`rahul@acmecorp.com`, Score=`9` | Row added. |
| 3 | Run **Lead Engine → Score new leads** | No crash. Row is processed. Check Apps Script Logs → confirm `resolveColumnMapping` ran (look for `_FieldMap` sheet appearing in the spreadsheet). |
| 4 | Open `_FieldMap` sheet (it's hidden — click the sheet tab arrows to find it) | One row: `Leads | {"email":..., "company":..., "industry":..., "first_name":...}` — company, industry, first_name all mapped to the correct column numbers. |
| 5 | Restore column names back to Apollo defaults (`Company`, `Industry`, `First Name`) | Columns renamed back. |

### Test 2.2 — Missing email is caught before scoring

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add a row with all fields populated but **leave Email blank**. Score=`9`. | Row added. |
| 2 | Run **Lead Engine → Score new leads** | That row's `Pipeline Stage` = `"Missing: email"`. It is NOT scored. No error in the batch. |

---

## SECTION 3 — AI Name Splitting (Feature 2)

### Test 3.1 — Apollo ingestion splits complex names

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Manually add a test row: `First Name` = `Karthikeyan Subramaniam`, `Last Name` = *(blank)* | Row added. |
| 2 | Run **Lead Engine → Process Leads (End-to-End)** | Check the row: `First Name` = `Karthikeyan`, `Last Name` = `Subramaniam`, `Raw Name Backup` = `Karthikeyan Subramaniam`. |

### Test 3.2 — Three South Indian name patterns

| # | First Name value (Last Name blank) | Expected First Name | Expected Last Name |
|---|------------------------------------|--------------------|--------------------|
| A | `R. Venkataraman` | `Venkataraman` | `R` (or similar) |
| B | `S Ramachandran` | `Ramachandran` | `S` |
| C | `Thiruvenkatam Pillai` | `Thiruvenkatam` | `Pillai` |

For each: add the row, run pipeline, check both name columns and `Raw Name Backup`.

### Test 3.3 — Last name already filled — AI skipped

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row: `First Name` = `Rajesh`, `Last Name` = `Kumar` | Row added. |
| 2 | Run pipeline | `Raw Name Backup` = *(empty)*. Name columns unchanged. No extra AI call for this row. |

### Test 3.4 — LinkedIn & GitHub ingestion also splits names

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Run **Lead Engine → Ingest Leads → LinkedIn X-Ray Ingest** (or GitHub) | After ingestion, check newly added rows. Names with spaces should have both First Name and Last Name populated, and `Raw Name Backup` filled where the original was a full name. |

---

## SECTION 4 — Manual Score Override (Feature 3)

### Test 4.1 — Manual score skips AI

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add a new row with Company, Email, and all required fields. Set `Score` = `9`. Leave `Score Reason` blank. | Row added. |
| 2 | Run **Lead Engine → Score new leads** | `Score` remains `9`. `Score Reason` = `"Manual Override"`. `Score Source` = `"Manual"`. `Pipeline Stage` = `"Scored"`. No Gemini/Groq API call made for this row. |

### Test 4.2 — Manual score flows to categorization correctly

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Run **Lead Engine → Process Leads (End-to-End)** on the row from Test 4.1 (score = 9) | Row is validated (ZeroBounce or Apollo-verified check runs). If email is valid: `Outreach Status` = `"Ready for outreach"`. |

### Test 4.3 — Blank score triggers AI normally

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add a row with Score left blank | Row added. |
| 2 | Run **Score new leads** | AI scores the row. `Score Source` = `"AI"`. |

### Test 4.4 — Invalid score falls back to AI

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add a row. Set `Score` = `"high"` (text) | Row added. |
| 2 | Run **Score new leads** | AI scoring triggers for this row (because `"high"` is not a valid 1–10 number). `Score Source` = `"AI"` after completion. |

---

## SECTION 5 — Recruitment Industry Filter (Feature 4)

### Test 5.1 — Exact keyword match disqualifies before AI

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add a row: `Company` = `TalentBridge`, `Industry` = `Staffing and Recruitment`, all other fields filled | Row added. |
| 2 | Run **Lead Engine → Score new leads** | `Pipeline Stage` = `"Disqualified — Recruitment Industry"`. `Outreach Status` = `"Disqualified"`. No AI scoring call made for this row. |

### Test 5.2 — High manual score doesn't override the filter

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row with `Industry` = `Executive Search`, `Score` = `9` | Row added. |
| 2 | Run pipeline | Row is still disqualified by recruitment filter. Never reaches outreach. |

### Test 5.3 — No false positive on unrelated fields

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row: `Title` = `Staff Software Engineer`, `Industry` = `Fintech` | Row added. |
| 2 | Run pipeline | Row is NOT disqualified. The word "staff" in the job title doesn't trigger the filter (filter only checks the industry/keyword fields). |

### Test 5.4 — IT Staffing in company keywords IS caught

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row: `Industry` = `IT Staffing Solutions` | Row added. |
| 2 | Run pipeline | Row disqualified with `"Disqualified — Recruitment Industry"` (because "staffing" is a keyword substring). |

---

## SECTION 6 — ZeroBounce Bypass for Apollo-Verified (Feature 6)

### Test 6.1 — Verified by Apollo skips ZeroBounce

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row: `Email` = `ceo@verified.com`, `Email Status` = `verified` | Row added. |
| 2 | Run **Process Leads (End-to-End)** with score ≥ 8 | `Validation Status` = `"Ready"`, `Validation Reason` = `"Verified (Apollo)"`. ZeroBounce API NOT called (check Apps Script Logs — no ZeroBounce URL fetch logged). |

### Test 6.2 — Guessed email calls ZeroBounce

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row: `Email Status` = `guessed`, Score ≥ 8 | Row added. |
| 2 | Run pipeline | ZeroBounce call made (or if key missing, row flagged as `"Risky"`). `Validation Reason` reflects ZeroBounce result. |

### Test 6.3 — Validation gates the send step

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Set up a row with `Validation Status` = `"Flagged"` and `Outreach Status` = `"Flagged - email risk"` | Row visible. |
| 2 | Run **Lead Engine → Generate outreach drafts (All Ready)** | Row is NOT selected for drafting (validation status is not "Ready"). |

---

## SECTION 7 — Score-Based Account Routing (Feature 5)

### Test 7.1 — High score (≥ 8) routes to high-band account

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | In Config tab, set `Score Band High Accounts` = `Account A` | Saved. |
| 2 | Add a row with Score = `9`, Validation Status = `Ready`, Outreach Status = `Ready for outreach` | Row ready. |
| 3 | Run **Generate outreach drafts (All Ready)** → choose Cancel on the popup (use row settings) | After drafting: `Send From Account` column = `"Account A"`. `Pipeline Stage` = `"Draft Created"`. |

### Test 7.2 — Mid score (6–7) routes to mid-band account

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Set `Score Band Mid Accounts` = `Account B` in Config | Saved. |
| 2 | Add row with Score = `7`, Validation = `Ready`, Outreach = `Ready for outreach` | Row ready. |
| 3 | Generate drafts → Cancel on popup | `Send From Account` = `"Account B"`. |

### Test 7.3 — Low score with Hold behavior stays off outreach

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Set `Score Band Low Behavior` = `Hold` in Config | Saved. |
| 2 | Add row with Score = `3`, Validation = `Ready`, Outreach = `Ready for outreach` | Row ready. |
| 3 | Generate drafts | Row's `Outreach Status` = `"On Hold (Low Score)"`, `Pipeline Stage` = `"Held"`. No draft created for it. |

### Test 7.4 — Per-row manual override beats score routing

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Set `Send From Account` = `Account B` manually on a row with Score = `9` (which would route to Account A) | Value set. |
| 2 | Generate drafts → Cancel on popup | `Send From Account` remains `"Account B"` — per-row override wins. |

---

## SECTION 8 — Manual Apollo Enrichment (Feature 7)

### Test 8.1 — Missing email row gets enriched

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add a row: `First Name` = `Priya`, `Last Name` = `Nair`, `Company` = `Zerodha`, `Website` = `zerodha.com`. Leave `Email` blank. | Row added. |
| 2 | Click the row to select it | Row highlighted. |
| 3 | **Lead Engine → Ingest Leads → Manual Apollo Enrichment (Selected)** | Popup shows "Enriching 1 contacts...". |
| 4 | After completion: check `Email`, `Email Status`, `Mobile Phone`, `Enrichment Status` columns | If Apollo found a match: `Enrichment Status` = `"Apollo Enrichment"`. If not found: `"No Match"`. Either way — no error crash. |

### Test 8.2 — Row with no domain is still handled gracefully

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Add row with only First Name and Last Name, no email, no website | Row added. |
| 2 | Select row → run enrichment | `Enrichment Status` = `"No Match"`. No crash. Alert shows correct counts. |

### Test 8.3 — Function is NOT triggered automatically

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Run **Score new leads** or **Process Leads (End-to-End)** on any row | Check Apps Script Logs — no `bulk_match` API call logged during automatic pipeline. Enrichment only fires from menu. |

---

## SECTION 9 — "Ready to Send" Tab (Feature 8)

### Test 9.1 — Draft routes to Ready to Send tab

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Confirm `Sender Source` in Config = `"MainSheet"` | Set if needed. |
| 2 | Ensure `DRY_RUN = true` in Outreach.gs | Confirmed. |
| 3 | Have one row with: Score ≥ 8, Validation Status = `Ready`, Outreach Status = `Ready for outreach` | Row ready. |
| 4 | Run **Lead Engine → Generate outreach drafts (All Ready)** | Alert confirms draft generated. |
| 5 | Open the **Ready to Send** tab | A new row is appended. Check all columns: Draft ID, Timestamp, Company, First Name, Last Name, Email, Selected Account, Lead Score, Draft Quality Score (1–10), Subject, Body, Ready for Send (checkbox). |
| 6 | Check the Body column | Should contain the actual email — NOT the `[TEST DRAFT]` prefixed version. |
| 7 | Check the Draft Quality Score | A number 1–10. If score ≥ 7 (threshold), the `Ready for Send` checkbox = TRUE. |

### Test 9.2 — Quality score < threshold defaults to unchecked

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Generate a draft for a row with minimal company context (thin body likely) | Row appears in Ready tab. |
| 2 | If Draft Quality Score < 7, check checkbox column | `Ready for Send` = FALSE (unchecked). |

### Test 9.3 — Manual flip of checkbox is honoured by sender

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | In Ready to Send tab, manually uncheck the checkbox on a row that was auto-checked TRUE | Checkbox = FALSE. |
| 2 | Set `Sender Source` = `"ReadyTab"` in Config | Saved. |
| 3 | Run **Lead Engine → Generate outreach drafts (All Ready)** (which triggers the hourly logic for manual testing) | That unchecked row is NOT sent. Other checked rows ARE processed. |

### Test 9.4 — After send, main sheet is updated

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | With `Sender Source` = `"ReadyTab"`, check a row in Ready tab = TRUE | Row ready to send. |
| 2 | Trigger the send by running `processReadyTabHourly` from the Apps Script editor directly | Ready tab row disappears (deleted after send). Go back to the **Leads** tab — find the corresponding lead row. `Pipeline Stage` = `"Sent"`, `Last Sent At` = today's timestamp, `Outreach Status` = `"Email Sent"`. |

---

## SECTION 10 — Multi-Account Routing (Batch 1 Feature)

### Test 10.1 — Account A / Account B assignment via menu

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Select 2-3 rows in the Leads tab | Rows highlighted. |
| 2 | **Lead Engine → Set Account A for selected** | `Send From Account` column = `"Account A"` for all selected rows. |
| 3 | **Lead Engine → Set Account B for selected** | `Send From Account` = `"Account B"` for selected rows. |

### Test 10.2 — Correct sender email used in draft

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Ensure Config has `Account A Email` = your email, `Account B Email` = second email | Set values. |
| 2 | Set one row to Account A, another to Account B. Generate drafts. | In Gmail Drafts (DRY_RUN mode): each draft shows the correct `From` address matching the account. |

---

## SECTION 11 — Follow-Up Workflow (Batch 1 Feature)

### Test 11.1 — Reply detection skips rows without a Thread ID

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Ensure some rows have `Thread Id` blank (unsent leads) | Visible. |
| 2 | Run **Lead Engine → Scan replies & send follow-ups** | Rows with no Thread Id are skipped. No error. |

### Test 11.2 — Follow-up preview shows correct template

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Open the Templates tab → find the row with name `"Follow-Up"` | Template visible with subject and body. |
| 2 | Select a sent row (has a Thread Id) | Row selected. |
| 3 | **Lead Engine → Preview follow-up (selected rows)** | Popup shows a draft follow-up rendered from the Follow-Up template with the lead's name and company substituted. Correct template — not the initial email style. |

---

## SECTION 12 — Full End-to-End Pipeline Trace

This is the capstone test. Run one lead through the complete pipeline manually.

| Step | Action | Expected Output |
|------|--------|-----------------|
| 1 | Add a fresh test row: `First Name` = `Arjun Subramaniam` (no last name), `Company` = `RazorpayX`, `Email` = `arjun@razorpayx.com`, `Email Status` = `verified`, `Industry` = `Fintech`, `Annual Revenue` = `$50M`, `Score` = *(blank)* | Row added. |
| 2 | Run **Score new leads** | Name splits: `First Name` = `Arjun`, `Last Name` = `Subramaniam`, `Raw Name Backup` = `Arjun Subramaniam`. Recruitment filter passes (Fintech ≠ recruitment). Score computed by AI. `Score Source` = `"AI"`. |
| 3 | Run **Process Leads (End-to-End)** | `Validation Status` = `"Ready"`, reason = `"Verified (Apollo)"` (no ZeroBounce call). `Outreach Status` = `"Ready for outreach"`. |
| 4 | Run **Generate outreach drafts (All Ready)** | Row appears in Ready to Send tab with Subject, Body, Draft Quality Score, and checkbox. |
| 5 | Check Body in Ready tab | No hard numbers (revenue figures, percentages) in email copy. Body matches the professional template style. |
| 6 | Check `Pipeline Stage` on Leads tab | `"Draft Created"`. |
| 7 | Manually check the checkbox in Ready tab if unchecked | Checkbox = TRUE. |
| 8 | Set `Sender Source` = `"ReadyTab"` in Config. Run `processReadyTabHourly` from Apps Script editor. | Row removed from Ready tab. On Leads tab: `Pipeline Stage` = `"Sent"`, `Last Sent At` = now, `Outreach Status` = `"Email Sent"`. |
| 9 | Check Gmail → Drafts | `[TEST DRAFT]` email visible (DRY_RUN mode). Correct From account. Correct recipient. |

---

## Quick Reference: What to Check After Each Menu Action

| Menu Action | Primary columns to inspect |
|-------------|--------------------------|
| Setup Sheets | All headers in row 1, Ready to Send tab exists |
| Score new leads | Score, Score Reason, Score Source, Pipeline Stage |
| Process Leads (End-to-End) | All of above + Validation Status, Validation Reason, Outreach Status |
| Generate outreach drafts | Ready to Send tab row + Pipeline Stage = "Draft Created" |
| Manual Apollo Enrichment | Email, Email Status, Mobile Phone, Enrichment Status |
| Scan replies & send follow-ups | Follow-up Status column on rows with Thread Id |
| processReadyTabHourly | Leads tab: Pipeline Stage = "Sent", Last Sent At; Ready tab: row deleted |

---

> **Tip:** After each section, reset your test rows (delete them or reset their Pipeline Stage) before running the next section, so tests don't interfere with each other.
