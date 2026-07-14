/**
 * AI Lead Generation Engine - Follow-up State Machine
 * File: FollowupEngine.gs
 *
 * Email Ops Upgrade. Drives the follow-up sequence off explicit due dates rather than a
 * single ad-hoc timer, and integrates with:
 *   - QuotaManager.gs        (getRemainingQuota / recordSend) so follow-ups respect the
 *                             per-account daily quota.
 *   - ResponseClassifier.gs  (Followup Cancelled / Response Status) so bounces stop the
 *                             sequence and out-of-office replies defer to a 10-day touch.
 *
 * Leads sheet columns used (added in Setup.gs):
 *   Send Date, Followup 1 Due Date, Followup 2 Due Date,
 *   Followup 1 Sent Date, Followup 2 Sent Date, Followup Cancelled, Response Status
 * plus existing: Email, First Name, Company, Pipeline Stage, Outreach Status, Replied,
 *   Follow-up Status, Thread Id, Sent From Account, Last Sent At.
 */

// First follow-up delay (days after the original send) — falls back to Config.
var FOLLOWUP_1_DELAY_DAYS_DEFAULT = 3;
// Out-of-office deferral (days). OOO cancels the 3-day touch and schedules this instead.
// Counted from the ORIGINAL send date (open question #4 — default assumption documented).
var FOLLOWUP_OOO_DELAY_DAYS = 10;

/**
 * Reads the configured 3-day follow-up delay, defaulting when unset.
 */
function followupDelayDays_(config) {
  var raw = config ? getConfigValue(config, "Follow-up Delay (Days)", FOLLOWUP_1_DELAY_DAYS_DEFAULT) : FOLLOWUP_1_DELAY_DAYS_DEFAULT;
  var n = parseInt(raw);
  return (isNaN(n) || n <= 0) ? FOLLOWUP_1_DELAY_DAYS_DEFAULT : n;
}

/**
 * Adds a whole number of days to a Date and returns a new Date.
 */
function addDays_(date, days) {
  var d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

/**
 * Coerces a cell value into a Date, or null if blank/unparseable.
 */
function toDateOrNull_(value) {
  if (value === "" || value === null || value === undefined) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") return value;
  var d = new Date(value.toString());
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Schedules the first follow-up for a lead by setting Followup 1 Due Date to
 * Send Date + delay (default 3 days). If Send Date is blank, it is stamped to now
 * so the schedule has an anchor. Intended to be called right after the initial send.
 *
 * @param {object} lead A lead object carrying at least `row` (1-indexed Leads row).
 *                       May also carry `sendDate`.
 * @return {object} { row, sendDate, followup1DueDate } describing what was scheduled,
 *                   or { error } on failure.
 */
function scheduleFollowups(lead) {
  if (!lead || !lead.row) return { error: "scheduleFollowups requires lead.row" };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) return { error: "Leads sheet not found" };
  var headersMap = getHeadersMap(leadsSheet);
  var config = safeGetConfig_();

  var row = lead.row;

  // Anchor: prefer an explicit Send Date, else Last Sent At, else now (stamped back).
  var sendDate = lead.sendDate ? toDateOrNull_(lead.sendDate) : null;
  if (!sendDate && headersMap["Send Date"]) {
    sendDate = toDateOrNull_(leadsSheet.getRange(row, headersMap["Send Date"]).getValue());
  }
  if (!sendDate && headersMap["Last Sent At"]) {
    sendDate = toDateOrNull_(leadsSheet.getRange(row, headersMap["Last Sent At"]).getValue());
  }
  if (!sendDate) {
    sendDate = new Date();
    if (headersMap["Send Date"]) {
      leadsSheet.getRange(row, headersMap["Send Date"]).setValue(sendDate);
    }
  }

  var delay = followupDelayDays_(config);
  var due1 = addDays_(sendDate, delay);

  if (headersMap["Followup 1 Due Date"]) {
    leadsSheet.getRange(row, headersMap["Followup 1 Due Date"]).setValue(due1);
  } else {
    Logger.log("scheduleFollowups: 'Followup 1 Due Date' column missing — run Setup Sheets.");
  }

  // Reset the follow-up tracking state for a fresh sequence.
  if (headersMap["Follow-up Status"]) {
    leadsSheet.getRange(row, headersMap["Follow-up Status"]).setValue("Pending");
  }
  if (headersMap["Followup Cancelled"]) {
    var current = leadsSheet.getRange(row, headersMap["Followup Cancelled"]).getValue();
    if (current === "" || current === null) {
      leadsSheet.getRange(row, headersMap["Followup Cancelled"]).setValue(false);
    }
  }

  Logger.log("scheduleFollowups: row " + row + " -> Followup 1 due " +
             Utilities.formatDate(due1, Session.getScriptTimeZone(), "yyyy-MM-dd") + ".");
  return { row: row, sendDate: sendDate, followup1DueDate: due1 };
}

/**
 * Scans the Leads sheet for pending follow-ups and processes each:
 *   - Skips rows that have replied or whose follow-ups are cancelled (e.g. bounced).
 *   - Handles out-of-office: cancels the 3-day touch and schedules a 10-day one.
 *   - Sends any follow-up whose due date has arrived, respecting per-account quota,
 *     and stamps the sent date / advances the state.
 *
 * Safe to run headless (from a time-based trigger).
 *
 * @return {object} { scanned, sent, skippedQuota, skippedCancelled } summary.
 */
function processFollowupQueue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) {
    Logger.log("processFollowupQueue: Leads sheet not found.");
    return { error: "Leads sheet not found" };
  }

  var config = safeGetConfig_();
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) return { scanned: 0, sent: 0, skippedQuota: 0, skippedCancelled: 0 };

  var now = new Date();
  var summary = { scanned: 0, sent: 0, skippedQuota: 0, skippedCancelled: 0 };

  for (var r = 2; r <= lastRow; r++) {
    var email = headersMap["Email"] ? leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim() : "";
    if (!email) continue;

    var replied = headersMap["Replied"] ? leadsSheet.getRange(r, headersMap["Replied"]).getValue().toString().trim() : "";
    if (replied === "Yes") continue; // Human reply ends the sequence.

    // Only follow up on rows that actually got the initial send.
    var pipelineStage = headersMap["Pipeline Stage"] ? leadsSheet.getRange(r, headersMap["Pipeline Stage"]).getValue().toString().trim() : "";
    var outreachStatus = headersMap["Outreach Status"] ? leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim() : "";
    var wasSent = (pipelineStage === "Sent" || pipelineStage === "Follow-up Sent" || outreachStatus === "Email Sent");
    if (!wasSent) continue;

    summary.scanned++;

    // Cancelled (e.g. hard bounce) — stop the sequence.
    var cancelled = headersMap["Followup Cancelled"] ? leadsSheet.getRange(r, headersMap["Followup Cancelled"]).getValue() : false;
    if (cancelled === true || cancelled.toString().toLowerCase() === "true") {
      if (headersMap["Follow-up Status"]) {
        leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("Skipped");
      }
      summary.skippedCancelled++;
      continue;
    }

    // Out-of-office handling: cancel the 3-day, ensure a 10-day touch is scheduled.
    var responseStatus = headersMap["Response Status"] ? leadsSheet.getRange(r, headersMap["Response Status"]).getValue().toString().trim() : "";
    if (responseStatus === "out_of_office") {
      handleOooReschedule_(leadsSheet, r, headersMap);
    }

    // Determine which follow-up (1 or 2) is due, if any.
    var dueInfo = nextDueFollowup_(leadsSheet, r, headersMap, now);
    if (!dueInfo) continue; // Nothing due yet.

    // Resolve the sending account and check its remaining quota BEFORE sending.
    var account = resolveLeadAccount_(leadsSheet, r, headersMap, config);
    var remaining = getRemainingQuota(account);
    if (remaining <= 0) {
      if (headersMap["Follow-up Status"]) {
        leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("skipped_quota_exhausted");
      }
      summary.skippedQuota++;
      Logger.log("processFollowupQueue: row " + r + " skipped — quota exhausted for " + account + ".");
      continue;
    }

    var lead = {
      row: r,
      email: email,
      firstName: headersMap["First Name"] ? leadsSheet.getRange(r, headersMap["First Name"]).getValue().toString().trim() : "",
      company: headersMap["Company"] ? leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim() : "",
      account: account,
      followupNumber: dueInfo.number
    };

    var sendResult = sendFollowup(lead, config);
    if (sendResult && sendResult.success) {
      recordFollowupSent_(leadsSheet, r, headersMap, dueInfo.number, config);
      recordSend(account); // Decrement the per-account daily quota.
      summary.sent++;
      Utilities.sleep(2000); // Gentle pacing to avoid rate limits.
    } else {
      Logger.log("processFollowupQueue: row " + r + " send failed: " + (sendResult && sendResult.reason));
    }
  }

  Logger.log("processFollowupQueue complete: " + JSON.stringify(summary));
  return summary;
}

/**
 * If a lead is out-of-office and hasn't had its 10-day touch scheduled yet, cancel the
 * pending 3-day (Followup 1) and set Followup 2 Due Date to Send Date + 10 days.
 */
// TODO(confirm-with-manager): OOO carve-out currently reschedules followup 2
// from ORIGINAL SEND DATE + 10 days. Alternative interpretation: from the
// followup-1 DUE date + 10 (same value today since due date = send + 3,
// but would diverge if that offset ever changes). Confirm which the manager means.
function handleOooReschedule_(leadsSheet, r, headersMap) {
  if (!headersMap["Followup 2 Due Date"]) return;

  var existing2 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).getValue());
  if (existing2) return; // Already rescheduled.

  // Anchor from the original send date (documented assumption for open question #4).
  var sendDate = null;
  if (headersMap["Send Date"]) sendDate = toDateOrNull_(leadsSheet.getRange(r, headersMap["Send Date"]).getValue());
  if (!sendDate && headersMap["Last Sent At"]) sendDate = toDateOrNull_(leadsSheet.getRange(r, headersMap["Last Sent At"]).getValue());
  if (!sendDate) sendDate = new Date();

  var due2 = addDays_(sendDate, FOLLOWUP_OOO_DELAY_DAYS);
  leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).setValue(due2);

  // Cancel the 3-day touch so it is not also sent.
  if (headersMap["Followup 1 Due Date"]) {
    leadsSheet.getRange(r, headersMap["Followup 1 Due Date"]).setValue("");
  }
  if (headersMap["Follow-up Status"]) {
    leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("OOO — 10-day scheduled");
  }
  Logger.log("handleOooReschedule_: row " + r + " -> Followup 2 due " +
             Utilities.formatDate(due2, Session.getScriptTimeZone(), "yyyy-MM-dd") + ".");
}

/**
 * Determines which follow-up (1 or 2) is due for a row, respecting sent-date guards.
 * @return {object|null} { number: 1|2, dueDate } or null if none due.
 */
function nextDueFollowup_(leadsSheet, r, headersMap, now) {
  // Follow-up 1
  if (headersMap["Followup 1 Due Date"] && headersMap["Followup 1 Sent Date"]) {
    var sent1 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 1 Sent Date"]).getValue());
    var due1 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 1 Due Date"]).getValue());
    if (!sent1 && due1 && due1.getTime() <= now.getTime()) {
      return { number: 1, dueDate: due1 };
    }
  }
  // Follow-up 2 (e.g. the OOO 10-day touch)
  if (headersMap["Followup 2 Due Date"] && headersMap["Followup 2 Sent Date"]) {
    var sent2 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 2 Sent Date"]).getValue());
    var due2 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).getValue());
    if (!sent2 && due2 && due2.getTime() <= now.getTime()) {
      return { number: 2, dueDate: due2 };
    }
  }
  return null;
}

/**
 * Stamps the sent date for the given follow-up number and advances the follow-up state.
 */
function recordFollowupSent_(leadsSheet, r, headersMap, followupNumber, config) {
  var now = new Date();
  var sentCol = followupNumber === 2 ? "Followup 2 Sent Date" : "Followup 1 Sent Date";
  if (headersMap[sentCol]) {
    leadsSheet.getRange(r, headersMap[sentCol]).setValue(now);
  }
  if (headersMap["Last Sent At"]) {
    leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(now);
  }
  if (headersMap["Pipeline Stage"]) {
    leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Follow-up Sent");
  }
  if (headersMap["Follow-up Status"]) {
    leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("Sent");
  }

  // If follow-up 1 just went out and there is no scheduled follow-up 2, schedule the next
  // touch so the sequence continues on the same cadence.
  if (followupNumber === 1 && headersMap["Followup 2 Due Date"]) {
    var existing2 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).getValue());
    if (!existing2) {
      var due2 = addDays_(now, followupDelayDays_(config));
      leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).setValue(due2);
    }
  }
}

/**
 * Resolves the account a lead's follow-up should be sent from, preferring the account
 * the original email was sent from, then the row override, then the default.
 */
function resolveLeadAccount_(leadsSheet, r, headersMap, config) {
  var account = "";
  if (headersMap["Sent From Account"]) {
    account = leadsSheet.getRange(r, headersMap["Sent From Account"]).getValue().toString().trim();
  }
  if (!account && headersMap["Send From Account"]) {
    account = leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim();
  }
  if (!account) {
    account = getConfigValue(config, "Default Send Account", "Account A").toString().trim();
  }
  return account;
}

/**
 * Sends a single follow-up email for a lead, threading onto the original conversation
 * where possible and honoring DRY_RUN / Outreach Mode / Test Email Recipient exactly
 * like the primary outreach path. The subject is prefixed with "RE: ".
 *
 * @param {object} lead { row, email, firstName, company, account, followupNumber }
 * @param {object} [config]
 * @return {object} { success, reason }
 */
function sendFollowup(lead, config) {
  if (!lead || !lead.email) return { success: false, reason: "Missing lead email" };
  config = config || safeGetConfig_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  var headersMap = leadsSheet ? getHeadersMap(leadsSheet) : {};

  var account = lead.account || getConfigValue(config, "Default Send Account", "Account A").toString().trim();
  var senderEmail = getConfigValue(config, account + " Email", Session.getActiveUser().getEmail());
  var senderLabel = getConfigValue(config, account + " Label", "");

  // Build the follow-up copy from the Templates tab (with a safe default fallback).
  var tpl = loadFollowupTemplate_(ss, account);
  var subject = "RE: " + tpl.subject.replace(/{Company}/g, lead.company || "");
  // Guard against a double "RE: RE:" if the template already carried a reply prefix.
  subject = subject.replace(/^(RE:\s*)+/i, "RE: ");

  var body = tpl.body
    .replace(/{First Name}/g, lead.firstName || "")
    .replace(/{Company}/g, lead.company || "")
    .replace(/{AI_INSIGHT}/g, ""); // Optional insight slot left blank if not generated.

  var outreachMode = getConfigValue(config, "Outreach Mode", "Draft").toString();
  var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();
  var dryRun = (typeof DRY_RUN !== "undefined") ? DRY_RUN : true;

  var recipient = lead.email;
  var finalSubject = subject;
  var finalBody = body;

  if (dryRun) {
    recipient = testRecipient || lead.email;
    finalSubject = "[TEST DRAFT] " + subject;
    finalBody = "=== DRY RUN FOLLOW-UP (Original Recipient: " + lead.email + ") ===\n\n" + body;
  } else if (testRecipient) {
    recipient = testRecipient;
    finalSubject = "[TEST] " + subject;
    finalBody = "=== TEST FOLLOW-UP (Original Recipient: " + lead.email + ") ===\n\n" + body;
  }

  var options = {
    htmlBody: finalBody.replace(/\n/g, "<br>"),
    from: senderEmail,
    name: senderLabel,
    replyTo: senderEmail
  };

  // Thread onto the original conversation when we have its id.
  var lastMsg = null;
  if (leadsSheet && lead.row && headersMap["Thread Id"]) {
    var threadId = leadsSheet.getRange(lead.row, headersMap["Thread Id"]).getValue().toString().trim();
    if (threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        if (thread) {
          var msgs = thread.getMessages();
          if (msgs.length > 0) lastMsg = msgs[msgs.length - 1];
        }
      } catch (e) {
        Logger.log("sendFollowup: could not load thread " + threadId + ": " + e.toString());
      }
    }
  }

  try {
    if (outreachMode.toLowerCase() === "send" && !dryRun) {
      if (lastMsg) {
        lastMsg.reply(finalBody, options);
      } else {
        GmailApp.sendEmail(recipient, finalSubject, finalBody, options);
      }
    } else {
      // Draft mode (or DRY_RUN) — create a draft instead of sending.
      if (lastMsg && !dryRun) {
        lastMsg.createDraftReply(finalBody, options);
      } else {
        GmailApp.createDraft(recipient, finalSubject, finalBody, options);
      }
    }
    return { success: true, reason: (dryRun || outreachMode.toLowerCase() !== "send") ? "Draft created" : "Sent" };
  } catch (e) {
    Logger.log("sendFollowup error for row " + (lead.row || "?") + ": " + e.toString());
    return { success: false, reason: e.toString() };
  }
}

/**
 * Loads the "Follow-up Template" from the Templates sheet for a given account,
 * falling back to the first available template and then a hardcoded default.
 */
function loadFollowupTemplate_(ss, account) {
  var def = {
    subject: "Top talent hiring at {Company}",
    body: "Hi {First Name},\n\nJust following up on my earlier note — completely understand things get busy.\n\n" +
          "Happy to connect for a quick call to explore how we might partner effectively.\n\nLooking forward to hearing from you."
  };

  var templatesSheet = ss.getSheetByName("Templates");
  if (!templatesSheet) return def;

  var rows = templatesSheet.getDataRange().getValues();
  var firstMatch = null;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().trim() === "Follow-up Template") {
      var candidate = {
        subject: (rows[i][1] || def.subject).toString(),
        body: (rows[i][2] || def.body).toString()
      };
      if (!firstMatch) firstMatch = candidate;
      var prefAccount = rows[i][3] ? rows[i][3].toString().trim() : "";
      if (prefAccount === account) return candidate;
    }
  }
  return firstMatch || def;
}
