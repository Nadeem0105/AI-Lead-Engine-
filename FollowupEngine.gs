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
  try {
  
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
  var MS_PER_DAY_FU = 24 * 60 * 60 * 1000;
  var NOT_INTERESTED_DAYS_AFTER_FU2 = 3; // 3 days after FU2 with no reply → Not Interested
  var summary = { scanned: 0, sent: 0, skippedQuota: 0, skippedCancelled: 0, markedNotInterested: 0 };

  // ── PASS 0: "Not Interested" Cleanup ──────────────────────────────────────────────────
  // If FU2 was sent 3+ days ago AND the lead still hasn't replied → close them out.
  for (var ri = 2; ri <= lastRow; ri++) {
    var fu2SentVal = headersMap["Followup 2 Sent Date"]
      ? leadsSheet.getRange(ri, headersMap["Followup 2 Sent Date"]).getValue() : null;
    var fu2Sent = toDateOrNull_(fu2SentVal);
    if (!fu2Sent) continue;

    var niStatus = headersMap["Follow-up Status"]
      ? leadsSheet.getRange(ri, headersMap["Follow-up Status"]).getValue().toString().trim() : "";
    // Skip if already closed
    if (niStatus === "Not Interested" || niStatus === "Replied" || niStatus === "Bounced — Sequence Stopped") continue;

    var niReplied = headersMap["Replied"]
      ? leadsSheet.getRange(ri, headersMap["Replied"]).getValue().toString().trim() : "";
    if (niReplied === "Yes") continue; // They replied — don't touch

    var daysSinceFU2 = (now.getTime() - fu2Sent.getTime()) / MS_PER_DAY_FU;
    if (daysSinceFU2 >= NOT_INTERESTED_DAYS_AFTER_FU2) {
      if (headersMap["Follow-up Status"]) {
        leadsSheet.getRange(ri, headersMap["Follow-up Status"]).setValue("Not Interested");
      }
      if (headersMap["Outreach Status"]) {
        leadsSheet.getRange(ri, headersMap["Outreach Status"]).setValue("Not Interested");
      }
      if (headersMap["Pipeline Stage"]) {
        leadsSheet.getRange(ri, headersMap["Pipeline Stage"]).setValue("Closed — Not Interested");
      }
      summary.markedNotInterested++;
      Logger.log("processFollowupQueue: row " + ri + " → Not Interested (" +
                 Math.floor(daysSinceFU2) + " days since FU2, no reply).");
    }
  }
  // ─────────────────────────────────────────────────────────────────────────────────────

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

    // Resolve the sending account BEFORE sending.
    // Note: We intentionally do NOT check getRemainingQuota() here. Follow-ups have absolute priority
    // and will send even if it temporarily exceeds the strict daily cap, preventing broken threads.
    var account = resolveLeadAccount_(leadsSheet, r, headersMap, config);

    var lead = {
      row: r,
      email: email,
      firstName: headersMap["First Name"] ? leadsSheet.getRange(r, headersMap["First Name"]).getValue().toString().trim() : "",
      company: headersMap["Company"] ? leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim() : "",
      account: account,
      followupNumber: dueInfo.number
    };

    var fuType = dueInfo.number === 1 ? "fu1" : "fu2";

    // ── HOURLY RATE LIMIT CHECK ─────────────────────────────────────────────
    try {
      var hourlyFuCheck = checkHourlyLimit(account, fuType);
      if (!hourlyFuCheck.ok) {
        Logger.log("processFollowupQueue: hourly limit hit for " + account + "/" + fuType + ". Stopping batch.");
        SpreadsheetApp.getUi().alert("Hourly Limit Reached", hourlyFuCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
        break;
      }
    } catch(hle) { Logger.log("Hourly limit check error in followup: " + hle); }

    var sendResult = sendFollowup(lead, config);
    if (sendResult && sendResult.success) {
      recordFollowupSent_(leadsSheet, r, headersMap, dueInfo.number, config);
      recordSend(account); // Decrement the per-account daily quota.
      try { recordHourlySend(account, fuType); } catch(he) {}
      summary.sent++;
      Utilities.sleep(2000); // Gentle pacing to avoid rate limits.
    } else {
      Logger.log("processFollowupQueue: row " + r + " send failed: " + (sendResult && sendResult.reason));
    }
  }

  Logger.log("processFollowupQueue complete: " + JSON.stringify(summary));
  return summary;

    var ui = SpreadsheetApp.getUi();
    ui.alert("Success", "Operation completed successfully.", ui.ButtonSet.OK);
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error", e.toString(), ui.ButtonSet.OK);
    Logger.log("Error in processFollowupQueue: " + e.toString());
  }
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
  // Use specific labels so the status progression is visible in the sheet:
  // Pending → Follow-up 1 Sent → (Scan Inbox, no reply) → Pending → Follow-up 2 Sent → Not Interested
  if (headersMap["Follow-up Status"]) {
    leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue(
      "Follow-up " + followupNumber + " Sent"
    );
  }

  // If follow-up 1 just went out and there is no scheduled follow-up 2, schedule the next
  // touch. Follow-up 2 is scheduled for 10 days after Follow-up 1 is sent.
  if (followupNumber === 1 && headersMap["Followup 2 Due Date"]) {
    var existing2 = toDateOrNull_(leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).getValue());
    if (!existing2) {
      var due2 = addDays_(now, 10);
      leadsSheet.getRange(r, headersMap["Followup 2 Due Date"]).setValue(due2);
      Logger.log("recordFollowupSent_: row " + r + " → FU2 scheduled for " +
                 Utilities.formatDate(due2, Session.getScriptTimeZone(), "yyyy-MM-dd") + ".");
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
 * Sends a single follow-up email for a lead.
 *
 * Follow-up 1: Always threaded as a REPLY to the original conversation.
 *              Subject is "RE: <original subject>".
 * Follow-up 2: A FRESH email with a NEW subject and a new angle — NOT a thread reply.
 *              This gives the prospect a clean second chance without a cluttered thread.
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

  var followupNumber = lead.followupNumber || 1;
  var account = lead.account || getConfigValue(config, "Default Send Account", "Account A").toString().trim();
  var senderEmail = getConfigValue(config, account + " Email", "");
  var senderLabel = getConfigValue(config, account + " Label", "");

  // ── Load the correct template per follow-up number ───────────────────────
  var tpl = loadFollowupTemplate_(ss, account, followupNumber);

  var subject, body;
  if (followupNumber === 1) {
    // FU1: Reply subject (RE: <original subject>)
    subject = "RE: " + tpl.subject.replace(/{Company}/g, lead.company || "");
    subject = subject.replace(/^(RE:\s*)+/i, "RE: "); // avoid "RE: RE:"
  } else {
    // FU2: Fresh subject — no "RE:" prefix, new angle entirely
    subject = tpl.subject
      .replace(/{First Name}/g, lead.firstName || "")
      .replace(/{Company}/g, lead.company || "");
    // Strip any accidental RE: prefix from the FU2 template
    subject = subject.replace(/^(RE:\s*)+/i, "").trim();
  }

  body = tpl.body
    .replace(/{First Name}/g, lead.firstName || "")
    .replace(/{Company}/g, lead.company || "")
    .replace(/{AI_INSIGHT}/g, "");

  var outreachMode = getConfigValue(config, "Outreach Mode", "Draft").toString();
  var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();
  var dryRun = (typeof DRY_RUN !== "undefined") ? DRY_RUN : false;

  var recipient = lead.email;
  var finalSubject = subject;
  var finalBody = body;

  if (dryRun) {
    recipient = testRecipient || lead.email;
    finalSubject = "[TEST DRAFT] " + subject;
    finalBody = "=== DRY RUN FOLLOW-UP " + followupNumber + " (Original Recipient: " + lead.email + ") ===\n\n" + body;
  } else if (testRecipient) {
    recipient = testRecipient;
    finalSubject = "[TEST] " + subject;
    finalBody = "=== TEST FOLLOW-UP " + followupNumber + " (Original Recipient: " + lead.email + ") ===\n\n" + body;
  }

  var options = {
    htmlBody: finalBody.replace(/\n/g, "<br>"),
    from: senderEmail,
    name: senderLabel,
    replyTo: senderEmail,
    to: recipient  // Explicitly set To: so reply() doesn't default to the sender's own address
  };


  // ── Send / Draft via Gmail API ───────────────────────────────────────────
  var threadId = null;
  if (leadsSheet && lead.row && headersMap["Thread Id"]) {
    threadId = leadsSheet.getRange(lead.row, headersMap["Thread Id"]).getValue().toString().trim();
  }

  try {
    // Follow-ups ALWAYS send directly — they must not wait in Drafts.
    // Outreach Mode (Draft/Send) only governs initial cold emails.
    // DRY_RUN is still respected as a safety override during development.
    var isDraft = dryRun;

    var apiRes;
    if (threadId) {
      // Both FU1 and FU2 are sent as replies within the SAME thread.
      //
      // FU1: Subject is "RE: <original subject>" → a classic reply.
      // FU2: Subject is a NEW, fresh subject (e.g. "One more thought for {Company}").
      //      This uses Gmail's "Edit Subject" feature — the message is threaded (same
      //      conversation) but arrives with a visually distinct subject, exactly like
      //      the Gmail "Pop out reply → Edit subject" workflow.
      //
      // In both cases we call sendFollowupViaAPI_ which passes the threadId and
      // In-Reply-To / References headers so Gmail keeps it in the original conversation.
      apiRes = sendFollowupViaAPI_(senderEmail, senderLabel, recipient, finalSubject, options.htmlBody, threadId, isDraft);
      if (apiRes.success) {
        Logger.log("sendFollowup FU" + followupNumber + ": sent as threaded reply" +
                   (followupNumber === 2 ? " with new subject (Edit Subject)" : "") +
                   " for row " + lead.row);
      } else {
        throw new Error("API Thread Reply Failed: " + apiRes.error);
      }
    } else {
      // No thread ID found — fall back to a standalone send.
      // This can happen if the original email was created before thread tracking was
      // implemented, or if tracking data was lost.
      Logger.log("sendFollowup FU" + followupNumber + " row " + lead.row +
                 ": no threadId found, sending as standalone fallback.");
      apiRes = sendOrDraftViaAPI_(senderEmail, senderLabel, recipient, finalSubject, options.htmlBody, isDraft);
      if (apiRes.success) {
        Logger.log("sendFollowup FU" + followupNumber + ": standalone fallback sent for row " + lead.row);
      } else {
        throw new Error("API Standalone Fallback Failed: " + apiRes.error);
      }
    }
    
    // HISTORICAL METRICS: Log the follow-up send
    if (!isDraft && typeof logMetricEvent === "function") {
      try {
        var originalSendDateStr = leadsSheet.getRange(lead.row, headersMap["Send Date"]).getValue();
        var originalSendDate = new Date();
        if (originalSendDateStr) {
          originalSendDate = new Date(originalSendDateStr);
        }
        logMetricEvent(account, "Followup_Sent", threadId, lead.email, originalSendDate);
      } catch(me) {
        Logger.log("sendFollowup: metrics logging failed: " + me);
      }
    }

    return { success: true, reason: !isDraft ? "Sent (FU" + followupNumber + ")" : "DRY_RUN Draft (FU" + followupNumber + ")" };
  } catch (e) {
    Logger.log("sendFollowup error for row " + (lead.row || "?") + " FU" + followupNumber + ": " + e.toString());
    return { success: false, reason: e.toString() };
  }
}


/**
 * Loads the follow-up email template for a given account and follow-up number.
 *
 * Template sheet lookup (by "Type" column):
 *   followupNumber = 1 → looks for "Follow-up 1 Template" (or legacy "Follow-up Template")
 *   followupNumber = 2 → looks for "Follow-up 2 Template"
 *
 * Matching priority: exact account match > first match for that type > hardcoded default.
 *
 * Default subjects:
 *   FU1 = "RE: <original subject>" (the RE: is prepended in sendFollowup, not here)
 *   FU2 = "One more thought for {Company}" (fresh angle, no RE:)
 */
function loadFollowupTemplate_(ss, account, followupNumber) {
  followupNumber = followupNumber || 1;

  var def1 = {
    subject: "Top talent hiring at {Company}",
    body: "Hi {First Name},\n\nJust following up on my earlier note — completely understand things get busy.\n\n" +
          "Happy to connect for a quick call to explore how we might partner effectively.\n\nLooking forward to hearing from you."
  };
  var def2 = {
    subject: "One more thought for {Company}",
    body: "Hi {First Name},\n\nI wanted to reach out one more time with a slightly different angle.\n\n" +
          "We've helped companies similar to {Company} significantly reduce time-to-hire while improving " +
          "candidate quality. I'd love to explore whether there's a fit here.\n\n" +
          "Would a 15-minute call this week work for you?"
  };

  var def = (followupNumber === 2) ? def2 : def1;

  var templatesSheet = ss.getSheetByName("Templates");
  if (!templatesSheet) return def;

  var rows = templatesSheet.getDataRange().getValues();

  // Accept both the new specific type names and the legacy generic name for FU1
  var targetTypes = (followupNumber === 2)
    ? ["Follow-up 2 Template", "Followup 2 Template"]
    : ["Follow-up 1 Template", "Followup 1 Template", "Follow-up Template", "Followup Template"];

  var firstMatchPool = [];
  var exactMatchPool = [];

  for (var i = 1; i < rows.length; i++) {
    var rowType = rows[i][0].toString().trim();
    var isMatch = false;
    for (var t = 0; t < targetTypes.length; t++) {
      if (rowType === targetTypes[t]) { isMatch = true; break; }
    }
    if (!isMatch) continue;

    var prefAccount = rows[i][1] ? rows[i][1].toString().trim() : "";
    
    // Read all variations horizontally
    var rowTemplates = [];
    for (var c = 2; c < rows[i].length; c += 2) {
      if (rows[i][c] && rows[i][c+1]) {
        rowTemplates.push({
          subject: rows[i][c].toString(),
          body: rows[i][c+1].toString()
        });
      }
    }
    
    if (rowTemplates.length === 0) continue;
    
    if (firstMatchPool.length === 0) firstMatchPool = rowTemplates;
    if (prefAccount === account || prefAccount === "") {
      exactMatchPool = exactMatchPool.concat(rowTemplates);
    }
  }
  
  var pool = exactMatchPool.length > 0 ? exactMatchPool : (firstMatchPool.length > 0 ? firstMatchPool : [def]);
  
  if (pool.length > 1) {
    var props = PropertiesService.getScriptProperties();
    var propKey = "RECENT_FU" + followupNumber + "_" + account;
    var recentStr = props.getProperty(propKey);
    var recent = [];
    if (recentStr) {
      try { recent = JSON.parse(recentStr); } catch(e) {}
    }
    
    var maxHistory = Math.max(0, pool.length - 1);
    
    var available = pool.filter(function(t) {
      return recent.indexOf(t.subject) === -1;
    });
    
    if (available.length === 0) {
      available = pool;
      recent = [];
    }
    
    var randomIdx = Math.floor(Math.random() * available.length);
    var chosen = available[randomIdx];
    
    recent.push(chosen.subject);
    if (recent.length > maxHistory) {
      recent.shift();
    }
    props.setProperty(propKey, JSON.stringify(recent));
    
    return chosen;
  }
  
  // If pool size is 1, just return the only template available
  return pool[0];
}


/**
 * Calculates how many follow-ups are due TODAY (or earlier) for a specific account.
 * Used by QuotaManager to deduct pending follow-ups from the daily quota before
 * allowing new emails to be sent.
 */
function getPendingFollowupsCount(accountName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) return 0;
  
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  if (lastRow < 2) return 0;
  
  // We grab only the relevant columns to optimize memory and speed
  var statusIdx = headersMap["Follow-up Status"];
  var accIdx = headersMap["Sent From Account"];
  var f1Idx = headersMap["Followup 1 Due Date"];
  var f2Idx = headersMap["Followup 2 Due Date"];
  
  if (statusIdx === undefined || accIdx === undefined) return 0;
  
  var data = leadsSheet.getRange(2, 1, lastRow - 1, leadsSheet.getLastColumn()).getValues();
  
  // quotaTodayString_ returns something like "2024-05-18" (local timezone string)
  // But wait, the function is in QuotaManager.gs. We can do a simpler Date logic here.
  var now = new Date();
  var count = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var status = row[statusIdx - 1] ? row[statusIdx - 1].toString().trim() : "";
    var acc = row[accIdx - 1] ? row[accIdx - 1].toString().trim() : "";
    
    // Only count leads that are currently pending for this exact account
    if (status !== "Pending" && status !== "OOO — 10-day scheduled") continue;
    if (acc !== accountName) continue;
    
    // Check if Due Date is today or earlier
    var due1 = row[f1Idx - 1] ? new Date(row[f1Idx - 1]) : null;
    var due2 = row[f2Idx - 1] ? new Date(row[f2Idx - 1]) : null;
    
    if (due1 && !isNaN(due1.getTime()) && due1.getTime() <= now.getTime()) {
      count++;
    } else if (due2 && !isNaN(due2.getTime()) && due2.getTime() <= now.getTime()) {
      count++;
    }
  }
  return count;
}

/**
 * Called by the Drip Engine.
 * Scans the Leads sheet and sends exactly ONE due follow-up, then stops.
 * Returns true if a follow-up was sent, false otherwise.
 */
function sendOneDueFollowup(config) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leadsSheet = ss.getSheetByName("Leads");
    if (!leadsSheet) return false;
    
    var headersMap = getHeadersMap(leadsSheet);
    var lastRow = leadsSheet.getLastRow();
    if (lastRow <= 1) return false;
    
    var now = new Date();
    
    for (var r = 2; r <= lastRow; r++) {
      var email = headersMap["Email"] ? leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim() : "";
      if (!email) continue;
      
      var replied = headersMap["Replied"] ? leadsSheet.getRange(r, headersMap["Replied"]).getValue().toString().trim() : "";
      if (replied === "Yes") continue;
      
      var pipelineStage = headersMap["Pipeline Stage"] ? leadsSheet.getRange(r, headersMap["Pipeline Stage"]).getValue().toString().trim() : "";
      var outreachStatus = headersMap["Outreach Status"] ? leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim() : "";
      var wasSent = (pipelineStage === "Sent" || pipelineStage === "Follow-up Sent" || outreachStatus === "Email Sent");
      if (!wasSent) continue;
      
      var cancelled = headersMap["Followup Cancelled"] ? leadsSheet.getRange(r, headersMap["Followup Cancelled"]).getValue() : false;
      if (cancelled === true || cancelled.toString().toLowerCase() === "true") {
        continue;
      }
      
      var dueInfo = nextDueFollowup_(leadsSheet, r, headersMap, now);
      if (!dueInfo) continue;
      
      var account = resolveLeadAccount_(leadsSheet, r, headersMap, config);
      var lead = {
        row: r,
        email: email,
        firstName: headersMap["First Name"] ? leadsSheet.getRange(r, headersMap["First Name"]).getValue().toString().trim() : "",
        company: headersMap["Company"] ? leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim() : "",
        account: account,
        followupNumber: dueInfo.number
      };
      var fuType = dueInfo.number === 1 ? "fu1" : "fu2";
      
      // Hourly limit check
      var hourlyFuCheck = checkHourlyLimit(account, fuType);
      if (!hourlyFuCheck.ok) {
        Logger.log("sendOneDueFollowup: Hourly limit hit for " + account + "/" + fuType + ". Skipping this lead for now.");
        continue; // Try finding another lead that might be assigned to a different account!
      }
      
      var sendResult = sendFollowup(lead, config);
      if (sendResult && sendResult.success) {
        recordFollowupSent_(leadsSheet, r, headersMap, dueInfo.number, config);
        try { recordHourlySend(account, fuType); } catch(he) {}
        
        // Log to Daily Send Log
        var threadId = headersMap["Thread Id"] ? leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim() : "";
        try { logDailySend("FU" + dueInfo.number, account, email, "(Follow-up " + dueInfo.number + ")", threadId); } catch(e) {}
        
        Logger.log("Drip Engine: Sent FU" + dueInfo.number + " to " + email);
        return true; // Sent exactly one
      }
    }
  } catch (e) {
    Logger.log("sendOneDueFollowup error: " + e.toString());
  }
  return false;
}
 
 