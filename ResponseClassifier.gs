/**
 * AI Lead Generation Engine - Inbound Response Classifier
 * File: ResponseClassifier.gs
 *
 * Email Ops Upgrade. Inspects incoming thread messages and classifies them as a hard/soft
 * bounce, an out-of-office (OOO) auto-reply, or a genuine human reply. Bounces and OOO
 * replies feed the follow-up state machine (FollowupEngine.gs): bounces cancel all
 * follow-ups and update the rolling bounce-rate metric; OOO replies cancel the 3-day
 * follow-up and defer to a 10-day one.
 *
 * A "lead" is a plain object; the relevant columns are written back onto it:
 *   { row, email, responseStatus, bounceType, followupCancelled }
 * When lead.row is present, the same fields are also persisted to the Leads sheet.
 */

// Response Status enum values written to the Leads sheet.
var RESPONSE_STATUS = {
  HUMAN_REPLY: "human_reply",
  OUT_OF_OFFICE: "out_of_office",
  BOUNCED: "bounced"
};

// Script Property key backing the bounce-rate metric. Stores a JSON array of timestamped
// events — [ [epochMillis, bounceFlag], ... ] where bounceFlag is 1 (bounce) or 0
// (delivered) — so the bounce rate can be computed over any rolling window after the fact.
var BOUNCE_METRIC_EVENTS_KEY = "bounce_metric_events";
// Legacy cumulative-counter keys (pre-rolling-window). Kept only so resetBounceMetric()
// can clear any stale values left behind by the previous implementation.
var BOUNCE_METRIC_SENT_KEY = "bounce_metric_sent_total";
var BOUNCE_METRIC_BOUNCED_KEY = "bounce_metric_bounced_total";

// Default rolling window (days) over which getBounceRate() is computed.
var BOUNCE_METRIC_WINDOW_DAYS = 7;
// How much history to retain on disk. Kept larger than the active window so the window can
// be widened later (e.g. to 30 days) without having discarded the events it would need.
var BOUNCE_METRIC_RETENTION_DAYS = 90;
// One day in milliseconds.
var MS_PER_DAY = 24 * 60 * 60 * 1000;

/**
 * Inbox Scanner Loop. Walks the Leads sheet for active threads (Thread Id present,
 * Response Status still empty), opens each Gmail thread, and classifies the first
 * genuine inbound message — a prospect reply or a mailer-daemon/postmaster bounce.
 *
 * Messages sent by our OWN sending accounts (e.g. automated follow-ups from
 * FollowupEngine.gs) are ignored so a follow-up we sent is never mistaken for a reply.
 * Own-account detection resolves the lead's "Sent From Account" label to its configured
 * email, and additionally excludes every "<Account> Email" configured in Config plus the
 * active user, so multi-account setups are covered.
 *
 * Safe to run headless (from a menu item or a time-based trigger).
 *
 * @return {object} { scanned, classified, skippedNoInbound, errors } summary.
 */
function processInboundResponses() {
  try {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) {
    Logger.log("processInboundResponses: Leads sheet not found.");
    return { error: "Leads sheet not found" };
  }

  var headersMap = getHeadersMap(leadsSheet);
  if (!headersMap["Thread Id"]) {
    Logger.log("processInboundResponses: 'Thread Id' column missing — run Setup Sheets.");
    return { error: "'Thread Id' column missing" };
  }

  var config = safeGetConfig_();
  var ownEmails = collectOwnAccountEmails_(config);

  var lastRow = leadsSheet.getLastRow();
  var summary = { scanned: 0, classified: 0, skippedNoInbound: 0, errors: 0 };
  if (lastRow <= 1) return summary;

  for (var r = 2; r <= lastRow; r++) {
    var threadId = leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim();
    if (!threadId) continue; // No conversation to inspect.

    // Only look at threads not yet classified.
    var responseStatus = headersMap["Response Status"]
      ? leadsSheet.getRange(r, headersMap["Response Status"]).getValue().toString().trim()
      : "";
    if (responseStatus) continue;

    summary.scanned++;

    var leadEmail = headersMap["Email"] ? leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim() : "";

    // Resolve the account this lead was sent from, then its email, so replies from our
    // own follow-up sends are not mistaken for prospect responses.
    var accountEmail = "";
    var rowOwnEmails = ownEmails.slice();
    if (headersMap["Sent From Account"]) {
      var accountLabel = leadsSheet.getRange(r, headersMap["Sent From Account"]).getValue().toString().trim();
      if (accountLabel) {
        accountEmail = getConfigValue(config, accountLabel + " Email", "").toString().trim().toLowerCase();
        if (accountEmail && rowOwnEmails.indexOf(accountEmail) === -1) rowOwnEmails.push(accountEmail);
      }
    }
    if (!accountEmail) return; // Skip if no email configured

    var apiRes;
    try {
      apiRes = checkForReplyViaAPI_(accountEmail, threadId);
    } catch (e) {
      Logger.log("processInboundResponses: could not load thread API " + threadId + " (row " + r + "): " + e.toString());
      summary.errors++;
      continue;
    }

    if (apiRes.error) {
      Logger.log("processInboundResponses: API returned error for thread " + threadId + ": " + apiRes.error);
      summary.errors++;
      continue;
    }

    if (apiRes.hasReply) {
       var lead = { row: r, email: leadEmail };
       lead.responseStatus = RESPONSE_STATUS.HUMAN_REPLY;
       lead.replied = "Yes";
       lead.followupCancelled = true;
       persistResponseFields_(lead);
       summary.classified++;
       Logger.log("API: " + lead.email + " -> HUMAN REPLY.");
    } else if (apiRes.isOOO) {
       var lead = { row: r, email: leadEmail };
       lead.responseStatus = RESPONSE_STATUS.OUT_OF_OFFICE;
       lead.followupCancelled = false; 
       persistResponseFields_(lead);
       summary.classified++;
       Logger.log("API: " + lead.email + " -> OUT OF OFFICE. 3-day follow-up cancelled, 10-day scheduled.");
    } else {
       summary.skippedNoInbound++;
    }
  } // end for loop

  Logger.log("processInboundResponses complete: " + JSON.stringify(summary));
  return summary;

  } catch (e) {
    Logger.log("Error in processInboundResponses: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * Collects the set of lowercased email addresses that belong to us (our sending
 * accounts + the active user), so inbound scanning can filter out our own messages.
 *
 * Pulls every Config key shaped like "<Account> Email" (e.g. "Account A Email").
 */
function collectOwnAccountEmails_(config) {
  var emails = [];
  var add = function (addr) {
    var e = (addr || "").toString().trim().toLowerCase();
    if (e && emails.indexOf(e) === -1) emails.push(e);
  };

  if (config) {
    for (var key in config) {
      if (!config.hasOwnProperty(key)) continue;
      if (/ Email$/.test(key)) add(config[key]);
    }
  }

  try {
    // add(Session.getActiveUser().getEmail()); // Removed to avoid permission issues if not configured
  } catch (e) {
    // Active user email may be unavailable in some execution contexts — ignore.
  }

  return emails;
}

/**
 * True if a message was sent by one of our own accounts. Matches the address parsed
 * out of the message's From header against the provided own-email list.
 */
function isOwnAccountMessage_(message, ownEmails) {
  if (!ownEmails || ownEmails.length === 0) return false;
  var from = "";
  try {
    from = (message.getFrom() || "").toString().toLowerCase();
  } catch (e) {
    return false;
  }
  var sender = extractEmailAddress_(from);
  for (var i = 0; i < ownEmails.length; i++) {
    // Compare on the parsed address, falling back to a substring check on the raw header.
    if ((sender && sender === ownEmails[i]) || from.indexOf(ownEmails[i]) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * Extracts the bare email address from a From header such as
 * '"Ayush" <sender1@company.com>' -> 'sender1@company.com'. Returns "" if none found.
 */
function extractEmailAddress_(from) {
  if (!from) return "";
  var angle = from.match(/<([^>]+)>/);
  if (angle && angle[1]) return angle[1].trim().toLowerCase();
  var bare = from.match(/[^\s<>@]+@[^\s<>@]+/);
  return bare ? bare[0].trim().toLowerCase() : "";
}

/**
 * Classifies a single inbound thread message for a lead and records the outcome.
 *
 * @param {GmailMessage} threadMessage The incoming message to classify.
 * @param {object} lead The lead object (optionally carrying a `row` for sheet writes).
 * @return {string} One of RESPONSE_STATUS.* ("bounced" | "out_of_office" | "human_reply").
 */
function classifyReply(threadMessage, lead) {
  lead = lead || {};

  // 1. Bounce detection takes precedence — a bounce means the address is dead.
  var bounceType = isBounceMessage(threadMessage);
  if (bounceType) {
    lead.responseStatus = RESPONSE_STATUS.BOUNCED;
    lead.bounceType = bounceType;            // "hard" | "soft"
    lead.followupCancelled = true;
    persistResponseFields_(lead);
    updateBounceMetric_(true);
    Logger.log("classifyReply: " + (lead.email || "") + " -> BOUNCE (" + bounceType + "). Follow-ups cancelled.");
    return RESPONSE_STATUS.BOUNCED;
  }

  // A delivered (non-bounce) message counts toward the rolling send/deliverability metric.
  updateBounceMetric_(false);

  // 2. Out-of-office auto-reply — cancel the 3-day follow-up, defer to the 10-day one.
  if (isOutOfOfficeMessage(threadMessage)) {
    lead.responseStatus = RESPONSE_STATUS.OUT_OF_OFFICE;
    lead.followupCancelled = false; // Not cancelled outright — FollowupEngine reschedules to 10 days.
    persistResponseFields_(lead);
    Logger.log("classifyReply: " + (lead.email || "") + " -> OUT OF OFFICE. 3-day follow-up cancelled, 10-day scheduled.");
    return RESPONSE_STATUS.OUT_OF_OFFICE;
  }

  // 3. Otherwise it is a genuine human reply.
  lead.responseStatus = RESPONSE_STATUS.HUMAN_REPLY;
  lead.replied = "Yes";
  lead.followupCancelled = true;
  persistResponseFields_(lead);
  Logger.log("classifyReply: " + (lead.email || "") + " -> HUMAN REPLY.");
  return RESPONSE_STATUS.HUMAN_REPLY;
}

/**
 * Writes the classification fields back to the Leads sheet if the lead carries a row.
 * No-op when there is no row (pure in-memory classification, e.g. unit tests).
 */
function persistResponseFields_(lead) {
  if (!lead || !lead.row) return;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leadsSheet = ss.getSheetByName("Leads");
    if (!leadsSheet) return;
    var headersMap = getHeadersMap(leadsSheet);

    if (lead.responseStatus !== undefined && headersMap["Response Status"]) {
      leadsSheet.getRange(lead.row, headersMap["Response Status"]).setValue(lead.responseStatus);
    }
    if (lead.bounceType !== undefined && headersMap["Bounce Type"]) {
      leadsSheet.getRange(lead.row, headersMap["Bounce Type"]).setValue(lead.bounceType);
    }
    if (lead.followupCancelled !== undefined && headersMap["Followup Cancelled"]) {
      leadsSheet.getRange(lead.row, headersMap["Followup Cancelled"]).setValue(lead.followupCancelled === true);
    }
    if (lead.replied !== undefined && headersMap["Replied"]) {
      leadsSheet.getRange(lead.row, headersMap["Replied"]).setValue(lead.replied);
    }
    
    // Update Follow-up Status based on classification outcome.
    // Human reply → "Replied" (stops all further follow-ups).
    // Bounce       → "Bounced — Sequence Stopped".
    if (headersMap["Follow-up Status"]) {
      if (lead.responseStatus === RESPONSE_STATUS.HUMAN_REPLY) {
        leadsSheet.getRange(lead.row, headersMap["Follow-up Status"]).setValue("Replied");
        // Also mark lead as Replied in Outreach Status for visibility
        if (headersMap["Outreach Status"]) {
          leadsSheet.getRange(lead.row, headersMap["Outreach Status"]).setValue("Replied");
        }
        if (headersMap["Pipeline Stage"]) {
          leadsSheet.getRange(lead.row, headersMap["Pipeline Stage"]).setValue("Replied");
        }
      } else if (lead.responseStatus === RESPONSE_STATUS.BOUNCED) {
        leadsSheet.getRange(lead.row, headersMap["Follow-up Status"]).setValue("Bounced — Sequence Stopped");
      }
    }
    
    // HISTORICAL METRICS: Log Replied or Bounced
    if (typeof logMetricEvent === "function" && typeof hasEventAlreadyLogged === "function") {
      var eventTypeToLog = "";
      if (lead.responseStatus === RESPONSE_STATUS.HUMAN_REPLY) {
        eventTypeToLog = "Replied";
      } else if (lead.responseStatus === RESPONSE_STATUS.BOUNCED && lead.bounceType === "hard") {
        eventTypeToLog = "Bounced"; // Only log hard bounces to metrics
      }
      
      if (eventTypeToLog) {
        var account = headersMap["Send From Account"] ? leadsSheet.getRange(lead.row, headersMap["Send From Account"]).getValue().toString().trim() : "";
        var threadId = headersMap["Thread Id"] ? leadsSheet.getRange(lead.row, headersMap["Thread Id"]).getValue().toString().trim() : "";
        var originalSendDateStr = headersMap["Send Date"] ? leadsSheet.getRange(lead.row, headersMap["Send Date"]).getValue() : "";
        var originalSendDate = originalSendDateStr ? new Date(originalSendDateStr) : new Date();
        
        if (threadId && !hasEventAlreadyLogged(threadId, eventTypeToLog)) {
          logMetricEvent(account, eventTypeToLog, threadId, lead.email, originalSendDate);
        }
      }
    }

  } catch (e) {
    Logger.log("ResponseClassifier.persistResponseFields_ error: " + e.toString());
  }
}


/**
 * Detects a delivery-failure (bounce) message and distinguishes hard vs soft bounces.
 *
 * Hard bounce  = permanent failure (5.x.x SMTP codes, "user unknown", "no such user").
 * Soft bounce  = transient failure (4.x.x SMTP codes, "mailbox full", "try again").
 *
 * @param {GmailMessage} threadMessage
 * @return {string|null} "hard" | "soft" if a bounce, otherwise null (truthy = bounce).
 */
function isBounceMessage(threadMessage) {
  if (!threadMessage) return null;

  var from = "";
  var subject = "";
  var body = "";
  try {
    from = (threadMessage.getFrom() || "").toString().toLowerCase();
    subject = (threadMessage.getSubject() || "").toString().toLowerCase();
    body = (threadMessage.getPlainBody() || "").toString().toLowerCase();
  } catch (e) {
    Logger.log("isBounceMessage: could not read message: " + e.toString());
    return null;
  }

  var raw = getRawContentSafe_(threadMessage).toLowerCase();

  // Sender-based signals: bounces come from the mail system, not a person.
  var senderIsDaemon =
    from.indexOf("mailer-daemon") !== -1 ||
    from.indexOf("postmaster") !== -1 ||
    from.indexOf("mail delivery") !== -1 ||
    from.indexOf("maildelivery") !== -1;

  // Subject-based signals commonly used by MTAs.
  var subjectSignals = [
    "delivery status notification", "undeliverable", "undelivered mail",
    "returned mail", "mail delivery failed", "delivery failure",
    "failure notice", "could not be delivered", "delivery incomplete"
  ];
  var subjectHit = matchesAny_(subject, subjectSignals);

  // Content-type marker present on DSN bounce reports.
  var isDsn = raw.indexOf("report-type=delivery-status") !== -1 ||
              raw.indexOf("message/delivery-status") !== -1;

  if (!senderIsDaemon && !subjectHit && !isDsn) {
    return null; // Not a bounce.
  }

  // Classify severity. Look at SMTP status codes and phrasing in body + raw content.
  var haystack = body + " " + raw;

  // Hard bounce indicators (permanent).
  var hardPhrases = [
    "user unknown", "no such user", "does not exist", "recipient address rejected",
    "address not found", "account has been disabled", "account disabled",
    "mailbox unavailable", "invalid recipient", "unrouteable address",
    "permanent failure", "550", "551", "553", "554"
  ];
  // A 5.x.x enhanced status code = permanent failure.
  var hasPermanentCode = /\b5\.\d\.\d\b/.test(haystack);

  // Soft bounce indicators (transient).
  var softPhrases = [
    "mailbox full", "over quota", "quota exceeded", "temporarily", "try again",
    "temporary failure", "greylist", "rate limit", "451", "452", "421"
  ];
  var hasTransientCode = /\b4\.\d\.\d\b/.test(haystack);

  if (hasPermanentCode || matchesAny_(haystack, hardPhrases)) {
    return "hard";
  }
  if (hasTransientCode || matchesAny_(haystack, softPhrases)) {
    return "soft";
  }

  // Recognized as a bounce but severity unclear — treat as hard (fail safe: stop sending).
  return "hard";
}

/**
 * Detects an automated out-of-office / vacation auto-reply.
 *
 * @param {GmailMessage} threadMessage
 * @return {boolean} True if the message is an OOO / auto-reply.
 */
function isOutOfOfficeMessage(threadMessage) {
  if (!threadMessage) return false;

  var subject = "";
  var body = "";
  try {
    subject = (threadMessage.getSubject() || "").toString().toLowerCase();
    body = (threadMessage.getPlainBody() || "").toString().toLowerCase();
  } catch (e) {
    Logger.log("isOutOfOfficeMessage: could not read message: " + e.toString());
    return false;
  }

  var raw = getRawContentSafe_(threadMessage).toLowerCase();

  // Header-based signals are the most reliable indicator of an auto-responder.
  var headerSignals = [
    "auto-submitted: auto-replied",
    "auto-submitted: auto-generated",
    "x-autoreply",
    "x-autorespond",
    "x-auto-response-suppress",
    "precedence: auto_reply",
    "precedence: bulk"
  ];
  if (matchesAny_(raw, headerSignals)) {
    return true;
  }

  // Subject / body phrasing signals.
  var phraseSignals = [
    "out of office", "out of the office", "ooo", "automatic reply", "auto-reply",
    "auto reply", "on leave", "annual leave", "on vacation", "on holiday",
    "away from my desk", "currently away", "i am currently out",
    "will be out of office", "limited access to email", "maternity leave",
    "paternity leave", "parental leave"
  ];
  return matchesAny_(subject, phraseSignals) || matchesAny_(body, phraseSignals);
}

/**
 * Returns true if the haystack contains any of the needle substrings.
 */
function matchesAny_(haystack, needles) {
  for (var i = 0; i < needles.length; i++) {
    if (haystack.indexOf(needles[i]) !== -1) return true;
  }
  return false;
}

/**
 * Safely reads a message's raw MIME content (headers + body) for deep inspection.
 * Returns "" if unavailable (large messages or permission edge cases).
 */
function getRawContentSafe_(threadMessage) {
  try {
    return threadMessage.getRawContent() || "";
  } catch (e) {
    return "";
  }
}

/**
 * Records one classified message as a timestamped event in the bounce-metric log, prunes
 * events older than the retention horizon, and returns the current rolling bounce rate.
 *
 * Storing per-event timestamps (rather than a running counter) means the reporting window
 * is a pure read-time decision: nothing here bakes in a particular cadence.
 *
 * // TODO(confirm-with-manager): bounce rate currently computed on a rolling
 * // 7-day window. Manager may want weekly/monthly/all-time instead — the
 * // timestamp-log storage above supports any window without a schema change,
 * // only the windowDays argument needs to change once confirmed.
 *
 * @param {boolean} isBounce True if the classified message was a bounce.
 * @return {number} The current bounce rate over the default window as a fraction (0-1).
 */
function updateBounceMetric_(isBounce) {
  var props = PropertiesService.getScriptProperties();
  var nowMs = new Date().getTime();

  var events = readBounceEvents_(props);
  events.push([nowMs, isBounce ? 1 : 0]);
  events = pruneBounceEvents_(events, nowMs, BOUNCE_METRIC_RETENTION_DAYS);
  props.setProperty(BOUNCE_METRIC_EVENTS_KEY, JSON.stringify(events));

  var rate = computeBounceRate_(events, nowMs, BOUNCE_METRIC_WINDOW_DAYS);
  var windowStats = countBounceEvents_(events, nowMs, BOUNCE_METRIC_WINDOW_DAYS);
  Logger.log("Bounce metric updated (" + BOUNCE_METRIC_WINDOW_DAYS + "-day window): " +
             windowStats.bounced + "/" + windowStats.total + " = " + (rate * 100).toFixed(2) + "%");
  return rate;
}

/**
 * Returns the bounce rate as a fraction (0-1) computed over a rolling window.
 * @param {number} [windowDays] Window size in days (defaults to BOUNCE_METRIC_WINDOW_DAYS).
 */
function getBounceRate(windowDays) {
  var props = PropertiesService.getScriptProperties();
  var nowMs = new Date().getTime();
  var events = readBounceEvents_(props);
  var w = (windowDays === undefined || windowDays === null) ? BOUNCE_METRIC_WINDOW_DAYS : windowDays;
  return computeBounceRate_(events, nowMs, w);
}

/**
 * Reads and parses the stored bounce-event log. Returns [] on any missing/corrupt data.
 */
function readBounceEvents_(props) {
  var raw = props.getProperty(BOUNCE_METRIC_EVENTS_KEY);
  if (!raw) return [];
  try {
    var arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    Logger.log("readBounceEvents_: could not parse event log, resetting: " + e.toString());
    return [];
  }
}

/**
 * Drops events older than retentionDays before the given reference time.
 */
function pruneBounceEvents_(events, nowMs, retentionDays) {
  var cutoff = nowMs - retentionDays * MS_PER_DAY;
  var kept = [];
  for (var i = 0; i < events.length; i++) {
    if (events[i] && events[i][0] >= cutoff) kept.push(events[i]);
  }
  return kept;
}

/**
 * Counts { total, bounced } for events falling within windowDays of the reference time.
 */
function countBounceEvents_(events, nowMs, windowDays) {
  var cutoff = nowMs - windowDays * MS_PER_DAY;
  var total = 0, bounced = 0;
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (!ev || ev[0] < cutoff) continue;
    total++;
    if (ev[1] === 1) bounced++;
  }
  return { total: total, bounced: bounced };
}

/**
 * Computes the bounce rate (0-1) over a rolling window from a stored event log.
 */
function computeBounceRate_(events, nowMs, windowDays) {
  var stats = countBounceEvents_(events, nowMs, windowDays);
  return stats.total > 0 ? (stats.bounced / stats.total) : 0;
}

/**
 * Clears the bounce-metric event log (and any legacy cumulative counters). Rarely needed
 * now that the rolling window ages events out on its own, but kept for a hard reset.
 */
function resetBounceMetric() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(BOUNCE_METRIC_EVENTS_KEY);
  props.deleteProperty(BOUNCE_METRIC_SENT_KEY);
  props.deleteProperty(BOUNCE_METRIC_BOUNCED_KEY);
  Logger.log("Bounce metric event log reset.");
}

function debugInboxScanner() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leadsSheet = ss.getSheetByName("Leads");
    var headersMap = getHeadersMap(leadsSheet);
    var lastRow = leadsSheet.getLastRow();
    var config = safeGetConfig_();
    var ownEmails = collectOwnAccountEmails_(config);
    var debugLog = "DEBUG INBOX SCANNER:\nOwn Emails: " + JSON.stringify(ownEmails) + "\n\n";

    for (var r = 2; r <= lastRow; r++) {
      var threadId = leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim();
      var responseStatus = headersMap["Response Status"] ? leadsSheet.getRange(r, headersMap["Response Status"]).getValue().toString().trim() : "";
      
      if (!threadId) continue;
      
      debugLog += "Row " + r + " [Thread ID: " + threadId + "] [Status: " + (responseStatus||"BLANK") + "]\n";
      
      if (responseStatus) {
         debugLog += "  -> Skipped: Already has Response Status.\n";
         continue;
      }
      
      try {
        var thread = GmailApp.getThreadById(threadId);
        if (!thread) {
           debugLog += "  -> ERROR: Thread not found in Gmail.\n";
           continue;
        }
        var messages = thread.getMessages();
        debugLog += "  -> Messages length: " + messages.length + "\n";
        
        if (messages.length <= 1) {
           debugLog += "  -> Skipped: No replies detected (length <= 1).\n";
           continue;
        }
        
        for (var m = 1; m < messages.length; m++) {
          var from = messages[m].getFrom();
          var isOwn = isOwnAccountMessage_(messages[m], ownEmails);
          debugLog += "  -> Msg " + m + " From: " + from + " | isOwnAccount: " + isOwn + "\n";
          if (!isOwn) {
             debugLog += "  -> ATTEMPTING CLASSIFY REPLY...\n";
             try {
                var result = classifyReply(messages[m], { row: r, email: "test@test.com" });
                debugLog += "  -> CLASSIFY RESULT: " + result + "\n";
             } catch(err2) {
                debugLog += "  -> CLASSIFY ERROR: " + err2.toString() + "\n";
             }
             break;
          }
        }
      } catch (e) {
        debugLog += "  -> THREAD ERROR: " + e.toString() + "\n";
      }
    }
    
    SpreadsheetApp.getUi().alert("Debug Output", debugLog, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(err) {
    SpreadsheetApp.getUi().alert("Error", err.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
