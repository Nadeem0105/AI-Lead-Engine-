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

// Script Property keys backing the rolling bounce-rate metric.
var BOUNCE_METRIC_SENT_KEY = "bounce_metric_sent_total";
var BOUNCE_METRIC_BOUNCED_KEY = "bounce_metric_bounced_total";

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
 * Updates the rolling bounce-rate metric stored in Script Properties.
 * Tracks cumulative delivered + bounced counts and logs the current rate.
 *
 * NOTE: window semantics (daily/weekly/monthly) are pending manager confirmation
 * (open question #3). This implements an all-time cumulative counter; callers can
 * reset it on whatever cadence is chosen via resetBounceMetric().
 *
 * @param {boolean} isBounce True if the classified message was a bounce.
 * @return {number} The current bounce rate as a fraction (0-1).
 */
function updateBounceMetric_(isBounce) {
  var props = PropertiesService.getScriptProperties();
  var sent = parseInt(props.getProperty(BOUNCE_METRIC_SENT_KEY) || "0");
  var bounced = parseInt(props.getProperty(BOUNCE_METRIC_BOUNCED_KEY) || "0");

  sent += 1;
  if (isBounce) bounced += 1;

  props.setProperty(BOUNCE_METRIC_SENT_KEY, sent.toString());
  props.setProperty(BOUNCE_METRIC_BOUNCED_KEY, bounced.toString());

  var rate = sent > 0 ? (bounced / sent) : 0;
  Logger.log("Bounce metric updated: " + bounced + "/" + sent + " = " + (rate * 100).toFixed(2) + "%");
  return rate;
}

/**
 * Returns the current rolling bounce rate as a fraction (0-1).
 */
function getBounceRate() {
  var props = PropertiesService.getScriptProperties();
  var sent = parseInt(props.getProperty(BOUNCE_METRIC_SENT_KEY) || "0");
  var bounced = parseInt(props.getProperty(BOUNCE_METRIC_BOUNCED_KEY) || "0");
  return sent > 0 ? (bounced / sent) : 0;
}

/**
 * Resets the rolling bounce-rate metric window (call on the chosen cadence).
 */
function resetBounceMetric() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(BOUNCE_METRIC_SENT_KEY, "0");
  props.setProperty(BOUNCE_METRIC_BOUNCED_KEY, "0");
  Logger.log("Bounce metric window reset.");
}
