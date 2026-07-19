/**
 * MULTI-MAILBOX GMAIL API INTEGRATION
 * Sends from 5 mailboxes (3x buttersearch.in, 2x buttersearch.io) from a single
 * Sheet/Apps Script project, regardless of which account owns the script itself.
 * Also handles reply detection, OOO detection, and bounce detection per mailbox.
 *
 * ============================================================
 * ONE-TIME SETUP
 * ============================================================
 * 1. Google Cloud Console (can reuse your existing Gemini project or make a new one):
 *    - Enable the Gmail API
 *    - Create an OAuth 2.0 Client ID (type: Web application)
 *    - Add the redirect URI given by the OAuth2 library below once you run getService() once
 *
 * 2. In the Apps Script editor: Libraries (+) -> add this script ID:
 *    1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkPQ9UPbw
 *
 * 3. Store your OAuth client credentials once:
 *    PropertiesService.getScriptProperties().setProperty('OAUTH_CLIENT_ID', '...');
 *    PropertiesService.getScriptProperties().setProperty('OAUTH_CLIENT_SECRET', '...');
 *
 * 4. Update MAILBOX_CONFIG below with your real 5 addresses.
 *
 * 5. Authorize each mailbox ONCE (see "ONE-TIME AUTHORIZATION" section at the bottom):
 *    - Run startAuth('in1') from the script editor, check Logger for the URL
 *    - Open that URL in a browser while logged into in1@buttersearch.in specifically
 *    - Click Allow
 *    - Repeat for in2, in3, io1, io2
 *    Scopes requested cover BOTH sending and reading (gmail.modify), so you only
 *    need to do this once per mailbox, not once for send + once for read.
 * ============================================================
 */

// ---------------------------------------------------------
// MAILBOX CONFIG — map internal keys to real addresses
// ---------------------------------------------------------
var MAILBOX_CONFIG = {
  'in1': 'mailbox1@buttersearch.in',
  'in2': 'mailbox2@buttersearch.in',
  'in3': 'mailbox3@buttersearch.in',
  'io1': 'mailbox1@buttersearch.io',
  'io2': 'mailbox2@buttersearch.io'
};

function getMailboxAddress(mailboxKey) {
  return MAILBOX_CONFIG[mailboxKey];
}

// ---------------------------------------------------------
// OAUTH SERVICE SETUP
// ---------------------------------------------------------
function getService(mailboxKey) {
  var props = PropertiesService.getScriptProperties();
  return OAuth2.createService('gmail_' + mailboxKey)
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(props.getProperty('OAUTH_CLIENT_ID'))
    .setClientSecret(props.getProperty('OAUTH_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    // gmail.modify is a superset covering send + read + label/archive,
    // so one scope covers sending, reply detection, and bounce detection.
    .setScope('https://www.googleapis.com/auth/gmail.modify')
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent');
}

// ---------------------------------------------------------
// SENDING
// Returns { success: bool, threadId: string } — store threadId
// on the lead's Sheet row so you can check for replies later.
// ---------------------------------------------------------
function sendFromMailbox(mailboxKey, toAddress, subject, bodyText) {
  var service = getService(mailboxKey);
  if (!service.hasAccess()) {
    Logger.log('Not authorized for ' + mailboxKey + ': ' + service.getLastError());
    return { success: false, threadId: null };
  }

  var rawMessage = Utilities.base64EncodeWebSafe(
    "To: " + toAddress + "\r\n" +
    "Subject: " + subject + "\r\n" +
    "Content-Type: text/plain; charset=UTF-8\r\n\r\n" +
    bodyText
  );

  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(mailboxKey) + "/messages/send";
  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    payload: JSON.stringify({ raw: rawMessage }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    var data = JSON.parse(response.getContentText());
    return { success: true, threadId: data.threadId };
  } else {
    Logger.log("Send failed for " + mailboxKey + ": " + response.getContentText());
    return { success: false, threadId: null };
  }
}

// Use this for followups — same thread, so it stays a proper reply chain
// rather than a disconnected new email. Requires the original threadId
// and Message-ID header (also worth storing at initial send time).
function sendFollowupInThread(mailboxKey, toAddress, subject, bodyText, threadId, originalMessageId) {
  var service = getService(mailboxKey);
  if (!service.hasAccess()) {
    Logger.log('Not authorized for ' + mailboxKey + ': ' + service.getLastError());
    return { success: false };
  }

  var rawMessage = Utilities.base64EncodeWebSafe(
    "To: " + toAddress + "\r\n" +
    "Subject: Re: " + subject + "\r\n" +
    "In-Reply-To: " + originalMessageId + "\r\n" +
    "References: " + originalMessageId + "\r\n" +
    "Content-Type: text/plain; charset=UTF-8\r\n\r\n" +
    bodyText
  );

  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(mailboxKey) + "/messages/send";
  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    payload: JSON.stringify({ raw: rawMessage, threadId: threadId }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  return { success: response.getResponseCode() === 200 };
}

// ---------------------------------------------------------
// REPLY + OOO DETECTION
// Call this before firing any scheduled followup.
// ---------------------------------------------------------
function checkForReply(mailboxKey, threadId) {
  var service = getService(mailboxKey);
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(mailboxKey) + "/threads/" + threadId;
  var options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log("Thread fetch failed: " + response.getContentText());
    return { hasReply: false, isOOO: false };
  }

  var thread = JSON.parse(response.getContentText());
  var messages = thread.messages;

  if (messages.length <= 1) return { hasReply: false, isOOO: false };

  var lastMessage = messages[messages.length - 1];
  var headers = lastMessage.payload.headers;
  var fromHeader = headers.filter(function(h) { return h.name === "From"; })[0];
  var autoSubmitted = headers.filter(function(h) { return h.name === "Auto-Submitted"; })[0];

  var isFromLead = fromHeader && fromHeader.value.indexOf(getMailboxAddress(mailboxKey)) === -1;
  var isAutoReply = !!(autoSubmitted && autoSubmitted.value !== "no");

  return {
    hasReply: isFromLead && !isAutoReply,
    isOOO: isFromLead && isAutoReply,
    messageId: lastMessage.id
  };
}

// ---------------------------------------------------------
// BOUNCE DETECTION
// Returns candidate bounce message IDs since a given date;
// fetch each one's body/headers separately to extract which
// original recipient bounced (Original-Recipient / Final-Recipient
// headers, or parsed from the DSN body).
// ---------------------------------------------------------
function checkForBounce(mailboxKey, sinceDate) {
  var service = getService(mailboxKey);
  var query = "from:mailer-daemon after:" + Utilities.formatDate(sinceDate, "GMT", "yyyy/MM/dd");
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(mailboxKey) + "/messages?q=" + encodeURIComponent(query);
  var options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log("Bounce check failed for " + mailboxKey + ": " + response.getContentText());
    return [];
  }

  var data = JSON.parse(response.getContentText());
  return data.messages || []; // array of { id, threadId } — fetch each for full DSN details
}

// ---------------------------------------------------------
// DAILY BATCH RUNNER (suggested pattern)
// Run once per morning, before the followup-priority sending window,
// looping through all mailboxes rather than checking on a per-send basis.
// ---------------------------------------------------------
function dailyReplyAndBounceCheck() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  for (var key in MAILBOX_CONFIG) {
    var bounces = checkForBounce(key, yesterday);
    Logger.log(key + ": " + bounces.length + " potential bounce message(s) found");
    // TODO: cross-reference bounces against your send log,
    // then loop your Sheet's pending-followup rows through checkForReply()
    // to decide send / skip / reschedule for today's followup batch.
  }
}

// ---------------------------------------------------------
// ONE-TIME AUTHORIZATION (run manually per mailbox, then delete/ignore)
// ---------------------------------------------------------
// authCallback removed to prevent conflict with GmailAPI.gs

function getServiceNameFromRequest(request) {
  // The OAuth2 library encodes the service name in the state param;
  // this helper just strips the library's own prefix back to your mailboxKey.
  return request.parameter.state.split('/')[0].replace('gmail_', '');
}

function startAuth(mailboxKey) {
  var service = getService(mailboxKey);
  Logger.log('Open this URL while logged into ' + getMailboxAddress(mailboxKey) + ':');
  Logger.log(service.getAuthorizationUrl());
}

function checkAuthStatus(mailboxKey) {
  var service = getService(mailboxKey);
  Logger.log(mailboxKey + ' (' + getMailboxAddress(mailboxKey) + '): ' + (service.hasAccess() ? 'AUTHORIZED' : 'NOT authorized'));
}

function checkAllAuthStatus() {
  for (var key in MAILBOX_CONFIG) {
    checkAuthStatus(key);
  }
}
