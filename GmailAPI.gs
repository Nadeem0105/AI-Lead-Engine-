/**
 * GMAIL API INTEGRATION (OAUTH2)
 * Handles multi-mailbox sending and reply detection using the Apps Script OAuth2 library.
 * Instead of Domain-Wide Delegation, this uses standard per-mailbox consent.
 */

/**
 * Returns the OAuth2 service for a specific email address.
 * The email address itself is used as the service name (e.g., 'gmail_test@example.com').
 */
function getGmailService_(senderEmail) {
  if (!senderEmail) throw new Error("senderEmail is required for getGmailService_");
  
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('OAUTH_CLIENT_ID');
  var clientSecret = props.getProperty('OAUTH_CLIENT_SECRET');
  
  if (!clientId || !clientSecret) {
    throw new Error("OAUTH_CLIENT_ID and OAUTH_CLIENT_SECRET must be set in Script Properties.");
  }

  // Sanitize email for the service name (remove special chars)
  var safeName = senderEmail.replace(/[^a-zA-Z0-9]/g, '_');
  
  return OAuth2.createService('gmail_' + safeName)
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setCache(CacheService.getScriptCache())
    .setScope('https://www.googleapis.com/auth/gmail.modify')
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent')
    .setParam('login_hint', senderEmail); // Forces the login screen to default to this email
}

/**
 * The callback that handles the OAuth2 response.
 */
function authCallback(request) {
  var state = request.parameter.state || '';
  var email = PropertiesService.getScriptProperties().getProperty('oauth_map_' + state);
  
  if (!email) {
    return HtmlService.createHtmlOutput('Denied. Could not find the mapping for this authorization request. Please try again.');
  }

  var service = getGmailService_(email);
  try {
    var authorized = service.handleCallback(request);
    if (authorized) {
      // Validate that the authenticated user's email matches the expected email
      var url = "https://gmail.googleapis.com/gmail/v1/users/me/profile";
      var options = {
        headers: { Authorization: "Bearer " + service.getAccessToken() },
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        var profile = JSON.parse(response.getContentText());
        var authedEmail = profile.emailAddress;
        if (authedEmail.toLowerCase() !== email.toLowerCase()) {
           service.reset(); // clear the invalid token
           return HtmlService.createHtmlOutput('<h3>Authorization Denied (Account Mismatch)</h3><p>You attempted to authorize <b>' + email + '</b> but signed in with Google Account <b>' + authedEmail + '</b>.</p><p>Please close this tab, try again, and ensure you select the correct Google account.</p>');
        }
      }

      // Clean up the map
      PropertiesService.getScriptProperties().deleteProperty('oauth_map_' + state);
      return HtmlService.createHtmlOutput('Success! You can close this tab. The mailbox authorized was: <b>' + email + '</b>');
    } else {
      return HtmlService.createHtmlOutput('Denied. Error: ' + service.getLastError());
    }
  } catch (e) {
    return HtmlService.createHtmlOutput('Denied. Exception: ' + e.toString());
  }
}

/**
 * Generates the authorization URL for a given email address.
 */
function getAuthorizationUrl(senderEmail) {
  var url = getGmailService_(senderEmail).getAuthorizationUrl();
  
  // Extract the state parameter to map it to the email
  var match = url.match(/state=([^&]+)/);
  if (match && match[1]) {
    var state = decodeURIComponent(match[1]);
    PropertiesService.getScriptProperties().setProperty('oauth_map_' + state, senderEmail);
  }
  
  return url;
}

/**
 * Checks if a specific email address is authorized.
 */
function isAuthorized(senderEmail) {
  return getGmailService_(senderEmail).hasAccess();
}

/**
 * Internal helper to base64url encode strings (RFC 4648)
 */
function base64EncodeWebSafe_(str) {
  return Utilities.base64EncodeWebSafe(str);
}

/**
 * Builds the raw RFC 2822 email string.
 */
function buildRawEmail_(senderEmail, senderName, toAddress, subject, bodyHtml, originalMessageId) {
  var boundary = "----=_Part_" + Utilities.getUuid().replace(/-/g, "");
  var raw = [];
  var fromStr = senderName ? '"' + senderName + '" <' + senderEmail + '>' : senderEmail;
  raw.push("From: " + fromStr);
  raw.push("To: " + toAddress);
  raw.push("Subject: " + subject);
  
  if (originalMessageId) {
    raw.push("In-Reply-To: " + originalMessageId);
    raw.push("References: " + originalMessageId);
  }
  
  raw.push("MIME-Version: 1.0");
  raw.push("Content-Type: multipart/alternative; boundary=\"" + boundary + "\"");
  raw.push("");
  
  // Plain text version (stripped from HTML)
  raw.push("--" + boundary);
  raw.push("Content-Type: text/plain; charset=\"UTF-8\"");
  raw.push("");
  raw.push(bodyHtml.replace(/<br>/g, "\n").replace(/<[^>]+>/g, ""));
  
  // HTML version
  raw.push("--" + boundary);
  raw.push("Content-Type: text/html; charset=\"UTF-8\"");
  raw.push("");
  raw.push(bodyHtml);
  
  raw.push("--" + boundary + "--");
  
  return base64EncodeWebSafe_(raw.join("\r\n"));
}

/**
 * Sends or drafts a fresh email via the Gmail API.
 * @returns { success: boolean, threadId: string, error: string }
 */
function sendOrDraftViaAPI_(senderEmail, senderName, toAddress, subject, bodyHtml, isDraft) {
  var service = getGmailService_(senderEmail);
  if (!service.hasAccess()) {
    return { success: false, error: "Not authorized for " + senderEmail };
  }

  var rawMessage = buildRawEmail_(senderEmail, senderName, toAddress, subject, bodyHtml, null);
  var endpoint = isDraft ? "drafts" : "messages/send";
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(senderEmail) + "/" + endpoint;
  
  var payload = { message: { raw: rawMessage } };
  if (!isDraft) {
    payload = { raw: rawMessage };
  }

  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    var data = JSON.parse(response.getContentText());
    // For drafts, data is { id: "draftId", message: { id: "msgId", threadId: "threadId" } }
    // For send, data is { id: "msgId", threadId: "threadId" }
    var threadId = isDraft ? data.message.threadId : data.threadId;
    var draftId = isDraft ? data.id : null;
    return { success: true, threadId: threadId, draftId: draftId };
  } else {
    return { success: false, error: response.getContentText() };
  }
}

/**
 * Sends an existing draft by its ID via the Gmail API.
 * @param {string} senderEmail The account to send from.
 * @param {string} draftId The ID of the draft to send.
 * @returns { success: boolean, threadId: string, error: string }
 */
function sendDraftByIdViaAPI_(senderEmail, draftId) {
  var service = getGmailService_(senderEmail);
  if (!service.hasAccess()) {
    return { success: false, error: "Not authorized for " + senderEmail };
  }
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(senderEmail) + "/drafts/send";
  var payload = { id: draftId };
  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    var data = JSON.parse(response.getContentText());
    return { success: true, threadId: data.threadId };
  } else {
    return { success: false, error: response.getContentText() };
  }
}

/**
 * Sends or drafts a reply in an existing thread via the Gmail API.
 */
function sendFollowupViaAPI_(senderEmail, senderName, toAddress, subject, bodyHtml, threadId, isDraft) {
  var service = getGmailService_(senderEmail);
  if (!service.hasAccess()) {
    return { success: false, error: "Not authorized for " + senderEmail };
  }

  // 1. Fetch the thread to get the original Message-ID
  var getUrl = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(senderEmail) + "/threads/" + threadId;
  var getOptions = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  var getResponse = UrlFetchApp.fetch(getUrl, getOptions);
  var originalMessageId = null;
  
  if (getResponse.getResponseCode() === 200) {
    var threadData = JSON.parse(getResponse.getContentText());
    if (threadData.messages && threadData.messages.length > 0) {
      var lastMsg = threadData.messages[threadData.messages.length - 1];
      var msgIdHeader = lastMsg.payload.headers.filter(function(h) { return h.name.toLowerCase() === "message-id"; })[0];
      if (msgIdHeader) originalMessageId = msgIdHeader.value;
    }
  } else {
    Logger.log("Warning: Could not fetch thread " + threadId + " to get Message-ID for reply threading.");
  }

  // 2. Build and send the reply
  var rawMessage = buildRawEmail_(senderEmail, senderName, toAddress, subject, bodyHtml, originalMessageId);
  var endpoint = isDraft ? "drafts" : "messages/send";
  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(senderEmail) + "/" + endpoint;
  
  var payload = { message: { raw: rawMessage, threadId: threadId } };
  if (!isDraft) {
    payload = { raw: rawMessage, threadId: threadId };
  }

  var options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    return { success: true };
  } else {
    return { success: false, error: response.getContentText() };
  }
}

/**
 * Checks a specific thread for replies (from the prospect) and OOO responses.
 */
function checkForReplyViaAPI_(senderEmail, threadId) {
  var service = getGmailService_(senderEmail);
  if (!service.hasAccess()) return { hasReply: false, isOOO: false, error: "Not auth" };

  var url = "https://gmail.googleapis.com/gmail/v1/users/" + encodeURIComponent(senderEmail) + "/threads/" + threadId;
  var options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    return { hasReply: false, isOOO: false, error: response.getContentText() };
  }

  var thread = JSON.parse(response.getContentText());
  var messages = thread.messages;

  // Thread only has 1 message? Then no reply.
  if (!messages || messages.length <= 1) return { hasReply: false, isOOO: false };

  var lastMessage = messages[messages.length - 1];
  var headers = lastMessage.payload.headers;
  
  var fromHeader = headers.filter(function(h) { return h.name.toLowerCase() === "from"; })[0];
  var autoSubmitted = headers.filter(function(h) { return h.name.toLowerCase() === "auto-submitted"; })[0];

  // If the last message was NOT sent by us, it's a reply from the prospect.
  var isFromUs = fromHeader && fromHeader.value.toLowerCase().indexOf(senderEmail.toLowerCase()) !== -1;
  var isFromProspect = !isFromUs;
  
  var isAutoReply = !!(autoSubmitted && autoSubmitted.value.toLowerCase() !== "no");

  return {
    hasReply: isFromProspect && !isAutoReply,
    isOOO: isFromProspect && isAutoReply
  };
}
