/**
 * AI Lead Generation Engine - Email Open Tracking (Invisible Pixel)
 * File: Tracking.gs
 */

/**
 * Handles incoming GET requests for the tracking pixel.
 * Increments the email open count for the corresponding lead.
 * 
 * @param {object} e Event parameter from Google Apps Script Web App execution
 * @return {TextOutput} Transparent 1x1 GIF image bytes
 */
function doGet(e) {
  // Tracking pixel functionality has been removed per user request.
  // We no longer track email opens.
  return ContentService.createTextOutput("Tracking disabled").setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Finds the lead row with the matching email and increments 'Mail Opens' and updates 'Last Opened At'.
 * 
 * @param {string} email The email address to look up
 */
function recordEmailOpen(email) {
  // This function is deprecated and no longer writes to the 'Mail Opens' or 'Last Opened At' columns.
  // Reply tracking is now the only engagement signal tracked.
}

/**
 * Logs a sent email to the 'Daily Send Log' sheet.
 * @param {string} type 'Fresh', 'FU1', or 'FU2'
 * @param {string} account The account name that sent the email
 * @param {string} email The lead's email address
 * @param {string} subject The subject line of the email
 * @param {string} threadId The Thread ID (for tying replies back later)
 */
function logDailySend(type, account, email, subject, threadId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Daily Send Log");
    if (!logSheet) return; // Silent fail if the tab doesn't exist
    
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    
    // Columns: "Date", "Time", "Type", "Send From Account", "Lead Email", "Subject", "Thread ID"
    logSheet.appendRow([dateStr, timeStr, type, account, email, subject, threadId]);
    
  } catch(e) {
    Logger.log("logDailySend Error: " + e.toString());
  }
}
