/**
 * AI Lead Generation Engine - Historical Metrics Tracker
 * File: MetricsManager.gs
 * 
 * Handles logging of sending, reply, and bounce events for cohort-based analytics.
 */

/**
 * Logs an event to the 'Metrics Log' sheet for historical dashboard tracking.
 * 
 * @param {string} account The email account that sent the email (e.g. "Account A")
 * @param {string} eventType "Fresh_Sent", "Followup_Sent", "Replied", or "Bounced"
 * @param {string} threadId The Gmail thread ID for deduplication and attribution
 * @param {string} leadEmail The recipient's email address
 * @param {Date} originalSendDate The date the *first* outreach email was sent in this thread
 */
function logMetricEvent(account, eventType, threadId, leadEmail, originalSendDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Metrics Log");
    if (!sheet) return; // Silent fail if dashboard not set up yet
    
    // Ensure we have a valid date object for the cohort math
    var cohortDate = originalSendDate;
    if (!cohortDate || !(cohortDate instanceof Date) || isNaN(cohortDate.getTime())) {
      cohortDate = new Date(); // Fallback to today if not provided or invalid
    }
    
    // Calculate cohort week and month based on the ORIGINAL SEND DATE
    var weekStr = getIsoWeekString_(cohortDate);
    var monthStr = Utilities.formatDate(cohortDate, Session.getScriptTimeZone(), "yyyy-MM");
    
    // Date of the actual event happening (today)
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Build row: ["Timestamp", "Date", "Week", "Month", "Account", "Event Type", "Lead Email", "Thread Id", "Original Send Date"]
    var newRow = [
      now,                    // Timestamp of event
      dateStr,                // Date of event
      weekStr,                // Cohort Week
      monthStr,               // Cohort Month
      account,
      eventType,
      leadEmail,
      threadId,
      Utilities.formatDate(cohortDate, Session.getScriptTimeZone(), "yyyy-MM-dd") // Cohort Date string
    ];
    
    sheet.appendRow(newRow);
    
  } catch(e) {
    Logger.log("Error in logMetricEvent: " + e.toString());
  }
}

/**
 * Checks if a specific event type has already been logged for a given thread.
 * Used to prevent duplicate "Replied" or "Bounced" logs when the inbox scanner runs hourly.
 * 
 * @param {string} threadId The Gmail thread ID
 * @param {string} eventType "Replied" or "Bounced"
 * @return {boolean} True if already logged, false otherwise
 */
function hasEventAlreadyLogged(threadId, eventType) {
  if (!threadId) return false;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Metrics Log");
    if (!sheet) return false;
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false; // Empty sheet (only headers)
    
    // Fetch just Event Type (Col F = 6) and Thread Id (Col H = 8) to save memory
    var eventData = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    var threadData = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
    
    for (var i = eventData.length - 1; i >= 0; i--) {
      // Search backwards since duplicate is likely recent
      if (threadData[i][0] === threadId && eventData[i][0] === eventType) {
        return true;
      }
    }
    
    return false;
  } catch(e) {
    Logger.log("Error in hasEventAlreadyLogged: " + e.toString());
    return false; // Fail open (allow log) if error
  }
}

/**
 * Calculates the ISO-8601 week string for a given date (e.g., "2026-W29").
 * 
 * @param {Date} date The date to calculate
 * @return {string} The ISO week string
 */
function getIsoWeekString_(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1)/7);
  var weekStr = weekNo < 10 ? "0" + weekNo : weekNo;
  return d.getUTCFullYear() + "-W" + weekStr;
}
