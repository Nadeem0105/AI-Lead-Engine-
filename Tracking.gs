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
