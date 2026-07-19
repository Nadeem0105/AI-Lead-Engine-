/**
 * AI Lead Generation Engine - Secondary Draft Verification
 * Loops through the "Ready to Send" tab, scores unchecked drafts,
 * writes the reason, and auto-checks them if they pass the threshold.
 */

function verifyDraftsInReadyTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var readySheet = ss.getSheetByName("Ready to Send");
  if (!readySheet) {
    Logger.log("Ready to Send sheet not found.");
    return;
  }

  var headersMap = getHeadersMap(readySheet);
  
  // Required columns
  if (!headersMap["Draft Quality Score"] || !headersMap["2nd Time AI Score"] || !headersMap["Ready for Send"] || !headersMap["AI Verification Notes"]) {
    Logger.log("Missing required columns in Ready to Send tab.");
    SpreadsheetApp.getUi().alert("Error", "Missing required columns in Ready to Send tab. Please ensure '2nd Time AI Score' and 'AI Verification Notes' exist.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var config = safeGetConfig_();
  var thresholdStr = getConfigValue(config, "Draft Quality Threshold", "7").toString();
  var threshold = parseFloat(thresholdStr);
  if (isNaN(threshold)) threshold = 7;

  var lastRow = readySheet.getLastRow();
  if (lastRow < 2) return;

  var summary = { processed: 0, passed: 0, failed: 0, errors: 0 };
  
  for (var r = 2; r <= lastRow; r++) {
    var draftId = headersMap["Draft ID"] ? readySheet.getRange(r, headersMap["Draft ID"]).getValue().toString().trim() : "";
    if (draftId === "") {
      break; // Checkboxes make getLastRow() return 1000, so we break when data ends
    }

    var isReady = readySheet.getRange(r, headersMap["Ready for Send"]).getValue();
    var existingNote = readySheet.getRange(r, headersMap["AI Verification Notes"]).getValue().toString().trim();
    
    // Only process unchecked drafts that haven't been reviewed yet
    if (isReady === true || existingNote !== "") {
      continue;
    }

    var company = headersMap["Company"] ? readySheet.getRange(r, headersMap["Company"]).getValue().toString().trim() : "";
    var firstName = headersMap["First Name"] ? readySheet.getRange(r, headersMap["First Name"]).getValue().toString().trim() : "";
    var subject = headersMap["Subject"] ? readySheet.getRange(r, headersMap["Subject"]).getValue().toString().trim() : "";
    var body = headersMap["Body"] ? readySheet.getRange(r, headersMap["Body"]).getValue().toString().trim() : "";

    if (!body || !subject) continue;
    
    try {
      summary.processed++;
      
      var prompt = "You are a QA manager reviewing a cold outreach email.\n" +
                   "Lead Context:\n" +
                   "- Name: " + firstName + "\n" +
                   "- Company: " + company + "\n\n" +
                   "Draft Subject: " + subject + "\n" +
                   "Draft Body:\n" + body + "\n\n" +
                   "Evaluate this draft strictly on a scale of 1-10. Look for placeholder errors (like '[Company Name]'), tone, and logical flow.\n" +
                   "Return a pure JSON object in exactly this format: {\"score\": 8, \"reason\": \"Short, professional, accurate.\"}\n" +
                   "Do not include any other text or markdown wrappers.";

      var modelName = getConfigValue(config, "Gemini Model", "gemini-2.0-flash").toString().trim();
      var aiResponse = callFailoverModelForOutreach(modelName, prompt, config);
      
      // Attempt to parse JSON
      var parsed = null;
      try {
        var cleanResponse = aiResponse.replace(/```json/gi, "").replace(/```/g, "").trim();
        parsed = JSON.parse(cleanResponse);
      } catch (parseErr) {
        // Fallback parsing if JSON fails
        var scoreMatch = aiResponse.match(/"score"\s*:\s*(\d+)/i);
        var score = scoreMatch ? parseInt(scoreMatch[1], 10) : 5;
        parsed = { score: score, reason: "Failed to parse pure JSON. Raw output: " + aiResponse };
      }

      var finalScore = parsed.score;
      var finalReason = parsed.reason;

      // Update the sheet
      readySheet.getRange(r, headersMap["2nd Time AI Score"]).setValue(finalScore);
      readySheet.getRange(r, headersMap["AI Verification Notes"]).setValue(finalReason);

      if (finalScore >= threshold) {
        readySheet.getRange(r, headersMap["Ready for Send"]).setValue(true);
        summary.passed++;
      } else {
        summary.failed++;
      }
      
      Utilities.sleep(2000); // Pacing for rate limits
      
    } catch (e) {
      Logger.log("Draft verification error for row " + r + ": " + e.toString());
      readySheet.getRange(r, headersMap["AI Verification Notes"]).setValue("Error: " + e.toString());
      summary.errors++;
    }
  }

  var msg = "Verification Complete.\n\n" +
            "Processed: " + summary.processed + "\n" +
            "Passed (Auto-checked): " + summary.passed + "\n" +
            "Failed (Needs manual review): " + summary.failed + "\n" +
            "Errors: " + summary.errors;
            
  SpreadsheetApp.getUi().alert("AI Verification Results", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
