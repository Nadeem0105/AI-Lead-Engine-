/**
 * AI Lead Generation Engine - DeepSeek Automated Email Outreach
 * File: Outreach.gs
 */

// Set to true during development/testing to prevent any real sends.
// If true, even if mode is 'Send', it will only create drafts with "[TEST DRAFT]" prepended.
var DRY_RUN = false; 

/**
 * Main pipeline to process and generate outreach emails for qualified and validated leads.
 */
function runOutreachPipeline() {
  processOutreachInternal(false, false, false);
}

/**
 * Manual testing action that ignores hourly trigger rules.
 */
function runOutreachPipelineManualBatch() {
  processOutreachInternal(false, false, true);
}

/**
 * Core outreach processing engine. Can be run manually (with UI) or hourly (headless).
 */
/**
 * Shows a YES/NO/CANCEL popup to let the user choose which account to send from.
 * YES  = Account A
 * NO   = Account B
 * CANCEL = use the row-level 'Send From Account' column setting
 *
 * @return {string|null} 'Account A', 'Account B', or null (use row settings)
 */
function promptAccountSelection() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Choose Sending Account",
    "Which email account should be used to send these emails?\n\n" +
    "\u2022 Yes  \u2192  Force Account A (Ayush)\n" +
    "\u2022 No   \u2192  Force Account B (Harshith)\n" +
    "\u2022 Cancel  \u2192  Use AUTOMATIC routing (Score Bands or Row settings)",
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (response === ui.Button.YES)  return "Account A";
  if (response === ui.Button.NO)   return "Account B";
  return null; // CANCEL = respect per-row setting
}

/**
 * Feature 5: Routes lead to an account pool based on their score.
 * 
 * @param {number} score 
 * @param {object} config 
 * @returns {string|null} The randomly selected account from the pool, or null if it should fall back to default.
 */
function selectAccountByScore(score, config) {
  if (isNaN(score) || score === "") return null;
  var scoreNum = parseFloat(score);
  
  var dailyCap = parseInt(getConfigValue(config, "Per Account Daily Cap", "40")) || 40;
  
  if (scoreNum >= 8) {
    var highPool = (config["Score Band High Accounts"] || "").split(",");
    var chosen = getAvailableAccountFromPool(highPool, dailyCap, config);
    if (chosen) return chosen;
    // All high-band accounts at cap — fall through to mid band
    Logger.log("All high-band accounts at daily cap. Falling back to mid band.");
  }
  
  if (scoreNum >= 6) {
    var midPool = (config["Score Band Mid Accounts"] || "").split(",");
    var chosen = getAvailableAccountFromPool(midPool, dailyCap, config);
    if (chosen) return chosen;
    Logger.log("All mid-band accounts at daily cap. No eligible account found.");
    return "HOLD"; // All accounts at cap — hold the row rather than exceed cap
  }
  
  var lowBehavior = (config["Score Band Low Behavior"] || "Default").trim();
  if (lowBehavior === "Hold") {
    return "HOLD"; // Special flag to prevent sending
  }
  return null;
}

/**
 * Returns a randomly selected account from the pool that is still under its daily send cap.
 * Returns null if all accounts in the pool are at or above the cap.
 */
function getAvailableAccountFromPool(pool, dailyCap, config) {
  var availableAccounts = [];
  for (var i = 0; i < pool.length; i++) {
    var acc = pool[i].toString().trim();
    if (!acc) continue;
    
    // Instead of just checking sentToday, we must use getRemainingQuota
    // because getRemainingQuota properly deducts pending follow-ups from the daily cap.
    var remainingSlots = 0;
    if (typeof getRemainingQuota === "function") {
      remainingSlots = getRemainingQuota(acc);
    } else {
      // Fallback if QuotaManager is missing (should never happen)
      var sentToday = getAccountSentToday(acc);
      remainingSlots = dailyCap - sentToday;
    }
    
    if (remainingSlots > 0) {
      availableAccounts.push(acc);
    } else {
      Logger.log("Account '" + acc + "' has no remaining fresh slots (cap hit or reserved for follow-ups). Skipping.");
    }
  }
  if (availableAccounts.length === 0) return null;
  return availableAccounts[Math.floor(Math.random() * availableAccounts.length)];
}

/**
 * Returns how many emails have been sent by this account today.
 * Count is stored in Script Properties under the key: account_sent_YYYY-MM-DD_AccountName
 */
function getAccountSentToday(accountName) {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var key = "account_sent_" + today + "_" + accountName;
  return parseInt(props.getProperty(key) || "0");
}

/**
 * Increments the per-account send counter for today.
 * Call this immediately after a successful send for an account.
 */
function recordAccountSend(accountName) {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var key = "account_sent_" + today + "_" + accountName;
  var current = parseInt(props.getProperty(key) || "0");
  props.setProperty(key, (current + 1).toString());
}

function getRandomAccountFromPool(pool) {
  var validAccounts = [];
  for (var i = 0; i < pool.length; i++) {
    var acc = pool[i].trim();
    if (acc) validAccounts.push(acc);
  }
  if (validAccounts.length === 0) return null;
  return validAccounts[Math.floor(Math.random() * validAccounts.length)];
}


/**
 * Applies a warning-only protection to the 'Send From Account' cell in a given row
 * so that accidental edits trigger a confirmation dialog.
 *
 * @param {Sheet}  sheet    The Leads sheet.
 * @param {number} rowNumber 1-indexed row.
 * @param {number} colIndex  1-indexed column for 'Send From Account'.
 */
function lockCellAfterSend(sheet, rowNumber, colIndex, label) {
  try {
    var range = sheet.getRange(rowNumber, colIndex);
    var existing = range.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    // Avoid duplicate protections on the same cell
    for (var i = 0; i < existing.length; i++) {
      existing[i].remove();
    }
    var protection = range.protect();
    protection.setWarningOnly(true);
    protection.setDescription("Locked after send — " + (label || range.getValue()));
  } catch (e) {
    Logger.log("Could not protect cell (row " + rowNumber + ", col " + colIndex + "): " + e);
  }
}

function processOutreachInternal(isHourly, selectedOnly, isManualBatch) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    if (!isHourly) {
      SpreadsheetApp.getUi().alert("Error", "Leads sheet not found. Please run 'Setup Sheets' first.", SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Error: Leads sheet not found.");
    }
    return;
  }
  
  var config = getConfig();
  
  if (isHourly) {
    var sendingActive = getConfigValue(config, "Sending Active", false);
    if (sendingActive === false || sendingActive.toString().toLowerCase() === "false") {
      Logger.log("Hourly run skipped: Sending Active is false.");
      return;
    }
  }
  
  // Verify API Key presence
  var nvidiaKeys = getApiKeysList("NVIDIA_API_KEYS", "NVIDIA_API_KEY");
  var groqKeys = getApiKeysList("GROQ_API_KEYS", "GROQ_API_KEY");
  var cerebrasKeys = getApiKeysList("CEREBRAS_API_KEYS", "CEREBRAS_API_KEY");
  var geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  
  if (nvidiaKeys.length === 0 && groqKeys.length === 0 && cerebrasKeys.length === 0 && !geminiKey) {
    if (!isHourly) {
      SpreadsheetApp.getUi().alert("API Key Missing", "Please set at least one of GEMINI_API_KEY, NVIDIA_API_KEY, GROQ_API_KEY, or CEREBRAS_API_KEY in your Script Properties.", SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Error: API Key Missing.");
    }
    return;
  }
  
  var headersMap = getHeadersMap(leadsSheet);
  
  // Verify required columns exist — literal check only for script-owned OUTPUT columns
  // (Feature 1: data columns like Company/Email resolve via canonical mapping downstream).
  var required = [
    "Validation Status", "Outreach Status", "Score", "Pipeline Stage"
  ];
  for (var i = 0; i < required.length; i++) {
    if (!headersMap[required[i]]) {
      if (!isHourly) {
        SpreadsheetApp.getUi().alert("Error", "Missing required column in Leads sheet: '" + required[i] + "'. Please run Setup Sheets.", SpreadsheetApp.getUi().ButtonSet.OK);
      } else {
        Logger.log("Error: Missing required column '" + required[i] + "'.");
      }
      return;
    }
  }
  
  var lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) {
    if (!isHourly) {
      SpreadsheetApp.getUi().alert("No Data", "Leads sheet is empty.", SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      Logger.log("Leads sheet is empty.");
    }
    return;
  }
  
  var startTime = new Date().getTime();
  var processed = 0;
  var successCount = 0;
  var errors = [];
  
  // Load outreach parameters from Config
  var outreachMode = getConfigValue(config, "Outreach Mode", "Draft");
  var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();
  var batchLimit = parseInt(getConfigValue(config, "Emails Per Run", "10")) || 10;
  var senderSource = getConfigValue(config, "Sender Source", "MainSheet").toString().trim();

  // ── STAGING MODE ──────────────────────────────────────────────────────────
  // When Staging Mode = true: emails go to REAL lead addresses, no [TEST] prefix,
  // outreach mode is forced to Draft, and batch size is capped.
  var stagingMode = getConfigValue(config, "Staging Mode", "false").toString().trim().toLowerCase() === "true";
  if (stagingMode) {
    outreachMode   = "Draft"; // Force Draft so nothing auto-sends
    batchLimit     = parseInt(getConfigValue(config, "Staging Batch Limit", "3")) || 3;
    
    var alertTitle = "⚡ Staging Mode Active";
    var alertMsg = "The engine is running in STAGING MODE.\n\n";
    
    if (testRecipient) {
      Logger.log("STAGING MODE active: Draft forced, batch capped at " + batchLimit + ". Using TEST RECIPIENT.");
      alertMsg += "⚠️ Emails will go to your TEST RECIPIENT: " + testRecipient + "\n";
      alertMsg += "✅ No [TEST] prefix on subject lines (Clean Test)\n";
    } else {
      Logger.log("STAGING MODE active: real recipients, Draft forced, batch capped at " + batchLimit);
      alertMsg += "✅ Emails will go to REAL lead addresses\n";
      alertMsg += "✅ No [TEST] prefix on subject lines\n";
    }
    
    alertMsg += "✅ Drafts will appear in Gmail (NOT auto-sent)\n";
    alertMsg += "✅ Batch is capped at " + batchLimit + " leads\n\n";
    alertMsg += "Review the Gmail drafts folder to see exactly what the leads would receive.";
    
    if (!isHourly) {
      SpreadsheetApp.getUi().alert(alertTitle, alertMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
  // ─────────────────────────────────────────────────────────────────────────

  if (isHourly && senderSource === "ReadyTab") {
    return processReadyTabHourly(config, ss, batchLimit, testRecipient, outreachMode);
  }
  
  var startRow = 2;
  var endRow = lastRow;
  
  if (selectedOnly && !isHourly) {
    var activeRange = leadsSheet.getActiveRange();
    if (!activeRange) {
      SpreadsheetApp.getUi().alert("No Selection", "Please select the rows you want to process.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    startRow = activeRange.getRow();
    endRow = startRow + activeRange.getNumRows() - 1;
    
    if (startRow < 2) startRow = 2; // skip header
    if (endRow > lastRow) endRow = lastRow;
    
    if (startRow > endRow) {
      SpreadsheetApp.getUi().alert("Invalid Selection", "Selected range contains no valid data rows.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }
  
  // Show account-selection popup before starting (manual runs only)
  var forcedAccount = null;
  if (!isHourly) {
    forcedAccount = promptAccountSelection();
  }

  for (var r = startRow; r <= endRow; r++) {
    // 5-minute timeout safety
    if (new Date().getTime() - startTime > 300000) {
      if (!isHourly) {
        SpreadsheetApp.getUi().alert("Timeout Warning", "Script has been running for 5 minutes. Stopping early to prevent Google timeout. Please run again to process remaining leads.", SpreadsheetApp.getUi().ButtonSet.OK);
      } else {
        Logger.log("Timeout Warning: Stopping early to prevent Google timeout.");
      }
      break;
    }
    
    if (successCount >= batchLimit) {
      if (!isHourly) {
        SpreadsheetApp.getUi().alert("Batch Limit Reached", "Processed the maximum limit of " + batchLimit + " emails in this run. Please run the menu action again to process the rest.", SpreadsheetApp.getUi().ButtonSet.OK);
      } else {
        Logger.log("Batch Limit Reached: processed " + batchLimit + " emails.");
      }
      break;
    }
    
    var validationStatus = leadsSheet.getRange(r, headersMap["Validation Status"]).getValue().toString().trim();
    var outreachStatus = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
    var scoreVal = leadsSheet.getRange(r, headersMap["Score"]).getValue().toString().trim();
    var lastSentAt = headersMap["Last Sent At"] ? leadsSheet.getRange(r, headersMap["Last Sent At"]).getValue().toString().trim() : "";
    
    // Gate 1: Validated (Ready), Outreach Status is Ready for outreach, and Not Sent
    if (validationStatus === "Ready" && outreachStatus === "Ready for outreach" && !lastSentAt) {
      processed++;
      
      var result = processSingleOutreach(leadsSheet, r, headersMap, config, outreachMode, testRecipient, forcedAccount);
      
      if (result.success) {
        successCount++;
      } else {
        errors.push("Row " + r + ": " + result.error);
      }
      
      // Sleep 3 seconds to avoid API spamming
      Utilities.sleep(3000);
    }
  }
  
  // Log the outreach run
  logRun(processed, 0, 0, "Outreach: processed=" + processed + ", success=" + successCount + ", errors=" + (errors.length > 0 ? errors.join("; ") : "None"));
  
  if (!isHourly) {
    var alertMsg = "Outreach generation completed!\n\n" +
                   "- Leads checked: " + processed + "\n" +
                   "- Emails successfully generated: " + successCount + "\n" +
                   "- Failures: " + errors.length;
                   
    if (DRY_RUN) {
      alertMsg += "\n\n(SAFETY: Script is in DRY_RUN mode. All emails were created as drafts with '[TEST DRAFT]'.)";
    } else if (testRecipient) {
      alertMsg += "\n\n(Note: All emails were redirected to test address: " + testRecipient + ")";
    }
    
    SpreadsheetApp.getUi().alert("Outreach Complete", alertMsg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    Logger.log("Outreach Complete. Checked: " + processed + ", Success: " + successCount + ", Failures: " + errors.length);
  }
}

/**
 * Generates 3-5 sample emails on the first few qualified leads for user review.
 * Does not send or create real drafts, just shows them in a popup dialog.
 */
function runOutreachTest() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  var config = getConfig();
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  
  var samples = [];
  var count = 0;
  
  for (var r = 2; r <= lastRow; r++) {
    if (count >= 5) break; // Limit to 5 samples
    
    var company = leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim();
    if (!company) continue;
    
    var rowOverride = headersMap["Send From Account"] ? leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim() : "";
    var defaultAccount = getConfigValue(config, "Default Send Account", "Account A").toString().trim();
    var selectedAccountName = rowOverride || defaultAccount;
    
    Logger.log("Generating test sample for row " + r + " (" + company + ")");
    var emailData = generatePersonalizedEmail(leadsSheet, r, headersMap, config, selectedAccountName);
    
    if (emailData.success) {
      samples.push(
        "<h3>Sample " + (count + 1) + ": " + company + "</h3>" +
        "<p><b>Subject:</b> " + emailData.subject + "</p>" +
        "<pre style='background:#f3f4f6;padding:10px;border-radius:4px;white-space:pre-wrap;font-family:monospace;'>" + emailData.body + "</pre>"
      );
      count++;
    }
  }
  
  if (samples.length === 0) {
    ui.alert("No Samples Generated", "Could not find any rows with valid company data to generate samples.", ui.ButtonSet.OK);
    return;
  }
  
  var htmlOutput = HtmlService.createHtmlOutput(
    "<div style='font-family:sans-serif;font-size:14px;color:#1f2937;'>" +
    "<h2>AI-Generated Outreach Samples</h2>" +
    "<p>Review the tone, personalization, and structure below. If you approve, we can proceed with generating drafts.</p>" +
    "<hr>" +
    samples.join("<hr>") +
    "</div>"
  ).setWidth(600).setHeight(500);
  
  ui.showModalDialog(htmlOutput, "Outreach Email Samples");
}

/**
 * Generates test drafts for the rows currently selected by the user.
 * Does not send or create real drafts, just shows them in a popup dialog.
 */
function previewSelectedDrafts() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  var activeRange = leadsSheet.getActiveRange();
  if (!activeRange) {
    ui.alert("Error", "Please select one or more rows to preview.", ui.ButtonSet.OK);
    return;
  }
  
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  var endRow = startRow + numRows - 1;
  var lastRow = leadsSheet.getLastRow();
  
  if (startRow < 2) startRow = 2; // skip header
  if (endRow > lastRow) endRow = lastRow;
  
  if (startRow > endRow) {
    ui.alert("Error", "Selected range contains no valid data rows.", ui.ButtonSet.OK);
    return;
  }
  
  var config = getConfig();
  var headersMap = getHeadersMap(leadsSheet);
  
  var samples = [];
  var count = 0;
  var maxPreview = 10; // Prevent UI freeze
  
  for (var r = startRow; r <= endRow; r++) {
    if (count >= maxPreview) break;
    
    var company = leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim();
    if (!company) continue;
    
    var rowOverride = headersMap["Send From Account"] ? leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim() : "";
    var defaultAccount = getConfigValue(config, "Default Send Account", "Account A").toString().trim();
    var selectedAccountName = rowOverride || defaultAccount;
    
    var emailData = generatePersonalizedEmail(leadsSheet, r, headersMap, config, selectedAccountName);
    if (emailData.success) {
      samples.push(
        "<h3>Row " + r + ": " + company + "</h3>" +
        "<p><b>Subject:</b> " + emailData.subject + "</p>" +
        "<pre style='background:#f3f4f6;padding:10px;border-radius:4px;white-space:pre-wrap;font-family:monospace;'>" + emailData.body + "</pre>"
      );
      count++;
    }
  }
  
  if (samples.length === 0) {
    ui.alert("No Samples Generated", "Could not find any selected rows with valid company data.", ui.ButtonSet.OK);
    return;
  }
  
  var htmlOutput = HtmlService.createHtmlOutput(
    "<div style='font-family:sans-serif;font-size:14px;color:#1f2937;'>" +
    "<h2>Selected Rows Outreach Preview</h2>" +
    "<p>Showing up to " + maxPreview + " selected rows.</p>" +
    "<hr>" +
    samples.join("<hr>") +
    "</div>"
  ).setWidth(600).setHeight(500);
  
  ui.showModalDialog(htmlOutput, "Selected Outreach Preview");
}

/**
 * Triggers the outreach pipeline but restricted only to rows the user has currently selected.
 */
function generateSelectedDrafts() {
  processOutreachInternal(false, true, false);
}

/**
 * Generates and saves a single lead outreach email.
 */
function processSingleOutreach(sheet, rowNumber, headersMap, config, outreachMode, testRecipient, forcedAccount) {
  try {
    // Resolve source-agnostic column mapping once for this call (Feature 1)
    var mapping = resolveColumnMapping(sheet, false);

    var originalEmail = getFieldValue(sheet, rowNumber, headersMap, mapping, "Email", "email");
    if (!originalEmail) {
      throw new Error("Missing email address.");
    }

    // 1. Determine sending account in priority order:
    // Priority 1: forcedAccount (user chose via popup)
    // Priority 2: Per-row Send From Account column
    // Priority 3: Score-Band Routing (Feature 5)
    // Priority 4: Default Send Account in Config
    var rowOverride = headersMap["Send From Account"] ? sheet.getRange(rowNumber, headersMap["Send From Account"]).getValue().toString().trim() : "";
    
    var scoreRange = headersMap["Score"] ? sheet.getRange(rowNumber, headersMap["Score"]).getValue() : "";
    var routedAccount = selectAccountByScore(scoreRange, config);
    if (routedAccount === "HOLD") {
      Logger.log("Row " + rowNumber + " on HOLD due to Score Band Low Behavior.");
      sheet.getRange(rowNumber, headersMap["Outreach Status"]).setValue("On Hold (Low Score)");
      sheet.getRange(rowNumber, headersMap["Pipeline Stage"]).setValue("Held");
      return { success: false, error: "Held back due to low score routing rules." };
    }
    
    var defaultAccount = getConfigValue(config, "Default Send Account", "Account A").toString().trim();
    var selectedAccountName = forcedAccount || rowOverride || routedAccount || defaultAccount;
    
    // 2. Generate the personalized subject, opener, and closer using selected account
    var emailData = generatePersonalizedEmail(sheet, rowNumber, headersMap, config, selectedAccountName, mapping);
    if (!emailData.success) {
      throw new Error(emailData.error);
    }
    
    if (emailData.flagged) {
      sheet.getRange(rowNumber, headersMap["Outreach Status"]).setValue("Flagged - Contains Numbers");
      sheet.getRange(rowNumber, headersMap["Pipeline Stage"]).setValue("Needs Review");
      Logger.log("Row " + rowNumber + " flagged for containing specific figures/numbers in generated text.");
      return { success: false, error: "Flagged - Contains Numbers" };
    }
    
    var subject = emailData.subject;
    var emailBody = emailData.body;
    
    var senderEmail = getConfigValue(config, selectedAccountName + " Email", "");
    var senderLabel = getConfigValue(config, selectedAccountName + " Label", "");
    
    // 3. Handle recipient routing and dry-run safety
    var recipient = originalEmail;
    var finalBody = emailBody;
    var finalSubject = subject;
    
    if (DRY_RUN) {
      recipient = testRecipient || originalEmail;
      finalSubject = "[TEST DRAFT] " + subject;
      finalBody = "=== DRY RUN (Original Recipient: " + originalEmail + ") ===\n\n" + emailBody;
    } else if (testRecipient) {
      recipient = testRecipient;
      finalSubject = "[TEST] " + subject;
      finalBody = "=== TEST RUN (Original Recipient: " + originalEmail + ") ===\n\n" + emailBody;
    }
    
    // 4. Build HTML version and options
    var options = {
      htmlBody: finalBody.replace(/\n/g, "<br>"),
      from: senderEmail,
      name: senderLabel,
      replyTo: senderEmail
    };
    
    // 5. Always route to "Ready to Send" tab for manual review, regardless of Outreach Mode
    var statusVal = "";
    var readySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ready to Send");
    if (readySheet) {
        var draftId = Utilities.getUuid();
        // Score the RAW draft body before any DRY_RUN / TEST prefix is applied,
        // so the quality score reflects actual email content.
        var qualityScore = scoreDraftQuality(emailBody, config);
        var threshold = parseFloat(getConfigValue(config, "Draft Quality Threshold", "7"));
        var isReady = (!isNaN(qualityScore) && qualityScore >= threshold);
        
        var aVals = readySheet.getRange("A:A").getValues();
        var insertRow = 1;
        for (var i = 0; i < aVals.length; i++) {
          if (aVals[i][0] === "") {
            insertRow = i + 1;
            break;
          }
        }
        if (insertRow === 1) insertRow = aVals.length + 1;
        
        readySheet.getRange(insertRow, 1, 1, 15).setValues([[
          draftId,
          new Date(),
          getFieldValue(sheet, rowNumber, headersMap, mapping, "Company", "company"),
          getFieldValue(sheet, rowNumber, headersMap, mapping, "First Name", "first_name"),
          getFieldValue(sheet, rowNumber, headersMap, mapping, "Last Name", "last_name"),
          originalEmail,        // always store the real recipient, not the test-redirected one
          selectedAccountName,
          scoreRange,           // the lead score (value read earlier)
          emailData.templateId || "1", // The Template ID used
          qualityScore,
          "",                   // 2nd Time AI Score (blank initially)
          "",                   // AI Verification Notes (blank initially)
          subject,              // raw subject (no TEST prefix)
          emailBody,            // raw body (no DRY_RUN wrapper)
          isReady
        ]]);
        statusVal = "Draft in Ready Tab";
      } else {
        var apiRes = sendOrDraftViaAPI_(senderEmail, senderLabel, recipient, finalSubject, finalBody, true);
        if (!apiRes.success) throw new Error("API Draft Failed: " + apiRes.error);
        statusVal = "Draft Created";
      }
      sheet.getRange(rowNumber, headersMap["Pipeline Stage"]).setValue("Draft Created");
      if (headersMap["Send From Account"]) {
        sheet.getRange(rowNumber, headersMap["Send From Account"]).setValue(selectedAccountName);
      }
    
    // Lock 'Send From Account' and 'Email' cells so they can't be

    // accidentally changed after the email has been drafted/sent.
    if (headersMap["Send From Account"]) {
      sheet.getRange(rowNumber, headersMap["Send From Account"]).setValue(selectedAccountName);
      lockCellAfterSend(sheet, rowNumber, headersMap["Send From Account"], "account: " + selectedAccountName);
    }
    if (headersMap["Email"]) {
      lockCellAfterSend(sheet, rowNumber, headersMap["Email"], "email: " + originalEmail);
    }

    sheet.getRange(rowNumber, headersMap["Outreach Status"]).setValue(statusVal);
    return { success: true };
    
  } catch (e) {
    var errorMsg = e.toString();
    Logger.log("Error processing outreach for row " + rowNumber + ": " + errorMsg);
    sheet.getRange(rowNumber, headersMap["Outreach Status"]).setValue("Outreach Error: " + errorMsg);
    return { success: false, error: errorMsg };
  }
}

/**
 * Generates the personalized opener and closer via LLM, and combines them
 * with the fixed introduction, middle section, and signature.
 * @param {object} mapping Result of resolveColumnMapping — used for source-agnostic field reads.
 */
function generatePersonalizedEmail(sheet, rowNumber, headersMap, config, selectedAccountName, mapping) {
  try {
    // Use canonical mapping where available (Feature 1 fix), fall back to headersMap for Apollo-specific fields
    var firstName  = (mapping ? getCanonical(sheet, rowNumber, mapping, "first_name")  : "") ||
                     (headersMap["First Name"] ? sheet.getRange(rowNumber, headersMap["First Name"]).getValue() : "") || "";
    var lastName   = (mapping ? getCanonical(sheet, rowNumber, mapping, "last_name")   : "") ||
                     (headersMap["Last Name"]  ? sheet.getRange(rowNumber, headersMap["Last Name"]).getValue()  : "") || "";
    var title      = (mapping ? getCanonical(sheet, rowNumber, mapping, "title")        : "") ||
                     (headersMap["Title"]      ? sheet.getRange(rowNumber, headersMap["Title"]).getValue()      : "") || "";
    var company    = (mapping ? getCanonical(sheet, rowNumber, mapping, "company")      : "") ||
                     (headersMap["Company"]    ? sheet.getRange(rowNumber, headersMap["Company"]).getValue()    : "") || "";
    var industry   = (mapping ? getCanonical(sheet, rowNumber, mapping, "industry")     : "") ||
                     (headersMap["Industry"]   ? sheet.getRange(rowNumber, headersMap["Industry"]).getValue()   : "") || "";
    var keywords   = headersMap["Keywords"]     ? sheet.getRange(rowNumber, headersMap["Keywords"]).getValue()     || "" : "";
    var website    = headersMap["Website"]      ? sheet.getRange(rowNumber, headersMap["Website"]).getValue()      || "" : "";
    var technologies = headersMap["Technologies"] ? sheet.getRange(rowNumber, headersMap["Technologies"]).getValue() || "" : "";
    
    var revenue           = (mapping ? getCanonical(sheet, rowNumber, mapping, "revenue")      : "") ||
                            (headersMap["Annual Revenue"] ? sheet.getRange(rowNumber, headersMap["Annual Revenue"]).getValue() : "") || "";
    var totalFunding      = headersMap["Total Funding"]        ? sheet.getRange(rowNumber, headersMap["Total Funding"]).getValue()        || "" : "";
    var latestFunding     = headersMap["Latest Funding"]       ? sheet.getRange(rowNumber, headersMap["Latest Funding"]).getValue()       || "" : "";
    var latestFundingAmount = headersMap["Latest Funding Amount"] ? sheet.getRange(rowNumber, headersMap["Latest Funding Amount"]).getValue() || "" : "";
    var lastRaisedAt      = (mapping ? getCanonical(sheet, rowNumber, mapping, "funding_date") : "") ||
                            (headersMap["Last Raised At"] ? sheet.getRange(rowNumber, headersMap["Last Raised At"]).getValue() : "") || "";
    
    // 1. Gather web search context if the existing columns are thin/generic
    var webSearchContext = "";
    var isContextThin = (industry.toString().length < 5 && keywords.toString().length < 5);
    
    if (isContextThin && company) {
      Logger.log("Context is thin for " + company + ". Performing web search...");
      webSearchContext = searchCompanyWeb(company, website);
    }
    
    // 2. Load templates & signature
    var signature = getConfigValue(config, selectedAccountName + " Signature", "Best regards,\nAyush\nButter Search");
    
    var styleReferenceSubject = "Top talent hiring at {Company}";
    var styleReferenceBody = "Hi {First Name},\n\nI'm Ayush from Butter Search - an executive recruitment firm founded by IIM Calcutta alumni (ex-Naukri, Alvarez & Marsal, PwC).\n\nI've been following {Company}'s impressive journey in {Industry/Keywords}. As you gear up for your next phase of growth, having the right set of people in place becomes critical.\n\nThat's where we come in - getting top talent connected with leading fintech and housing finance platforms, working directly with founders, CXOs and business leaders.\n\nWould you be open to a quick connect to explore how we can support your hiring needs?";
    
    var templatesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Templates");
    if (templatesSheet) {
      var rows = templatesSheet.getDataRange().getValues();
      var defaultStyleSubject = "";
      var defaultStyleBody = "";
      var firstMatchPool = [];
      var exactMatchPool = [];
      var finalTemplateId = "";
      
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString().trim() === "Style Reference") {
          var prefAccount = rows[i][1] ? rows[i][1].toString().trim() : "";
          
          var rowTemplates = [];
          var templateNum = 1;
          for (var c = 2; c < rows[i].length; c += 2) {
            if (rows[i][c] && rows[i][c+1]) {
              rowTemplates.push({
                subject: rows[i][c].toString(),
                body: rows[i][c+1].toString(),
                templateId: templateNum.toString()
              });
              templateNum++;
            }
          }
          
          if (rowTemplates.length === 0) continue;
          
          if (firstMatchPool.length === 0) firstMatchPool = rowTemplates;
          if (prefAccount === selectedAccountName || prefAccount === "") {
            exactMatchPool = exactMatchPool.concat(rowTemplates);
          }
        }
      }
      
      var matchingTemplates = exactMatchPool.length > 0 ? exactMatchPool : firstMatchPool;
      
      if (matchingTemplates.length > 0) {
        defaultStyleSubject = matchingTemplates[0].subject;
        defaultStyleBody = matchingTemplates[0].body;
        
        var props = PropertiesService.getScriptProperties();
        var propKey = "RECENT_TEMPLATES_" + selectedAccountName;
        var recentStr = props.getProperty(propKey);
        var recent = [];
        if (recentStr) {
          try { recent = JSON.parse(recentStr); } catch(e) {}
        }
        
        var maxHistory = Math.max(0, matchingTemplates.length - 1);
        
        var available = matchingTemplates.filter(function(t) {
          return recent.indexOf(t.subject) === -1;
        });
        
        if (available.length === 0) {
           available = matchingTemplates;
           recent = [];
        }
        
        var randomIdx = Math.floor(Math.random() * available.length);
        var chosen = available[randomIdx];
        
        styleReferenceSubject = chosen.subject;
        styleReferenceBody = chosen.body;
        finalTemplateId = chosen.templateId || "";
        
        recent.push(chosen.subject);
        if (recent.length > maxHistory) {
           recent.shift();
        }
        props.setProperty(propKey, JSON.stringify(recent));
        
        Logger.log("Selected template variation " + finalTemplateId + " for " + selectedAccountName);
      } else if (defaultStyleSubject) {
        styleReferenceSubject = defaultStyleSubject;
        styleReferenceBody = defaultStyleBody;
        finalTemplateId = "1";
      }
    }
    
    // 3. Formulate the LLM prompt for full email personalization
    var prompt = "You are a professional B2B copywriter. We need to generate a highly personalized, direct outreach email for a lead.\n\n" +
    "Here is the style and tone reference email we are modeling. Match its tone, sentence rhythm, and formality, but you MUST rewrite it to personalize it for this lead:\n" +
    "```\n" +
    "Subject: " + styleReferenceSubject + "\n\n" +
    styleReferenceBody + "\n" +
    "```\n\n" +
    "Lead Details:\n" +
    "- Lead Name: " + firstName + " " + lastName + "\n" +
    "- Title: " + title + "\n" +
    "- Company: " + company + "\n" +
    "- Industry/Keywords: " + industry + " / " + keywords + "\n" +
    "- Technologies: " + technologies + "\n" +
    "- Annual Revenue: " + revenue + "\n" +
    "- Total Funding: " + totalFunding + "\n" +
    "- Latest Funding: " + latestFunding + " (" + latestFundingAmount + ") raised in " + lastRaisedAt + "\n" +
    "- Web Search Context: " + webSearchContext + "\n\n" +
    "Global rules:\n" +
    "- Keep the EXACT same sender name, signature, and general structure as the reference email.\n" +
    "- Do not include specific figures, amounts, dates, employee counts, or numeric values. Reference signals qualitatively only (e.g. 'following your recent funding milestone').\n" +
    "- Never fabricate facts. If information is thin, keep it safe and general.\n" +
    "- Tone: direct, concise, professional. No fluff, no exclamation marks.\n\n" +
    "Your Tasks:\n" +
    "1. Replace {First Name} and {Company} with the actual details.\n" +
    "2. Personalize the middle of the email to reference what the company actually does, specifically and accurately, based on the details or search context.\n" +
    "3. Personalize the final CTA to reference the qualifying signal (recent funding event or revenue growth stage) naturally.\n\n" +
    "OUTPUT FORMAT:\n" +
    "Respond ONLY with a JSON object in this format (do not include markdown blocks or any other text):\n" +
    "{\n" +
    "  \"subject\": \"<Generated subject line>\",\n" +
    "  \"body\": \"<Full generated email body with all paragraphs, using \\n\\n for line breaks>\"\n" +
    "}";

    var modelName = config.model || "gemini-2.0-flash";
    var responseText = callFailoverModelForOutreach(modelName, prompt, config);

    var cleanedJson = cleanJsonResponseText(responseText);
    var result = JSON.parse(cleanedJson);

    if (!result.body || !result.subject) {
      throw new Error("AI response missing body or subject. Got: " + responseText);
    }

    var containsNumbers = /[\d₹\$%]/i.test(result.body);
               
    return {
      success: true,
      subject: result.subject,
      body: result.body,
      templateId: finalTemplateId,
      flagged: containsNumbers
    };
    
  } catch (e) {
    return {
      success: false,
      error: "Email generation error: " + e.toString()
    };
  }
}

/**
 * Searches Google/Serper for company background to assist email personalization.
 */
function searchCompanyWeb(company, website) {
  var props = PropertiesService.getScriptProperties();
  var serperApiKey = props.getProperty("SERPER_API_KEY");
  var googleApiKey = props.getProperty("GOOGLE_API_KEY");
  var cseId = props.getProperty("GOOGLE_CSE_ID");
  
  var query = '"' + company + '"';
  if (website) {
    var domain = website.replace(/^(https?:\/\/)?(www\.)?/, "").split("/")[0];
    query += ' site:' + domain;
  } else {
    query += ' "about us" OR "what we do"';
  }
  
  try {
    if (serperApiKey) {
      var url = "https://google.serper.dev/search";
      var payload = { "q": query, "num": 3 };
      var response = UrlFetchApp.fetch(url, {
        "method": "post",
        "contentType": "application/json",
        "headers": { "X-API-KEY": serperApiKey },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });
      if (response.getResponseCode() === 200) {
        var data = JSON.parse(response.getContentText());
        var snippets = (data.organic || []).map(function(item) {
          return item.title + ": " + item.snippet;
        });
        return snippets.join("\n");
      }
    } else if (googleApiKey && cseId) {
      var url = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(googleApiKey) +
                  "&cx=" + encodeURIComponent(cseId) +
                  "&q=" + encodeURIComponent(query) +
                  "&num=3";
      var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
      if (response.getResponseCode() === 200) {
        var data = JSON.parse(response.getContentText());
        var snippets = (data.items || []).map(function(item) {
          return item.title + ": " + item.snippet;
        });
        return snippets.join("\n");
      }
    }
  } catch (e) {
    Logger.log("Search error in outreach personalization: " + e.toString());
  }
  return "";
}

/**
 * Finds all drafts in Gmail addressed to leads marked as 'Draft Created',
 * sends them, and updates the sheet status to 'Email Sent'.
 */
function sendDraftsFromPipeline() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Send Drafts", 
    "This will find all emails currently marked as 'Draft Created' in your sheet and send them live from your Gmail. Are you sure you want to send them?", 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  
  var startRow = 2;
  var endRow = lastRow;
  var activeRange = leadsSheet.getActiveRange();
  
  if (activeRange) {
    var rStart = activeRange.getRow();
    var rEnd = rStart + activeRange.getNumRows() - 1;
    if (rStart > 1) {
      startRow = rStart;
      endRow = Math.min(rEnd, lastRow);
    }
  }
  
  ui.alert("Sending Selected...", "Please wait while the system checks your selected rows for drafts to send.", ui.ButtonSet.OK);
  
  var sentCount = 0;
  var notFoundCount = 0;
  var skippedQuotaCount = 0;
  var config = getConfig();
  
  for (var r = startRow; r <= endRow; r++) {
    var status = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
    if (status === "Draft Created" || status === "Draft Created (Dry Run)") {
      var email = leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim();
      var accountName = headersMap["Send From Account"] ? leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim() : "";
      if (!accountName) accountName = "Account A"; // Fallback if empty
      var threadIdColVal = headersMap["Thread Id"] ? leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim() : "";
      
      var senderEmail = getConfigValue(config, accountName + " Email", "").toString().trim();
      var draftId = "";
      
      // We expect the Thread Id column to store 'DRAFT:<id>' for unsent drafts
      if (threadIdColVal.indexOf("DRAFT:") === 0) {
        draftId = threadIdColVal.substring(6);
      }
      
      if (draftId && senderEmail) {
        // ── QUOTA GATE: check remaining allowance for this account ──────────────
        var remaining = getRemainingQuota(accountName);
        if (remaining <= 0) {
          Logger.log("sendDraftsFromPipeline: quota exhausted for " + accountName + ", skipping row " + r);
          skippedQuotaCount++;
          continue;
        }
        
        var sendRes = sendDraftByIdViaAPI_(senderEmail, draftId);
        
        if (sendRes.success) {
          sentCount++;
          
          // ── HOURLY RATE LIMIT: check before confirming send ──────────────────
          try {
            var hourlyCheck = checkHourlyLimit(accountName, "fresh");
            if (!hourlyCheck.ok) {
              SpreadsheetApp.getUi().alert("Hourly Limit Reached", hourlyCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
              Logger.log("sendDraftsFromPipeline: hourly limit hit for " + accountName + ". Stopping.");
              break;
            }
          } catch(he) { Logger.log("Hourly check error: " + he); }

          // ── QUOTA TRACKING: record the send against this account ────────────
          try { recordSend(accountName); } catch(qe) {
            Logger.log("sendDraftsFromPipeline: recordSend failed for " + accountName + ": " + qe);
          }
          try { recordHourlySend(accountName, "fresh"); } catch(he) {}
          
          // Update sheet — status, pipeline stage, sent-at timestamp
          leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Email Sent");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Sent");
          
          var now = new Date();
          leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(now);
          if (headersMap["Send Date"]) {
            leadsSheet.getRange(r, headersMap["Send Date"]).setValue(now);
          }
          
          // CRITICAL: write real Thread Id so Scan Inbox can track this lead
          var actualThreadId = "";
          if (headersMap["Thread Id"] && sendRes.threadId) {
            actualThreadId = sendRes.threadId;
            try { leadsSheet.getRange(r, headersMap["Thread Id"]).setValue(actualThreadId); } catch(te) {}
          }
          
          // Ensure Send From Account is stamped
          if (headersMap["Send From Account"]) {
            var existingAcc = leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim();
            if (!existingAcc) leadsSheet.getRange(r, headersMap["Send From Account"]).setValue(accountName);
          }
          
          try { logDailySend("Fresh", accountName, email, "(Sent via Pipeline Draft)", actualThreadId); } catch(e) {}
          
          // HISTORICAL METRICS: Log the fresh send
          if (typeof logMetricEvent === "function" && actualThreadId) {
            try { logMetricEvent(accountName, "Fresh_Sent", actualThreadId, email, now); } catch(me) {
              Logger.log("sendDraftsFromPipeline: metrics logging failed: " + me);
            }
          }
          
          // Auto-schedule follow-up sequence
          if (typeof scheduleFollowups === "function") {
            try { scheduleFollowups({ row: r }); } catch(se) {}
          }
        } else {
          Logger.log("Failed to send draft " + draftId + " for row " + r + ": " + sendRes.error);
          notFoundCount++;
        }
      } else {
        notFoundCount++;
      }
    }
  }
  
  var msg = "Sent " + sentCount + " drafts successfully!";
  if (skippedQuotaCount > 0) {
    msg += "\n⚠️ " + skippedQuotaCount + " draft(s) were skipped because the daily quota for their account is exhausted. They will remain as drafts until tomorrow.";
  }
  if (notFoundCount > 0) {
    msg += "\nCould not find Gmail API draft IDs for " + notFoundCount + " leads. They may have been created before this update, or sent manually.";
  }
  
  if (typeof updateDailyForecast === "function") updateDailyForecast();
  
  ui.alert("Sending Complete", msg, ui.ButtonSet.OK);
}

/**
 * Scans sent emails for replies and automatically sends follow-up emails
 * to leads who haven't replied after N days.
 */
function detectRepliesAndFollowUp() {
  var ui = SpreadsheetApp.getUi();
  var result = detectRepliesAndFollowUpInternal(false);
  if (result && result.error) {
    ui.alert("Error", result.error, ui.ButtonSet.OK);
    return;
  }
  ui.alert(
    "Follow-up Process Complete",
    "Scan complete!\n\n" +
    "- New replies detected: " + result.repliesDetected + "\n" +
    "- Follow-up emails sent/drafted: " + result.followUpsSent,
    ui.ButtonSet.OK
  );
}

/**
 * Headless entry point for the daily follow-up trigger. Never calls the UI.
 */
function runFollowUpDaily() {
  var result = detectRepliesAndFollowUpInternal(true);
  Logger.log(
    "Daily follow-up run complete. Replies detected: " + (result.repliesDetected || 0) +
    ", Follow-ups sent/drafted: " + (result.followUpsSent || 0) +
    (result.error ? (", Error: " + result.error) : "")
  );
}

/**
 * Core reply-detection and follow-up logic. Safe to run with or without a UI.
 *
 * @param {boolean} isHeadless True when invoked from a time-based trigger (suppresses UI).
 * @return {object} { repliesDetected, followUpsSent } on success, or { error } on failure.
 */
function detectRepliesAndFollowUpInternal(isHeadless) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");

  if (!leadsSheet) {
    return { error: "Leads sheet not found." };
  }

  var config = getConfig();
  var headersMap = getHeadersMap(leadsSheet);

  // Verify required columns exist
  var required = [
    "First Name", "Company", "Email", "Pipeline Stage", "Outreach Status",
    "Last Sent At", "Replied", "Follow-up Status"
  ];
  for (var i = 0; i < required.length; i++) {
    if (!headersMap[required[i]]) {
      return { error: "Missing required column: '" + required[i] + "'. Please run Setup Sheets." };
    }
  }
  
  var lastRow = leadsSheet.getLastRow();
  var myEmail = ""; // Fallback not used when multi-mailbox is configured properly
  var delayDays = parseInt(getConfigValue(config, "Follow-up Delay (Days)", "3")) || 3;
  var outreachMode = getConfigValue(config, "Outreach Mode", "Draft");
  var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();
  
  // Load follow-up templates from Templates sheet
  var defaultFollowUp = {
    subject: "Re: Top talent hiring at {Company}",
    body: "Hi {First Name},\n\nI wanted to quickly follow up on my previous email. I know you're busy, but I'd love to see if you have 5 minutes for a quick chat about supporting your hiring needs at {Company}.\n\nBest,\nAyush"
  };
  var followUpTemplates = {};
  
  var templatesSheet = ss.getSheetByName("Templates");
  if (templatesSheet) {
    var rows = templatesSheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0].toString().trim() === "Follow-up Template") {
        var prefAccount = rows[i][3] ? rows[i][3].toString().trim() : "";
        var subj = rows[i][1] || defaultFollowUp.subject;
        var bdy = rows[i][2] || defaultFollowUp.body;
        
        if (prefAccount) {
          followUpTemplates[prefAccount] = { subject: subj, body: bdy };
        }
        
        if (!followUpTemplates["__default__"]) {
          followUpTemplates["__default__"] = { subject: subj, body: bdy };
        }
      }
    }
  }
  
  var repliesDetected = 0;
  var followUpsSent = 0;
  
  for (var r = 2; r <= lastRow; r++) {
    var email = leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim();
    if (!email) continue;
    
    var pipelineStage = leadsSheet.getRange(r, headersMap["Pipeline Stage"]).getValue().toString().trim();
    var outreachStatus = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
    var repliedStatus = leadsSheet.getRange(r, headersMap["Replied"]).getValue().toString().trim();
    var firstName = leadsSheet.getRange(r, headersMap["First Name"]).getValue().toString().trim();
    var company = leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim();

    // 0. Auto-detect if Draft was manually sent (Phase Upgrade)
    if (pipelineStage === "Draft Created" || outreachStatus.indexOf("Draft Created") !== -1) {
      var sentSearchQuery = "to:" + email + " in:sent newer_than:14d";
      var sentThreads = GmailApp.search(sentSearchQuery);
      if (sentThreads.length > 0) {
        var messages = sentThreads[0].getMessages();
        var lastMsg = messages[messages.length - 1];
        
        leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Email Sent");
        leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Sent");
        leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(lastMsg.getDate());
        
        // Update local variables so reply check can happen immediately
        pipelineStage = "Sent";
        outreachStatus = "Email Sent";
      }
    }
    
    // 1. Check for replies if the email was sent
    if (repliedStatus !== "Yes" && (pipelineStage === "Sent" || outreachStatus === "Email Sent")) {
      // Search for threads involving this email address
      var searchQuery = email;
      var threads = GmailApp.search(searchQuery);
      var hasReplied = false;
      
      for (var t = 0; t < threads.length; t++) {
        var messages = threads[t].getMessages();
        for (var m = 0; m < messages.length; m++) {
          var from = messages[m].getFrom().toLowerCase();
          var subject = messages[m].getSubject() || "";
          var isTestMessage = (subject.indexOf("[TEST]") !== -1 || subject.indexOf("[TEST DRAFT]") !== -1);
          var expectedFrom = (isTestMessage && testRecipient) ? testRecipient.toLowerCase() : email.toLowerCase();
          
          // A valid reply must come from the target (or test recipient) and NOT from our own email address.
          // This prevents our own follow-ups or identically named test drafts from being counted as replies.
          if (from.indexOf(expectedFrom) !== -1 && from.indexOf(myEmail.toLowerCase()) === -1) {
            hasReplied = true;
            break;
          }
        } // Close the inner loop over messages
        if (hasReplied) break;
      }
      
      if (hasReplied) {
        leadsSheet.getRange(r, headersMap["Replied"]).setValue("Yes");
        leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Replied");
        leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Replied");
        repliesDetected++;
        continue; // Skip follow-up if they replied
      }
    }
    
    // 2. Check if eligible for follow-up
    var lastSentVal = leadsSheet.getRange(r, headersMap["Last Sent At"]).getValue();
    var followUpSentVal = leadsSheet.getRange(r, headersMap["Follow-up Status"]).getValue().toString().trim();
    
    if (
      (pipelineStage === "Sent" || outreachStatus === "Email Sent") &&
      repliedStatus !== "Yes" &&
      followUpSentVal !== "Sent" &&
      followUpSentVal.indexOf("Follow-up Needed") === -1 &&
      lastSentVal
    ) {
      var lastSentDate = new Date(lastSentVal);
      var now = new Date();
      var diffTime = Math.abs(now - lastSentDate);
      var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      if (diffDays >= delayDays) {
        // Sender Identity
        var sentFromAccount = headersMap["Send From Account"] ? leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim() : "";
        var selectedAccountName = sentFromAccount || getConfigValue(config, "Default Send Account", "Account A").toString().trim();
        var senderEmail = getConfigValue(config, selectedAccountName + " Email", "");
        var senderLabel = getConfigValue(config, selectedAccountName + " Label", "");
        
        // Select correct template
        var templateToUse = followUpTemplates[selectedAccountName] || followUpTemplates["__default__"] || defaultFollowUp;
        
        var subject = templateToUse.subject.replace(/{Company}/g, company);
        var body = templateToUse.body
          .replace(/{First Name}/g, firstName)
          .replace(/{Company}/g, company);
          
        var recipient = email;
        var finalSubject = subject;
        var finalBody = body;
        
        if (DRY_RUN) {
          recipient = testRecipient || email;
          finalSubject = "[TEST DRAFT] " + subject;
          finalBody = "=== DRY RUN FOLLOW-UP (Original Recipient: " + email + ") ===\n\n" + body;
        } else if (testRecipient) {
          recipient = testRecipient;
          finalSubject = "[TEST] " + subject;
          finalBody = "=== TEST FOLLOW-UP (Original Recipient: " + email + ") ===\n\n" + body;
        }
        
        var options = {
          htmlBody: finalBody.replace(/\n/g, "<br>"),
          from: senderEmail,
          name: senderLabel,
          replyTo: senderEmail
        };
        
        // Threading
        var threadId = headersMap["Thread Id"] ? leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim() : "";
        var thread = threadId ? GmailApp.getThreadById(threadId) : null;
        var lastMsg = null;
        if (thread) {
          var msgs = thread.getMessages();
          if (msgs.length > 0) lastMsg = msgs[msgs.length - 1];
        }
        
        // Send or Draft
        if (outreachMode.toLowerCase() === "send" && !DRY_RUN) {
          if (lastMsg) {
            lastMsg.reply(finalBody, options);
          } else {
            GmailApp.sendEmail(recipient, finalSubject, finalBody, options);
          }
          leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("Sent");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Follow-up Sent");
          leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(new Date());
        } else {
          if (lastMsg) {
            lastMsg.createDraftReply(finalBody, options);
          } else {
            GmailApp.createDraft(recipient, finalSubject, finalBody, options);
          }
          leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("Follow-up Needed (Draft Created)");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Follow-up Drafted");
        }
        
        followUpsSent++;
        Utilities.sleep(2000); // Avoid rate limits
      }
    }
  }
  
  return { repliesDetected: repliesDetected, followUpsSent: followUpsSent };
}

/**
 * Builds the follow-up email body and subject for a specific row.
 * Uses an LLM to generate a personalised middle paragraph based on real company/lead facts.
 * @return {object} { success, subject, body, email, company, senderEmail, senderLabel, selectedAccountName }
 *                  or { success: false, reason }
 */
function buildFollowUpForRow(leadsSheet, r, headersMap, config, followUpTemplates, defaultFollowUp, forcedAccount) {
  var email = leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim();
  if (!email) return { success: false, reason: "No email address" };

  var pipelineStage  = leadsSheet.getRange(r, headersMap["Pipeline Stage"]).getValue().toString().trim();
  var outreachStatus = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
  var repliedStatus  = leadsSheet.getRange(r, headersMap["Replied"]).getValue().toString().trim();
  var firstName      = leadsSheet.getRange(r, headersMap["First Name"]).getValue().toString().trim();
  var company        = leadsSheet.getRange(r, headersMap["Company"]).getValue().toString().trim();

  // Only generate follow-ups for rows where an email was sent and no reply yet
  var emailWasSent = (pipelineStage === "Sent" || pipelineStage === "Follow-up Drafted" ||
                      outreachStatus === "Email Sent");
  if (!emailWasSent) return { success: false, reason: "Email not yet sent (Stage: " + pipelineStage + ")" };
  if (repliedStatus === "Yes") return { success: false, reason: "Lead already replied" };

  // Sender identity — priority: forced (popup) > Send From Account column > Default
  var sentFromAccount     = headersMap["Send From Account"] ? leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim() : "";
  var selectedAccountName = forcedAccount || sentFromAccount || getConfigValue(config, "Default Send Account", "Account A").toString().trim();
  var senderEmail         = getConfigValue(config, selectedAccountName + " Email", "");
  var senderLabel         = getConfigValue(config, selectedAccountName + " Label", "");
  var senderFirstName     = senderLabel ? senderLabel.split(" ")[0] : "Ayush";
  var signature           = getConfigValue(config, selectedAccountName + " Signature", "Best,\n" + senderFirstName + "\nButter Search");

  // ── Rich company context from the lead row ──────────────────────────────
  function safeGet(col) {
    return headersMap[col] ? leadsSheet.getRange(r, headersMap[col]).getValue().toString().trim() : "";
  }
  var title         = safeGet("Title");
  var industry      = safeGet("Industry");
  var revenue       = safeGet("Annual Revenue");
  var employees     = safeGet("# Employees");
  var totalFunding  = safeGet("Total Funding");
  var latestFunding = safeGet("Latest Funding");
  var fundingAmt    = safeGet("Latest Funding Amount");
  var lastRaisedAt  = safeGet("Last Raised At");
  var hiringStatus  = safeGet("Hiring Status");
  var technologies  = safeGet("Technologies");
  var country       = safeGet("Country");

  // ── Generate personalised middle paragraph via LLM ──────────────────────
  var aiInsight = "";
  try {
    var llmPrompt =
      "You are writing a follow-up paragraph for a B2B recruiting/staffing agency called Butter Search.\n" +
      "The previous cold email reached out about supporting this company's hiring needs.\n\n" +
      "Lead details:\n" +
      "  - Name: " + (firstName || "the prospect") + "\n" +
      "  - Title: " + (title || "N/A") + "\n" +
      "  - Company: " + (company || "N/A") + "\n" +
      "  - Industry: " + (industry || "N/A") + "\n" +
      "  - Country: " + (country || "N/A") + "\n" +
      "  - Annual Revenue: " + (revenue || "N/A") + "\n" +
      "  - Employees: " + (employees || "N/A") + "\n" +
      "  - Total Funding: " + (totalFunding || "N/A") + "\n" +
      "  - Latest Funding Round: " + (latestFunding || "N/A") + "\n" +
      "  - Latest Funding Amount: " + (fundingAmt || "N/A") + "\n" +
      "  - Last Raised At: " + (lastRaisedAt || "N/A") + "\n" +
      "  - Hiring Intent: " + (hiringStatus || "N/A") + "\n" +
      "  - Technologies: " + (technologies || "N/A") + "\n\n" +
      "Write ONLY a single, punchy paragraph (maximum 2 sentences) to go in the MIDDLE of the follow-up email.\n" +
      "Requirements:\n" +
      "  1. Reference ONE specific, real fact about the company (e.g., funding stage, revenue scale, or headcount).\n" +
      "  2. Tie it naturally to how Butter Search (a specialized talent network) can help them attract top candidates right now.\n" +
      "  3. Tone must be warm, direct, human, and highly concise. AVOID formal, stiff, or lengthy corporate speak (do NOT say 'As a Talent Acquisition Lead...'). Keep it punchy.\n" +
      "  4. Do NOT include greetings, opening lines, closing lines, or sign-offs. ONLY the middle paragraph.\n" +
      "  5. ABSOLUTELY NO JSON. Return plain text only.\n\n" +
      "Example output style:\n" +
      "We'd love to support the hiring efforts at " + (company || "your company") + " as you continue to scale. " +
      "With our strong network across high-impact talent pools, we're confident we can bring meaningful value to your growth plans.";

    var rawText = callFailoverModelForOutreach(
      getConfigValue(config, "Gemini Model", "gemini-2.0-flash"),
      llmPrompt,
      config
    );

    var textStr = rawText.toString().trim();

    // The LLM sometimes wraps the paragraph in a markdown code fence and/or JSON
    // despite instructions. Strip the fence first, then unwrap the JSON if present.
    textStr = textStr.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim();

    if (textStr.charAt(0) === "{") {
      try {
        var parsed = JSON.parse(textStr);
        // Prefer known paragraph keys, then fall back to the first non-empty
        // string value so ANY wrapper key (e.g. follow_up_paragraph) is unwrapped.
        var extracted = parsed.follow_up_paragraph || parsed.followUpParagraph ||
                        parsed.paragraph || parsed.response || parsed.message ||
                        parsed.text || parsed.body || parsed.content;
        if (!extracted) {
          for (var key in parsed) {
            if (parsed.hasOwnProperty(key) && typeof parsed[key] === "string" && parsed[key].trim()) {
              extracted = parsed[key];
              break;
            }
          }
        }
        if (extracted) textStr = extracted.toString().trim();
      } catch (e) {
        // Not valid JSON — leave textStr as-is.
      }
    }

    // Strip any surrounding quotes/backticks left over.
    aiInsight = textStr.replace(/^["'`]+|["'`]+$/g, "").trim();
  } catch (e) {
    Logger.log("Follow-up LLM personalisation failed for row " + r + ": " + e);
    // Graceful fallback — use a sensible generic paragraph
    aiInsight =
      "We'd love to support the hiring efforts at " + company +
      " as you continue to grow. With our strong network across high-impact talent pools, " +
      "we're confident we can bring meaningful value to your team-building plans.";
  }

  // ── Assemble the final email body ────────────────────────────────────────
  var templateToUse = followUpTemplates[selectedAccountName] || followUpTemplates["__default__"] || defaultFollowUp;
  var subject = templateToUse.subject.replace(/{Company}/g, company);

  var body;
  if (templateToUse.body && templateToUse.body.indexOf("{AI_INSIGHT}") !== -1) {
    // Template has the AI placeholder — substitute it
    body = templateToUse.body
      .replace(/{First Name}/g, firstName || "there")
      .replace(/{Company}/g,    company)
      .replace(/{AI_INSIGHT}/g, aiInsight);
  } else {
    // Build the full structured body — matching the approved sample template
    body =
      "Hi " + (firstName || "there") + ",\n\n" +
      "Just following up on my earlier note \u2014 completely understand things get busy.\n\n" +
      aiInsight + "\n\n" +
      "Happy to connect for a quick call to explore how we might partner effectively.\n\n" +
      "Looking forward to hearing from you.";
  }

  // Append signature only if not already present
  var signatureFirstLine = signature.split("\n")[0].trim();
  if (signatureFirstLine && body.indexOf(signatureFirstLine) === -1) {
    body = body + "\n\n" + signature;
  }

  return {
    success:             true,
    email:               email,
    company:             company,
    firstName:           firstName,
    subject:             subject,
    body:                body,
    senderEmail:         senderEmail,
    senderLabel:         senderLabel,
    selectedAccountName: selectedAccountName
  };
}

/**
 * Loads follow-up templates from the Templates sheet.
 * Templates can optionally include {AI_INSIGHT} as a placeholder for the
 * LLM-generated personalised paragraph; if omitted, the full structural frame
 * is built automatically in buildFollowUpForRow.
 */
function loadFollowUpTemplates(ss) {
  var defaultFollowUp = {
    subject: "Re: Partnering with {Company} on hiring",
    // body intentionally left empty — buildFollowUpForRow builds the full frame
    body: ""
  };
  var followUpTemplates = {};

  var templatesSheet = ss.getSheetByName("Templates");
  if (templatesSheet) {
    var rows = templatesSheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0].toString().trim() === "Follow-up Template") {
        var prefAccount = rows[i][3] ? rows[i][3].toString().trim() : "";
        var subj = rows[i][1] || defaultFollowUp.subject;
        var bdy  = rows[i][2] || defaultFollowUp.body;
        if (prefAccount) followUpTemplates[prefAccount] = { subject: subj, body: bdy };
        if (!followUpTemplates["__default__"]) followUpTemplates["__default__"] = { subject: subj, body: bdy };
      }
    }
  }
  return { templates: followUpTemplates, defaultFollowUp: defaultFollowUp };
}

/**
 * Previews follow-up email content for the currently selected rows in the Leads sheet.
 * Shows a popup dialog with the draft subject and body — no email is sent or created.
 */
function previewFollowUpSelected() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
  }

  var activeRange = leadsSheet.getActiveRange();
  if (!activeRange) {
    ui.alert("No Selection", "Please select one or more rows on the Leads sheet to preview.", ui.ButtonSet.OK);
    return;
  }

  var startRow = Math.max(activeRange.getRow(), 2);
  var endRow = Math.min(startRow + activeRange.getNumRows() - 1, leadsSheet.getLastRow());

  var config = getConfig();
  var headersMap = getHeadersMap(leadsSheet);
  var loaded = loadFollowUpTemplates(ss);

  var samples = [];
  var maxPreview = 10;

  for (var r = startRow; r <= endRow && samples.length < maxPreview; r++) {
    var data = buildFollowUpForRow(leadsSheet, r, headersMap, config, loaded.templates, loaded.defaultFollowUp);
    if (data.success) {
      samples.push(
        "<h3>Row " + r + ": " + data.company + " (" + data.email + ")</h3>" +
        "<p><b>From:</b> " + data.senderLabel + " &lt;" + data.senderEmail + "&gt;</p>" +
        "<p><b>Subject:</b> " + data.subject + "</p>" +
        "<pre style='background:#f3f4f6;padding:10px;border-radius:4px;white-space:pre-wrap;font-family:monospace;font-size:12px;'>" + data.body + "</pre>"
      );
    } else {
      samples.push(
        "<h3>Row " + r + " — Skipped</h3>" +
        "<p style='color:#6b7280;'><i>Reason: " + data.reason + "</i></p>"
      );
    }
  }

  if (samples.length === 0) {
    ui.alert("Nothing to Preview", "No valid rows found in your selection.", ui.ButtonSet.OK);
    return;
  }

  var html = HtmlService.createHtmlOutput(
    "<div style='font-family:sans-serif;font-size:14px;color:#1f2937;'>" +
    "<h2>Follow-up Email Preview</h2>" +
    "<p style='color:#6b7280;'>Showing previews for up to " + maxPreview + " selected rows. No emails will be sent.</p>" +
    "<hr>" +
    samples.join("<hr>") +
    "</div>"
  ).setWidth(640).setHeight(520);

  ui.showModalDialog(html, "Follow-up Preview");
}

/**
 * Immediately sends/drafts follow-up emails for the currently selected rows in the Leads sheet.
 * Bypasses the day-delay requirement — useful for manual, on-demand follow-ups.
 */
function sendFollowUpForSelected() {
  _sendFollowUpManual(true);
}

/**
 * Immediately sends/drafts follow-up emails for ALL eligible rows (sent but not replied).
 * Bypasses the day-delay requirement.
 */
function sendFollowUpForAll() {
  _sendFollowUpManual(false);
}

/**
 * Core helper for manual follow-up sending. If selectedOnly is true, processes only
 * the user's selected rows; otherwise processes all rows.
 *
 * @param {boolean} selectedOnly True to process only the active selection.
 */
function _sendFollowUpManual(selectedOnly) {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");

  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found. Please run 'Setup Sheets' first.", ui.ButtonSet.OK);
    return;
  }

  var headersMap = getHeadersMap(leadsSheet);
  var required = ["Email", "First Name", "Company", "Pipeline Stage", "Outreach Status", "Replied", "Follow-up Status", "Last Sent At"];
  for (var i = 0; i < required.length; i++) {
    if (!headersMap[required[i]]) {
      ui.alert("Error", "Missing required column: '" + required[i] + "'. Please run Setup Sheets.", ui.ButtonSet.OK);
      return;
    }
  }

  var startRow = 2;
  var endRow = leadsSheet.getLastRow();

  if (selectedOnly) {
    var activeRange = leadsSheet.getActiveRange();
    if (!activeRange) {
      ui.alert("No Selection", "Please select one or more rows on the Leads sheet.", ui.ButtonSet.OK);
      return;
    }
    startRow = Math.max(activeRange.getRow(), 2);
    endRow = Math.min(startRow + activeRange.getNumRows() - 1, leadsSheet.getLastRow());
  }

  var config = getConfig();
  var loaded = loadFollowUpTemplates(ss);
  var outreachMode = getConfigValue(config, "Outreach Mode", "Draft");
  var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();

  // Ask user which account to send from before processing
  var forcedAccount = promptAccountSelection();

  var sentCount = 0;
  var skippedCount = 0;
  var errors = [];

  for (var r = startRow; r <= endRow; r++) {
    var data = buildFollowUpForRow(leadsSheet, r, headersMap, config, loaded.templates, loaded.defaultFollowUp, forcedAccount);

    if (!data.success) {
      skippedCount++;
      continue;
    }

    try {
      var recipient = data.email;
      var finalSubject = data.subject;
      var finalBody = data.body;

      if (DRY_RUN) {
        recipient = testRecipient || data.email;
        finalSubject = "[TEST DRAFT] " + data.subject;
        finalBody = "=== DRY RUN FOLLOW-UP (Original Recipient: " + data.email + ") ===\n\n" + data.body;
      } else if (testRecipient) {
        recipient = testRecipient;
        finalSubject = "[TEST] " + data.subject;
        finalBody = "=== TEST FOLLOW-UP (Original Recipient: " + data.email + ") ===\n\n" + data.body;
      }

      var options = {
        htmlBody: finalBody.replace(/\n/g, "<br>"),
        from: data.senderEmail,
        name: data.senderLabel,
        replyTo: data.senderEmail
      };

      // Attempt to thread the reply into the original email thread
      var threadId = headersMap["Thread Id"] ? leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim() : "";
      var thread = threadId ? GmailApp.getThreadById(threadId) : null;
      var lastMsg = null;
      if (thread) {
        var msgs = thread.getMessages();
        if (msgs.length > 0) lastMsg = msgs[msgs.length - 1];
      }

      if (outreachMode.toLowerCase() === "send" && !DRY_RUN) {
        if (lastMsg) {
          lastMsg.reply(finalBody, options);
        } else {
          GmailApp.sendEmail(recipient, finalSubject, finalBody, options);
        }
        leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("Sent");
        leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Follow-up Sent");
        leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(new Date());
      } else {
        if (lastMsg) {
          lastMsg.createDraftReply(finalBody, options);
        } else {
          GmailApp.createDraft(recipient, finalSubject, finalBody, options);
        }
        leadsSheet.getRange(r, headersMap["Follow-up Status"]).setValue("Follow-up Needed (Draft Created)");
        leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Follow-up Drafted");
      }

      // Freeze 'Send From Account' and 'Email' cells so they can't be accidentally changed.
      if (headersMap["Send From Account"]) {
        leadsSheet.getRange(r, headersMap["Send From Account"]).setValue(data.selectedAccountName);
        lockCellAfterSend(leadsSheet, r, headersMap["Send From Account"], "account: " + data.selectedAccountName);
      }
      if (headersMap["Email"]) {
        lockCellAfterSend(leadsSheet, r, headersMap["Email"], "email: " + data.email);
      }

      sentCount++;
      Utilities.sleep(1500);
    } catch (e) {
      errors.push("Row " + r + ": " + e.toString());
    }
  }

  var modeLabel = (outreachMode.toLowerCase() === "send" && !DRY_RUN) ? "sent" : "drafted";
  var msg = "Follow-up run complete!\n\n" +
            "- Emails " + modeLabel + ": " + sentCount + "\n" +
            "- Rows skipped (not eligible): " + skippedCount + "\n" +
            "- Errors: " + errors.length;

  if (DRY_RUN) {
    msg += "\n\n(SAFETY: DRY_RUN mode is ON. All follow-ups were drafted with '[TEST DRAFT]'.)";
  }
  if (errors.length > 0) {
    msg += "\n\nErrors:\n" + errors.join("\n");
  }

  ui.alert("Follow-up Complete", msg, ui.ButtonSet.OK);
}

/**
 * Set up the daily trigger that scans for replies and sends follow-ups automatically.
 */
function setupFollowUpTrigger() {
  var ui = SpreadsheetApp.getUi();
  clearFollowUpTriggerInternal();
  setConfigValue("Follow-up Active", true);

  ScriptApp.newTrigger("runFollowUpDaily")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  ui.alert(
    "Follow-up Trigger Activated",
    "The system will scan for replies once per day (around 9 AM) and send a follow-up to any lead that was emailed but hasn't replied after the configured 'Follow-up Delay (Days)'. 'Follow-up Active' is now true.",
    ui.ButtonSet.OK
  );
}

/**
 * Deactivate the daily follow-up trigger.
 */
function deactivateFollowUpTrigger() {
  var ui = SpreadsheetApp.getUi();
  clearFollowUpTriggerInternal();
  setConfigValue("Follow-up Active", false);
  ui.alert("Follow-up Trigger Deactivated", "The daily follow-up trigger has been removed. 'Follow-up Active' is now false.", ui.ButtonSet.OK);
}

/**
 * Helper to clear all existing daily follow-up triggers.
 */
function clearFollowUpTriggerInternal() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runFollowUpDaily") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Set up the 1-minute trigger for the Drip Engine.
 */
function setupDripEngineTrigger() {
  var ui = SpreadsheetApp.getUi();
  clearDripEngineTriggerInternal();
  setConfigValue("Sending Active", true);
  
  ScriptApp.newTrigger("runDripEngine")
    .timeBased()
    .everyMinutes(1)
    .create();
    
  ui.alert("Drip Engine Activated", "The Drip Engine has been set up successfully. 'Sending Active' is now true. It will check every minute and send 1 email only if the Drip Gap has passed.", ui.ButtonSet.OK);
}

/**
 * Deactivate the Drip Engine trigger.
 */
function deactivateDripEngineTrigger() {
  var ui = SpreadsheetApp.getUi();
  clearDripEngineTriggerInternal();
  setConfigValue("Sending Active", false);
  ui.alert("Drip Engine Deactivated", "The Drip Engine trigger has been deactivated. 'Sending Active' is now false.", ui.ButtonSet.OK);
}

/**
 * Helper to clear all existing Drip Engine triggers.
 */
function clearDripEngineTriggerInternal() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runDripEngine" || triggers[i].getHandlerFunction() === "runOutreachPipelineHourly") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Headless entry point for the Drip Engine (runs every 1 minute).
 */
function runDripEngine() {
  try {
    var config = getConfig();
    var gapMinutes = parseInt(getConfigValue(config, "Drip Gap (Minutes)", "5")) || 5;
    
    var lastDripStr = PropertiesService.getScriptProperties().getProperty("LAST_DRIP_TIME");
    var now = new Date();
    
    if (lastDripStr) {
      var lastDrip = new Date(parseInt(lastDripStr));
      var diffMinutes = (now.getTime() - lastDrip.getTime()) / 60000;
      if (diffMinutes < gapMinutes) {
        Logger.log("Drip Engine: " + gapMinutes + " minutes haven't passed yet. Skipping execution.");
        return;
      }
    }
    
    Logger.log("Drip Engine: Executing. Priority 1: Follow-ups. Priority 2: Ready Drafts. Priority 3: Pipeline Drafts.");
    
    // Priority 1: Try sending ONE due follow-up
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof sendOneDueFollowup === "function") {
      if (sendOneDueFollowup(config)) {
        PropertiesService.getScriptProperties().setProperty("LAST_DRIP_TIME", now.getTime().toString());
        if (typeof updateDailyForecast === "function") updateDailyForecast();
        return;
      }
    }
    
    // Priority 2: Try sending ONE approved draft from 'Ready to Send'
    if (sendOneReadyDraft_(config, ss)) {
      PropertiesService.getScriptProperties().setProperty("LAST_DRIP_TIME", now.getTime().toString());
      if (typeof updateDailyForecast === "function") updateDailyForecast();
      return;
    }
    
    // Priority 3: Try sending ONE pending draft from the Pipeline (Draft Created)
    if (sendOnePipelineDraft_(config, ss)) {
      PropertiesService.getScriptProperties().setProperty("LAST_DRIP_TIME", now.getTime().toString());
      if (typeof updateDailyForecast === "function") updateDailyForecast();
      return;
    }
    
    Logger.log("Drip Engine: Nothing to send right now.");
    
  } catch (e) {
    Logger.log("runDripEngine Error: " + e.toString());
  }
}

/**
 * Helper to retrieve configuration values with a fallback default.
 */
function getConfigValue(config, key, defaultValue) {
  if (config && config[key] !== undefined && config[key] !== null && config[key] !== "") {
    return config[key];
  }
  return defaultValue;
}

/**
 * Helper to call the LLM for outreach generation. Uses Gemini 2.0 Flash as primary,
 * with failover to other providers if needed.
 */
/**
 * Helper to call the LLM for outreach generation. Uses Gemini 2.0 Flash as primary,
 * with failover to other providers if needed.
 */
function callFailoverModelForOutreach(modelName, prompt, config) {
  var isGemini = (modelName.indexOf("gemini") !== -1);
  
  if (isGemini) {
    var geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!geminiKey) {
      throw new Error("GEMINI_API_KEY is not defined in Script Properties.");
    }
    if (modelName.indexOf("models/") === 0) {
      modelName = modelName.substring(7);
    }
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + geminiKey;
    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": {
        "responseMimeType": "application/json",
        "temperature": parseFloat(config.temperature) || 0.2
      }
    };
    
    var options = {
      "method": "post",
      "headers": { "Content-Type": "application/json" },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error("Gemini API error: " + response.getContentText());
    }
    
    var resObj = JSON.parse(response.getContentText());
    return resObj.candidates[0].content.parts[0].text;
  } else {
    // Fallback to failover providers (NVIDIA, Groq, Cerebras)
    var nvidiaKeys = getApiKeysList("NVIDIA_API_KEYS", "NVIDIA_API_KEY");
    var groqKeys = getApiKeysList("GROQ_API_KEYS", "GROQ_API_KEY");
    var cerebrasKeys = getApiKeysList("CEREBRAS_API_KEYS", "CEREBRAS_API_KEY");
    
    var providers = [
      {
        name: "NVIDIA",
        url: "https://integrate.api.nvidia.com/v1/chat/completions",
        keys: nvidiaKeys,
        defaultModel: "deepseek-ai/deepseek-r1"
      },
      {
        name: "Groq",
        url: "https://api.groq.com/openai/v1/chat/completions",
        keys: groqKeys,
        defaultModel: "llama-3.3-70b-versatile"
      },
      {
        name: "Cerebras",
        url: "https://api.cerebras.ai/v1/chat/completions",
        keys: cerebrasKeys,
        defaultModel: "zai-glm-4.7"
      }
    ];
    
    // Sort providers based on the configured model
    var primaryName = "Groq";
    var modelNameLower = modelName.toLowerCase();
    if (modelNameLower.indexOf("deepseek") !== -1 || modelNameLower.indexOf("nvidia") !== -1) {
      primaryName = "NVIDIA";
    } else if (modelNameLower.indexOf("cerebras") !== -1 || modelNameLower.indexOf("zai-glm") !== -1) {
      primaryName = "Cerebras";
    } else if (modelNameLower.indexOf("groq") !== -1 || modelNameLower.indexOf("llama") !== -1) {
      primaryName = "Groq";
    }
    
    providers.sort(function(a, b) {
      if (a.name === primaryName) return -1;
      if (b.name === primaryName) return 1;
      return 0;
    });
    
    var lastApiError = "";
    for (var i = 0; i < providers.length; i++) {
      var provider = providers[i];
      if (provider.keys.length === 0) {
        lastApiError += (lastApiError ? " | " : "") + "No keys for " + provider.name;
        continue;
      }
      
      var modelToUse = (provider.name === primaryName) ? modelName : provider.defaultModel;
      var shuffledKeys = shuffleArray(provider.keys);
      
      for (var k = 0; k < shuffledKeys.length; k++) {
        var key = shuffledKeys[k];
        try {
          return callOpenAiCompatibleApi(provider.url, modelToUse, key, prompt, config.temperature, true);
        } catch (err) {
          lastApiError += (lastApiError ? " | " : "") + provider.name + " Error: " + err.toString();
          Utilities.sleep(500);
        }
      }
    }
    
    throw new Error("All outreach LLM providers failed. Details: " + lastApiError);
  }
}
/**
 * Feature 8: Score the draft quality 1-10 using the secondary Gemini API Key if available.
 */
function scoreDraftQuality(draftBody, config) {
  var prompt = "You are a QA editor for outbound sales emails. Rate the following draft email on a scale of 1-10, where 10 is highly personalized, professional, and free of weird placeholders or AI hallucinations. 1 is robotic, contains placeholders like [Company Name], or uses weird formatting.\n\n" +
               "Draft Email:\n" + draftBody + "\n\n" +
               "Reply with ONLY a number from 1 to 10.";
               
  var modelName = getConfigValue(config, "Gemini Model", "gemini-2.0-flash").toString().trim();
  try {
    var response = callFailoverModelForOutreach(modelName, prompt, config);
    var numMatch = response.match(/\d+/);
    if (numMatch) {
      return parseInt(numMatch[0], 10);
    }
  } catch(e) {
    Logger.log("Draft quality scoring error: " + e.toString());
  }
  
  return 0; // default failure
}
function pushReadyDraftsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig();
  var batchLimit = parseInt(getConfigValue(config, "Emails Per Run", "10")) || 10;
  var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();
  var outreachMode = getConfigValue(config, "Outreach Mode", "Draft").toString().trim();
  
  processReadyTabHourly(config, ss, batchLimit, testRecipient, outreachMode);
  
  if (typeof updateDailyForecast === "function") updateDailyForecast();
  
  SpreadsheetApp.getUi().alert("Success", "Finished pushing approved drafts to Gmail.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Feature 8: Hourly sender pulling from the Ready to Send tab instead of the Leads tab.
 */
function processReadyTabHourly(config, ss, batchLimit, testRecipient, outreachMode) {
  Logger.log("processReadyTabHourly: Routing hourly execution to Ready to Send tab.");
  var readySheet = ss.getSheetByName("Ready to Send");
  if (!readySheet) {
    Logger.log("Ready to Send sheet not found.");
    return;
  }
  
  var lastRow = readySheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("Ready to Send sheet is empty.");
    return;
  }
  
  // Map headers
  var headersMap = getHeadersMap(readySheet);
  
  // Pre-load the Leads sheet for post-send status updates
  var leadsSheet = ss.getSheetByName("Leads");
  var leadsHeadersMap = leadsSheet ? getHeadersMap(leadsSheet) : null;
  
  var processed = 0;
  var successCount = 0;
  var dataRange = readySheet.getRange(2, 1, lastRow - 1, readySheet.getLastColumn());
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (successCount >= batchLimit) {
      Logger.log("Batch Limit Reached: processed " + batchLimit + " emails from Ready to Send tab.");
      break;
    }
    
    var row = data[i];
    var r = i + 2; // actual row number in the sheet
    var isReady = row[headersMap["Ready for Send"] - 1]; // boolean
    
    if (isReady === true || isReady === "true" || isReady === "TRUE") {
      var recipient = row[headersMap["Email"] - 1];
      var selectedAccountName = row[headersMap["Selected Account"] - 1];
      var finalSubject = row[headersMap["Subject"] - 1];
      var finalBody = row[headersMap["Body"] - 1];
      
      if (!recipient) {
        Logger.log("Skipping Ready tab row " + r + ": no recipient email.");
        continue;
      }
      
      var senderEmail = getConfigValue(config, selectedAccountName + " Email", "");
      var senderLabel = getConfigValue(config, selectedAccountName + " Label", "");
      var senderSignature = getConfigValue(config, selectedAccountName + " Signature", "");
      
      var originalEmail = recipient;
      var stagingMode = getConfigValue(config, "Staging Mode", "false").toString().trim().toLowerCase() === "true";
      
      if (DRY_RUN) {
        recipient = testRecipient || originalEmail;
        finalSubject = "[TEST DRAFT] " + finalSubject;
        finalBody = "=== DRY RUN (Original Recipient: " + originalEmail + ") ===\n\n" + finalBody;
      } else if (testRecipient) {
        recipient = testRecipient;
        if (!stagingMode) {
          finalSubject = "[TEST] " + finalSubject;
          finalBody = "=== TEST RUN (Original Recipient: " + originalEmail + ") ===\n\n" + finalBody;
        }
      }
      
      var finalBodyWithSig = finalBody;
      if (senderSignature) {
        finalBodyWithSig += "\n\n" + senderSignature;
      }
      
      var options = {
        htmlBody: finalBodyWithSig.replace(/\n/g, "<br>"),
        from: senderEmail,
        name: senderLabel,
        replyTo: senderEmail
      };
      
      try {
        var isDraft = (outreachMode.toLowerCase() !== "send" || DRY_RUN);
        
        // Quota check before sending directly
        if (!isDraft) {
          var remaining = getRemainingQuota(selectedAccountName);
          if (remaining <= 0) {
            Logger.log("Skipping Ready tab row " + r + ": Quota exhausted for " + selectedAccountName);
            continue;
          }
        }
        
        var apiRes = sendOrDraftViaAPI_(senderEmail, senderLabel, recipient, finalSubject, options.htmlBody, isDraft);
        if (!apiRes.success) {
          throw new Error("Gmail API Error: " + apiRes.error);
        }
        
        if (!isDraft) {
          try { recordSend(selectedAccountName); } catch(qe) {}
        }
        successCount++;
        processed++;
        
        // Update the main Leads sheet so Pipeline Stage reflects the send or draft
        if (leadsSheet && leadsHeadersMap) {
          var leadsLastRow = leadsSheet.getLastRow();
          if (leadsLastRow > 1 && leadsHeadersMap["Email"]) {
            var emailCol = leadsHeadersMap["Email"];
            var emailVals = leadsSheet.getRange(2, emailCol, leadsLastRow - 1, 1).getValues();
            for (var li = 0; li < emailVals.length; li++) {
              if (emailVals[li][0].toString().trim().toLowerCase() === originalEmail.toLowerCase()) {
                var leadsRow = li + 2;
                if (isDraft) {
                  if (leadsHeadersMap["Pipeline Stage"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Pipeline Stage"]).setValue("Draft Created");
                  if (leadsHeadersMap["Outreach Status"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Outreach Status"]).setValue(DRY_RUN ? "Draft Created (Dry Run)" : "Draft Created");
                  if (leadsHeadersMap["Thread Id"] && apiRes.draftId) {
                    try { leadsSheet.getRange(leadsRow, leadsHeadersMap["Thread Id"]).setValue("DRAFT:" + apiRes.draftId); } catch(te) {}
                  }
                } else {
                  if (leadsHeadersMap["Pipeline Stage"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Pipeline Stage"]).setValue("Sent");
                  if (leadsHeadersMap["Last Sent At"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Last Sent At"]).setValue(new Date());
                  if (leadsHeadersMap["Outreach Status"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Outreach Status"]).setValue("Email Sent");
                  // CRITICAL: write Thread Id so Scan Inbox can track this lead
                  if (leadsHeadersMap["Thread Id"] && apiRes.threadId) {
                    try { leadsSheet.getRange(leadsRow, leadsHeadersMap["Thread Id"]).setValue(apiRes.threadId); } catch(te) {}
                  }
                  // Auto-schedule follow-up sequence: FU1 = 3 days from now, FU2 = 10 days after FU1.
                  // Also sets Follow-up Status = "Pending" automatically.
                  if (typeof scheduleFollowups === "function") {
                    try { scheduleFollowups({ row: leadsRow }); } catch(se) {
                      Logger.log("processReadyTabHourly: scheduleFollowups failed for leadsRow " + leadsRow + ": " + se);
                    }
                  }
                }
                
                // Stamp the account name (whether draft or send)
                if (leadsHeadersMap["Send From Account"]) {
                  var existingAcc2 = leadsSheet.getRange(leadsRow, leadsHeadersMap["Send From Account"]).getValue().toString().trim();
                  if (!existingAcc2) leadsSheet.getRange(leadsRow, leadsHeadersMap["Send From Account"]).setValue(selectedAccountName);
                }

                break;
              }
            }
          }
        }
        
        // Remove row from Ready tab after successful send
        readySheet.deleteRow(r);
        
        // After deletion, re-check how many rows remain before re-fetching
        var newLastRow = readySheet.getLastRow();
        if (newLastRow <= 1) break; // sheet is now empty — stop
        
        // Re-fetch data snapshot since row indices shifted
        i--;
        dataRange = readySheet.getRange(2, 1, newLastRow - 1, readySheet.getLastColumn());
        data = dataRange.getValues();
        
      } catch (e) {
        Logger.log("Error sending from Ready tab row " + r + ": " + e.toString());
      }
    }
  }
  
  Logger.log("Outreach Complete (Ready Tab). Checked: " + processed + ", Success: " + successCount);
}

/**
 * Sends exactly ONE approved draft from the 'Ready to Send' tab.
 * Used exclusively by the Drip Engine.
 */
function sendOneReadyDraft_(config, ss) {
  var readySheet = ss.getSheetByName("Ready to Send");
  var leadsSheet = ss.getSheetByName("Leads");
  if (!readySheet || !leadsSheet) return false;
  
  var lastRow = readySheet.getLastRow();
  if (lastRow <= 1) return false;
  
  var headers = readySheet.getRange(1, 1, 1, readySheet.getLastColumn()).getValues()[0];
  var readyForSendIdx = headers.indexOf("Ready for Send");
  var emailIdx = headers.indexOf("Email");
  var accountIdx = headers.indexOf("Selected Account");
  var subjectIdx = headers.indexOf("Subject");
  var bodyIdx = headers.indexOf("Body");
  
  if (readyForSendIdx === -1 || emailIdx === -1 || accountIdx === -1) return false;
  
  var dataRange = readySheet.getRange(2, 1, lastRow - 1, readySheet.getLastColumn());
  var data = dataRange.getValues();
  
  var leadsHeadersMap = getHeadersMap(leadsSheet);
  
  for (var i = 0; i < data.length; i++) {
    var isReady = data[i][readyForSendIdx];
    if (isReady === true || isReady.toString().toLowerCase() === "true") {
      
      var r = i + 2;
      var recipient = data[i][emailIdx].toString().trim();
      var selectedAccountName = data[i][accountIdx].toString().trim();
      var finalSubject = data[i][subjectIdx].toString().trim();
      var finalBody = data[i][bodyIdx].toString().trim();
      
      if (!recipient || !selectedAccountName) continue;
      
      var senderEmail = getConfigValue(config, selectedAccountName + " Email", "").toString().trim();
      var senderLabel = getConfigValue(config, selectedAccountName + " Label", "").toString().trim();
      var senderSignature = getConfigValue(config, selectedAccountName + " Signature", "").toString().trim();
      
      var finalBodyWithSig = finalBody;
      if (senderSignature) {
        finalBodyWithSig += "\n\n" + senderSignature;
      }
      
      var options = {
        htmlBody: finalBodyWithSig.replace(/\n/g, "<br>"),
        from: senderEmail,
        name: senderLabel,
        replyTo: senderEmail
      };
      
      // Check Hourly Quota
      var hourlyCheck = checkHourlyLimit(selectedAccountName, "fresh");
      if (!hourlyCheck.ok) {
        Logger.log("sendOneReadyDraft: Hourly limit hit for " + selectedAccountName + ". Skipping.");
        continue;
      }
      
      // Check Daily Quota
      var remaining = getRemainingQuota(selectedAccountName);
      if (remaining <= 0) {
        Logger.log("sendOneReadyDraft: Daily quota exhausted for " + selectedAccountName + ". Skipping.");
        continue;
      }
      
      var apiRes = sendOrDraftViaAPI_(senderEmail, senderLabel, recipient, finalSubject, options.htmlBody, false);
      if (!apiRes.success) {
        Logger.log("sendOneReadyDraft API Error: " + apiRes.error);
        continue;
      }
      
      try { recordSend(selectedAccountName); } catch(qe) {}
      try { recordHourlySend(selectedAccountName, "fresh"); } catch(he) {}
      
      // Update Leads sheet
      var leadsLastRow = leadsSheet.getLastRow();
      if (leadsLastRow > 1 && leadsHeadersMap["Email"]) {
        var emailCol = leadsHeadersMap["Email"];
        var emailVals = leadsSheet.getRange(2, emailCol, leadsLastRow - 1, 1).getValues();
        for (var li = 0; li < emailVals.length; li++) {
          if (emailVals[li][0].toString().trim().toLowerCase() === recipient.toLowerCase()) {
            var leadsRow = li + 2;
            if (leadsHeadersMap["Pipeline Stage"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Pipeline Stage"]).setValue("Sent");
            if (leadsHeadersMap["Last Sent At"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Last Sent At"]).setValue(new Date());
            if (leadsHeadersMap["Outreach Status"]) leadsSheet.getRange(leadsRow, leadsHeadersMap["Outreach Status"]).setValue("Email Sent");
            if (leadsHeadersMap["Send From Account"]) {
              var existingAcc2 = leadsSheet.getRange(leadsRow, leadsHeadersMap["Send From Account"]).getValue().toString().trim();
              if (!existingAcc2) leadsSheet.getRange(leadsRow, leadsHeadersMap["Send From Account"]).setValue(selectedAccountName);
            }
            if (leadsHeadersMap["Thread Id"] && apiRes.threadId) {
              try { leadsSheet.getRange(leadsRow, leadsHeadersMap["Thread Id"]).setValue(apiRes.threadId); } catch(te) {}
            }
            if (typeof scheduleFollowups === "function") {
              try { scheduleFollowups({ row: leadsRow }); } catch(se) {}
            }
            break;
          }
        }
      }
      
      try { logDailySend("Fresh", selectedAccountName, recipient, finalSubject, apiRes.threadId || ""); } catch(e) {}
      
      readySheet.deleteRow(r);
      Logger.log("Drip Engine: Sent Ready Draft to " + recipient);
      return true; // Sent one!
    }
  }
  return false;
}

/**
 * Sends exactly ONE pending draft from the Leads pipeline.
 * Used exclusively by the Drip Engine.
 */
function sendOnePipelineDraft_(config, ss) {
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) return false;
  
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) return false;
  
  for (var r = 2; r <= lastRow; r++) {
    var status = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
    if (status === "Draft Created" || status === "Draft Created (Dry Run)") {
      var email = leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim();
      var accountName = headersMap["Send From Account"] ? leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim() : "";
      if (!accountName) accountName = "Account A";
      
      var threadIdColVal = headersMap["Thread Id"] ? leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim() : "";
      var senderEmail = getConfigValue(config, accountName + " Email", "").toString().trim();
      var draftId = "";
      
      if (threadIdColVal.indexOf("DRAFT:") === 0) {
        draftId = threadIdColVal.substring(6);
      }
      
      if (draftId && senderEmail) {
        // Quota check
        var remaining = getRemainingQuota(accountName);
        if (remaining <= 0) continue;
        
        // Hourly check
        var hourlyCheck = checkHourlyLimit(accountName, "fresh");
        if (!hourlyCheck.ok) continue;
        
        var sendRes = sendDraftByIdViaAPI_(senderEmail, draftId);
        
        if (sendRes.success) {
          try { recordSend(accountName); } catch(qe) {}
          try { recordHourlySend(accountName, "fresh"); } catch(he) {}
          
          leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Email Sent");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Sent");
          
          var now = new Date();
          leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(now);
          if (headersMap["Send Date"]) leadsSheet.getRange(r, headersMap["Send Date"]).setValue(now);
          
          var actualThreadId = "";
          if (headersMap["Thread Id"] && sendRes.threadId) {
            actualThreadId = sendRes.threadId;
            try { leadsSheet.getRange(r, headersMap["Thread Id"]).setValue(actualThreadId); } catch(te) {}
          }
          
          if (headersMap["Send From Account"]) {
            var existingAcc = leadsSheet.getRange(r, headersMap["Send From Account"]).getValue().toString().trim();
            if (!existingAcc) leadsSheet.getRange(r, headersMap["Send From Account"]).setValue(accountName);
          }
          
          try { logDailySend("Fresh", accountName, email, "(Sent via Pipeline Draft)", actualThreadId); } catch(e) {}
          
          if (typeof scheduleFollowups === "function") {
            try { scheduleFollowups({ row: r }); } catch(se) {}
          }
          
          Logger.log("Drip Engine: Sent Pipeline Draft to " + email);
          return true; // Sent one!
        }
      }
    }
  }
  return false;
}
