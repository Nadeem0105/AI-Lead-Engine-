/**
 * AI Lead Generation Engine - Main Controller & Menu
 * File: Main.gs
 */

/**
 * Triggered automatically when the spreadsheet is opened.
 * Adds a custom menu to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Lead Engine")
    .addItem("Setup Sheets", "setupSheetStructure")
    .addItem("Upgrade to Dual Accounts", "upgradeToDualAccount")
    .addItem("Apply Column Dropdowns", "applyLeadDropdownValidations")
    .addItem("Highlight Required Columns", "applyHeaderFormattingMenu")
    .addSeparator()
    .addSubMenu(ui.createMenu("Ingest Leads")
      .addItem("Apollo Ingest (Sync Saved)", "runLeadIngestionPipeline")
      .addItem("LinkedIn X-Ray Ingest (Google CSE)", "runLinkedInXRayIngestionPipeline")
      .addItem("GitHub Search Ingest", "runGitHubIngestionPipeline")
      .addItem("Google Maps Ingest", "runGoogleMapsIngestionPipeline")
      .addSeparator()
      .addItem("Manual Apollo Enrichment (Selected)", "enrichSelectedRows")
    )
    .addSeparator()
    .addItem("Score new leads", "runScoringOnlyMenu")
    .addItem("Process Leads (End-to-End)", "runFullPipelineMenu")
    .addSeparator()
    .addItem("Test Outreach Tone (3 Samples)", "runOutreachTest")
    .addItem("Preview drafts for selected rows", "previewSelectedDrafts")
    .addItem("Generate outreach drafts (All Ready)", "runOutreachPipeline")
    .addItem("Generate drafts for selected rows", "generateSelectedDrafts")
    .addItem("Send all generated drafts", "sendDraftsFromPipeline")
    .addItem("Scan replies & send follow-ups", "detectRepliesAndFollowUp")
    .addItem("Preview follow-up (selected rows)", "previewFollowUpSelected")
    .addItem("Send follow-up (selected rows)", "sendFollowUpForSelected")
    .addItem("Send follow-up (all eligible)", "sendFollowUpForAll")
    .addSeparator()
    .addSubMenu(ui.createMenu("Automation Triggers")
      .addItem("Start Hourly Sending", "setupOutreachTrigger")
      .addItem("Stop Sending", "deactivateOutreachTrigger")
      .addItem("Send Now (Manual)", "runOutreachPipelineManualBatch")
      .addSeparator()
      .addItem("Start Daily Follow-ups", "setupFollowUpTrigger")
      .addItem("Stop Follow-ups", "deactivateFollowUpTrigger")
    )
    .addSeparator()
    .addSubMenu(ui.createMenu("Assign Send Account")
      .addItem("Assign to Account A (Selected rows)", "setAccountAForSelected")
      .addItem("Assign to Account B (Selected rows)", "setAccountBForSelected")
    )
    .addToUi();
}

/**
 * Loads configuration parameters from the "Config" sheet.
 * 
 * @return {object} Config values parsed into an object
 */
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    throw new Error("Config sheet not found. Please run 'Setup Sheets' first.");
  }
  
  var data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  var config = {};
  
  data.forEach(function(row) {
    var key = row[0].toString().trim();
    var val = row[1];
    
    // Store raw key-value pair for easy lookup
    config[key] = val;
    
    if (key === "Revenue Cutoff (INR)") {
      config.revenueCutoff = val;
    } else if (key === "Funding Recency Window (Months)") {
      config.fundingWindow = val;
    } else if (key === "Email Confidence Threshold") {
      config.emailConfidenceThreshold = parseFloat(val) || 0.7;
    } else if (key === "Gemini Model") {
      config.model = val.toString().trim();
    } else if (key === "Gemini Temperature") {
      config.temperature = parseFloat(val) || 0.2;
    } else if (key === "Scoring Prompt Template") {
      config.promptTemplate = val.toString();
    } else if (key === "Apollo Ingest Job Titles") {
      config.apolloTitles = val.toString().split(",").map(function(s) { return s.trim(); }).filter(Boolean);
    } else if (key === "Apollo Ingest Locations") {
      config.apolloLocations = val.toString().split(",").map(function(s) { return s.trim(); }).filter(Boolean);
    } else if (key === "Apollo Ingest Industries") {
      config.apolloIndustries = val.toString().split(",").map(function(s) { return s.trim(); }).filter(Boolean);
    } else if (key === "Apollo Ingest Limit") {
      config.apolloLimit = parseInt(val) || 10;
    } else if (key === "Google CSE ID") {
      config.googleCseId = val.toString().trim();
    } else if (key === "GitHub API Key (Optional)") {
      config.githubKey = val.toString().trim();
    } else if (key === "Google Maps API Key") {
      config.mapsKey = val.toString().trim();
    } else if (key === "Hunter.io API Key (Optional)") {
      config.hunterKey = val.toString().trim();
    } else if (key === "ZeroBounce API Key (Optional)") {
      config.zeroBounceKey = val.toString().trim();
    }
  });
  
  return config;
}

/**
 * Sets a configuration value in the Config sheet.
 */
function setConfigValue(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) return;
  var data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === key) {
      configSheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  // If not found, append it
  configSheet.appendRow([key, value]);
}

/**
 * Helper to build a mapping of header names to 1-indexed column positions.
 * 
 * @param {Sheet} sheet The sheet to read headers from
 * @return {object} Key-value map of column header name -> 1-based index
 */
function getHeadersMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(header, idx) {
    map[header.toString().trim()] = idx + 1;
  });
  return map;
}

/**
 * Logs a run event to the "Log" sheet.
 * 
 * @param {number} processedCount Number of rows scanned/processed
 * @param {number} scoredCount Number of rows scored by AI
 * @param {number} flaggedCount Number of rows flagged on validation
 * @param {string} errors Concatenated error messages or "None"
 */
function logRun(processedCount, scoredCount, flaggedCount, errors) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log");
  if (!logSheet) return;
  
  logSheet.appendRow([
    new Date(),
    processedCount,
    scoredCount,
    flaggedCount,
    errors || "None"
  ]);
}

/**
 * Helper to check if the lead is in the recruitment industry.
 * Feature 4: Recruitment industry hard filter.
 */
function isRecruitmentIndustry(industryText, config) {
  if (!industryText || typeof industryText !== "string") return false;
  
  var keywords = (config["Recruitment Industry Keywords"] || "").split(",");
  if (keywords.length === 0 || (keywords.length === 1 && keywords[0] === "")) return false;
  
  var textLower = industryText.toLowerCase();
  for (var i = 0; i < keywords.length; i++) {
    var kw = keywords[i].toString().trim().toLowerCase();
    if (kw && textLower.indexOf(kw) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * Manual menu action to only run lead scoring (Step 2).
 */
function runScoringOnlyMenu() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found. Please run 'Setup Sheets' first.", ui.ButtonSet.OK);
    return;
  }
  
  var config = getConfig();
  if (!checkApiKeys(ui, config)) {
    return;
  }
  var headersMap = getHeadersMap(leadsSheet);
  var mapping = resolveColumnMapping(leadsSheet, true);
  
  if (!validateHeadersPresent(ui, headersMap, ["Company", "Annual Revenue", "Total Funding", "Latest Funding", "Latest Funding Amount", "Last Raised At", "Score", "Score Reason", "Company Validation Status", "Company Validation Reason", "Research Status", "Pipeline Stage"])) {
    return;
  }
  
  var lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert("No Data", "Leads sheet is empty. Please add lead rows below the header.", ui.ButtonSet.OK);
    return;
  }
  
  var startTime = new Date().getTime();
  var processed = 0;
  var scored = 0;
  var errors = [];
  
  for (var r = 2; r <= lastRow; r++) {
    if (new Date().getTime() - startTime > 300000) {
      ui.alert("Timeout Warning", "Script has been running for 5 minutes. Stopping early to prevent Google timeout. Please run again to process remaining leads.", ui.ButtonSet.OK);
      break;
    }
    
    if (scored >= 80) {
      ui.alert("Batch Limit Reached", "Processed the maximum limit of 80 leads in this run. Please run the menu action again to score more leads.", ui.ButtonSet.OK);
      break;
    }
    
    // Check Pipeline Stage first
    var pipelineStageRange = leadsSheet.getRange(r, headersMap["Pipeline Stage"]);
    var pipelineStage = pipelineStageRange.getValue().toString().trim();
    if (pipelineStage.indexOf("Disqualified") !== -1) {
      continue;
    }
    
    // Feature 4: Recruitment Filter Gate
    var industryText = getCanonical(leadsSheet, r, mapping, "industry");
    if (isRecruitmentIndustry(industryText, config)) {
      pipelineStageRange.setValue("Disqualified — Recruitment Industry");
      if (headersMap["Outreach Status"]) leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Disqualified");
      continue;
    }
    
    // Feature 1: Flag rows with no resolvable email before wasting API credits
    if (flagIfMissingEmail(leadsSheet, r, mapping, headersMap)) continue;
    
    // Ensure company is validated first
    var companyValStatus = leadsSheet.getRange(r, headersMap["Company Validation Status"]).getValue().toString().trim();
    if (companyValStatus === "") {
      var valResult = validateCompanyDetails(leadsSheet, r, headersMap, config);
      companyValStatus = valResult.status;
      if (headersMap["Company Validation Reason"] && valResult.reason) {
        leadsSheet.getRange(r, headersMap["Company Validation Reason"]).setValue(valResult.reason);
      }
      pipelineStageRange.setValue(companyValStatus === "Verified" ? "Company Verified" : "Needs Review");
    }
    
    if (companyValStatus !== "Verified") {
      // Skip scoring for unverified companies
      continue;
    }
    
    var existingScore = leadsSheet.getRange(r, headersMap["Score"]).getValue().toString().trim();
    var existingReason = leadsSheet.getRange(r, headersMap["Score Reason"]).getValue().toString().trim();
    
    // Feature 3: A valid numeric score (1-10) is treated as a manual override — skip AI entirely.
    var scoreAsNum = parseFloat(existingScore);
    var isValidManualScore = (!isNaN(scoreAsNum) && scoreAsNum >= 1 && scoreAsNum <= 10 &&
                              existingScore.indexOf("Error") === -1);
    var isScoreEmptyOrError = (existingScore === "" || existingScore.indexOf("Error") !== -1 || isNaN(scoreAsNum));
    var isReasonEmptyOrError = (existingReason === "" || existingReason.indexOf("Error") !== -1);
    
    if (isValidManualScore) {
      // Manual score path: ensure Score Source and Reason are stamped
      var scoreSourceRange = headersMap["Score Source"] ? leadsSheet.getRange(r, headersMap["Score Source"]) : null;
      if (scoreSourceRange && scoreSourceRange.getValue().toString().trim() === "") {
        scoreSourceRange.setValue("Manual");
      }
      if (existingReason === "") {
        leadsSheet.getRange(r, headersMap["Score Reason"]).setValue("Manual Override");
      }
      pipelineStageRange.setValue("Scored");
    } else if (isScoreEmptyOrError || isReasonEmptyOrError) {
      processed++;
      var result = scoreSingleLead(leadsSheet, r, headersMap, config, mapping);
      if (result.success) {
        scored++;
        pipelineStageRange.setValue("Scored");
        if (headersMap["Score Source"]) leadsSheet.getRange(r, headersMap["Score Source"]).setValue("AI");
      } else {
        errors.push("Row " + r + ": " + result.error);
      }
      // Delay 3 seconds to optimize for Groq's 30 RPM and TPM limits
      Utilities.sleep(3000);
    }
  }
  
  var errorText = errors.length > 0 ? errors.join("; ") : "None";
  logRun(processed, scored, 0, errorText);
  
  var remainingUnscored = countUnscoredRows(leadsSheet, headersMap, lastRow);
  var alertMsg = "Scoring run completed!\n\n" +
                 "- Rows checked: " + processed + "\n" +
                 "- Rows successfully scored: " + scored + "\n" +
                 "- Failures: " + errors.length + "\n" +
                 "- Remaining unscored rows: " + remainingUnscored;
                 
  if (remainingUnscored > 0) {
    alertMsg += "\n\n(Tip: Run 'Score new leads' again to process remaining leads)";
  }
  
  ui.alert("Scoring Complete", alertMsg, ui.ButtonSet.OK);
}

/**
 * End-to-End pipeline menu action: Scores and Validates leads in one sequence (Step 4).
 */
function runFullPipelineMenu() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found. Please run 'Setup Sheets' first.", ui.ButtonSet.OK);
    return;
  }
  
  var config = getConfig();
  if (!checkApiKeys(ui, config)) {
    return;
  }
  var headersMap = getHeadersMap(leadsSheet);
  var mapping = resolveColumnMapping(leadsSheet, true);
  
  // Verify all required headers exist
  var requiredHeaders = [
    "Company", "Annual Revenue", "Total Funding", "Latest Funding", "Latest Funding Amount", "Last Raised At",
    "Email", "Email Status", "Email Confidence", "Primary Email Catch-all Status",
    "Score", "Score Reason", "Validation Status", "Validation Reason", "Outreach Status",
    "Company Validation Status", "Company Validation Reason", "Research Status", "Pipeline Stage"
  ];
  
  if (!validateHeadersPresent(ui, headersMap, requiredHeaders)) {
    return;
  }
  
  var lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert("No Data", "Leads sheet is empty. Please add lead rows below the header.", ui.ButtonSet.OK);
    return;
  }
  
  var startTime = new Date().getTime();
  var processed = 0;
  var scoredThisRun = 0;
  var validatedThisRun = 0;
  var flaggedThisRun = 0;
  var errors = [];
  
  for (var r = 2; r <= lastRow; r++) {
    // 5-minute timeout safety
    if (new Date().getTime() - startTime > 300000) {
      ui.alert("Timeout Warning", "Script has been running for 5 minutes. Stopping early to prevent Google timeout. Please run again to process remaining leads.", ui.ButtonSet.OK);
      break;
    }
    
    // Safety cap of 80 new AI calls per run
    if (scoredThisRun >= 80) {
      ui.alert("Batch Limit Reached", "Processed the maximum limit of 80 AI scoring calls in this run. Please run the menu action again to process the rest.", ui.ButtonSet.OK);
      break;
    }
    
    var pipelineStageRange = leadsSheet.getRange(r, headersMap["Pipeline Stage"]);
    var pipelineStage = pipelineStageRange.getValue().toString().trim();
    if (pipelineStage.indexOf("Disqualified") !== -1) {
      continue;
    }
    
    // Feature 4: Recruitment Filter Gate
    var industryText = getCanonical(leadsSheet, r, mapping, "industry");
    if (isRecruitmentIndustry(industryText, config)) {
      pipelineStageRange.setValue("Disqualified — Recruitment Industry");
      if (headersMap["Outreach Status"]) leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Disqualified");
      continue;
    }
    
    // Feature 1: Flag rows with no resolvable email before wasting API credits
    if (flagIfMissingEmail(leadsSheet, r, mapping, headersMap)) continue;
    
    // Ensure company is validated first
    var companyValStatus = leadsSheet.getRange(r, headersMap["Company Validation Status"]).getValue().toString().trim();
    if (companyValStatus === "") {
      var valResult = validateCompanyDetails(leadsSheet, r, headersMap, config);
      companyValStatus = valResult.status;
      if (headersMap["Company Validation Reason"] && valResult.reason) {
        leadsSheet.getRange(r, headersMap["Company Validation Reason"]).setValue(valResult.reason);
      }
      pipelineStageRange.setValue(companyValStatus === "Verified" ? "Company Verified" : "Needs Review");
    }
    
    if (companyValStatus !== "Verified") {
      // Skip scoring and validation for unverified companies
      continue;
    }

    var scoreRange = leadsSheet.getRange(r, headersMap["Score"]);
    var scoreVal = scoreRange.getValue().toString().trim();
    var existingReason = leadsSheet.getRange(r, headersMap["Score Reason"]).getValue().toString().trim();
    
    // Feature 3: A valid numeric score (1-10) is treated as a manual override — skip AI entirely.
    var scoreAsNum = parseFloat(scoreVal);
    var isValidManualScore = (!isNaN(scoreAsNum) && scoreAsNum >= 1 && scoreAsNum <= 10 &&
                              scoreVal.indexOf("Error") === -1);
    var isScoreEmptyOrError = (scoreVal === "" || scoreVal.indexOf("Error") !== -1 || isNaN(scoreAsNum));
    var isReasonEmptyOrError = (existingReason === "" || existingReason.indexOf("Error") !== -1);
    
    // Idempotency Step 1: AI Scoring or Manual Override
    if (isValidManualScore) {
      // Manual score path
      var scoreSourceRange = headersMap["Score Source"] ? leadsSheet.getRange(r, headersMap["Score Source"]) : null;
      if (scoreSourceRange && scoreSourceRange.getValue().toString().trim() === "") {
        scoreSourceRange.setValue("Manual");
      }
      if (existingReason === "") {
        leadsSheet.getRange(r, headersMap["Score Reason"]).setValue("Manual Override");
      }
      pipelineStageRange.setValue("Scored");
    } else if (isScoreEmptyOrError || isReasonEmptyOrError) {
      processed++;
      
      // Look up hiring status if it hasn't been fetched yet
      if (headersMap["Hiring Status"]) {
        var currentHiring = leadsSheet.getRange(r, headersMap["Hiring Status"]).getValue().toString().trim();
        if (currentHiring === "" || currentHiring === "Skipped (No API Config)") {
          checkAndRecordHiringStatus(leadsSheet, r, headersMap, config);
        }
      }
      
      var scoreResult = scoreSingleLead(leadsSheet, r, headersMap, config, mapping);
      if (scoreResult.success) {
        scoredThisRun++;
        scoreVal = scoreResult.score; // Use newly generated score
        pipelineStageRange.setValue("Scored");
        if (headersMap["Score Source"]) leadsSheet.getRange(r, headersMap["Score Source"]).setValue("AI");
      } else {
        errors.push("Row " + r + " Scoring: " + scoreResult.error);
        Utilities.sleep(3000); // Sleep before continuing to prevent API spamming
        continue; // Skip validation since scoring failed
      }
      Utilities.sleep(3000); // Sleep after a successful API call
    } else {
      // Feature 3: Manual Score
      var scoreSourceRange = headersMap["Score Source"] ? leadsSheet.getRange(r, headersMap["Score Source"]) : null;
      if (scoreSourceRange && scoreSourceRange.getValue().toString().trim() === "") {
        scoreSourceRange.setValue("Manual");
        pipelineStageRange.setValue("Scored");
      }
    }
    
    // Idempotency Step 2: Email Validation & Outreach Status (if score exists)
    if (scoreVal !== "" && !isNaN(scoreVal)) {
      var scoreNum = parseInt(scoreVal);
      var validationRange = leadsSheet.getRange(r, headersMap["Validation Status"]);
      var existingValidation = validationRange.getValue().toString().trim();
      
      var outreachRange = leadsSheet.getRange(r, headersMap["Outreach Status"]);
      var existingOutreach = outreachRange.getValue().toString().trim();
      
      var validationReasonRange = headersMap["Validation Reason"] ? leadsSheet.getRange(r, headersMap["Validation Reason"]) : null;
      
      // Step A: Run validation if missing
      if (existingValidation === "" && scoreNum >= 8) {
        var validationResult = validateEmailQuality(leadsSheet, r, headersMap, config);
        existingValidation = validationResult.status;
        validationRange.setValue(existingValidation);
        if (validationReasonRange && validationResult.reason) {
          validationReasonRange.setValue(validationResult.reason);
        }
        validatedThisRun++;
      } else if (existingValidation === "" && scoreNum >= 4 && scoreNum <= 7) {
        existingValidation = "Skipped (Nurture)";
        validationRange.setValue(existingValidation);
        if (validationReasonRange) validationReasonRange.setValue("Score between 4 and 7 (Nurture)");
      } else if (existingValidation === "" && scoreNum < 4) {
        existingValidation = "Skipped (Disqualified)";
        validationRange.setValue(existingValidation);
        if (validationReasonRange) validationReasonRange.setValue("Score below 4 (Disqualified)");
      }
      
      // Step B: Set Outreach Status if missing
      if (existingOutreach === "") {
        if (scoreNum >= 8) {
          if (existingValidation === "Ready") {
            outreachRange.setValue("Ready for outreach");
            leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Validated");
          } else {
            outreachRange.setValue("Flagged - email risk");
            leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Email Flagged");
            flaggedThisRun++;
          }
        } else if (scoreNum >= 4 && scoreNum <= 7) {
          outreachRange.setValue("Nurture");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Nurture");
        } else {
          outreachRange.setValue("Low Priority");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Disqualified");
        }
      }
    }
  }
  
  // Log the run
  var errorText = errors.length > 0 ? errors.join("; ") : "None";
  logRun(processed, scoredThisRun, flaggedThisRun, errorText);
  
  var remainingUnscored = countUnscoredRows(leadsSheet, headersMap, lastRow);
  var alertMsg = "Pipeline completed successfully!\n\n" +
                 "- Total leads scanned: " + processed + "\n" +
                 "- Leads scored by AI this run: " + scoredThisRun + "\n" +
                 "- Leads validated this run: " + validatedThisRun + "\n" +
                 "- Email risk flags raised: " + flaggedThisRun + "\n" +
                 "- Failures: " + errors.length + "\n" +
                 "- Remaining unscored rows: " + remainingUnscored;
                 
  if (remainingUnscored > 0) {
    alertMsg += "\n\n(Tip: Run 'Process Leads (End-to-End)' again to process the remaining leads)";
  }
  
  ui.alert("Pipeline Run Complete", alertMsg, ui.ButtonSet.OK);
}

/**
 * Helper to count remaining unscored rows.
 */
function countUnscoredRows(sheet, headersMap, lastRow) {
  var count = 0;
  for (var r = 2; r <= lastRow; r++) {
    var val = sheet.getRange(r, headersMap["Score"]).getValue().toString().trim();
    var reason = sheet.getRange(r, headersMap["Score Reason"]).getValue().toString().trim();
    var isScoreEmptyOrError = (val === "" || val.indexOf("Error") !== -1 || isNaN(val));
    var isReasonEmptyOrError = (reason === "" || reason.indexOf("Error") !== -1);
    if (isScoreEmptyOrError || isReasonEmptyOrError) {
      count++;
    }
  }
  return count;
}

/**
 * Helper to alert user about API key setup.
 */
function showApiKeyAlert(ui, missingKeyType) {
  ui.alert(
    "API Key Missing",
    "Please set your API keys in Apps Script project settings:\n\n" +
    "1. In the Apps Script editor, click the gear icon (Project Settings).\n" +
    "2. Scroll down to 'Script Properties' and click 'Add script property'.\n" +
    "3. Add at least one of these properties:\n" +
    "   - NVIDIA_API_KEYS (for DeepSeek via NVIDIA, e.g. key1,key2)\n" +
    "   - GROQ_API_KEYS (for Groq keys, e.g. key1,key2)\n" +
    "   - CEREBRAS_API_KEYS (for Cerebras keys, e.g. key1,key2)\n" +
    "   - GEMINI_API_KEY (for Gemini)\n" +
    "4. Click 'Save script properties'.",
    ui.ButtonSet.OK
  );
}

/**
 * Validates that the appropriate API keys exist in script properties.
 */
function checkApiKeys(ui, config) {
  var modelName = config.model || "deepseek-ai/deepseek-r1";
  var modelNameLower = modelName.toLowerCase();
  var isGemini = (modelNameLower.indexOf("gemini") !== -1);
  var isNvidia = (modelNameLower.indexOf("deepseek") !== -1 || modelNameLower.indexOf("nvidia") !== -1);
  var props = PropertiesService.getScriptProperties();
  
  if (isGemini) {
    var geminiKey = props.getProperty("GEMINI_API_KEY");
    if (!geminiKey) {
      showApiKeyAlert(ui, "GEMINI_API_KEY");
      return false;
    }
  } else if (isNvidia) {
    var hasNvidia = props.getProperty("NVIDIA_API_KEYS") || props.getProperty("NVIDIA_API_KEY");
    if (!hasNvidia) {
      showApiKeyAlert(ui, "NVIDIA_API_KEYS / NVIDIA_API_KEY");
      return false;
    }
  } else {
    var hasGroq = props.getProperty("GROQ_API_KEYS") || props.getProperty("GROQ_API_KEY");
    var hasCerebras = props.getProperty("CEREBRAS_API_KEYS") || props.getProperty("CEREBRAS_API_KEY");
    
    if (!hasGroq && !hasCerebras) {
      showApiKeyAlert(ui, "GROQ_API_KEYS / CEREBRAS_API_KEYS");
      return false;
    }
  }
  return true;
}

/**
 * Helper to validate all required headers are present.
 */
function validateHeadersPresent(ui, headersMap, requiredHeaders) {
  for (var i = 0; i < requiredHeaders.length; i++) {
    if (!headersMap[requiredHeaders[i]]) {
      ui.alert("Error", "Missing required column in Leads sheet: '" + requiredHeaders[i] + "'. Please run Setup Sheets to restore defaults.", ui.ButtonSet.OK);
      return false;
    }
  }
  return true;
}

/**
 * Assigns Account A to the 'Send From Account' column for the currently selected rows.
 */
function setAccountAForSelected() {
  setAccountForSelected("Account A");
}

/**
 * Assigns Account B to the 'Send From Account' column for the currently selected rows.
 */
function setAccountBForSelected() {
  setAccountForSelected("Account B");
}

/**
 * Helper to assign a given account string to selected rows in the Leads sheet.
 */
function setAccountForSelected(accountName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== "Leads") {
    SpreadsheetApp.getUi().alert("Error", "Please run this from the 'Leads' tab.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var headersMap = getHeadersMap(sheet);
  var accountCol = headersMap["Send From Account"];
  
  if (!accountCol) {
    SpreadsheetApp.getUi().alert("Error", "'Send From Account' column not found. Please click 'Setup Sheets' from the Lead Engine menu to add it first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  
  if (startRow < 2) {
    SpreadsheetApp.getUi().alert("Error", "Please select lead rows below the header.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  sheet.getRange(startRow, accountCol, numRows, 1).setValue(accountName);
  SpreadsheetApp.getUi().alert("Success", "Assigned '" + accountName + "' to " + numRows + " selected row(s).", SpreadsheetApp.getUi().ButtonSet.OK);
}
