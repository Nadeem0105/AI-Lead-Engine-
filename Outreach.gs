/**
 * AI Lead Generation Engine - DeepSeek Automated Email Outreach
 * File: Outreach.gs
 */

// Set to true during development/testing to prevent any real sends.
// If true, even if mode is 'Send', it will only create drafts with "[TEST DRAFT]" prepended.
var DRY_RUN = true; 

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
  
  // Verify required columns exist
  var required = [
    "First Name", "Last Name", "Title", "Company", "Email", "Validation Status", 
    "Outreach Status", "Annual Revenue", "Latest Funding", "Score", "Pipeline Stage"
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
    
    var scoreNum = 0;
    if (scoreVal !== "" && !isNaN(scoreVal)) {
      scoreNum = parseInt(scoreVal);
    }
    
    // Gate 1: Qualified leads (Score >= 8), Validated (Ready), Outreach Status is Ready, and Not Sent
    if (validationStatus === "Ready" && scoreNum >= 8 && outreachStatus === "Ready for outreach" && !lastSentAt) {
      processed++;
      
      var result = processSingleOutreach(leadsSheet, r, headersMap, config, outreachMode, testRecipient);
      
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
function processSingleOutreach(sheet, rowNumber, headersMap, config, outreachMode, testRecipient) {
  try {
    var originalEmail = sheet.getRange(rowNumber, headersMap["Email"]).getValue().toString().trim();
    if (!originalEmail) {
      throw new Error("Missing email address.");
    }
    
    // 1. Determine sending account in priority order:
    // Priority 1: Per-row Send From Account
    // Priority 2: Default Send Account in Config
    var rowOverride = headersMap["Send From Account"] ? sheet.getRange(rowNumber, headersMap["Send From Account"]).getValue().toString().trim() : "";
    var defaultAccount = getConfigValue(config, "Default Send Account", "Account A").toString().trim();
    var selectedAccountName = rowOverride || defaultAccount;
    
    // 2. Generate the personalized subject, opener, and closer using selected account
    var emailData = generatePersonalizedEmail(sheet, rowNumber, headersMap, config, selectedAccountName);
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
    
    var senderEmail = getConfigValue(config, selectedAccountName + " Email", Session.getActiveUser().getEmail());
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
    
    // 5. Send or Draft the email (In DRY_RUN mode, we ALWAYS create drafts)
    var statusVal = "";
    if (outreachMode.toLowerCase() === "send" && !DRY_RUN) {
      var draft = GmailApp.createDraft(recipient, finalSubject, finalBody, options);
      var msg = draft.send();
      statusVal = "Email Sent";
      sheet.getRange(rowNumber, headersMap["Pipeline Stage"]).setValue("Sent");
      if (headersMap["Last Sent At"]) {
        sheet.getRange(rowNumber, headersMap["Last Sent At"]).setValue(new Date());
      }
      if (headersMap["Thread Id"]) {
        sheet.getRange(rowNumber, headersMap["Thread Id"]).setValue(msg.getThread().getId());
      }
      if (headersMap["Sent From Account"]) {
        sheet.getRange(rowNumber, headersMap["Sent From Account"]).setValue(selectedAccountName);
      }
    } else {
      GmailApp.createDraft(recipient, finalSubject, finalBody, options);
      statusVal = "Draft Created";
      sheet.getRange(rowNumber, headersMap["Pipeline Stage"]).setValue("Draft Created");
      if (headersMap["Sent From Account"]) {
        sheet.getRange(rowNumber, headersMap["Sent From Account"]).setValue(selectedAccountName);
      }
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
 */
function generatePersonalizedEmail(sheet, rowNumber, headersMap, config, selectedAccountName) {
  try {
    var firstName = sheet.getRange(rowNumber, headersMap["First Name"]).getValue() || "";
    var lastName = sheet.getRange(rowNumber, headersMap["Last Name"]).getValue() || "";
    var title = sheet.getRange(rowNumber, headersMap["Title"]).getValue() || "";
    var company = sheet.getRange(rowNumber, headersMap["Company"]).getValue() || "";
    var industry = sheet.getRange(rowNumber, headersMap["Industry"]).getValue() || "";
    var keywords = headersMap["Keywords"] ? sheet.getRange(rowNumber, headersMap["Keywords"]).getValue() || "" : "";
    var website = headersMap["Website"] ? sheet.getRange(rowNumber, headersMap["Website"]).getValue() || "" : "";
    var technologies = headersMap["Technologies"] ? sheet.getRange(rowNumber, headersMap["Technologies"]).getValue() || "" : "";
    
    var revenue = sheet.getRange(rowNumber, headersMap["Annual Revenue"]).getValue() || "";
    var totalFunding = sheet.getRange(rowNumber, headersMap["Total Funding"]).getValue() || "";
    var latestFunding = sheet.getRange(rowNumber, headersMap["Latest Funding"]).getValue() || "";
    var latestFundingAmount = sheet.getRange(rowNumber, headersMap["Latest Funding Amount"]).getValue() || "";
    var lastRaisedAt = sheet.getRange(rowNumber, headersMap["Last Raised At"]).getValue() || "";
    
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
      
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0].toString().trim() === "Style Reference") {
          var prefAccount = rows[i][3] ? rows[i][3].toString().trim() : "";
          
          if (!defaultStyleSubject) {
            defaultStyleSubject = rows[i][1] || styleReferenceSubject;
            defaultStyleBody = rows[i][2] || styleReferenceBody;
          }
          
          if (prefAccount === selectedAccountName) {
            styleReferenceSubject = rows[i][1] || styleReferenceSubject;
            styleReferenceBody = rows[i][2] || styleReferenceBody;
            break;
          }
        }
      }
      
      if (styleReferenceSubject === "Top talent hiring at {Company}" && defaultStyleSubject) {
        styleReferenceSubject = defaultStyleSubject;
        styleReferenceBody = defaultStyleBody;
      }
    }
    
    // 3. Formulate the LLM prompt for opener/closer generation
    var prompt = "You are a professional B2B copywriter. We need to generate parts of a highly personalized, direct outreach email for a lead.\n\n" +
                 "Here is the style and tone reference email we are modeling:\n" +
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
                 "Your Tasks:\n" +
                 "1. Generate an 'opener' (1-2 sentences). It must reference what the company actually does specifically and accurately based on the details or search context. Avoid generic flattery. Never fabricate facts. If information is thin, write a safe, general opener.\n" +
                 "2. Generate a 'closer' (1 sentence). It must act as a CTA and reference the qualifying signal (recent funding event or revenue growth stage) naturally.\n\n" +
                 "IMPORTANT INSTRUCTION: Do not include any specific figures, amounts, dates, employee counts, or numeric values in the email body. Reference signals qualitatively only (e.g. 'following your recent funding milestone' instead of 'raised $50M').\n\n" +
                 "Guidelines:\n" +
                 "- Keep the tone direct, concise, and professional (no fluff, no exclamation marks).\n" +
                 "- Do not output the entire email. Only output the requested JSON object.\n\n" +
                 "OUTPUT FORMAT:\n" +
                 "Respond ONLY with a JSON object in this format:\n" +
                 "{\n" +
                 "  \"opener\": \"<Your generated opener>\",\n" +
                 "  \"closer\": \"<Your generated closer>\"\n" +
                 "}";
                 
    var modelName = config.model || "gemini-2.0-flash";
    var responseText = callFailoverModelForOutreach(modelName, prompt, config);
    
    var cleanedJson = cleanJsonResponseText(responseText);
    var result = JSON.parse(cleanedJson);
    
    if (!result.opener || !result.closer) {
      throw new Error("AI response missing opener or closer. Got: " + responseText);
    }
    
    // 4. Assemble the email body
    var subject = "Top talent hiring at " + company;
    var body = "Hi " + (firstName || "there") + ",\n\n" +
               "I'm Ayush from Butter Search - an executive recruitment firm founded by IIM Calcutta alumni (ex-Naukri, Alvarez & Marsal, PwC).\n\n" +
               result.opener.trim() + "\n\n" +
               "That's where we come in - getting top talent connected with leading fintech and housing finance platforms, working directly with founders, CXOs and business leaders.\n\n" +
               result.closer.trim() + "\n\n" +
               signature;
               
    var containsNumbers = /[\d₹\$%]/i.test(result.opener) || /[\d₹\$%]/i.test(result.closer);
               
    return {
      success: true,
      subject: subject,
      body: body,
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
  
  ui.alert("Sending...", "Please wait while the system finds and sends your drafts. This may take a minute.", ui.ButtonSet.OK);
  
  // Get all drafts in Gmail once to avoid calling getDrafts() in a loop
  var allDrafts = GmailApp.getDrafts();
  var sentCount = 0;
  var notFoundCount = 0;
  
  for (var r = 2; r <= lastRow; r++) {
    var status = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
    if (status === "Draft Created" || status === "Draft Created (Dry Run)") {
      var email = leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim();
      var foundDraft = false;
      
      // Look for a draft addressed to this email OR containing the original recipient tag
      for (var d = 0; d < allDrafts.length; d++) {
        var draft = allDrafts[d];
        var msg = draft.getMessage();
        var draftTo = msg.getTo() || "";
        var draftBody = msg.getBody() || "";
        var draftPlain = msg.getPlainBody() || "";
        
        var isMatch = false;
        if (draftTo.toLowerCase().indexOf(email.toLowerCase()) !== -1) {
          isMatch = true;
        } else if (draftBody.indexOf(email) !== -1 || draftPlain.indexOf(email) !== -1) {
          // Fallback for Test Drafts where the email is in the "Original Recipient: ..." header
          isMatch = true;
        }
        
        if (isMatch) {
          draft.send();
          foundDraft = true;
          
          // Update sheet
          leadsSheet.getRange(r, headersMap["Outreach Status"]).setValue("Email Sent");
          leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Sent");
          leadsSheet.getRange(r, headersMap["Last Sent At"]).setValue(new Date());
          sentCount++;
          
          // Remove from our array so we don't process it again
          allDrafts.splice(d, 1);
          break;
        }
      }
      
      if (!foundDraft) {
        notFoundCount++;
      }
    }
  }
  
  var msg = "Sent " + sentCount + " drafts successfully!";
  if (notFoundCount > 0) {
    msg += "\nCould not find Gmail drafts for " + notFoundCount + " leads (they may have been deleted, altered, or sent manually).";
  }
  ui.alert("Sending Complete", msg, ui.ButtonSet.OK);
}

/**
 * Scans sent emails for replies and automatically sends follow-up emails
 * to leads who haven't replied after N days.
 */
function detectRepliesAndFollowUp() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
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
      ui.alert("Error", "Missing required column: '" + required[i] + "'. Please run Setup Sheets.", ui.ButtonSet.OK);
      return;
    }
  }
  
  var lastRow = leadsSheet.getLastRow();
  var myEmail = Session.getActiveUser().getEmail();
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
        var sentFromAccount = headersMap["Sent From Account"] ? leadsSheet.getRange(r, headersMap["Sent From Account"]).getValue().toString().trim() : "";
        var selectedAccountName = sentFromAccount || getConfigValue(config, "Default Send Account", "Account A").toString().trim();
        var senderEmail = getConfigValue(config, selectedAccountName + " Email", Session.getActiveUser().getEmail());
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
  
  ui.alert(
    "Follow-up Process Complete",
    "Scan complete!\n\n" +
    "- New replies detected: " + repliesDetected + "\n" +
    "- Follow-up emails sent/drafted: " + followUpsSent,
    ui.ButtonSet.OK
  );
}

/**
 * Set up the hourly trigger for automated outreach.
 */
function setupOutreachTrigger() {
  var ui = SpreadsheetApp.getUi();
  clearOutreachTriggerInternal();
  setConfigValue("Sending Active", true);
  
  ScriptApp.newTrigger("runOutreachPipelineHourly")
    .timeBased()
    .everyHours(1)
    .create();
    
  ui.alert("Trigger Activated", "Hourly outreach trigger has been set up successfully. 'Sending Active' is now true. The system will automatically send/draft up to the batch limit every hour.", ui.ButtonSet.OK);
}

/**
 * Deactivate the hourly trigger for automated outreach.
 */
function deactivateOutreachTrigger() {
  var ui = SpreadsheetApp.getUi();
  clearOutreachTriggerInternal();
  setConfigValue("Sending Active", false);
  ui.alert("Trigger Deactivated", "Hourly outreach trigger has been deactivated. 'Sending Active' is now false.", ui.ButtonSet.OK);
}

/**
 * Helper to clear all existing outreach triggers.
 */
function clearOutreachTriggerInternal() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runOutreachPipelineHourly") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Headless entry point for the hourly trigger.
 */
function runOutreachPipelineHourly() {
  processOutreachInternal(true, false, false);
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
