/**
 * AI Lead Generation Engine - Sheet Setup
 * File: Setup.gs
 */

// Define standard Apollo export columns
var APOLLO_COLUMNS = [
  "First Name", "Last Name", "Title", "Company", "Company Name for Emails",
  "Email", "Email Status", "Primary Email Source", "Primary Email Verification Source",
  "Email Confidence", "Primary Email Catch-all Status", "Primary Email Last Verified At",
  "Seniority", "Departments", "Contact Owner", "Work Direct Phone", "Home Phone",
  "Mobile Phone", "Corporate Phone", "Other Phone", "Stage", "Lists", "Last Contacted",
  "Account Owner", "# Employees", "Industry", "Keywords", "Person Linkedin Url", "Website",
  "Company Linkedin Url", "Facebook Url", "Twitter Url", "City", "State", "Country",
  "Company Address", "Company City", "Company State", "Company Country", "Company Phone",
  "Technologies", "Annual Revenue", "Total Funding", "Latest Funding", "Latest Funding Amount",
  "Last Raised At"
];

// Define our new custom tracking columns
var CUSTOM_COLUMNS = [
  "Score", "Score Reason", "Validation Status", "Validation Reason", "Outreach Status", "Hiring Status",
  "Company Validation Status", "Company Validation Reason", "Research Status", "Pipeline Stage", "Last Sent At", "Replied", "Follow-up Status", "Thread Id", "Send From Account",
  "Raw Name Backup", "Score Source", "Enrichment Status",
  "ZB Status", "ZB Score", "Send Priority", "Response Status", "Bounce Type", "Followup 1 Sent Date", "Followup 2 Sent Date", "Followup 1 Due Date", "Followup 2 Due Date", "Followup Cancelled", "Send Date"
];

// Define default config values
var DEFAULT_CONFIG = [
  ["Parameter", "Value", "Description"],
  ["Revenue Cutoff (INR)", "50000000", "Minimum annual revenue threshold in INR (e.g. 50000000 = 5 Crore)"],
  ["Funding Recency Window (Months)", "12", "Maximum age of funding in months to be considered recent"],
  ["Email Confidence Threshold", "0.7", "Minimum confidence score required to avoid being marked 'Risky' (range: 0 to 1)"],
  ["Gemini Model", "deepseek-ai/deepseek-r1", "Model to use for scoring (e.g., deepseek-ai/deepseek-r1, zai-glm-4.7, or gemini-2.0-flash)"],
  ["Gemini Temperature", "0.2", "Sampling temperature for the scoring API (lower is more consistent)"],
  ["Scoring Prompt Template", 
   "You are an expert B2B lead scoring agent. Analyze the provided lead data and assign a qualification score (1-10) and category based on the following rules:\n\n" +
   "CRITERIA:\n" +
   "1. Revenue Fit: Annual revenue above ₹5 Crore ({Revenue Cutoff} INR) is a strong positive signal.\n" +
   "2. Funding Recency: Any funding event within the last {Funding Recency Window} months is a strong positive signal (check Latest Funding and Last Raised At).\n" +
   "3. Funding Scale: Check Total Funding and Latest Funding Amount. Significant funding scale is a positive indicator.\n" +
   "4. Hiring Intent: Active job postings or recruitment signals for software/tech roles indicate high growth and immediate need (check Hiring Status).\n" +
   "5. Size & Stack: Estimate company size using the number of Employees and Technologies used.\n\n" +
   "ESTIMATION RULES:\n" +
   "- If explicit Annual Revenue or Funding data is blank, estimate the company size and revenue potential using the number of Employees ({Employees}), Technologies ({Technologies}), Location ({Location}), and your own knowledge of the company.\n" +
   "- For example, if a company has more than 50 employees, it is likely above ₹5 Crore in revenue (Score 8+). If it is a massive public/private enterprise (like the MTA with thousands of employees), score it highly as Qualified (Score 8-10).\n" +
   "- DO NOT return 'insufficient data' or generic text. Always provide a numeric score (1-10) and a category, using your knowledge of the company as a fallback.\n\n" +
   "SCORING RULES:\n" +
   "- Qualified (Score 8-10): Meets the revenue cutoff (>= ₹5 Crore) AND/OR has a recent funding event (within {Funding Recency Window} months) with high funding scale, or is a medium-to-large enterprise.\n" +
   "- Nurture (Score 4-7): Moderate fit. E.g., has decent revenue/size but no recent funding, or has recent funding but is very early-stage with low revenue/size.\n" +
   "- Disqualified (Score 1-3): Revenue/size is well below the cutoff AND there are no recent funding events.\n\n" +
   "LEAD DATA:\n" +
   "Company: {Company}\n" +
   "Annual Revenue: {Annual Revenue}\n" +
   "Total Funding: {Total Funding}\n" +
   "Latest Funding: {Latest Funding}\n" +
   "Latest Funding Amount: {Latest Funding Amount}\n" +
   "Last Raised At: {Last Raised At}\n" +
   "Hiring Status: {Hiring Status}\n" +
   "Employees: {Employees}\n" +
   "Technologies: {Technologies}\n" +
   "Location: {Location}\n\n" +
   "OUTPUT FORMAT:\n" +
   "Respond ONLY with a JSON object in this format:\n" +
   "{\n" +
   "  \"score\": <1-10 integer>,\n" +
   "  \"category\": \"qualified|nurture|disqualified\",\n" +
   "  \"reason\": \"<Provide a concise, 1-sentence explanation focusing on the revenue, company size, hiring/funding scale signals>\"\n" +
   "}", 
   "Template used for prompt generation. Do not change placeholders in curly brackets."
  ],
  ["Outreach Mode", "Draft", "Outreach generation mode: 'Draft' to create Gmail drafts (recommended for safety) or 'Send' to send emails directly."],
  ["Test Email Recipient", "", "Optional: Redirect all outreach emails to this address for testing. Leave blank to use the lead's email."],
  ["Outreach Subject Template", "Quick question for {Company}", "Subject line template for outreach emails. Placeholder {Company} is supported."],
  ["Outreach Prompt Template",
   "You are an expert sales representative. Generate a short, highly personalized B2B cold email to a prospect.\n\n" +
   "PROSPECT DATA:\n" +
   "First Name: {First Name}\n" +
   "Last Name: {Last Name}\n" +
   "Title: {Title}\n" +
   "Company: {Company}\n" +
   "Annual Revenue: {Annual Revenue}\n" +
   "Latest Funding: {Latest Funding}\n\n" +
   "INSTRUCTIONS:\n" +
   "1. Keep the email under 120 words.\n" +
   "2. Focus on how we can help them scale based on their revenue or recent funding.\n" +
   "3. Tone should be professional, direct, and conversational (no corporate jargon or fake flattery).\n" +
   "4. End with a soft call to action asking for a brief chat.\n" +
   "5. Do not include subject line or placeholder brackets like [Your Name] in the email body.",
    "Prompt template used to generate the email body using DeepSeek."
  ],
  ["Apollo Ingest Job Titles", "CTO, Co-Founder, Founder, CEO", "Comma-separated list of target job titles for lead ingestion"],
  ["Apollo Ingest Locations", "India, United States", "Comma-separated list of target locations (cities, states, or countries)"],
  ["Apollo Ingest Industries", "Software, Technology, Financial Services", "Comma-separated list of target industries for organization filter"],
  ["Apollo Ingest Limit", "10", "Number of leads to import per run (maximum 100)"],
  ["Google CSE ID", "", "Google Custom Search Engine ID (CX) for LinkedIn X-Ray search"],
  ["GitHub API Key (Optional)", "", "GitHub Personal Access Token for search (increases rate limit)"],
  ["Google Maps API Key", "", "Google Places API Key (can be the same as Google Search API Key)"],
  ["Hunter.io API Key (Optional)", "", "Hunter.io API key for second-pass email finding and verification"],
  ["ZeroBounce API Key (Optional)", "", "ZeroBounce API key for second-pass email verification"],
  ["Emails Per Run", "10", "Number of emails to send per hourly batch run (recommended: 10)"],
  ["Follow-up Delay (Days)", "3", "Number of days to wait before sending a follow-up email if no reply is received"],
  ["Account A Email", "sender1@company.com", "Email address for Account A"],
  ["Account A Label", "Ayush - Butter Search", "Display name for Account A"],
  ["Account A Signature", "Best regards,\nAyush\nButter Search", "Email signature for Account A"],
  ["Account B Email", "sender2@company.com", "Email address for Account B"],
  ["Account B Label", "Harshith - Butter Search", "Display name for Account B"],
  ["Account B Signature", "Best regards,\nHarshith\nButter Search", "Email signature for Account B"],
  ["Default Send Account", "Account A", "Default account to use for sending (Account A or Account B)"],
  // --- Batch 2 additions ---
  ["Recruitment Industry Keywords", "recruitment,staffing,talent acquisition,executive search,hr consulting,headhunting,recruiting,staffing agency", "Comma-separated keywords. Any lead whose Industry/Keywords matches is hard-blocked from outreach regardless of score (Feature 4)."],
  ["Score Band High Accounts", "Account A,Account B", "Comma-separated account pool for leads scoring >= 8 (Feature 5). One is chosen at random, respecting send caps."],
  ["Score Band Mid Accounts", "Account C,Account D,Account E", "Comma-separated account pool for leads scoring 6-7.9 (Feature 5)."],
  ["Score Band Low Behavior", "Default", "What to do with leads scoring < 6 for account routing: 'Default' (use Default Send Account) or 'Hold' (do not route) (Feature 5)."],
  ["Per Account Daily Cap", "40", "Maximum emails sent per account per day (Feature 5). If an account hits this cap, the system auto-routes to another account in the same band. Gmail free limit is ~500/day; Workspace is ~2000/day. 40 is a safe conservative default."],
  ["Draft Quality Threshold", "7", "Ready to Send tab: drafts scoring above this (1-10) default to Ready for Send = TRUE (Feature 8)."],
  ["Sender Source", "MainSheet", "Where the hourly sender pulls candidates: 'MainSheet' (existing behavior) or 'ReadyTab' (Ready to Send tab, Feature 8)."],
  ["Account C Email", "", "Email address for Account C (Feature 5 scaffold — fill in to activate)."],
  ["Account C Label", "", "Display name for Account C."],
  ["Account C Signature", "", "Email signature for Account C."],
  ["Account D Email", "", "Email address for Account D (Feature 5 scaffold — fill in to activate)."],
  ["Account D Label", "", "Display name for Account D."],
  ["Account D Signature", "", "Email signature for Account D."],
  ["Account E Email", "", "Email address for Account E (Feature 5 scaffold — fill in to activate)."],
  ["Account E Label", "", "Display name for Account E."],
  ["Account E Signature", "", "Email signature for Account E."]
];

function setupSheetStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Setup "Leads" Sheet
    var leadsSheet = ss.getSheetByName("Leads");
    if (!leadsSheet) {
      leadsSheet = ss.insertSheet("Leads");
      var fullHeaders = APOLLO_COLUMNS.concat(CUSTOM_COLUMNS);
      leadsSheet.appendRow(fullHeaders);
      formatHeaderRow(leadsSheet);
      Logger.log("Created 'Leads' sheet with default headers.");
    } else {
      // If leads sheet exists, make sure custom columns exist and append if missing
      var existingHeaders = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
      var missingColumns = [];
      
      CUSTOM_COLUMNS.forEach(function(col) {
        if (existingHeaders.indexOf(col) === -1) {
          missingColumns.push(col);
        }
      });
      
      if (missingColumns.length > 0) {
        var startCol = leadsSheet.getLastColumn() + 1;
        leadsSheet.getRange(1, startCol, 1, missingColumns.length).setValues([missingColumns]);
        formatHeaderRow(leadsSheet);
        Logger.log("Added missing columns to 'Leads' sheet: " + missingColumns.join(", "));
      } else {
        Logger.log("'Leads' sheet already contains all tracking columns.");
      }
    }
    
    // Apply dropdown validations to fixed-option columns
    applyLeadDropdownValidations(leadsSheet);
    Logger.log("Applied dropdown validations on Leads sheet.");
    
    // 2. Setup "Config" Sheet
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
      configSheet = ss.insertSheet("Config");
      configSheet.getRange(1, 1, DEFAULT_CONFIG.length, 3).setValues(DEFAULT_CONFIG);
      formatHeaderRow(configSheet);
      configSheet.autoResizeColumns(1, 3);
      Logger.log("Created 'Config' sheet with default parameters.");
    } else {
      Logger.log("'Config' sheet already exists. Checking for missing parameters...");
      var configRows = configSheet.getDataRange().getValues();
      var existingKeys = configRows.map(function(row) { return row[0].toString().trim(); });
      
      for (var i = 1; i < DEFAULT_CONFIG.length; i++) {
        var defaultKey = DEFAULT_CONFIG[i][0];
        if (existingKeys.indexOf(defaultKey) === -1) {
          configSheet.appendRow(DEFAULT_CONFIG[i]);
          Logger.log("Added missing config key: " + defaultKey);
        }
      }
    }
    
    // Automatically set the model parameter to deepseek-ai/deepseek-r1 and configure formatting
    var finalConfigRows = configSheet.getDataRange().getValues();
    for (var i = 0; i < finalConfigRows.length; i++) {
      var key = finalConfigRows[i][0].toString().trim();
      
      if (key === "Gemini Model") {
        configSheet.getRange(i + 1, 2).setValue("deepseek-ai/deepseek-r1");
        Logger.log("Automatically set Model to deepseek-ai/deepseek-r1.");
      }
      if (key.indexOf("Prompt Template") !== -1) {
        configSheet.getRange(i + 1, 2).setWrap(true);
      }
    }
    
    configSheet.autoResizeColumns(1, 3);
    
    // 5. Setup "Log" Sheet
    var logSheet = ss.getSheetByName("Log");
    if (!logSheet) {
      logSheet = ss.insertSheet("Log");
      logSheet.appendRow(["Timestamp", "Rows Processed", "Rows Scored", "Validation Flags", "Errors"]);
      logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d1d5db");
      logSheet.setColumnWidth(1, 150);
      logSheet.setColumnWidth(5, 400);
      Logger.log("Created 'Log' sheet.");
    } else {
      Logger.log("'Log' sheet already exists.");
    }
    
    // 6. Setup "Ready to Send" Sheet (Feature 8)
    var readySheet = ss.getSheetByName("Ready to Send");
    if (!readySheet) {
      readySheet = ss.insertSheet("Ready to Send");
      readySheet.appendRow([
        "Draft ID", "Timestamp", "Company", "First Name", "Last Name", "Email", 
        "Selected Account", "Score", "Draft Quality Score", "Subject", "Body", "Ready for Send"
      ]);
      formatHeaderRow(readySheet);
      // Add checkboxes to the Ready for Send column
      readySheet.getRange(2, 12, 999, 1).insertCheckboxes();
      Logger.log("Created 'Ready to Send' sheet.");
    }

    // 7. Setup "Account Config" Sheet (Email Ops Upgrade)
    var accountConfigSheet = ss.getSheetByName("Account Config");
    if (!accountConfigSheet) {
      accountConfigSheet = ss.insertSheet("Account Config");
      accountConfigSheet.appendRow([
        "Account", "Daily Quota", "Sent Today Count", "Last Reset Date", "Quota Confirmed"
      ]);
      formatHeaderRow(accountConfigSheet);
      // Populate defaults from Account A..E if they exist
      accountConfigSheet.appendRow(["Account A", 40, 0, "", false]);
      accountConfigSheet.appendRow(["Account B", 40, 0, "", false]);
      
      // Add checkboxes to the Quota Confirmed column
      accountConfigSheet.getRange(2, 5, 999, 1).insertCheckboxes();
      
      accountConfigSheet.autoResizeColumns(1, 5);
      Logger.log("Created 'Account Config' sheet.");
    } else {
       var configHeaders = accountConfigSheet.getRange(1, 1, 1, accountConfigSheet.getLastColumn()).getValues()[0];
       if (configHeaders.indexOf("Quota Confirmed") === -1) {
          var startCol = accountConfigSheet.getLastColumn() + 1;
          accountConfigSheet.getRange(1, startCol, 1, 1).setValues([["Quota Confirmed"]]);
          accountConfigSheet.getRange(2, startCol, accountConfigSheet.getMaxRows() - 1, 1).insertCheckboxes();
          formatHeaderRow(accountConfigSheet);
          Logger.log("Added 'Quota Confirmed' column to 'Account Config' sheet.");
       }
    }

    ui.alert("Setup Complete", "All sheets and standard structures have been initialized.", ui.ButtonSet.OK);
    
    // 3b. Setup "Templates" Sheet
    var templatesSheet = ss.getSheetByName("Templates");
    var defaultTemplates = [
      ["Template Type", "Subject Template", "Body Template / Reference", "Preferred Account"],
      ["Style Reference", "Top talent hiring at {Company}", 
       "Hi {First Name},\n\n" +
       "I'm Ayush from Butter Search - an executive recruitment firm founded by IIM Calcutta alumni (ex-Naukri, Alvarez & Marsal, PwC).\n\n" +
       "I've been following {Company}'s impressive journey in {Industry/Keywords}. As you gear up for your next phase of growth, having the right set of people in place becomes critical.\n\n" +
       "That's where we come in - getting top talent connected with leading fintech and housing finance platforms, working directly with founders, CXOs and business leaders.\n\n" +
       "Would you be open to a quick connect to explore how we can support your hiring needs?",
       "Account A"],
      ["Style Reference", "Top talent hiring at {Company}", 
       "Hi {First Name},\n\n" +
       "I'm Harshith from Butter Search - an executive recruitment firm founded by IIM Calcutta alumni.\n\n" +
       "I've been following {Company}'s impressive journey in {Industry/Keywords}. As you gear up for your next phase of growth, having the right set of people in place becomes critical.\n\n" +
       "That's where we come in - getting top talent connected with leading fintech and housing finance platforms, working directly with founders, CXOs and business leaders.\n\n" +
       "Would you be open to a quick connect to explore how we can support your hiring needs?",
       "Account B"],
      ["Follow-up Template", "Re: Top talent hiring at {Company}",
       "Hi {First Name},\n\n" +
       "Just following up on my earlier note \u2014 completely understand things get busy.\n\n" +
       "{AI_INSIGHT}\n\n" +
       "Happy to connect for a quick call to explore how we might partner effectively.\n\n" +
       "Looking forward to hearing from you.",
       "Account A"],
      ["Follow-up Template", "Re: Top talent hiring at {Company}",
       "Hi {First Name},\n\n" +
       "Just following up on my earlier note \u2014 completely understand things get busy.\n\n" +
       "{AI_INSIGHT}\n\n" +
       "Happy to connect for a quick call to explore how we might partner effectively.\n\n" +
       "Looking forward to hearing from you.",
       "Account B"]
    ];
    if (!templatesSheet) {
      templatesSheet = ss.insertSheet("Templates");
      templatesSheet.getRange(1, 1, defaultTemplates.length, 4).setValues(defaultTemplates);
      formatHeaderRow(templatesSheet);
      templatesSheet.getRange(2, 3, 2, 1).setWrap(true);
      templatesSheet.autoResizeColumns(1, 4);
      Logger.log("Created 'Templates' sheet.");
    } else {
      Logger.log("'Templates' sheet already exists.");
    }
    
    // 4. Setup "Dashboard" Sheet
    setupDashboardStructure(ss);
    Logger.log("Created 'Dashboard' sheet.");
    
    ui.alert("Success", "Sheet structure has been set up successfully!\n\n1. 'Dashboard' tab is ready with analytics.\n2. 'Leads' tab is ready.\n3. 'Config' tab has been updated with the new detailed templates.\n4. 'Log' tab is ready.", ui.ButtonSet.OK);
    
  } catch(e) {
    Logger.log("Error setting up sheet structure: " + e.toString());
    ui.alert("Error", "Failed to set up sheets: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Utility to style the header row of a sheet for premium visual design.
 * Highlights essential/required columns with a different color on the Leads sheet.
 */
function formatHeaderRow(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  
  // Set clipping wrap strategy for the entire sheet so text doesn't overflow into adjacent empty columns
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#f3f4f6"); // Cool grey background by default
  headerRange.setFontColor("#1f2937"); // Dark charcoal font
  headerRange.setBorder(null, null, true, null, null, null, "#d1d5db", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 26);

  // If this is the Leads sheet, highlight the bare minimum essential columns
  if (sheet.getName() === "Leads") {
    var headers = headerRange.getValues()[0];
    var essentialCols = [
      "First Name", "Company", "Email", 
      "Annual Revenue", "# Employees", 
      "Total Funding", "Latest Funding", 
      "Hiring Status"
    ];
    
    for (var c = 0; c < headers.length; c++) {
      if (essentialCols.indexOf(headers[c].toString().trim()) !== -1) {
        // Highlight with a soft yellow to indicate importance/required
        sheet.getRange(1, c + 1).setBackground("#fef08a");
      }
    }
  }
}

/**
 * Menu wrapper to re-apply header formatting manually.
 */
function applyHeaderFormattingMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (leadsSheet) {
    formatHeaderRow(leadsSheet);
    SpreadsheetApp.getUi().alert("Success", "Essential columns have been highlighted in yellow.", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("Error", "Leads sheet not found.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Sets up a beautiful and premium visual Analytics Dashboard in Google Sheets.
 */
function setupDashboardStructure(ss) {
  var ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName("Dashboard");
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet("Dashboard", 0); // Put it first
  }
  
  // Set grid lines
  dashboardSheet.setHiddenGridlines(false);
  
  // Reset sheet to prevent duplicate ranges overlapping
  dashboardSheet.clear();
  dashboardSheet.clearFormats();
  
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) return;
  
  var headersMap = {};
  var headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    headersMap[headers[i]] = i + 1;
  }
  
  // Get Column Letters for key metrics
  var scoreCol = getColumnLetter(headersMap["Score"] || 47);
  var emailCol = getColumnLetter(headersMap["Email"] || 6);
  var outreachCol = getColumnLetter(headersMap["Outreach Status"] || 50);
  var websiteCol = getColumnLetter(headersMap["Website"] || 29);
  var emailStatusCol = getColumnLetter(headersMap["Email Status"] || 7);
  
  // Set Column Widths for spacing
  dashboardSheet.setColumnWidth(1, 20); // Column A is spacing spacer
  dashboardSheet.setColumnWidth(2, 220); // B
  dashboardSheet.setColumnWidth(3, 110); // C
  dashboardSheet.setColumnWidth(4, 20);  // D (spacer)
  dashboardSheet.setColumnWidth(5, 220); // E
  dashboardSheet.setColumnWidth(6, 110); // F
  dashboardSheet.setColumnWidth(7, 20);  // G (spacer)
  dashboardSheet.setColumnWidth(8, 220); // H
  dashboardSheet.setColumnWidth(9, 140); // I
  
  // Row Heights
  dashboardSheet.setRowHeight(2, 40); // Title Header
  dashboardSheet.setRowHeight(4, 25); // KPI Title
  dashboardSheet.setRowHeight(5, 45); // KPI Value
  
  // Title Block
  var titleRange = dashboardSheet.getRange("B2:I2");
  titleRange.merge();
  titleRange.setValue("AI LEAD GENERATION ENGINE - REAL-TIME DASHBOARD");
  titleRange.setFontWeight("bold");
  titleRange.setFontSize(14);
  titleRange.setBackground("#0F172A"); // Slate-900 (Obsidian dark theme)
  titleRange.setFontColor("#F8FAFC"); // Slate-50
  titleRange.setHorizontalAlignment("center");
  titleRange.setVerticalAlignment("middle");
  
  // --- KPI CARDS ---
  // We place 3 KPI cards side by side in row 4 & 5
  
  // Card 1: Total Leads
  setupKpiCard(dashboardSheet, "B4:C4", "B5:C5", "TOTAL LEADS INGESTED", "=COUNTA(Leads!A2:A)");
  // Card 2: Scored Leads
  setupKpiCard(dashboardSheet, "E4:F4", "E5:F5", "LEADS EVALUATED (AI)", "=COUNT(Leads!" + scoreCol + "2:" + scoreCol + ")");
  // Card 3: Qualified Leads
  setupKpiCard(dashboardSheet, "H4:I4", "H5:I5", "QUALIFIED LEADS (SCORE >= 8)", "=COUNTIF(Leads!" + scoreCol + "2:" + scoreCol + ", \">=8\")");
  
  // --- TABLES SECTION (Row 7 to 15) ---
  
  // 1. Table: Lead Quality Distribution (B7:C14)
  dashboardSheet.getRange("B7:C7").merge().setValue("LEAD QUALITY SEGMENTATION").setFontWeight("bold").setBackground("#334155").setFontColor("#F8FAFC").setHorizontalAlignment("center");
  dashboardSheet.getRange("B8").setValue("Segment");
  dashboardSheet.getRange("C8").setValue("Count");
  
  dashboardSheet.getRange("B9").setValue("Qualified (Score 8-10)");
  dashboardSheet.getRange("C9").setValue("=COUNTIF(Leads!" + scoreCol + "2:" + scoreCol + ", \">=8\")");
  
  dashboardSheet.getRange("B10").setValue("Nurture (Score 4-7)");
  dashboardSheet.getRange("C10").setValue("=COUNTIFS(Leads!" + scoreCol + "2:" + scoreCol + ", \">=4\", Leads!" + scoreCol + "2:" + scoreCol + ", \"<=7\")");
  
  dashboardSheet.getRange("B11").setValue("Disqualified (Score 1-3)");
  dashboardSheet.getRange("C11").setValue("=COUNTIFS(Leads!" + scoreCol + "2:" + scoreCol + ", \">=1\", Leads!" + scoreCol + "2:" + scoreCol + ", \"<=3\")");
  
  dashboardSheet.getRange("B12").setValue("Pending AI Scoring");
  dashboardSheet.getRange("C12").setValue("=COUNTIFS(Leads!A2:A, \"<>\", Leads!" + scoreCol + "2:" + scoreCol + ", \"\")");
  
  dashboardSheet.getRange("B13").setValue("Total Scored");
  dashboardSheet.getRange("C13").setValue("=SUM(C9:C11)");
  
  // Format Lead Quality table
  var qualityTableRange = dashboardSheet.getRange("B7:C13");
  qualityTableRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.getRange("B8:C8").setFontWeight("bold").setBackground("#E2E8F0");
  dashboardSheet.getRange("B13:C13").setFontWeight("bold").setBackground("#F1F5F9");
  dashboardSheet.getRange("C9:C13").setHorizontalAlignment("right");
  
  // 2. Table: Outreach & Open Tracking (E7:F14)
  dashboardSheet.getRange("E7:F7").merge().setValue("OUTREACH & INTERACTION").setFontWeight("bold").setBackground("#0284C7").setFontColor("#F8FAFC").setHorizontalAlignment("center");
  dashboardSheet.getRange("E8").setValue("Outreach Action");
  dashboardSheet.getRange("F8").setValue("Stats");
  
  dashboardSheet.getRange("E9").setValue("Emails Sent");
  dashboardSheet.getRange("F9").setValue("=COUNTIF(Leads!" + outreachCol + "2:" + outreachCol + ", \"Email Sent\")");
  
  dashboardSheet.getRange("E10").setValue("Drafts Created");
  dashboardSheet.getRange("F10").setValue("=COUNTIF(Leads!" + outreachCol + "2:" + outreachCol + ", \"Draft Created\")");
  
  // Format Outreach table
  var outreachTableRange = dashboardSheet.getRange("E7:F10");
  outreachTableRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.getRange("E8:F8").setFontWeight("bold").setBackground("#E2E8F0");
  dashboardSheet.getRange("F9:F10").setHorizontalAlignment("right");
  
  // 3. Table: Cost Estimations (H7:I14)
  dashboardSheet.getRange("H7:I7").merge().setValue("ESTIMATED API SPEND (USD)").setFontWeight("bold").setBackground("#059669").setFontColor("#F8FAFC").setHorizontalAlignment("center");
  dashboardSheet.getRange("H8").setValue("Provider/API");
  dashboardSheet.getRange("I8").setValue("Est. Cost");
  
  dashboardSheet.getRange("H9").setValue("DeepSeek API (Scoring/Outreach)");
  dashboardSheet.getRange("I9").setValue("=(COUNT(Leads!" + scoreCol + "2:" + scoreCol + ") + COUNTIF(Leads!" + outreachCol + "2:" + outreachCol + ", \"*Sent*\") + COUNTIF(Leads!" + outreachCol + "2:" + outreachCol + ", \"*Draft*\")) * 0.00015");
  
  dashboardSheet.getRange("H10").setValue("Google Maps Ingest");
  dashboardSheet.getRange("I10").setValue("=COUNTIF(Leads!" + websiteCol + "2:" + websiteCol + ", \"<>\") * 0.017");
  
  dashboardSheet.getRange("H11").setValue("Hunter.io Validation");
  dashboardSheet.getRange("I11").setValue("=COUNTIF(Leads!" + emailStatusCol + "2:" + emailStatusCol + ", \"*Hunter*\") * 0.002");
  
  dashboardSheet.getRange("H12").setValue("ZeroBounce Validation");
  dashboardSheet.getRange("I12").setValue("=COUNTIF(Leads!" + emailStatusCol + "2:" + emailStatusCol + ", \"*ZeroBounce*\") * 0.0008");
  
  dashboardSheet.getRange("H13").setValue("Total Est. Spend");
  dashboardSheet.getRange("I13").setValue("=SUM(I9:I12)");
  
  // Format Cost table
  var costTableRange = dashboardSheet.getRange("H7:I13");
  costTableRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.getRange("H8:I8").setFontWeight("bold").setBackground("#E2E8F0");
  dashboardSheet.getRange("H13:I13").setFontWeight("bold").setBackground("#F1F5F9");
  dashboardSheet.getRange("I9:I13").setNumberFormat("$#,##0.000").setHorizontalAlignment("right");
  
  // 4. Table: Pipeline Stage Breakdown (B15:C27)
  var pipelineCol = getColumnLetter(headersMap["Pipeline Stage"] || 46);
  dashboardSheet.getRange("B15:C15").merge().setValue("PIPELINE STAGE BREAKDOWN").setFontWeight("bold").setBackground("#475569").setFontColor("#F8FAFC").setHorizontalAlignment("center");
  dashboardSheet.getRange("B16").setValue("Stage");
  dashboardSheet.getRange("C16").setValue("Count");
  
  var stages = [
    ["Ingested", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Ingested\")"],
    ["Company Verified", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Company Verified\")"],
    ["Needs Review", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Needs Review\")"],
    ["Scored", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Scored\")"],
    ["Validated", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Validated\")"],
    ["Email Flagged", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Email Flagged\")"],
    ["Nurture", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Nurture\")"],
    ["Disqualified", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Disqualified\")"],
    ["Draft Created", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Draft Created\")"],
    ["Sent", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Sent\")"],
    ["Replied", "=COUNTIF(Leads!" + pipelineCol + "2:" + pipelineCol + ", \"Replied\")"]
  ];
  
  for (var i = 0; i < stages.length; i++) {
    dashboardSheet.getRange(17 + i, 2).setValue(stages[i][0]);
    dashboardSheet.getRange(17 + i, 3).setValue(stages[i][1]);
  }
  
  var pipelineTableRange = dashboardSheet.getRange("B15:C" + (16 + stages.length));
  pipelineTableRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.getRange("B16:C16").setFontWeight("bold").setBackground("#E2E8F0");
  dashboardSheet.getRange("C17:C" + (16 + stages.length)).setHorizontalAlignment("right");
}

/**
 * Helper to setup a KPI Card
 */
function setupKpiCard(sheet, titleRangeStr, valueRangeStr, title, formula) {
  var titleRange = sheet.getRange(titleRangeStr);
  titleRange.merge();
  titleRange.setValue(title);
  titleRange.setFontWeight("bold");
  titleRange.setFontSize(9);
  titleRange.setFontColor("#64748B"); // Slate-500
  titleRange.setBackground("#F8FAFC"); // Slate-50
  titleRange.setHorizontalAlignment("center");
  titleRange.setVerticalAlignment("bottom");
  
  var valueRange = sheet.getRange(valueRangeStr);
  valueRange.merge();
  valueRange.setValue(formula);
  valueRange.setFontWeight("bold");
  valueRange.setFontSize(20);
  valueRange.setFontColor("#0F172A"); // Slate-900
  valueRange.setBackground("#F8FAFC"); // Slate-50
  valueRange.setHorizontalAlignment("center");
  valueRange.setVerticalAlignment("top");
  
  // Draw card border
  var cardRange = sheet.getRange(titleRange.getRow(), titleRange.getColumn(), 2, titleRange.getNumColumns());
  cardRange.setBorder(true, true, true, true, false, false, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Converts a 1-indexed column number to standard A-Z letter reference.
 */
function getColumnLetter(colIndex) {
  var temp, letter = "";
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

/**
 * Automates the migration process to the dual account setup.
 * Appends new config keys and modifies the Templates tab to support Account A and B.
 */
function upgradeToDualAccount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Update Config
  var configSheet = ss.getSheetByName("Config");
  var newConfigs = [
    ["Account A Email", "sender1@company.com", "Email address for Account A"],
    ["Account A Label", "Ayush - Butter Search", "Display name for Account A"],
    ["Account A Signature", "Best regards,\nAyush\nButter Search", "Email signature for Account A"],
    ["Account B Email", "sender2@company.com", "Email address for Account B"],
    ["Account B Label", "Harshith - Butter Search", "Display name for Account B"],
    ["Account B Signature", "Best regards,\nHarshith\nButter Search", "Email signature for Account B"],
    ["Default Send Account", "Account A", "Default account to use for sending (Account A or Account B)"]
  ];
  
  if (configSheet) {
    var configRows = configSheet.getDataRange().getValues();
    var existingKeys = configRows.map(function(row) { return row[0].toString().trim(); });
    for (var i = 0; i < newConfigs.length; i++) {
      if (existingKeys.indexOf(newConfigs[i][0]) === -1) {
        configSheet.appendRow(newConfigs[i]);
      }
    }
  }
  
  // 2. Update Templates
  var templatesSheet = ss.getSheetByName("Templates");
  if (templatesSheet) {
    var lastRow = templatesSheet.getLastRow();
    
    // Add 'Preferred Account' header if missing
    var existingHeader = templatesSheet.getRange(1, 4).getValue();
    if (existingHeader !== "Preferred Account") {
      templatesSheet.getRange(1, 4).setValue("Preferred Account");
      templatesSheet.getRange(1, 4).setFontWeight("bold").setBackground("#f3f4f6");
    }
    
    // Default the existing ones to Account A if they are blank
    var rows = templatesSheet.getDataRange().getValues();
    var hasAccountB = false;
    for (var r = 1; r < rows.length; r++) {
      if (!rows[r][3]) {
        templatesSheet.getRange(r + 1, 4).setValue("Account A");
      }
      if (rows[r][3] === "Account B") {
        hasAccountB = true;
      }
    }
    
    // If we don't have Account B templates, add them
    if (!hasAccountB) {
      templatesSheet.appendRow([
        "Style Reference", 
        "Top talent hiring at {Company}", 
        "Hi {First Name},\n\nI'm Harshith from Butter Search - an executive recruitment firm founded by IIM Calcutta alumni.\n\nI've been following {Company}'s impressive journey in {Industry/Keywords}. As you gear up for your next phase of growth, having the right set of people in place becomes critical.\n\nThat's where we come in - getting top talent connected with leading fintech and housing finance platforms, working directly with founders, CXOs and business leaders.\n\nWould you be open to a quick connect to explore how we can support your hiring needs?",
        "Account B"
      ]);
      templatesSheet.appendRow([
        "Follow-up Template", 
        "Re: Top talent hiring at {Company}",
        "Hi {First Name},\n\nI wanted to quickly follow up on my previous email. I know you're busy, but I'd love to see if you have 5 minutes for a quick chat about supporting your hiring needs at {Company}.\n\nBest,\nHarshith",
        "Account B"
      ]);
      templatesSheet.getRange(templatesSheet.getLastRow() - 1, 3, 2, 1).setWrap(true);
    }
  }
  
  SpreadsheetApp.getUi().alert("Upgrade Complete", "Config and Templates tabs have been automatically updated for dual accounts!\n\nPlease fill out your actual Account A and Account B email addresses in the Config tab.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Applies dropdown data validations to all fixed-option columns in the Leads sheet.
 * This makes it easy to manually update lead statuses without typing.
 * Safe to call multiple times — existing values are preserved.
 *
 * @param {Sheet} leadsSheet The Leads sheet object
 */
function applyLeadDropdownValidations(leadsSheet) {
  leadsSheet = leadsSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads");
  if (!leadsSheet) return;

  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = Math.max(leadsSheet.getMaxRows(), 1000); // Apply to up to 1000 rows for future leads

  // Helper: build and apply a dropdown rule to an entire column (data rows only, starting row 2)
  function applyDropdown(colName, options) {
    var colIndex = headersMap[colName];
    if (!colIndex) return; // Column doesn't exist yet, skip
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)     // true = show dropdown arrow
      .setAllowInvalid(true)                 // Allow existing/script-written values that are not in the list
      .setHelpText("Select a value from the list or leave blank.")
      .build();
    leadsSheet.getRange(2, colIndex, lastRow - 1, 1).setDataValidation(rule);
    Logger.log("Applied dropdown to column: " + colName);
  }

  // --- Column: Company Validation Status ---
  applyDropdown("Company Validation Status", [
    "Verified",
    "Needs Review",
    "Rejected"
  ]);

  // --- Column: Validation Status (Email Quality) ---
  applyDropdown("Validation Status", [
    "Ready",
    "Risky",
    "Skipped (Nurture)",
    "Skipped (Disqualified)",
    "Flagged"
  ]);

  // --- Column: Outreach Status ---
  applyDropdown("Outreach Status", [
    "Ready for outreach",
    "Draft Created",
    "Email Sent",
    "Nurture",
    "Low Priority",
    "Flagged - email risk"
  ]);

  // --- Column: Pipeline Stage ---
  applyDropdown("Pipeline Stage", [
    "Ingested",
    "Company Verified",
    "Needs Review",
    "Scored",
    "Validated",
    "Email Flagged",
    "Nurture",
    "Disqualified",
    "Draft Created",
    "Sent",
    "Replied"
  ]);

  // --- Column: Replied ---
  applyDropdown("Replied", [
    "Yes",
    "No"
  ]);

  // --- Column: Follow-up Status ---
  applyDropdown("Follow-up Status", [
    "Pending",
    "Sent",
    "Skipped"
  ]);

  // --- Column: Send From Account ---
  applyDropdown("Send From Account", [
    "Account A",
    "Account B",
    "Account C",
    "Account D",
    "Account E"
  ]);

  // --- Column: Hiring Status ---
  applyDropdown("Hiring Status", [
    "Hiring",
    "Not Hiring",
    "Unknown",
    "Skipped (No API Config)"
  ]);

  // --- Column: Research Status ---
  applyDropdown("Research Status", [
    "Pending",
    "Done",
    "Skipped"
  ]);

  // --- Email Ops Upgrade Enums ---
  applyDropdown("ZB Status", [
    "verified",
    "catch-all",
    "invalid",
    "unknown",
    "spamtrap",
    "abuse",
    "do_not_mail"
  ]);
  
  applyDropdown("Send Priority", [
    "high",
    "capped",
    "blocked"
  ]);
  
  applyDropdown("Response Status", [
    "none",
    "replied",
    "bounced",
    "out_of_office"
  ]);
  
  applyDropdown("Bounce Type", [
    "none",
    "hard",
    "soft"
  ]);

  // --- Column: Score (Feature 3) ---
  // Score is an OPEN INPUT (blank = run AI scoring; 1-10 = manual override).
  // Apply a numeric 1-10 validation but allow invalid so the script and manual
  // entries are never blocked; clear any legacy dropdown rule on this column.
  if (headersMap["Score"]) {
    var scoreRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 10)
      .setAllowInvalid(true)
      .setHelpText("Leave blank to let AI score this lead, or enter a number 1-10 to set the score manually.")
      .build();
    leadsSheet.getRange(2, headersMap["Score"], lastRow - 1, 1).setDataValidation(scoreRule);
  }

  // --- Hidden internal columns (Batch 2) ---
  ["Raw Name Backup", "Score Source"].forEach(function(col) {
    if (headersMap[col]) {
      leadsSheet.hideColumns(headersMap[col]);
    }
  });

  Logger.log("All dropdown validations applied successfully.");
}
