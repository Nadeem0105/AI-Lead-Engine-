/**
 * AI Lead Generation Engine - Apollo.io Ingestion
 * File: Ingestion.gs
 */

/**
 * Main entry point for the lead ingestion menu action.
 * Queries Apollo.io for new contacts matching the ICP filters in the config,
 * enriches them, runs duplicate checks, and appends them to the Leads sheet.
 */
function runLeadIngestionPipeline() {
  var ui = SpreadsheetApp.getUi();
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found. Please run 'Setup Sheets' first.", ui.ButtonSet.OK);
    return;
  }
  
  // 1. Retrieve the Apollo API Key
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty("APOLLO_API_KEY");
  if (!apiKey) {
    ui.alert(
      "Apollo API Key Missing",
      "Please set your Apollo API Key in Script Properties:\n\n" +
      "1. In the Apps Script editor, click the gear icon (Project Settings).\n" +
      "2. Scroll down to 'Script Properties' and click 'Add script property'.\n" +
      "3. Add a property named 'APOLLO_API_KEY'.\n" +
      "4. Paste your Apollo API key and click 'Save script properties'.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  // 2. Load Config & Headers
  var config = getConfig();
  var headersMap = getHeadersMap(leadsSheet);
  
  var searchLimit = config.apolloLimit || 10;
  
  try {
    // 3. Call Apollo Search API
    ui.showModelessDialog(
      HtmlService.createHtmlOutput("<p style='font-family:sans-serif;font-size:14px;color:#374151;'>Searching Apollo.io for ICP matching leads...</p>")
        .setWidth(300).setHeight(80),
      "Ingestion in progress"
    );
    
    var searchPayload = {
      "per_page": searchLimit,
      "page": 1
    };
    
    if (config.apolloTitles && config.apolloTitles.length > 0) {
      searchPayload["person_titles"] = config.apolloTitles;
    }
    if (config.apolloLocations && config.apolloLocations.length > 0) {
      searchPayload["person_locations"] = config.apolloLocations;
    }
    if (config.apolloIndustries && config.apolloIndustries.length > 0) {
      searchPayload["organization_industries"] = config.apolloIndustries;
    }
    
    var searchOptions = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Cache-Control": "no-cache",
        "X-Api-Key": apiKey
      },
      "payload": JSON.stringify(searchPayload),
      "muteHttpExceptions": true
    };
    
    var searchUrl = "https://api.apollo.io/api/v1/mixed_people/api_search";
    var searchResponse = UrlFetchApp.fetch(searchUrl, searchOptions);
    var searchCode = searchResponse.getResponseCode();
    var matches = [];
    var isCrmFallback = false;
    var totalEntries = null;
    
    // If forbidden (403), unprocessable (422), or unauthorized (401), fall back to CRM Contacts Search
    if (searchCode === 403 || searchCode === 422 || searchCode === 401) {
      Logger.log("Mixed People Search is restricted on this plan. Trying CRM Contacts Search instead...");
      isCrmFallback = true;
    } else if (searchCode !== 200) {
      var errorMsg = "Search API failed (HTTP " + searchCode + "): " + searchResponse.getContentText();
      logRun(0, 0, 0, errorMsg);
      ui.alert("Apollo API Error", errorMsg, ui.ButtonSet.OK);
      return;
    }
    
    if (isCrmFallback) {
      // Fallback: search saved contacts directly.
      // Saved contacts already have emails revealed, so no bulk match is required.
      var crmUrl = "https://api.apollo.io/api/v1/contacts/search";
      var crmPayload = {
        "per_page": searchLimit,
        "page": 1
      };
      
      var crmOptions = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
          "Cache-Control": "no-cache",
          "X-Api-Key": apiKey
        },
        "payload": JSON.stringify(crmPayload),
        "muteHttpExceptions": true
      };
      
      var crmResponse = UrlFetchApp.fetch(crmUrl, crmOptions);
      var crmCode = crmResponse.getResponseCode();
      
      if (crmCode !== 200) {
        var errorMsg = "CRM Contacts Search failed (HTTP " + crmCode + "): " + crmResponse.getContentText();
        logRun(0, 0, 0, errorMsg);
        ui.alert(
          "Apollo Plan Restriction",
          "Your Apollo plan doesn't support the Global Prospect Search API, and the fallback to CRM Contacts Search also failed:\n\n" + errorMsg,
          ui.ButtonSet.OK
        );
        return;
      }
      
      var crmResult = JSON.parse(crmResponse.getContentText());
      matches = crmResult.contacts || [];
      if (crmResult.pagination && crmResult.pagination.total_entries !== undefined) {
        totalEntries = crmResult.pagination.total_entries;
      }
      
    } else {
      // Standard paid flow: Parse api_search and call bulk_match
      var searchResult = JSON.parse(searchResponse.getContentText());
      var contacts = searchResult.contacts || [];
      
      if (contacts.length === 0) {
        logRun(0, 0, 0, "No leads found matching criteria.");
        ui.alert("No Results", "No new leads found matching your criteria in Apollo.io.", ui.ButtonSet.OK);
        return;
      }
      
      // 4. Call Apollo Bulk Match API to enrich selected contacts
      var personIds = contacts.map(function(c) { return c.id; });
      
      var matchPayload = {
        "person_ids": personIds
      };
      
      var matchOptions = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
          "Cache-Control": "no-cache",
          "X-Api-Key": apiKey
        },
        "payload": JSON.stringify(matchPayload),
        "muteHttpExceptions": true
      };
      
      var matchUrl = "https://api.apollo.io/api/v1/people/bulk_match";
      var matchResponse = UrlFetchApp.fetch(matchUrl, matchOptions);
      var matchCode = matchResponse.getResponseCode();
      
      if (matchCode !== 200) {
        var errorMsg = "Bulk Match API failed (HTTP " + matchCode + "): " + matchResponse.getContentText();
        logRun(0, 0, 0, errorMsg);
        ui.alert("Apollo Enrichment Error", errorMsg, ui.ButtonSet.OK);
        return;
      }
      
      var matchResult = JSON.parse(matchResponse.getContentText());
      matches = matchResult.matches || matchResult.contacts || [];
    }
    
    if (matches.length === 0) {
      var msg = "No profiles were returned by the API.";
      if (isCrmFallback) {
        msg = "Your Apollo CRM returned 0 saved contacts (Total Saved Contacts in your Apollo account: " + (totalEntries !== null ? totalEntries : "0") + ").\n\n" +
              "To import leads on the free tier:\n" +
              "1. Go to the Apollo.io website.\n" +
              "2. Search for prospects, select them, and click 'Save as Contact'.\n" +
              "3. Once you have saved contacts in your account, run this sync again.";
      }
      logRun(0, 0, 0, "No profiles returned from search or enrichment. fallback=" + isCrmFallback + " total=" + totalEntries);
      ui.alert("No Results", msg, ui.ButtonSet.OK);
      return;
    }
    
    // 5. Build lookup lists for duplicate detection
    var lastRow = leadsSheet.getLastRow();
    var existingEmails = {};
    var existingLinkedIns = {};
    
    if (lastRow > 1) {
      var emailColIdx = headersMap["Email"];
      var linkedinColIdx = headersMap["Person Linkedin Url"];
      
      var emailVals = leadsSheet.getRange(2, emailColIdx, lastRow - 1, 1).getValues();
      var linkedinVals = linkedinColIdx ? leadsSheet.getRange(2, linkedinColIdx, lastRow - 1, 1).getValues() : [];
      
      for (var i = 0; i < emailVals.length; i++) {
        var em = emailVals[i][0].toString().trim().toLowerCase();
        if (em) existingEmails[em] = true;
        
        if (linkedinVals.length > i) {
          var li = linkedinVals[i][0].toString().trim().toLowerCase();
          if (li) existingLinkedIns[li] = true;
        }
      }
    }
    
    // 6. Map matches and write to the sheet
    var addedCount = 0;
    var skippedCount = 0;
    
    matches.forEach(function(match) {
      var email = (match.email || "").toString().trim().toLowerCase();
      var linkedin = (match.linkedin_url || "").toString().trim().toLowerCase();
      
      // Duplicate check
      if (email && existingEmails[email]) {
        skippedCount++;
        return;
      }
      if (linkedin && existingLinkedIns[linkedin]) {
        skippedCount++;
        return;
      }
      
      // Map properties to sheet columns
      var row = mapApolloContactToRow(match, headersMap);
      leadsSheet.appendRow(row);
      addedCount++;
    });
    
    // Log run stats
    logRun(matches.length, 0, 0, "Ingestion: Added " + addedCount + ", skipped " + skippedCount + " duplicates.");
    
    ui.alert(
      "Ingestion Complete",
      "Apollo Ingestion finished successfully!\n\n" +
      "- Scanned from Apollo: " + matches.length + "\n" +
      "- Successfully Added: " + addedCount + " new leads\n" +
      "- Skipped (Duplicates): " + skippedCount + " leads\n\n" +
      "These new leads will be scored and validated on the next pipeline execution.",
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    var errorMsg = "Exception during Ingestion: " + e.toString();
    logRun(0, 0, 0, errorMsg);
    ui.alert("Ingestion Exception", errorMsg, ui.ButtonSet.OK);
  }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Maps an Apollo match object's properties to the standard sheet column positions.
 * 
 * @param {object} contact The contact/match object returned by Apollo API
 * @param {object} headersMap Mapping of header names to column index
 * @return {Array} Single row array ordered by sheet columns
 */
function mapApolloContactToRow(contact, headersMap) {
  // Determine the maximum column index to size the array correctly
  var maxColIdx = 0;
  for (var key in headersMap) {
    if (headersMap[key] > maxColIdx) {
      maxColIdx = headersMap[key];
    }
  }
  
  var row = new Array(maxColIdx);
  for (var i = 0; i < maxColIdx; i++) {
    row[i] = "";
  }
  
  // Helper to safely write value to column by name
  function setVal(colName, value) {
    var colIdx = headersMap[colName];
    if (colIdx) {
      row[colIdx - 1] = value !== undefined && value !== null ? value : "";
    }
  }
  
  // Helper to clean phone numbers which can be objects or stringified objects in Apollo CRM
  function cleanPhoneValue(ph) {
    if (!ph) return "";
    if (typeof ph === "object") {
      return ph.number || ph.sanitized_number || "";
    }
    
    var phStr = ph.toString().trim();
    if (phStr === "{}" || phStr === "{ }") {
      return "";
    }
    
    if (phStr.indexOf("{") === 0) {
      try {
        var jsonStr = phStr.replace(/=/g, ':').replace(/([{\s,])(\w+)(:)/g, '$1"$2"$3');
        var parsed = JSON.parse(jsonStr);
        return parsed.number || parsed.sanitized_number || phStr;
      } catch (e) {
        // Fallback: Regex matching with word boundaries
        var numberMatch = phStr.match(/\bnumber=([^,}]+)/);
        if (numberMatch && numberMatch[1]) {
          return numberMatch[1].trim();
        }
        var sanitizedMatch = phStr.match(/\bsanitized_number=([^,}]+)/);
        if (sanitizedMatch && sanitizedMatch[1]) {
          return sanitizedMatch[1].trim();
        }
      }
    }
    return phStr;
  }
  
  var org = contact.organization || {};
  
  // Basic personal details (Feature 2: AI Name Splitting)
  var firstName = contact.first_name || "";
  var lastName = contact.last_name || "";
  
  if (!lastName && firstName.indexOf(" ") !== -1) {
    var split = splitNameWithAI(firstName);
    if (split) {
      setVal("Raw Name Backup", firstName); // Saves the unparsed string to Backup column
      firstName = split.first_name || firstName;
      lastName = split.last_name || "";
    }
  }
  
  setVal("First Name", firstName);
  setVal("Last Name", lastName);
  setVal("Title", contact.title);
  setVal("Company", org.name || contact.organization_name);
  setVal("Company Name for Emails", org.name || contact.organization_name);
  
  // Email fields
  setVal("Email", contact.email);
  setVal("Email Status", contact.email_status);
  setVal("Primary Email Source", contact.email_source);
  setVal("Primary Email Verification Source", contact.email_verification_source);
  setVal("Email Confidence", contact.email_confidence);
  setVal("Primary Email Catch-all Status", contact.email_catch_all_status);
  setVal("Primary Email Last Verified At", contact.email_last_verified_at);
  
  // Seniority & Departments
  setVal("Seniority", contact.seniority);
  
  var depts = "";
  if (contact.departments && contact.departments.length > 0) {
    depts = contact.departments.join(", ");
  }
  setVal("Departments", depts);
  setVal("Contact Owner", "");
  
  // Phone numbers mapping
  var workPhone = cleanPhoneValue(contact.direct_phone);
  var mobilePhone = cleanPhoneValue(contact.mobile_phone);
  var corporatePhone = cleanPhoneValue(org.primary_phone || contact.corporate_phone);
  
  if (contact.phone_numbers && contact.phone_numbers.length > 0) {
    for (var i = 0; i < contact.phone_numbers.length; i++) {
      var ph = contact.phone_numbers[i];
      if (ph.type === "mobile" && !mobilePhone) mobilePhone = cleanPhoneValue(ph);
      if (ph.type === "work" && !workPhone) workPhone = cleanPhoneValue(ph);
      if (ph.type === "corporate" && !corporatePhone) corporatePhone = cleanPhoneValue(ph);
    }
  }
  
  setVal("Work Direct Phone", workPhone);
  setVal("Home Phone", "");
  setVal("Mobile Phone", mobilePhone);
  setVal("Corporate Phone", corporatePhone);
  setVal("Other Phone", "");
  
  // Stage and Lists defaults
  setVal("Stage", "New");
  setVal("Lists", "Apollo Ingest");
  setVal("Last Contacted", "");
  setVal("Account Owner", "");
  
  // Organization firmographics
  setVal("# Employees", org.estimated_num_employees);
  
  var industries = "";
  if (org.industries && org.industries.length > 0) {
    industries = org.industries.join(", ");
  } else if (contact.organization_industries && contact.organization_industries.length > 0) {
    industries = contact.organization_industries.join(", ");
  }
  setVal("Industry", industries);
  
  var keywords = "";
  if (org.keywords && org.keywords.length > 0) {
    keywords = org.keywords.join(", ");
  }
  setVal("Keywords", keywords);
  
  setVal("Person Linkedin Url", contact.linkedin_url);
  setVal("Website", org.website_url || contact.organization_website);
  setVal("Company Linkedin Url", org.linkedin_url);
  setVal("Facebook Url", org.facebook_url);
  setVal("Twitter Url", org.twitter_url);
  
  // Location
  setVal("City", contact.city);
  setVal("State", contact.state);
  setVal("Country", contact.country);
  
  // Company Location
  var companyAddress = "";
  if (org.street_address) {
    companyAddress = org.street_address;
    if (org.postal_code) companyAddress += ", " + org.postal_code;
  }
  setVal("Company Address", companyAddress);
  setVal("Company City", org.city);
  setVal("Company State", org.state);
  setVal("Company Country", org.country);
  setVal("Company Phone", cleanPhoneValue(org.primary_phone));
  
  // Technologies
  var techs = "";
  if (org.technology_names && org.technology_names.length > 0) {
    techs = org.technology_names.join(", ");
  }
  setVal("Technologies", techs);
  
  // Revenue & Funding (Crucial for scoring!)
  setVal("Annual Revenue", org.annual_revenue);
  setVal("Total Funding", org.funding_total_amount || org.total_funding);
  setVal("Latest Funding", org.latest_funding_round_date || org.latest_funding_round);
  setVal("Latest Funding Amount", org.latest_funding_round_amount || org.latest_funding_amount);
  setVal("Last Raised At", org.latest_funding_round_date || org.latest_funding_round);
  
  // Custom columns stay empty for scoring step to fill
  setVal("Score", "");
  setVal("Score Reason", "");
  setVal("Validation Status", "");
  setVal("Outreach Status", "");
  
  return row;
}

/**
 * Ingests leads using Google Custom Search Engine (CSE) to target public LinkedIn profiles.
 * Requires GOOGLE_API_KEY and GOOGLE_CSE_ID in Script Properties.
 */
function runLinkedInXRayIngestionPipeline() {
  var ui = SpreadsheetApp.getUi();
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found. Run 'Setup Sheets' first.", ui.ButtonSet.OK);
    return;
  }
  
  var props = PropertiesService.getScriptProperties();
  var serperApiKey = props.getProperty("SERPER_API_KEY");
  var googleApiKey = props.getProperty("GOOGLE_API_KEY");
  var cseId = props.getProperty("GOOGLE_CSE_ID") || getConfig().googleCseId;
  
  if (!serperApiKey && (!googleApiKey || !cseId)) {
    ui.alert(
      "Configuration Missing",
      "Please set your search credentials:\n\n" +
      "1. Add 'SERPER_API_KEY' (Serper.dev API Key) to your Script Properties in Settings (Recommended).\n" +
      "2. Or add 'GOOGLE_API_KEY' and 'GOOGLE_CSE_ID' for Google Custom Search.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  var config = getConfig();
  var titles = config.apolloTitles || ["CTO", "Founder"];
  var locations = config.apolloLocations || ["India"];
  var limit = config.apolloLimit || 10;
  
  // Format query: e.g. site:linkedin.com/in "CTO" "India"
  var query = 'site:linkedin.com/in "' + titles[0] + '" "' + locations[0] + '"';
             
  try {
    ui.showModelessDialog(
      HtmlService.createHtmlOutput("<p style='font-family:sans-serif;font-size:14px;color:#374151;'>Searching LinkedIn profiles...</p>")
        .setWidth(300).setHeight(80),
      "X-Ray Search in progress"
    );

    var items = [];
    if (serperApiKey) {
      // Use Serper.dev API (Fully automated, no restrictions)
      var serperUrl = "https://google.serper.dev/search";
      var payload = {
        "q": query,
        "num": Math.min(limit, 20)
      };
      var response = UrlFetchApp.fetch(serperUrl, {
        "method": "post",
        "contentType": "application/json",
        "headers": {
          "X-API-KEY": serperApiKey
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });
      if (response.getResponseCode() !== 200) {
        throw new Error("Serper.dev returned " + response.getResponseCode() + ": " + response.getContentText());
      }
      var data = JSON.parse(response.getContentText());
      items = data.organic || [];
    } else {
      // Fallback to Google Custom Search Engine
      var googleUrl = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(googleApiKey) +
                      "&cx=" + encodeURIComponent(cseId) +
                      "&q=" + encodeURIComponent(query) +
                      "&num=" + Math.min(limit, 10);
      var response = UrlFetchApp.fetch(googleUrl, { "muteHttpExceptions": true });
      if (response.getResponseCode() !== 200) {
        throw new Error("Google Search returned " + response.getResponseCode() + ": " + response.getContentText());
      }
      var data = JSON.parse(response.getContentText());
      items = data.items || [];
    }

    if (items.length === 0) {
      ui.alert("No Results", "No LinkedIn profiles found matching query: " + query, ui.ButtonSet.OK);
      return;
    }
    
    var headersMap = getHeadersMap(leadsSheet);
    var addedCount = 0;
    
    items.forEach(function(item) {
      var titleText = item.title || "";
      var link = item.link || "";
      
      var name = "";
      var role = titles[0];
      var company = "";
      
      var parts = titleText.split(" - ");
      if (parts.length >= 1) name = parts[0].replace(" | LinkedIn", "").trim();
      if (parts.length >= 2) role = parts[1].trim();
      if (parts.length >= 3) company = parts[2].split(" | ")[0].split(" - ")[0].trim();
      
      // Feature 2: AI name splitting for LinkedIn X-Ray
      var firstName = name;
      var lastName = "";
      if (name.indexOf(" ") !== -1) {
        var aiSplit = splitNameWithAI(name);
        if (aiSplit) {
          firstName = aiSplit.first_name || name;
          lastName = aiSplit.last_name || "";
        } else {
          // Graceful fallback: naive split if AI unavailable
          var np = name.split(" ");
          firstName = np[0];
          lastName = np.slice(1).join(" ");
        }
      }
      
      var maxColIdx = 0;
      for (var key in headersMap) {
        if (headersMap[key] > maxColIdx) {
          maxColIdx = headersMap[key];
        }
      }
      var row = new Array(maxColIdx);
      for (var k = 0; k < maxColIdx; k++) {
        row[k] = "";
      }
      function setVal(col, val) {
        var idx = headersMap[col];
        if (idx) row[idx - 1] = val !== undefined && val !== null ? val : "";
      }
      
      if (name !== firstName) { setVal("Raw Name Backup", name); }
      setVal("First Name", firstName);
      setVal("Last Name", lastName);
      setVal("Title", role);
      setVal("Company", company);
      setVal("Company Name for Emails", company);
      setVal("Person Linkedin Url", link);
      setVal("Lists", "LinkedIn X-Ray");
      setVal("Stage", "New");
      setVal("Email", "");
      setVal("Email Status", "Missing");
      
      leadsSheet.appendRow(row);
      addedCount++;
    });
    
    ui.alert("LinkedIn Ingestion Done", "Successfully ingested " + addedCount + " leads from LinkedIn search.", ui.ButtonSet.OK);
  } catch(e) {
    ui.alert("LinkedIn Ingestion Error", e.toString(), ui.ButtonSet.OK);
  }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Ingests leads from GitHub's search API. Good for technical recruitment/outreach.
 * Optional: GITHUB_TOKEN in Script Properties.
 */
function runGitHubIngestionPipeline() {
  var ui = SpreadsheetApp.getUi();
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  var props = PropertiesService.getScriptProperties();
  var gitHubToken = props.getProperty("GITHUB_TOKEN") || getConfig().githubKey;
  
  var config = getConfig();
  var role = (config.apolloTitles && config.apolloTitles[0]) || "CTO";
  var loc = (config.apolloLocations && config.apolloLocations[0]) || "India";
  var limit = Math.min(config.apolloLimit || 10, 30);
  
  var query = 'location:' + encodeURIComponent(loc) + ' ' + encodeURIComponent(role);
  var url = "https://api.github.com/search/users?q=" + encodeURIComponent(query) + "&per_page=" + limit;
  
  var headers = {
    "User-Agent": "AppsScript-Lead-Engine"
  };
  if (gitHubToken) {
    headers["Authorization"] = "token " + gitHubToken;
  }
  
  try {
    ui.showModelessDialog(
      HtmlService.createHtmlOutput("<p style='font-family:sans-serif;font-size:14px;color:#374151;'>Searching GitHub profiles...</p>")
        .setWidth(300).setHeight(80),
      "GitHub Search in progress"
    );

    var response = UrlFetchApp.fetch(url, { "headers": headers, "muteHttpExceptions": true });
    if (response.getResponseCode() !== 200) {
      throw new Error("GitHub Search returned HTTP " + response.getResponseCode() + ": " + response.getContentText());
    }
    
    var data = JSON.parse(response.getContentText());
    var items = data.items || [];
    if (items.length === 0) {
      ui.alert("No Results", "No GitHub users found matching: " + query, ui.ButtonSet.OK);
      return;
    }
    
    var headersMap = getHeadersMap(leadsSheet);
    var addedCount = 0;
    
    items.forEach(function(item) {
      var profileUrl = item.url;
      var profRes = UrlFetchApp.fetch(profileUrl, { "headers": headers, "muteHttpExceptions": true });
      if (profRes.getResponseCode() === 200) {
        var prof = JSON.parse(profRes.getContentText());
        
        var fullName = prof.name || prof.login || "";
        
        // Feature 2: AI name splitting for GitHub profiles
        var firstName = fullName;
        var lastName = "";
        if (fullName.indexOf(" ") !== -1) {
          var aiSplit = splitNameWithAI(fullName);
          if (aiSplit) {
            firstName = aiSplit.first_name || fullName;
            lastName = aiSplit.last_name || "";
          } else {
            var np = fullName.split(" ");
            firstName = np[0];
            lastName = np.slice(1).join(" ");
          }
        }
        
        var company = prof.company || "";
        if (company.startsWith("@")) company = company.substring(1);
        
        var maxColIdx = 0;
        for (var key in headersMap) {
          if (headersMap[key] > maxColIdx) {
            maxColIdx = headersMap[key];
          }
        }
        var row = new Array(maxColIdx);
        for (var k = 0; k < maxColIdx; k++) {
          row[k] = "";
        }
        function setVal(col, val) {
          var idx = headersMap[col];
          if (idx) row[idx - 1] = val !== undefined && val !== null ? val : "";
        }
        
        if (fullName !== firstName) { setVal("Raw Name Backup", fullName); }
        setVal("First Name", firstName);
        setVal("Last Name", lastName);
        setVal("Title", role);
        setVal("Company", company || "GitHub Profile " + prof.login);
        setVal("Company Name for Emails", company || "GitHub Profile " + prof.login);
        setVal("Email", prof.email);
        setVal("Email Status", prof.email ? "Verified" : "Missing");
        setVal("Website", prof.blog);
        setVal("Person Linkedin Url", "https://github.com/" + prof.login);
        setVal("Lists", "GitHub Search");
        setVal("Stage", "New");
        
        leadsSheet.appendRow(row);
        addedCount++;
      }
    });
    
    ui.alert("GitHub Ingestion Done", "Successfully ingested " + addedCount + " leads from GitHub.", ui.ButtonSet.OK);
  } catch(e) {
    ui.alert("GitHub Ingestion Error", e.toString(), ui.ButtonSet.OK);
  }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Ingests local business data from Google Places Text Search.
 * Requires GOOGLE_MAPS_API_KEY or GOOGLE_API_KEY in Script Properties.
 */
function runGoogleMapsIngestionPipeline() {
  var ui = SpreadsheetApp.getUi();
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty("GOOGLE_MAPS_API_KEY") || props.getProperty("GOOGLE_API_KEY") || getConfig().mapsKey;
  
  if (!apiKey) {
    ui.alert(
      "Configuration Missing",
      "Please set your Google Maps Places API key:\n\n" +
      "1. Add 'GOOGLE_MAPS_API_KEY' or 'GOOGLE_API_KEY' to your Script Properties in settings.\n" +
      "2. Or enter the API key in the 'Google Maps API Key' parameter in the Config tab.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  var config = getConfig();
  var title = (config.apolloTitles && config.apolloTitles[0]) || "Software Development";
  var location = (config.apolloLocations && config.apolloLocations[0]) || "Bangalore";
  var limit = Math.min(config.apolloLimit || 10, 20);
  
  var query = title + " in " + location;
  var url = "https://maps.googleapis.com/maps/api/place/textsearch/json?query=" + encodeURIComponent(query) + "&key=" + encodeURIComponent(apiKey);
  
  try {
    ui.showModelessDialog(
      HtmlService.createHtmlOutput("<p style='font-family:sans-serif;font-size:14px;color:#374151;'>Searching Google Places for local businesses...</p>")
        .setWidth(300).setHeight(80),
      "Maps Ingest in progress"
    );

    var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (response.getResponseCode() !== 200) {
      throw new Error("Google Places API returned HTTP " + response.getResponseCode() + ": " + response.getContentText());
    }
    
    var data = JSON.parse(response.getContentText());
    var results = data.results || [];
    if (results.length === 0) {
      ui.alert("No Results", "No businesses found matching: " + query, ui.ButtonSet.OK);
      return;
    }
    
    var headersMap = getHeadersMap(leadsSheet);
    var addedCount = 0;
    
    for (var i = 0; i < Math.min(results.length, limit); i++) {
      var place = results[i];
      
      // Fetch details for each business to retrieve website & phone number
      var detailsUrl = "https://maps.googleapis.com/maps/api/place/details/json?place_id=" + place.place_id +
                       "&fields=name,formatted_address,formatted_phone_number,website&key=" + encodeURIComponent(apiKey);
      
      var detRes = UrlFetchApp.fetch(detailsUrl, { "muteHttpExceptions": true });
      var companyName = place.name;
      var website = "";
      var phone = "";
      var address = place.formatted_address || "";
      
      if (detRes.getResponseCode() === 200) {
        var detData = JSON.parse(detRes.getContentText());
        var details = detData.result || {};
        website = details.website || "";
        phone = details.formatted_phone_number || "";
      }
      
      var maxColIdx = 0;
      for (var key in headersMap) {
        if (headersMap[key] > maxColIdx) {
          maxColIdx = headersMap[key];
        }
      }
      var row = new Array(maxColIdx);
      for (var k = 0; k < maxColIdx; k++) {
        row[k] = "";
      }
      function setVal(col, val) {
        var idx = headersMap[col];
        if (idx) row[idx - 1] = val !== undefined && val !== null ? val : "";
      }
      
      setVal("First Name", "Founder");
      setVal("Last Name", "/" + companyName);
      setVal("Title", "Founder / Owner");
      setVal("Company", companyName);
      setVal("Company Name for Emails", companyName);
      setVal("Website", website);
      setVal("Corporate Phone", phone);
      setVal("Company Address", address);
      setVal("Lists", "Google Maps Ingest");
      setVal("Stage", "New");
      setVal("Email Status", "Missing");
      
      leadsSheet.appendRow(row);
      addedCount++;
    }
    
    ui.alert("Google Maps Ingestion Done", "Successfully ingested " + addedCount + " companies from Google Maps.", ui.ButtonSet.OK);
  } catch(e) {
    ui.alert("Google Maps Ingestion Error", e.toString(), ui.ButtonSet.OK);
  }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Feature 2: Helper to intelligently split South Indian or complex names using Gemini API.
 */
function splitNameWithAI(fullName) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey || !fullName) return null;
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;
  var prompt = "I have a full name from a lead generation list that was placed entirely in the First Name field. I need to split it into a logical 'first_name' and 'last_name'.\n\n" +
               "Name: " + fullName + "\n\n" +
               "Note: Handle South Indian names properly (e.g. initial as last name or first name depending on standard conventions). Reply ONLY with valid JSON in this format: {\"first_name\": \"...\", \"last_name\": \"...\"}";
               
  var payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "temperature": 0.1
    }
  };
  
  var options = {
    "method": "post",
    "headers": { "Content-Type": "application/json" },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      var json = JSON.parse(response.getContentText());
      if (json.candidates && json.candidates[0] && json.candidates[0].content) {
        var text = json.candidates[0].content.parts[0].text;
        return JSON.parse(text);
      }
    }
  } catch(e) {
    Logger.log("Name split AI error: " + e.toString());
  }
  return null;
}

/**
 * Feature 7: Manual Apollo enrichment for selected rows.
 * Reads the emails/domains of selected rows and hits the bulk_match endpoint.
 */
function enrichSelectedRows() {
  var ui = SpreadsheetApp.getUi();
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  
  if (!leadsSheet) {
    ui.alert("Error", "Leads sheet not found.", ui.ButtonSet.OK);
    return;
  }
  
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty("APOLLO_API_KEY");
  if (!apiKey) {
    ui.alert("Apollo API Key Missing", "Please set APOLLO_API_KEY in Script Properties.", ui.ButtonSet.OK);
    return;
  }
  
  var activeRange = leadsSheet.getActiveRange();
  if (!activeRange) {
    ui.alert("No Selection", "Please select the rows you want to enrich.", ui.ButtonSet.OK);
    return;
  }
  
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  var endRow = startRow + numRows - 1;
  var lastRow = leadsSheet.getLastRow();
  
  if (startRow < 2) startRow = 2; // skip header
  if (endRow > lastRow) endRow = lastRow;
  
  if (startRow > endRow) {
    ui.alert("Invalid Selection", "Selected range contains no valid data rows.", ui.ButtonSet.OK);
    return;
  }
  
  var headersMap = getHeadersMap(leadsSheet);
  var mapping = resolveColumnMapping(leadsSheet, true);
  
  var toEnrich = [];
  var rowIndexMap = [];
  
  for (var r = startRow; r <= endRow; r++) {
    var email = getCanonical(leadsSheet, r, mapping, "email");
    var domain = getCanonical(leadsSheet, r, mapping, "company"); // Usually website or company domain is better but we try email domains first if available or website
    var website = headersMap["Website"] ? leadsSheet.getRange(r, headersMap["Website"]).getValue() : "";
    var fName = getCanonical(leadsSheet, r, mapping, "first_name");
    var lName = getCanonical(leadsSheet, r, mapping, "last_name");
    
    var obj = {};
    if (email) obj.email = email;
    if (fName) obj.first_name = fName;
    if (lName) obj.last_name = lName;
    if (website) obj.domain = website.replace(/^(https?:\/\/)?(www\.)?/, "").split("/")[0];
    
    if (Object.keys(obj).length > 0) {
      toEnrich.push(obj);
      rowIndexMap.push(r);
    }
  }
  
  if (toEnrich.length === 0) {
    ui.alert("No Data", "No valid data found in selected rows to match against Apollo.", ui.ButtonSet.OK);
    return;
  }
  
  if (toEnrich.length > 50) {
    ui.alert("Limit Exceeded", "Please select a maximum of 50 rows per batch for bulk enrichment.", ui.ButtonSet.OK);
    return;
  }
  
  ui.showModelessDialog(HtmlService.createHtmlOutput("<p>Enriching " + toEnrich.length + " contacts in batches of 10...</p>").setWidth(300).setHeight(80), "Apollo Enrichment");
  
  var BATCH_SIZE = 10;
  var totalEnrichedCount = 0;
  var totalNoMatchCount = 0;
  var hasApiError = false;
  
  for (var batchStart = 0; batchStart < toEnrich.length; batchStart += BATCH_SIZE) {
    var batchEnd = Math.min(batchStart + BATCH_SIZE, toEnrich.length);
    var batchDetails = toEnrich.slice(batchStart, batchEnd);
    var batchRows = rowIndexMap.slice(batchStart, batchEnd);
    
    Logger.log("Enrichment batch " + (Math.floor(batchStart / BATCH_SIZE) + 1) + ": rows " + batchStart + " to " + (batchEnd - 1));
    
    var matchPayload = { "details": batchDetails };
    var matchOptions = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Cache-Control": "no-cache",
        "X-Api-Key": apiKey
      },
      "payload": JSON.stringify(matchPayload),
      "muteHttpExceptions": true
    };
    
    var matchUrl = "https://api.apollo.io/api/v1/people/bulk_match";
    
    try {
      var matchResponse = UrlFetchApp.fetch(matchUrl, matchOptions);
      
      if (matchResponse.getResponseCode() !== 200) {
        Logger.log("Apollo batch error (HTTP " + matchResponse.getResponseCode() + "): " + matchResponse.getContentText());
        hasApiError = true;
        // Tag all rows in this batch as error so user can retry them
        for (var bi = 0; bi < batchRows.length; bi++) {
          if (headersMap["Enrichment Status"]) {
            leadsSheet.getRange(batchRows[bi], headersMap["Enrichment Status"]).setValue("API Error — Retry");
          }
        }
        Utilities.sleep(2000);
        continue; // move to next batch
      }
      
      var matchResult = JSON.parse(matchResponse.getContentText());
      var matches = matchResult.matches || [];
      
      // Apollo returns matches in the same order as the input `details` array.
      // Nulls or empty objects mean no match was found for that input.
      for (var i = 0; i < batchRows.length; i++) {
        var r = batchRows[i];
        var match = matches[i]; // may be null/undefined if Apollo found nothing
        
        if (!match || !match.id) {
          // No match — tag the row so the user knows enrichment ran but found nothing
          if (headersMap["Enrichment Status"]) {
            leadsSheet.getRange(r, headersMap["Enrichment Status"]).setValue("No Match");
          }
          totalNoMatchCount++;
          continue;
        }
        
        // Match found — write back missing fields
        var existingEmail = getCanonical(leadsSheet, r, mapping, "email");
        if (!existingEmail && match.email) {
          if (headersMap["Email"]) leadsSheet.getRange(r, headersMap["Email"]).setValue(match.email);
          if (headersMap["Email Status"]) leadsSheet.getRange(r, headersMap["Email Status"]).setValue(match.email_status || "Found (Apollo)");
          totalEnrichedCount++;
        }
        
        // Write mobile phone if missing
        var mobile = (match.phone_numbers || []).filter(function(p){ return p.type === "mobile"; })[0];
        if (mobile && headersMap["Mobile Phone"]) {
          var existingPhone = leadsSheet.getRange(r, headersMap["Mobile Phone"]).getValue().toString().trim();
          if (!existingPhone) {
            leadsSheet.getRange(r, headersMap["Mobile Phone"]).setValue(mobile.sanitized_number || mobile.number);
          }
        }
        
        // Tag the row with enrichment source
        if (headersMap["Enrichment Status"]) {
          leadsSheet.getRange(r, headersMap["Enrichment Status"]).setValue("Apollo Enrichment");
        }
      }
      
      // Pause between batches to avoid rate-limiting
      if (batchEnd < toEnrich.length) {
        Utilities.sleep(1500);
      }
      
    } catch (e) {
      Logger.log("Enrichment batch exception: " + e.toString());
      hasApiError = true;
      for (var bi = 0; bi < batchRows.length; bi++) {
        if (headersMap["Enrichment Status"]) {
          leadsSheet.getRange(batchRows[bi], headersMap["Enrichment Status"]).setValue("API Error — Retry");
        }
      }
    }
  }
  
  var alertMsg = "Apollo enrichment finished!\n\n" +
    "- Rows submitted: " + toEnrich.length + "\n" +
    "- Contacts enriched with email/phone: " + totalEnrichedCount + "\n" +
    "- No match found: " + totalNoMatchCount;
  
  if (hasApiError) {
    alertMsg += "\n\n⚠️ Some batches encountered API errors. Rows tagged 'API Error — Retry' can be re-selected and re-run.";
  }
  
  ui.alert("Enrichment Complete", alertMsg, ui.ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
