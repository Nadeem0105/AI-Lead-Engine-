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
  
  // Basic personal details
  setVal("First Name", contact.first_name);
  setVal("Last Name", contact.last_name);
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
      
      var nameParts = name.split(" ");
      var firstName = nameParts[0] || "";
      var lastName = nameParts.slice(1).join(" ") || "";
      
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
}

/**
 * Ingests leads from GitHub's search API. Good for technical recruitment/outreach.
 * Optional: GITHUB_TOKEN in Script Properties.
 */
function runGitHubIngestionPipeline() {
  var ui = SpreadsheetApp.getUi();
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
        var nameParts = fullName.split(" ");
        var firstName = nameParts[0] || "";
        var lastName = nameParts.slice(1).join(" ") || "";
        
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
}

/**
 * Ingests local business data from Google Places Text Search.
 * Requires GOOGLE_MAPS_API_KEY or GOOGLE_API_KEY in Script Properties.
 */
function runGoogleMapsIngestionPipeline() {
  var ui = SpreadsheetApp.getUi();
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
}
