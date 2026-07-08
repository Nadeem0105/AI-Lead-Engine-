/**
 * AI Lead Generation Engine - Hiring Intent Signals
 * File: Hiring.gs
 */

/**
 * Checks if a company is actively hiring by querying Google Custom Search.
 * Updates the 'Hiring Status' column for the lead.
 * 
 * @param {Sheet} sheet The Leads sheet reference
 * @param {number} r The row number (1-indexed)
 * @param {object} headersMap Mapping of header names to 1-indexed column numbers
 * @param {object} config Configuration parameters
 * @return {string} The detected hiring status
 */
function checkAndRecordHiringStatus(sheet, r, headersMap, config) {
  var hiringColIdx = headersMap["Hiring Status"];
  if (!hiringColIdx) return "";
  
  var company = (sheet.getRange(r, headersMap["Company"]).getValue() || "").toString().trim();
  if (!company) {
    sheet.getRange(r, hiringColIdx).setValue("No Company Name");
    return "No Company Name";
  }
  
  var props = PropertiesService.getScriptProperties();
  var serperApiKey = props.getProperty("SERPER_API_KEY");
  var googleApiKey = props.getProperty("GOOGLE_API_KEY");
  var cseId = props.getProperty("GOOGLE_CSE_ID") || config.googleCseId;
  
  if (!serperApiKey && (!googleApiKey || !cseId)) {
    Logger.log("Hiring Check skipped: Neither Serper.dev API Key nor Google Search API Key configured.");
    sheet.getRange(r, hiringColIdx).setValue("Skipped (No API Config)");
    return "Skipped (No API Config)";
  }
  
  try {
    Logger.log("Checking hiring signals for: " + company);
    
    // Target common ATS portals and LinkedIn jobs
    var query = '"' + company + '" (site:greenhouse.io OR site:lever.co OR site:workable.com OR site:linkedin.com/jobs/view)';
    
    var items = [];
    if (serperApiKey) {
      // Use Serper.dev
      var serperUrl = "https://google.serper.dev/search";
      var payload = {
        "q": query,
        "num": 5
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
      if (response.getResponseCode() === 200) {
        var data = JSON.parse(response.getContentText());
        items = data.organic || [];
      }
    } else {
      // Fallback to Google CSE
      var googleUrl = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(googleApiKey) +
                      "&cx=" + encodeURIComponent(cseId) +
                      "&q=" + encodeURIComponent(query) +
                      "&num=3";
      var response = UrlFetchApp.fetch(googleUrl, { "muteHttpExceptions": true });
      if (response.getResponseCode() === 200) {
        var data = JSON.parse(response.getContentText());
        items = data.items || [];
      }
    }

    var status = "No Active Signals";
    
    if (items.length > 0) {
      var rolesFound = [];
      
      items.forEach(function(item) {
        var title = item.title || "";
        var snippet = item.snippet || "";
        
        // Try to extract job title patterns
        var match = title.match(/Job Application for (.+?) at/i) || 
                    title.match(/(.+?) at [^|]+/i) ||
                    title.match(/(.+?) Job/i);
                    
        if (match && match[1]) {
          var role = match[1].trim();
          if (role.length < 50 && rolesFound.indexOf(role) === -1) {
            rolesFound.push(role);
          }
        }
      });
      
      if (rolesFound.length > 0) {
        status = "Hiring: " + rolesFound.join(", ");
      } else {
        status = "Active Job Boards Found";
      }
    } else {
      // Run a broad check: "company" careers page / hiring
      var broadQuery = '"' + company + '" (careers OR jobs OR "we are hiring")';
      var broadItems = [];
      
      if (serperApiKey) {
        var broadUrl = "https://google.serper.dev/search";
        var broadPayload = {
          "q": broadQuery,
          "num": 3
        };
        var broadRes = UrlFetchApp.fetch(broadUrl, {
          "method": "post",
          "contentType": "application/json",
          "headers": {
            "X-API-KEY": serperApiKey
          },
          "payload": JSON.stringify(broadPayload),
          "muteHttpExceptions": true
        });
        if (broadRes.getResponseCode() === 200) {
          var broadData = JSON.parse(broadRes.getContentText());
          broadItems = broadData.organic || [];
        }
      } else {
        var broadUrl = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(googleApiKey) +
                        "&cx=" + encodeURIComponent(cseId) +
                        "&q=" + encodeURIComponent(broadQuery) +
                        "&num=2";
        var broadRes = UrlFetchApp.fetch(broadUrl, { "muteHttpExceptions": true });
        if (broadRes.getResponseCode() === 200) {
          var broadData = JSON.parse(broadRes.getContentText());
          broadItems = broadData.items || [];
        }
      }
      
      if (broadItems.length > 0) {
        status = "Possible Hiring (Careers Page)";
      }
    }
    
    sheet.getRange(r, hiringColIdx).setValue(status);
    Logger.log("Hiring check done for " + company + ": " + status);
    return status;
    
  } catch (err) {
    Logger.log("Hiring Check Error: " + err.toString());
    sheet.getRange(r, hiringColIdx).setValue("Error: " + err.toString());
    return "Error";
  }
}
