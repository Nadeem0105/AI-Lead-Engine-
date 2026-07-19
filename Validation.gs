/**
 * AI Lead Generation Engine - Email Quality Filtering
 * File: Validation.gs
 */

/**
 * Validates the email quality of a single lead, performing a second-pass API verification
 * if the initial check indicates a missing, catch-all, or low-confidence email.
 * 
 * @param {Sheet} sheet The Leads sheet reference
 * @param {number} rowNumber The row number to process (1-indexed)
 * @param {object} headersMap Mapping of header names to 1-indexed column numbers
 * @param {object} config Configuration parameters read from Config sheet
 * @return {object} Object with {status, reason} where status is 'Ready', 'Risky', or 'Flagged'
 */
function validateEmailQuality(sheet, rowNumber, headersMap, config, mapping) {
  try {
    // Feature 1: tolerate generic-source sheets where Apollo-specific columns are absent.
    var email = getFieldValue(sheet, rowNumber, headersMap, mapping, "Email", "email");

    var emailStatus = headersMap["Email Status"] ? (sheet.getRange(rowNumber, headersMap["Email Status"]).getValue() || "") : "";
    emailStatus = emailStatus.toString().trim();

    var emailConfidence = headersMap["Email Confidence"] ? sheet.getRange(rowNumber, headersMap["Email Confidence"]).getValue() : "";

    var catchAllStatus = headersMap["Primary Email Catch-all Status"] ? (sheet.getRange(rowNumber, headersMap["Primary Email Catch-all Status"]).getValue() || "") : "";
    catchAllStatus = catchAllStatus.toString().trim();
    
    var emailStatusLower = emailStatus.toLowerCase();
    var catchAllLower = catchAllStatus.toLowerCase();
    var confidenceThreshold = config.emailConfidenceThreshold || 0.7;
    
    // 1. Check if email is missing entirely -> Try to find it via Hunter.io first
    if (!email) {
      var findResult = runSecondPassEnrichment(sheet, rowNumber, headersMap, config);
      if (findResult.status) {
        return findResult;
      }
      return { 
        status: "Flagged", 
        reason: "Missing email address" 
      };
    }
    
    // 2. If Email Status is "Unavailable" -> Flagged
    if (emailStatusLower === "unavailable") {
      return { 
        status: "Flagged", 
        reason: "Email is unavailable" 
      };
    }
    
    // Feature 6: If Email Status is "Verified", immediately return Ready without hitting ZeroBounce
    if (emailStatusLower === "verified") {
      return { 
        status: "Ready", 
        reason: "Verified (Apollo)" 
      };
    }
    
    // 3. Check if email needs verification (Catch-all, Guessed, or Low Confidence)
    var needsVerification = false;
    var riskReason = "";
    
    if (catchAllLower === "catch-all" || catchAllLower === "catch_all") {
      needsVerification = true;
      riskReason = "Catch-all email domain";
    } else if (emailStatusLower === "guessed") {
      needsVerification = true;
      riskReason = "Guessed email address";
    } else if (emailConfidence !== "" && emailConfidence !== null && !isNaN(emailConfidence)) {
      var confidenceNum = parseFloat(emailConfidence);
      if (confidenceNum < confidenceThreshold) {
        needsVerification = true;
        riskReason = "Email confidence (" + confidenceNum + ") below threshold (" + confidenceThreshold + ")";
      }
    }

    // Feature 6: anything other than an explicit Apollo "verified" must go through
    // second-pass verification (ZeroBounce/Hunter) — including a blank or unrecognized
    // status, which is the norm for non-Apollo sources.
    if (!needsVerification) {
      needsVerification = true;
      riskReason = emailStatus === "" ? "No provider email status — needs verification"
                                      : "Unrecognized email status: " + emailStatus;
    }

    if (needsVerification) {
      var verifyResult = runSecondPassEnrichment(sheet, rowNumber, headersMap, config);
      if (verifyResult.status) {
        return verifyResult;
      }
      // No second-pass provider configured/reachable — mark Risky, never assume verified
      return {
        status: "Risky",
        reason: riskReason
      };
    }

  } catch (e) {
    return { 
      status: "Flagged", 
      reason: "Validation exception: " + e.toString() 
    };
  }
}

/**
 * Performs a second-pass API enrichment to find missing emails or verify risky ones.
 * Supports Hunter.io (Finder/Verifier) and ZeroBounce (Verifier).
 * 
 * @param {Sheet} sheet The Leads sheet reference
 * @param {number} r The row number (1-indexed)
 * @param {object} headersMap Mapping of header names to 1-indexed column numbers
 * @param {object} config Configuration parameters
 * @return {object} Object with {status, reason, emailUpdated}
 */
function runSecondPassEnrichment(sheet, r, headersMap, config) {
  var props = PropertiesService.getScriptProperties();
  var hunterKey = props.getProperty("HUNTER_API_KEY") || config.hunterKey;
  var zeroBounceKey = props.getProperty("ZEROBOUNCE_API_KEY") || config.zeroBounceKey;
  
  // Feature 1: resolve columns tolerantly; generic sheets may lack Apollo-specific headers.
  var mapping = resolveColumnMapping(sheet, false);
  var emailCol = headersMap["Email"] || mapping["email"];
  if (!emailCol) return { status: null, reason: null, emailUpdated: false };
  var emailRange = sheet.getRange(r, emailCol);
  var email = (emailRange.getValue() || "").toString().trim();

  var statusRange = headersMap["Email Status"] ? sheet.getRange(r, headersMap["Email Status"]) : null;

  var firstName = getFieldValue(sheet, r, headersMap, mapping, "First Name", "first_name");
  var lastName = getFieldValue(sheet, r, headersMap, mapping, "Last Name", "last_name");
  var company = getFieldValue(sheet, r, headersMap, mapping, "Company", "company");
  var website = getFieldValue(sheet, r, headersMap, mapping, "Website", null);
  
  // CASE A: Email is missing -> Try Hunter.io Email Finder
  if (!email && hunterKey && firstName && lastName && (website || company)) {
    try {
      var domain = website;
      if (!domain && company) {
        domain = company.toLowerCase().replace(/[^a-z0-9]/g, "") + ".com";
      }
      
      // Clean domain formatting
      domain = domain.replace(/^(https?:\/\/)?(www\.)?/, "").split("/")[0];
      
      Logger.log("Running Hunter.io Email Finder for " + firstName + " " + lastName + " at " + domain);
      
      var finderUrl = "https://api.hunter.io/v2/email-finder?domain=" + encodeURIComponent(domain) +
                      "&first_name=" + encodeURIComponent(firstName) +
                      "&last_name=" + encodeURIComponent(lastName) +
                      "&api_key=" + encodeURIComponent(hunterKey);
                      
      var response = UrlFetchApp.fetch(finderUrl, { "muteHttpExceptions": true });
      if (response.getResponseCode() === 200) {
        var resObj = JSON.parse(response.getContentText());
        if (resObj.data && resObj.data.email) {
          var foundEmail = resObj.data.email;
          var confidence = resObj.data.score || 0;
          
          emailRange.setValue(foundEmail);
          if (statusRange) statusRange.setValue("Found (Hunter)");
          if (headersMap["Email Confidence"]) sheet.getRange(r, headersMap["Email Confidence"]).setValue(confidence / 100);
          
          Logger.log("Hunter.io found email: " + foundEmail + " (confidence: " + confidence + "%)");
          email = foundEmail; // Update variable for verification block below
        }
      }
    } catch (err) {
      Logger.log("Hunter Finder Error: " + err.toString());
    }
  }
  
  // CASE B: Email exists -> Run Second-Pass Verification
  if (email) {
    // 1. ZeroBounce Verification
    if (zeroBounceKey) {
      try {
        Logger.log("Running ZeroBounce validation for: " + email);
        var zbUrl = "https://api.zerobounce.net/v2/validate?api_key=" + encodeURIComponent(zeroBounceKey) +
                    "&email=" + encodeURIComponent(email) + "&ip_address=";
                    
        var response = UrlFetchApp.fetch(zbUrl, { "muteHttpExceptions": true });
        if (response.getResponseCode() === 200) {
          var resObj = JSON.parse(response.getContentText());
          var zbStatus = (resObj.status || "").toLowerCase();
          
          Logger.log("ZeroBounce status: " + zbStatus);
          
          if (zbStatus === "valid") {
            if (statusRange) statusRange.setValue("Verified (ZeroBounce)");
            if (headersMap["Primary Email Catch-all Status"]) sheet.getRange(r, headersMap["Primary Email Catch-all Status"]).setValue("No");
            return {
              status: "Ready",
              reason: "ZeroBounce verified deliverable",
              emailUpdated: true
            };
          } else if (zbStatus === "invalid" || zbStatus === "do_not_mail") {
            if (statusRange) statusRange.setValue("Invalid (ZeroBounce)");
            return {
              status: "Flagged",
              reason: "ZeroBounce flagged: " + (resObj.sub_status || "invalid"),
              emailUpdated: true
            };
          } else {
            if (statusRange) statusRange.setValue("Risky (ZeroBounce: " + zbStatus + ")");
            return {
              status: "Risky",
              reason: "ZeroBounce labeled risky: " + zbStatus,
              emailUpdated: true
            };
          }
        }
      } catch (err) {
        Logger.log("ZeroBounce API Error: " + err.toString());
      }
    }
    
    // 2. Hunter.io Verification (Fallback/Alternative)
    if (hunterKey) {
      try {
        Logger.log("Running Hunter.io verification for: " + email);
        var verifyUrl = "https://api.hunter.io/v2/email-verifier?email=" + encodeURIComponent(email) +
                        "&api_key=" + encodeURIComponent(hunterKey);
                        
        var response = UrlFetchApp.fetch(verifyUrl, { "muteHttpExceptions": true });
        if (response.getResponseCode() === 200) {
          var resObj = JSON.parse(response.getContentText());
          var data = resObj.data || {};
          var hunterResult = (data.result || "").toLowerCase();
          
          Logger.log("Hunter.io verify result: " + hunterResult);
          
          if (hunterResult === "deliverable") {
            if (statusRange) statusRange.setValue("Verified (Hunter)");
            if (headersMap["Primary Email Catch-all Status"]) sheet.getRange(r, headersMap["Primary Email Catch-all Status"]).setValue("No");
            return {
              status: "Ready",
              reason: "Hunter.io verified deliverable",
              emailUpdated: true
            };
          } else if (hunterResult === "undeliverable") {
            if (statusRange) statusRange.setValue("Invalid (Hunter)");
            return {
              status: "Flagged",
              reason: "Hunter.io verified undeliverable",
              emailUpdated: true
            };
          } else {
            if (statusRange) statusRange.setValue("Risky (Hunter: " + hunterResult + ")");
            return {
              status: "Risky",
              reason: "Hunter.io verify result: " + hunterResult,
              emailUpdated: true
            };
          }
        }
      } catch (err) {
        Logger.log("Hunter.io Verifier Error: " + err.toString());
      }
    }
  }
  
  return {
    status: null,
    reason: null,
    emailUpdated: false
  };
}

/**
 * Validates a company's legitimacy and industry classification.
 * Runs before scoring.
 * 
 * @param {Sheet} sheet The Leads sheet reference
 * @param {number} rowNumber The row number to process (1-indexed)
 * @param {object} headersMap Mapping of header names to 1-indexed column numbers
 * @param {object} config Configuration parameters
 * @return {object} Object with {status, reason}
 */
function validateCompanyDetails(sheet, rowNumber, headersMap, config, mapping) {
  try {
    // Feature 1: read via literal header first, canonical mapping as fallback,
    // and never throw when a column is absent on generic-source sheets.
    var company = getFieldValue(sheet, rowNumber, headersMap, mapping, "Company", "company");
    var website = getFieldValue(sheet, rowNumber, headersMap, mapping, "Website", null);
    var industry = getFieldValue(sheet, rowNumber, headersMap, mapping, "Industry", "industry");
    var keywords = headersMap["Keywords"] ? (sheet.getRange(rowNumber, headersMap["Keywords"]).getValue() || "").toString().trim() : "";
    
    var statusRange = sheet.getRange(rowNumber, headersMap["Company Validation Status"]);
    
    // 1. Basic checks for missing/suspicious data
    if (!company) {
      statusRange.setValue("Rejected");
      return { status: "Rejected", reason: "Missing company name" };
    }
    
    var suspiciousNames = ["test", "placeholder", "unknown", "n/a", "none", "null", "company name", "asdf"];
    if (suspiciousNames.indexOf(company.toLowerCase()) !== -1 || company.length < 2) {
      statusRange.setValue("Needs Review");
      return { status: "Needs Review", reason: "Suspicious or generic company name: " + company };
    }
    
    if (!website) {
      statusRange.setValue("Needs Review");
      return { status: "Needs Review", reason: "Missing website URL" };
    }
    
    // Clean website URL
    var targetUrl = website;
    if (!/^https?:\/\//i.test(targetUrl)) {
      targetUrl = "https://" + targetUrl;
    }
    
    // 2. Try to fetch website description/HTML
    var webContent = "";
    var fetchSuccess = false;
    try {
      Logger.log("Fetching website for company validation: " + targetUrl);
      var response = UrlFetchApp.fetch(targetUrl, {
        "muteHttpExceptions": true,
        "followRedirects": true,
        "validateHttpsCertificates": false,
        "headers": {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
      });
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        // Extract title and meta description
        var titleMatch = html.match(/<title>([^<]+)<\/title>/i);
        var descMatch = html.match(/<meta[^>]*name=["']description["'][^>]*content=["']([^"']+)["']/i) || 
                          html.match(/<meta[^>]*content=["']([^"']+)["'][^>]*name=["']description["']/i);
                          
        var title = titleMatch ? titleMatch[1].trim() : "";
        var description = descMatch ? descMatch[1].trim() : "";
        
        webContent = "Title: " + title + "\nDescription: " + description;
        if (html.length > 0 && !description) {
          // Fallback to first 500 chars of body text (stripped of tags)
          var bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
          if (bodyMatch) {
            var bodyText = bodyMatch[1].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
            webContent += "\nSnippet: " + bodyText.substring(0, 500);
          }
        }
        fetchSuccess = true;
      } else {
        webContent = "Failed to fetch website (HTTP Status " + response.getResponseCode() + ")";
      }
    } catch (err) {
      webContent = "Failed to fetch website: " + err.toString();
    }
    
    // 3. Call LLM to validate company and industry
    var prompt = "You are an expert business analyst. Validate the following company details:\n\n" +
                 "Company Name: " + company + "\n" +
                 "Website: " + website + "\n" +
                 "Categorized Industry: " + industry + "\n" +
                 "Keywords: " + keywords + "\n" +
                 "Fetched Web Content:\n" + webContent + "\n\n" +
                 "Determine:\n" +
                 "1. Is this a legitimate, real company? (Respond false if it is a placeholder, fake, or cannot be verified)\n" +
                 "2. Is the categorized industry '" + industry + "' accurate? (Respond true if it is correct or reasonably close, false if it is completely wrong or misleading)\n" +
                 "3. Provide a brief 1-sentence explanation of your decision.\n\n" +
                 "OUTPUT FORMAT:\n" +
                 "Respond ONLY with a JSON object in this format:\n" +
                 "{\n" +
                 "  \"legitimate\": <true|false>,\n" +
                 "  \"industry_accurate\": <true|false>,\n" +
                 "  \"corrected_industry\": \"<suggested industry if inaccurate, else null>\",\n" +
                 "  \"reason\": \"<1-sentence explanation>\"\n" +
                 "}";
                 
    var textContent = callLlmForValidation(prompt, config);
    var cleanedJsonText = cleanJsonResponseText(textContent);
    var result = JSON.parse(cleanedJsonText);
    
    var finalStatus = "Needs Review";
    if (result.legitimate === false) {
      finalStatus = "Rejected";
    } else if (result.legitimate === true && result.industry_accurate === true) {
      finalStatus = "Verified";
    } else if (result.legitimate === true && result.industry_accurate === false) {
      finalStatus = "Needs Review"; // Flag for manual review since industry is inaccurate
    }
    
    statusRange.setValue(finalStatus);
    
    // Log the validation result
    Logger.log("Company validation for " + company + ": " + finalStatus + ". Reason: " + result.reason);
    
    return {
      status: finalStatus,
      reason: result.reason
    };
    
  } catch (e) {
    Logger.log("Error in validateCompanyDetails: " + e.toString());
    sheet.getRange(rowNumber, headersMap["Company Validation Status"]).setValue("Needs Review");
    return {
      status: "Needs Review",
      reason: "Validation error: " + e.toString()
    };
  }
}

/**
 * Placeholder for deeper company research validation (Feature 2).
 * Currently returns Unconfirmed as requested by the user until data source is confirmed.
 */
function researchCompanyDetails(sheet, rowNumber, headersMap, config) {
  try {
    var statusRange = sheet.getRange(rowNumber, headersMap["Research Status"]);
    statusRange.setValue("Unconfirmed");
    return {
      status: "Unconfirmed",
      reason: "Deeper research validation placeholder"
    };
  } catch (e) {
    Logger.log("Error in researchCompanyDetails: " + e.toString());
    return {
      status: "Flagged",
      reason: e.toString()
    };
  }
}

/**
 * Helper to call the LLM for company validation. Uses Gemini 2.0 Flash as primary,
 * with failover to other providers if needed.
 */
function callLlmForValidation(prompt, config) {
  var modelName = config.model || "gemini-2.0-flash";
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
        "temperature": 0.1
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
          return callOpenAiCompatibleApi(provider.url, modelToUse, key, prompt, 0.1, true);
        } catch (err) {
          lastApiError += (lastApiError ? " | " : "") + provider.name + " Error: " + err.toString();
          Utilities.sleep(500);
        }
      }
    }
    
    throw new Error("All validation LLM providers failed. Details: " + lastApiError);
  }
}

// HELLO
