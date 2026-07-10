/**
 * AI Lead Generation Engine - Lead Scoring via Gemini, Groq, & Cerebras APIs
 * File: Scoring.gs
 */

/**
 * Scores a single lead row using Gemini, Groq, or Cerebras API.
 * Handles automatic key rotation, key-to-key retry on rate limit, and provider-level failover.
 * 
 * @param {Sheet} sheet The Leads sheet reference
 * @param {number} rowNumber The row number to process (1-indexed)
 * @param {object} headersMap Mapping of header names to 1-indexed column numbers
 * @param {object} config Configuration parameters read from Config sheet
 * @param {object} mapping Result of resolveColumnMapping
 * @return {object} Object with {success, score, reason, error}
 */
function scoreSingleLead(sheet, rowNumber, headersMap, config, mapping) {
  try {
    // 1. Extract values for the row using generic mapping where possible
    var company = getCanonical(sheet, rowNumber, mapping, "company") ||
                  (headersMap["Company"] ? sheet.getRange(rowNumber, headersMap["Company"]).getValue() : "");
    var annualRevenue = getCanonical(sheet, rowNumber, mapping, "revenue") ||
                        (headersMap["Annual Revenue"] ? sheet.getRange(rowNumber, headersMap["Annual Revenue"]).getValue() : "");
    var totalFunding = headersMap["Total Funding"] ? sheet.getRange(rowNumber, headersMap["Total Funding"]).getValue() : "";
    var latestFunding = headersMap["Latest Funding"] ? sheet.getRange(rowNumber, headersMap["Latest Funding"]).getValue() : "";
    var latestFundingAmount = headersMap["Latest Funding Amount"] ? sheet.getRange(rowNumber, headersMap["Latest Funding Amount"]).getValue() : "";
    var lastRaisedAt = getCanonical(sheet, rowNumber, mapping, "funding_date") || (headersMap["Last Raised At"] ? sheet.getRange(rowNumber, headersMap["Last Raised At"]).getValue() : "");
    var hiringStatus = headersMap["Hiring Status"] ? sheet.getRange(rowNumber, headersMap["Hiring Status"]).getValue() : "";
    var employees = getCanonical(sheet, rowNumber, mapping, "employee_count");
    var technologies = headersMap["Technologies"] ? sheet.getRange(rowNumber, headersMap["Technologies"]).getValue() : "";
    
    // City and Country aren't canonicalized yet, use headersMap
    var city = headersMap["City"] ? sheet.getRange(rowNumber, headersMap["City"]).getValue() : "";
    var country = headersMap["Country"] ? sheet.getRange(rowNumber, headersMap["Country"]).getValue() : "";
    
    // 2. Build the prompt using the template and values
    var prompt = config.promptTemplate
      .replace(/{Revenue Cutoff}/g, config.revenueCutoff)
      .replace(/{Funding Recency Window}/g, config.fundingWindow)
      .replace(/{Company}/g, company)
      .replace(/{Annual Revenue}/g, annualRevenue)
      .replace(/{Total Funding}/g, totalFunding)
      .replace(/{Latest Funding}/g, latestFunding)
      .replace(/{Latest Funding Amount}/g, latestFundingAmount)
      .replace(/{Last Raised At}/g, lastRaisedAt)
      .replace(/{Hiring Status}/g, hiringStatus);
      
    // Handle new placeholders
    var location = [city, country].filter(function(v) { return v && v.toString().trim() !== ""; }).join(", ");
    prompt = prompt.replace(/{Employees}/g, employees)
                   .replace(/{Technologies}/g, technologies)
                   .replace(/{Location}/g, location);

    // Feature 1: drop LEAD DATA lines whose value is empty so the model only
    // ever sees fields that are actually populated for this row.
    prompt = prompt.split("\n").filter(function(line) {
      return !/^(Company|Annual Revenue|Total Funding|Latest Funding|Latest Funding Amount|Last Raised At|Hiring Status|Employees|Technologies|Location):\s*$/.test(line.trim());
    }).join("\n");


    // 3. Determine primary and fallback providers
    var modelName = config.model || "gemini-2.0-flash";
    var isGemini = (modelName.indexOf("gemini") !== -1);
    
    var textContent;
    
    if (isGemini) {
      // --- GEMINI FLOW ---
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
      var responseCode = response.getResponseCode();
      var responseText = response.getContentText();
      
      if (responseCode !== 200) {
        throw new Error("Gemini API returned status code " + responseCode + ": " + responseText);
      }
      
      var responseJson = JSON.parse(responseText);
      if (!responseJson.candidates || responseJson.candidates.length === 0) {
        throw new Error("No response candidates returned from Gemini API.");
      }
      textContent = responseJson.candidates[0].content.parts[0].text;
      
    } else {
      // --- THREE-WAY FAILOVER FLOW (NVIDIA DEEPSEEK / GROQ / CEREBRAS) WITH MULTI-KEY ROTATION ---
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
      
      // Determine the primary provider based on the model name
      var primaryName = "Groq"; // Default
      var modelNameLower = modelName.toLowerCase();
      if (modelNameLower.indexOf("deepseek") !== -1 || modelNameLower.indexOf("nvidia") !== -1) {
        primaryName = "NVIDIA";
      } else if (modelNameLower.indexOf("cerebras") !== -1 || modelNameLower.indexOf("zai-glm") !== -1) {
        primaryName = "Cerebras";
      } else if (modelNameLower.indexOf("groq") !== -1 || modelNameLower.indexOf("llama") !== -1) {
        primaryName = "Groq";
      }
      
      // Sort providers so the primary provider is tried first
      providers.sort(function(a, b) {
        if (a.name === primaryName) return -1;
        if (b.name === primaryName) return 1;
        return 0;
      });
      
      var callSuccess = false;
      var lastApiError = "";
      
      for (var i = 0; i < providers.length; i++) {
        var provider = providers[i];
        if (provider.keys.length === 0) {
          lastApiError += (lastApiError ? " | " : "") + "No keys for " + provider.name;
          Logger.log("Row " + rowNumber + ": Skipping " + provider.name + " due to missing keys.");
          continue;
        }
        
        // Use the configured model name if this provider is the primary one,
        // otherwise use the provider's default fallback model
        var modelToUse = (provider.name === primaryName) ? modelName : provider.defaultModel;
        var shuffledKeys = shuffleArray(provider.keys);
        var providerSuccess = false;
        
        for (var k = 0; k < shuffledKeys.length; k++) {
          var key = shuffledKeys[k];
          try {
            Logger.log("Row " + rowNumber + ": Attempting " + provider.name + " (" + modelToUse + ") - Key " + (k + 1) + " of " + shuffledKeys.length);
            textContent = callOpenAiCompatibleApi(provider.url, modelToUse, key, prompt, config.temperature);
            providerSuccess = true;
            callSuccess = true;
            break; // Exit key loop on success
          } catch (err) {
            var errStr = err.toString();
            lastApiError += (lastApiError ? " | " : "") + provider.name + " Error: " + errStr;
            Logger.log("Row " + rowNumber + ": " + provider.name + " (Key " + (k + 1) + ") failed: " + errStr);
            Utilities.sleep(500); // Short pause before next key
          }
        }
        
        if (providerSuccess) {
          break; // Exit provider loop on success
        }
      }
      
      if (!callSuccess) {
        throw new Error("All Primary and Fallback providers failed. Details: " + lastApiError);
      }
    }
    
    // 4. Clean & Parse JSON response
    var cleanedJsonText = cleanJsonResponseText(textContent);
    var result = JSON.parse(cleanedJsonText);
    
    // Validate required fields in the JSON response
    if (result.score === undefined || result.category === undefined || !result.reason) {
      throw new Error("AI response missing required JSON properties (score, category, reason). Got: " + textContent);
    }
    
    // Write score and reason back to the sheet
    sheet.getRange(rowNumber, headersMap["Score"]).setValue(result.score);
    sheet.getRange(rowNumber, headersMap["Score Reason"]).setValue(result.reason);
    
    return {
      success: true,
      score: result.score,
      reason: result.reason,
      category: result.category
    };
    
  } catch (e) {
    var errorMsg = e.toString();
    Logger.log("Error scoring row " + rowNumber + ": " + errorMsg);
    // Write error to both Score and Score Reason cells
    sheet.getRange(rowNumber, headersMap["Score"]).setValue("Error");
    sheet.getRange(rowNumber, headersMap["Score Reason"]).setValue("Error: " + errorMsg);
    return {
      success: false,
      error: errorMsg
    };
  }
}

/**
 * Calls an OpenAI-compatible Chat Completions API (Groq/Cerebras/NVIDIA).
 */
function callOpenAiCompatibleApi(url, model, apiKey, prompt, temperature, isJsonMode) {
  if (isJsonMode === undefined) {
    isJsonMode = true;
  }
  var headers = {
    "Authorization": "Bearer " + apiKey,
    "Content-Type": "application/json"
  };
  var payload = {
    "model": model,
    "messages": [
      {
        "role": "user",
        "content": prompt
      }
    ],
    "temperature": parseFloat(temperature) || 0.2
  };
  
  if (isJsonMode) {
    payload["response_format"] = { "type": "json_object" };
  }
  
  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error("HTTP " + responseCode + ": " + responseText);
  }
  
  var responseJson = JSON.parse(responseText);
  if (!responseJson.choices || responseJson.choices.length === 0) {
    throw new Error("No response choices returned from API.");
  }
  
  return responseJson.choices[0].message.content;
}

/**
 * Returns a list of API keys from script properties (supporting comma-separated strings).
 */
function getApiKeysList(primaryProp, fallbackProp) {
  var props = PropertiesService.getScriptProperties();
  var primaryVal = props.getProperty(primaryProp) || "";
  var fallbackVal = props.getProperty(fallbackProp) || "";
  
  var combined = (primaryVal + "," + fallbackVal).split(",");
  var keys = [];
  combined.forEach(function(k) {
    var trimmed = k.trim();
    if (trimmed) {
      keys.push(trimmed);
    }
  });
  return keys;
}

/**
 * Randomly shuffles an array (Fisher-Yates Shuffle).
 */
function shuffleArray(array) {
  var shuffled = array.slice();
  for (var i = shuffled.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = shuffled[i];
    shuffled[i] = shuffled[j];
    shuffled[j] = temp;
  }
  return shuffled;
}

/**
 * Picks a random element from an array.
 */
function getRandomElement(arr) {
  if (!arr || arr.length === 0) return null;
  var idx = Math.floor(Math.random() * arr.length);
  return arr[idx];
}

/**
 * Cleans the model's text response to ensure valid, parsable JSON
 * by stripping markdown backticks if present.
 */
function cleanJsonResponseText(text) {
  var cleaned = text.trim();
  if (cleaned.indexOf("```json") === 0) {
    cleaned = cleaned.substring(7);
  } else if (cleaned.indexOf("```") === 0) {
    cleaned = cleaned.substring(3);
  }
  if (cleaned.lastIndexOf("```") === cleaned.length - 3 && cleaned.length >= 3) {
    cleaned = cleaned.substring(0, cleaned.length - 3);
  }
  return cleaned.trim();
}
