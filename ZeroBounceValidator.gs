/**
 * AI Lead Generation Engine - ZeroBounce Validation & Send-Priority Queue
 * File: ZeroBounceValidator.gs
 *
 * Email Ops Upgrade. Wraps the ZeroBounce API to (1) classify each lead's email into
 * a ZB Status, (2) map that status onto a send priority, and (3) build a deliverability-
 * protected daily queue that puts verified emails first and only lets a capped share of
 * risky "catch-all" addresses through — the mechanism that drives the 6% -> 2% bounce
 * target.
 *
 * A "lead" here is a plain object. Recognized fields (all optional except email):
 *   { email, score, row, zbStatus, zbScore, sendPriority }
 * Functions mutate the object in place AND return the relevant result so they compose
 * whether callers work with objects or return values.
 */

// Default ZeroBounce catch-all AI score below which a catch-all lead is not worth the
// deliverability risk. Overridable via the Config key "CATCH_ALL_SCORE_THRESHOLD".
var CATCH_ALL_SCORE_THRESHOLD_DEFAULT = 7;

// Share of the final daily queue that may be catch-all ("capped") emails.
// The remaining capacity is reserved for verified ("high") emails.
var CATCH_ALL_QUEUE_RATIO = 0.20; // 20% rule (per day, across the whole queue)

/**
 * Resolves the ZeroBounce API key from Script Properties (preferred) or Config.
 * @return {string} The key, or "" if none configured.
 */
function getZeroBounceKey_(config) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty("ZEROBOUNCE_API_KEY") || "";
  if (!key && config) {
    key = getConfigValue(config, "zeroBounceKey", "") ||
          getConfigValue(config, "ZeroBounce API Key (Optional)", "");
  }
  return (key || "").toString().trim();
}

/**
 * Reads the catch-all score threshold from Config, falling back to the module default.
 */
function getCatchAllScoreThreshold_(config) {
  var raw = config ? getConfigValue(config, "CATCH_ALL_SCORE_THRESHOLD", CATCH_ALL_SCORE_THRESHOLD_DEFAULT) : CATCH_ALL_SCORE_THRESHOLD_DEFAULT;
  var num = parseFloat(raw);
  return isNaN(num) ? CATCH_ALL_SCORE_THRESHOLD_DEFAULT : num;
}

/**
 * Extracts the email address from a lead (object or bare string).
 */
function leadEmail_(lead) {
  if (lead === null || lead === undefined) return "";
  if (typeof lead === "string") return lead.trim();
  return (lead.email || lead.Email || "").toString().trim();
}

/**
 * Calls ZeroBounce for a single lead and maps the response onto a ZB Status
 * ("Valid" | "Catch-all" | "Invalid" | "Spamtrap" | "Abuse" | "Do Not Mail" | "Unknown")
 * plus a numeric ZB Score (0-10). For catch-all/unknown addresses the AI scoring
 * endpoint is queried to obtain a confidence score used later by the 20% rule.
 *
 * If no API key is configured, the email is treated as "Unknown" (never silently
 * assumed deliverable) so downstream priority assignment stays conservative.
 *
 * @param {object|string} lead
 * @param {object} [config] Optional pre-loaded config (loaded on demand otherwise).
 * @return {object} { status, score, subStatus } — also written back onto the lead object.
 */
function validateWithZeroBounce(lead, config) {
  config = config || (typeof getConfig === "function" ? safeGetConfig_() : null);
  var email = leadEmail_(lead);
  var result = { status: "Unknown", score: 0, subStatus: "" };

  if (!email) {
    result.status = "Invalid";
    result.subStatus = "missing_email";
    return applyZbResult_(lead, result);
  }

  var apiKey = getZeroBounceKey_(config);
  if (!apiKey) {
    Logger.log("ZeroBounceValidator: no ZEROBOUNCE_API_KEY configured; marking '" + email + "' as Unknown.");
    return applyZbResult_(lead, result);
  }

  try {
    var url = "https://api.zerobounce.net/v2/validate?api_key=" + encodeURIComponent(apiKey) +
              "&email=" + encodeURIComponent(email) + "&ip_address=";
    var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (response.getResponseCode() === 200) {
      var obj = JSON.parse(response.getContentText());
      var raw = (obj.status || "").toString().toLowerCase();
      result.subStatus = (obj.sub_status || "").toString();
      result.status = mapZeroBounceStatus_(raw);

      if (result.status === "Valid") {
        result.score = 10;
      } else if (result.status === "Catch-all" || result.status === "Unknown") {
        // Query the AI scoring endpoint for a 0-10 confidence on the catch-all/unknown domain.
        result.score = getZeroBounceAiScore_(email, apiKey);
      } else {
        result.score = 0; // Invalid / Spamtrap / Abuse / Do Not Mail
      }
    } else {
      Logger.log("ZeroBounceValidator: HTTP " + response.getResponseCode() + " for '" + email + "'. Marking Unknown.");
    }
  } catch (e) {
    Logger.log("ZeroBounceValidator.validateWithZeroBounce error for '" + email + "': " + e.toString());
  }

  return applyZbResult_(lead, result);
}

/**
 * Writes the validation result back onto the lead object (if it is an object) and
 * returns the result.
 */
function applyZbResult_(lead, result) {
  if (lead && typeof lead === "object") {
    lead.zbStatus = result.status;
    lead.zbScore = result.score;
  }
  return result;
}

/**
 * Wrapper so validateWithZeroBounce can run standalone without a caller-supplied config.
 */
function safeGetConfig_() {
  try { return getConfig(); } catch (e) { return null; }
}

/**
 * Maps a raw ZeroBounce status string to our normalized ZB Status enum.
 */
function mapZeroBounceStatus_(raw) {
  switch (raw) {
    case "valid":       return "Valid";
    case "catch-all":
    case "catch_all":   return "Catch-all";
    case "invalid":     return "Invalid";
    case "spamtrap":    return "Spamtrap";
    case "abuse":       return "Abuse";
    case "do_not_mail": return "Do Not Mail";
    case "unknown":
    default:            return "Unknown";
  }
}

/**
 * Calls the ZeroBounce AI scoring endpoint to grade a catch-all/unknown address 0-10.
 * ZeroBounce field names vary by plan, so several candidate fields are probed.
 * Returns 0 on any failure (fails safe — low score keeps risky mail out of the queue).
 */
function getZeroBounceAiScore_(email, apiKey) {
  try {
    var url = "https://api.zerobounce.net/v2/scoring?api_key=" + encodeURIComponent(apiKey) +
              "&email=" + encodeURIComponent(email);
    var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (response.getResponseCode() === 200) {
      var obj = JSON.parse(response.getContentText());
      var candidates = [obj.score, obj.quality_score, obj.ZeroBounceQualityScore, obj.ZBAIScore, obj.ai_score];
      for (var i = 0; i < candidates.length; i++) {
        var val = parseFloat(candidates[i]);
        if (!isNaN(val)) return val;
      }
    }
  } catch (e) {
    Logger.log("ZeroBounceValidator.getZeroBounceAiScore_ error for '" + email + "': " + e.toString());
  }
  return 0;
}

/**
 * Maps a lead's ZB Status onto a send priority:
 *   "high"    — verified deliverable (ZB Status = Valid)
 *   "capped"  — risky-but-mailable catch-all/unknown (subject to the 20% rule)
 *   "blocked" — invalid / spamtrap / abuse / do-not-mail (never send)
 *
 * If the lead has not yet been validated (no zbStatus), it is validated first.
 *
 * @param {object} lead
 * @param {object} [config]
 * @return {string} "high" | "capped" | "blocked" — also written to lead.sendPriority.
 */
function assignSendPriority(lead, config) {
  var status = (lead && typeof lead === "object") ? lead.zbStatus : null;
  if (!status) {
    validateWithZeroBounce(lead, config);
    status = (lead && typeof lead === "object") ? lead.zbStatus : "Unknown";
  }

  var priority;
  switch (status) {
    case "Valid":
      priority = "high";
      break;
    case "Catch-all":
    case "Unknown":
      priority = "capped";
      break;
    case "Invalid":
    case "Spamtrap":
    case "Abuse":
    case "Do Not Mail":
    default:
      priority = "blocked";
      break;
  }

  if (lead && typeof lead === "object") {
    lead.sendPriority = priority;
  }
  return priority;
}

/**
 * Builds the deliverability-protected daily send queue.
 *
 * Rules:
 *   1. All "high" (verified) leads are included, ordered by descending score.
 *   2. "capped" (catch-all/unknown) leads are eligible only if their ZB Score meets the
 *      CATCH_ALL_SCORE_THRESHOLD; they are then added highest-score-first, but only up to
 *      the point where catch-alls make up at most CATCH_ALL_QUEUE_RATIO (20%) of the final
 *      queue for the day.
 *   3. "blocked" leads are dropped entirely.
 *
 * Given H verified leads, the max catch-alls C that keeps C/(H+C) <= 20% is
 *   C = floor( H * ratio / (1 - ratio) )   ->  floor(H * 0.25) for ratio = 0.20.
 *
 * @param {Array<object>} leads Array of lead objects (each with email and optional score).
 * @param {object} [config]
 * @return {Array<object>} The ordered send queue.
 */
function buildPrioritizedQueue(leads, config) {
  config = config || safeGetConfig_();
  var threshold = getCatchAllScoreThreshold_(config);

  if (!leads || leads.length === 0) return [];

  var high = [];
  var capped = [];

  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    // Ensure each lead has a priority (validates via ZeroBounce if needed).
    var priority = (lead && lead.sendPriority) ? lead.sendPriority : assignSendPriority(lead, config);

    if (priority === "high") {
      high.push(lead);
    } else if (priority === "capped") {
      var zbScore = parseFloat(lead && lead.zbScore);
      if (!isNaN(zbScore) && zbScore >= threshold) {
        capped.push(lead);
      } else {
        Logger.log("buildPrioritizedQueue: dropping catch-all '" + leadEmail_(lead) +
                   "' (ZB Score " + (isNaN(zbScore) ? "n/a" : zbScore) + " < threshold " + threshold + ").");
      }
    }
    // "blocked" leads are silently excluded.
  }

  // Sort helper: descending by lead score (falls back to zbScore, then 0).
  function byScoreDesc(a, b) {
    return leadSortScore_(b) - leadSortScore_(a);
  }
  high.sort(byScoreDesc);
  capped.sort(function(a, b) {
    // For catch-alls, prioritise by ZeroBounce confidence first, then lead score.
    var d = (parseFloat(b.zbScore) || 0) - (parseFloat(a.zbScore) || 0);
    if (d !== 0) return d;
    return byScoreDesc(a, b);
  });

  // Apply the 20% cap: catch-alls may be at most ratio/(1-ratio) of the verified count.
  var maxCatchAll = Math.floor(high.length * CATCH_ALL_QUEUE_RATIO / (1 - CATCH_ALL_QUEUE_RATIO));
  var acceptedCapped = capped.slice(0, maxCatchAll);

  if (capped.length > acceptedCapped.length) {
    Logger.log("buildPrioritizedQueue: 20% rule capped catch-alls at " + acceptedCapped.length +
               " of " + capped.length + " eligible (verified=" + high.length + ").");
  }

  return high.concat(acceptedCapped);
}

/**
 * Numeric sort key for a lead: prefers an explicit lead score, else ZB score, else 0.
 */
function leadSortScore_(lead) {
  if (!lead) return 0;
  var s = parseFloat(lead.score);
  if (!isNaN(s)) return s;
  var z = parseFloat(lead.zbScore);
  return isNaN(z) ? 0 : z;
}
