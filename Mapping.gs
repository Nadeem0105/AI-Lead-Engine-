/**
 * AI Lead Generation Engine - Generic, source-agnostic column mapping (Batch 2, Feature 1)
 * File: Mapping.gs
 *
 * Leads arrive from Apollo, LinkedIn X-Ray, GitHub, Google Maps, and manual DB imports,
 * each with different header names. This layer resolves the ACTUAL header row of a sheet
 * against a canonical field dictionary ONCE (cached in the hidden "_FieldMap" sheet), so
 * downstream logic can read canonical fields without assuming any source's exact layout.
 */

/**
 * Canonical field dictionary: canonical key -> list of accepted header aliases.
 * Matching is fuzzy (case-insensitive, non-alphanumeric stripped), so "Email Address",
 * "email_address", and "E-mail Address" all resolve to the same canonical key.
 */
var CANONICAL_FIELDS = {
  "email":          ["Email", "Email Address", "Work Email", "Contact Email", "E-mail", "Primary Email"],
  "company":        ["Company", "Company Name", "Organization", "Org", "Company Name for Emails", "Employer"],
  "first_name":     ["First Name", "Firstname", "Given Name", "Fname"],
  "last_name":      ["Last Name", "Lastname", "Surname", "Family Name", "Lname"],
  "full_name":      ["Full Name", "Contact Name", "Name", "Person Name"],
  // NOTE: "Keywords" is intentionally NOT included here — Apollo Keywords are product/service
  // tags (e.g. "it staffing", "bpo") and not the primary industry field. Including them caused
  // IT outsourcing companies to be falsely flagged by the recruitment filter.
  "industry":       ["Industry", "Sector", "Vertical", "Industry/Keywords"],
  // Keywords gets its own canonical key for use in scoring prompts only.
  "keywords":       ["Keywords", "Company Keywords", "Tags"],
  "revenue":        ["Annual Revenue", "Revenue", "Company Revenue", "Est. Revenue"],
  "funding_date":   ["Last Raised At", "Funding Date", "Latest Funding Date", "Last Funding Date"],
  "linkedin_url":   ["LinkedIn URL", "Linkedin", "LinkedIn Profile", "Person Linkedin Url", "Linkedin Url"],
  "title":          ["Title", "Job Title", "Position", "Role"],
  "employee_count": ["# Employees", "Employees", "Employee Count", "Headcount", "Company Size", "Num Employees"],
  "source":         ["Source", "Lead Source", "Origin", "Data Source"]
};

/**
 * Normalizes a header string for fuzzy comparison: lowercase, strip everything
 * except letters and digits. "Work Email" -> "workemail", "email_address" -> "emailaddress".
 *
 * @param {string} s
 * @return {string}
 */
function normalizeHeader(s) {
  return (s === undefined || s === null ? "" : s.toString()).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Resolves the canonical column mapping for a sheet's actual header row.
 * Result is cached in the hidden "_FieldMap" sheet keyed by sheet name so it is
 * not recomputed on every row. Pass forceRefresh = true to rebuild after a header change.
 *
 * @param {Sheet} sheet The sheet whose header row to resolve.
 * @param {boolean} forceRefresh Rebuild the cache even if a stored mapping exists.
 * @return {object} Map of canonical key -> 1-indexed column number (only mapped keys present).
 */
function resolveColumnMapping(sheet, forceRefresh) {
  var sheetName = sheet.getName();

  if (!forceRefresh) {
    var cached = readCachedMapping(sheetName);
    if (cached) return cached;
  }

  // Build a normalized-alias lookup once: normalizedAlias -> canonicalKey.
  var aliasLookup = {};
  for (var canonicalKey in CANONICAL_FIELDS) {
    if (!CANONICAL_FIELDS.hasOwnProperty(canonicalKey)) continue;
    var aliases = CANONICAL_FIELDS[canonicalKey];
    for (var a = 0; a < aliases.length; a++) {
      aliasLookup[normalizeHeader(aliases[a])] = canonicalKey;
    }
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var mapping = {};
  for (var c = 0; c < headers.length; c++) {
    var norm = normalizeHeader(headers[c]);
    if (!norm) continue;
    var canonical = aliasLookup[norm];
    // First mapped column wins for a given canonical key (don't overwrite).
    if (canonical && !mapping[canonical]) {
      mapping[canonical] = c + 1; // 1-indexed
    }
  }

  writeCachedMapping(sheetName, mapping);
  return mapping;
}

/**
 * Reads a stored mapping for a sheet from the hidden "_FieldMap" sheet.
 * @return {object|null} The mapping, or null if none stored.
 */
function readCachedMapping(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapSheet = ss.getSheetByName("_FieldMap");
  if (!mapSheet || mapSheet.getLastRow() < 1) return null;

  var data = mapSheet.getRange(1, 1, mapSheet.getLastRow(), 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === sheetName) {
      try {
        var parsed = JSON.parse(data[i][1]);
        return (parsed && typeof parsed === "object") ? parsed : null;
      } catch (e) {
        return null;
      }
    }
  }
  return null;
}

/**
 * Writes/updates the stored mapping for a sheet in the hidden "_FieldMap" sheet,
 * creating and hiding the sheet on first use.
 */
function writeCachedMapping(sheetName, mapping) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapSheet = ss.getSheetByName("_FieldMap");
  if (!mapSheet) {
    mapSheet = ss.insertSheet("_FieldMap");
    mapSheet.getRange(1, 1, 1, 2).setValues([["Sheet Name", "Mapping JSON"]]);
    mapSheet.hideSheet();
  }

  var json = JSON.stringify(mapping);
  var lastRow = mapSheet.getLastRow();
  if (lastRow >= 1) {
    var keys = mapSheet.getRange(1, 1, lastRow, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0].toString().trim() === sheetName) {
        mapSheet.getRange(i + 1, 2).setValue(json);
        return;
      }
    }
  }
  mapSheet.appendRow([sheetName, json]);
}

/**
 * Reads a canonical field value for a given row using a resolved mapping.
 * Returns "" for any field that is not mapped or is empty — so callers can
 * safely include only fields that actually have data (no empty placeholders).
 *
 * @param {Sheet} sheet
 * @param {number} rowNumber 1-indexed row.
 * @param {object} mapping Result of resolveColumnMapping.
 * @param {string} canonicalKey e.g. "email", "company".
 * @return {string} Trimmed string value, or "".
 */
function getCanonical(sheet, rowNumber, mapping, canonicalKey) {
  var col = mapping[canonicalKey];
  if (!col) return "";
  var val = sheet.getRange(rowNumber, col).getValue();
  return (val === undefined || val === null) ? "" : val.toString().trim();
}

/**
 * Builds a "Field: value" block containing ONLY canonical fields that have a
 * non-empty value for this row. Used by prompt builders so we never send empty
 * placeholders or reference missing fields (Feature 1 requirement).
 *
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {object} mapping
 * @param {Array<Array<string>>} fields Array of [canonicalKey, label] pairs to include in order.
 * @return {string} Newline-joined "- Label: value" lines for present fields only.
 */
function buildPresentFieldsBlock(sheet, rowNumber, mapping, fields) {
  var lines = [];
  for (var i = 0; i < fields.length; i++) {
    var key = fields[i][0];
    var label = fields[i][1];
    var val = getCanonical(sheet, rowNumber, mapping, key);
    if (val !== "") {
      lines.push("- " + label + ": " + val);
    }
  }
  return lines.join("\n");
}

/**
 * Reads a field value by literal header name first, then by canonical mapping.
 * Lets Apollo-format sheets keep working while generic sheets (e.g. "Organization"
 * instead of "Company") resolve through the alias layer (Feature 1).
 *
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {object} headersMap Result of getHeadersMap.
 * @param {object} mapping Result of resolveColumnMapping (may be null).
 * @param {string} literalHeader Exact header name to try first, e.g. "Company".
 * @param {string} canonicalKey Canonical fallback key, e.g. "company" (may be null).
 * @return {string} Trimmed string value, or "" if neither column exists.
 */
function getFieldValue(sheet, rowNumber, headersMap, mapping, literalHeader, canonicalKey) {
  if (literalHeader && headersMap[literalHeader]) {
    var val = sheet.getRange(rowNumber, headersMap[literalHeader]).getValue();
    var str = (val === undefined || val === null) ? "" : val.toString().trim();
    if (str !== "") return str;
  }
  if (canonicalKey && mapping) {
    return getCanonical(sheet, rowNumber, mapping, canonicalKey);
  }
  return "";
}

/**
 * Guard used at the top of processing: if a row has no mappable email, flag it in
 * Pipeline Stage and signal the caller to skip it (rather than erroring the batch).
 *
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {object} mapping
 * @param {object} headersMap Standard getHeadersMap for writing Pipeline Stage.
 * @return {boolean} true if the row is missing email and was flagged (caller should skip).
 */
function flagIfMissingEmail(sheet, rowNumber, mapping, headersMap) {
  var email = getCanonical(sheet, rowNumber, mapping, "email");
  if (email === "") {
    if (headersMap["Pipeline Stage"]) {
      sheet.getRange(rowNumber, headersMap["Pipeline Stage"]).setValue("Missing: email");
    }
    return true;
  }
  return false;
}
