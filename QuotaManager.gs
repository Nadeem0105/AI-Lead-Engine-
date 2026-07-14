/**
 * AI Lead Generation Engine - Per-Account Daily Quota Manager
 * File: QuotaManager.gs
 *
 * Email Ops Upgrade. Reads and writes the "Account Config" sheet which tracks a
 * per-mailbox daily send allowance. Columns (created in Setup.gs):
 *   Account | Daily Quota | Sent Today Count | Last Reset Date
 *
 * Quota values (Daily Quota) are entered manually per mailbox based on age/warmup
 * (see open question #5 in the implementation spec). This module never invents a
 * quota — it only resets counters, reads remaining allowance, and records sends.
 */

// Name of the sheet that stores per-account quota state.
var ACCOUNT_CONFIG_SHEET_NAME = "Account Config";

// Header order for the Account Config sheet. Kept here so this module can
// self-heal (recreate the sheet) even before Setup.gs has been run.
var ACCOUNT_CONFIG_HEADERS = ["Account", "Daily Quota", "Sent Today Count", "Last Reset Date"];

// Quota an account is seeded with when the Account Config sheet is first bootstrapped.
// It is intentionally a placeholder: a per-mailbox Daily Quota must be set manually
// (based on mailbox age/warmup) before the daily run is allowed to proceed. Any account
// left sitting at exactly this value is treated as "unconfigured" by
// validateAccountConfigBeforeRun().
var BOOTSTRAP_DEFAULT_QUOTA = 40;

/**
 * Returns the Account Config sheet, creating it with default headers + one row per
 * configured account if it does not exist yet. This keeps QuotaManager usable even
 * if Setup.gs has not (yet) initialized the sheet.
 *
 * @return {Sheet} The Account Config sheet.
 */
function getAccountConfigSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ACCOUNT_CONFIG_SHEET_NAME);
  if (sheet) return sheet;

  // Bootstrap: create the sheet and seed it from the account pools in Config.
  sheet = ss.insertSheet(ACCOUNT_CONFIG_SHEET_NAME);
  sheet.appendRow(ACCOUNT_CONFIG_HEADERS);
  sheet.getRange(1, 1, 1, ACCOUNT_CONFIG_HEADERS.length).setFontWeight("bold").setBackground("#f3f4f6");

  var defaultQuota = BOOTSTRAP_DEFAULT_QUOTA;
  var accounts = [];
  try {
    var config = getConfig();
    defaultQuota = parseInt(getConfigValue(config, "Per Account Daily Cap", BOOTSTRAP_DEFAULT_QUOTA.toString())) || BOOTSTRAP_DEFAULT_QUOTA;
    var pools = [
      getConfigValue(config, "Score Band High Accounts", ""),
      getConfigValue(config, "Score Band Mid Accounts", ""),
      getConfigValue(config, "Default Send Account", "Account A")
    ].join(",");
    pools.split(",").forEach(function(a) {
      var acc = a.toString().trim();
      if (acc && accounts.indexOf(acc) === -1) accounts.push(acc);
    });
  } catch (e) {
    Logger.log("QuotaManager: could not read Config while bootstrapping Account Config: " + e);
  }
  if (accounts.length === 0) {
    accounts = ["Account A", "Account B", "Account C", "Account D", "Account E"];
  }

  var today = quotaTodayString_();
  accounts.forEach(function(acc) {
    sheet.appendRow([acc, defaultQuota, 0, today]);
  });
  Logger.log("QuotaManager: created '" + ACCOUNT_CONFIG_SHEET_NAME + "' sheet with " + accounts.length + " accounts.");
  return sheet;
}

/**
 * Today's date as a yyyy-MM-dd string in the script timezone.
 */
function quotaTodayString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Normalizes a date-ish cell value (Date object or string) to a yyyy-MM-dd string.
 * Returns "" for blank/unparseable values.
 */
function quotaDateString_(value) {
  if (value === "" || value === null || value === undefined) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var str = value.toString().trim();
  if (!str) return "";
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  // Already a yyyy-MM-dd style string we can't parse into a Date — compare literally.
  return str;
}

/**
 * Builds a header-name -> 0-indexed column position map for the Account Config sheet,
 * tolerating column reordering.
 */
function accountConfigColMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) { map[h.toString().trim()] = i; });
  return map;
}

/**
 * Checks the Last Reset Date for every account and, if it is not today, resets the
 * Sent Today Count to 0 and stamps today's date. Idempotent — safe to run repeatedly;
 * once reset for the day, subsequent calls are no-ops. Intended to be called at the
 * start of the daily prep window (Scheduler.prepareDailyQueue).
 *
 * @return {number} How many account rows were reset this call.
 */
function resetDailyQuotasIfNeeded() {
  var sheet = getAccountConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var cols = accountConfigColMap_(sheet);
  var countCol = cols["Sent Today Count"];
  var resetCol = cols["Last Reset Date"];
  if (countCol === undefined || resetCol === undefined) {
    Logger.log("QuotaManager.resetDailyQuotasIfNeeded: Account Config missing required columns.");
    return 0;
  }

  var today = quotaTodayString_();
  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
  var resetCount = 0;

  for (var i = 0; i < data.length; i++) {
    var lastReset = quotaDateString_(data[i][resetCol]);
    if (lastReset !== today) {
      // Not reset today — zero the counter and stamp today's date.
      sheet.getRange(i + 2, countCol + 1).setValue(0);
      sheet.getRange(i + 2, resetCol + 1).setValue(today);
      resetCount++;
    }
  }

  if (resetCount > 0) {
    Logger.log("QuotaManager: reset daily quotas for " + resetCount + " account(s) on " + today + ".");
  }
  return resetCount;
}

/**
 * Locates the sheet row (1-indexed) for a given account name, case-insensitive.
 * Returns -1 if not found.
 */
function findAccountRow_(sheet, cols, accountName) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  var accountCol = cols["Account"];
  if (accountCol === undefined) return -1;

  var wanted = accountName.toString().trim().toLowerCase();
  var names = sheet.getRange(2, accountCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0].toString().trim().toLowerCase() === wanted) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Returns the remaining send allowance for an account today: Daily Quota - Sent Today Count.
 * Ensures the daily reset has run first so a stale counter never blocks sending.
 *
 * @param {string} account The account name, e.g. "Account A".
 * @return {number} Remaining quota (>= 0). Returns 0 if the account is unknown.
 */
function getRemainingQuota(account) {
  if (!account) return 0;
  resetDailyQuotasIfNeeded();

  var sheet = getAccountConfigSheet_();
  var cols = accountConfigColMap_(sheet);
  var row = findAccountRow_(sheet, cols, account);
  if (row === -1) {
    Logger.log("QuotaManager.getRemainingQuota: account '" + account + "' not found in Account Config.");
    return 0;
  }

  var quotaCol = cols["Daily Quota"];
  var countCol = cols["Sent Today Count"];
  if (quotaCol === undefined || countCol === undefined) return 0;

  var quota = parseInt(sheet.getRange(row, quotaCol + 1).getValue()) || 0;
  var sent = parseInt(sheet.getRange(row, countCol + 1).getValue()) || 0;
  var remaining = quota - sent;
  return remaining > 0 ? remaining : 0;
}

/**
 * Records a single successful send against an account by incrementing its
 * Sent Today Count. Runs the daily reset first so the increment lands on a
 * fresh counter at the start of a new day.
 *
 * @param {string} account The account name, e.g. "Account A".
 * @return {number} The new Sent Today Count, or -1 if the account was not found.
 */
function recordSend(account) {
  if (!account) return -1;
  resetDailyQuotasIfNeeded();

  var sheet = getAccountConfigSheet_();
  var cols = accountConfigColMap_(sheet);
  var row = findAccountRow_(sheet, cols, account);

  var countCol = cols["Sent Today Count"];
  var resetCol = cols["Last Reset Date"];
  if (countCol === undefined) {
    Logger.log("QuotaManager.recordSend: Account Config missing 'Sent Today Count' column.");
    return -1;
  }

  // Self-heal: if the account row is missing, append it so the send is still tracked.
  if (row === -1) {
    var newRow = [];
    ACCOUNT_CONFIG_HEADERS.forEach(function(h) {
      if (h === "Account") newRow.push(account);
      else if (h === "Daily Quota") newRow.push(0);
      else if (h === "Sent Today Count") newRow.push(0);
      else if (h === "Last Reset Date") newRow.push(quotaTodayString_());
      else newRow.push("");
    });
    sheet.appendRow(newRow);
    row = sheet.getLastRow();
    Logger.log("QuotaManager.recordSend: added missing account '" + account + "' to Account Config.");
  }

  var current = parseInt(sheet.getRange(row, countCol + 1).getValue()) || 0;
  var updated = current + 1;
  sheet.getRange(row, countCol + 1).setValue(updated);
  if (resetCol !== undefined) {
    sheet.getRange(row, resetCol + 1).setValue(quotaTodayString_());
  }
  return updated;
}

/**
 * Startup guard for the daily run. Scans the Account Config sheet and reports any account
 * whose Daily Quota is still sitting at the untouched BOOTSTRAP_DEFAULT_QUOTA placeholder —
 * i.e. a mailbox that was auto-seeded but never given a real, warmup-appropriate quota.
 *
 * prepareDailyQueue() should call this first and refuse to run when any offender is found,
 * so a placeholder quota can never silently drive real sends.
 *
 * CAVEAT: this cannot distinguish an untouched placeholder from a mailbox a human
 * deliberately set to BOOTSTRAP_DEFAULT_QUOTA. If 40 is a legitimate manual value for a
 * mailbox, nudge it (e.g. set then reset) or change BOOTSTRAP_DEFAULT_QUOTA to a value no
 * real mailbox would use.
 *
 * @return {object} { ok: boolean, offenders: string[] } — ok=true means every account has
 *                   a quota that differs from the bootstrap default.
 */
function validateAccountConfigBeforeRun() {
  var sheet = getAccountConfigSheet_();
  var cols = accountConfigColMap_(sheet);
  var accountCol = cols["Account"];
  var quotaCol = cols["Daily Quota"];
  var lastRow = sheet.getLastRow();
  var offenders = [];

  if (accountCol === undefined || quotaCol === undefined) {
    Logger.log("validateAccountConfigBeforeRun: Account Config missing required columns; allowing run.");
    return { ok: true, offenders: offenders };
  }
  if (lastRow <= 1) {
    Logger.log("validateAccountConfigBeforeRun: no accounts configured yet; allowing run.");
    return { ok: true, offenders: offenders };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    var account = data[i][accountCol].toString().trim();
    if (!account) continue;
    var quota = parseInt(data[i][quotaCol]);
    if (!isNaN(quota) && quota === BOOTSTRAP_DEFAULT_QUOTA) {
      offenders.push(account);
    }
  }

  var ok = offenders.length === 0;
  if (!ok) {
    Logger.log("validateAccountConfigBeforeRun: REFUSING to run — account(s) still at the bootstrap " +
               "default quota (" + BOOTSTRAP_DEFAULT_QUOTA + "): " + offenders.join(", ") +
               ". Set a real per-mailbox Daily Quota in the '" + ACCOUNT_CONFIG_SHEET_NAME + "' sheet first.");
  }
  return { ok: ok, offenders: offenders };
}
