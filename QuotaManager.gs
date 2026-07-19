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
  
  // Deduct follow-ups from the daily quota before allowing new emails.
  var pendingFollowups = 0;
  if (typeof getPendingFollowupsCount === 'function') {
    pendingFollowups = getPendingFollowupsCount(account);
  }
  
  var remaining = quota - sent - pendingFollowups;
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

/**
 * Calculates the pending follow-ups due today and updates the Daily Quota Forecast tab.
 * Schema: Account | Account Daily Limit | Fresh Mails Sent Today | FU1 Due Today |
 *         FU2 Due Today | Total Emails Due | Remaining Fresh Slots | Last Updated
 */
function updateDailyForecast() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var forecastSheet = ss.getSheetByName("Daily Quota Forecast");
    if (!forecastSheet) {
      if (typeof setupDailyForecastSheet_ === "function") setupDailyForecastSheet_(ss);
      forecastSheet = ss.getSheetByName("Daily Quota Forecast");
      if (!forecastSheet) { Logger.log("updateDailyForecast: cannot create sheet."); return; }
    }
    var leadsSheet = ss.getSheetByName("Leads");
    if (!leadsSheet) return;
    var headersMap = getHeadersMap(leadsSheet);
    resetDailyQuotasIfNeeded();
    var configSheet = getAccountConfigSheet_();
    var configCols = accountConfigColMap_(configSheet);
    var lastConfigRow = configSheet.getLastRow();
    if (lastConfigRow <= 1) return;
    var configData = configSheet.getRange(2, 1, lastConfigRow - 1, configSheet.getLastColumn()).getValues();
    var accountsData = {};
    for (var i = 0; i < configData.length; i++) {
      var acc = configData[i][configCols["Account"]].toString().trim();
      if (!acc) continue;
      accountsData[acc] = {
        dailyLimit: parseInt(configData[i][configCols["Daily Quota"]]) || 0,
        sentToday: parseInt(configData[i][configCols["Sent Today Count"]]) || 0,
        fu1Due: 0,
        fu2Due: 0
      };
    }
  
    var lastRow = leadsSheet.getLastRow();
    if (lastRow > 1) {
      var data = leadsSheet.getRange(2, 1, lastRow - 1, leadsSheet.getLastColumn()).getValues();
      var now = new Date();
      for (var r = 0; r < data.length; r++) {
        var rowData = data[r];
        var replied = headersMap["Replied"] ? rowData[headersMap["Replied"] - 1].toString().trim() : "";
        if (replied.toLowerCase() === "yes") continue;
        var pipelineStage = headersMap["Pipeline Stage"] ? rowData[headersMap["Pipeline Stage"] - 1].toString().trim() : "";
        var outreachStatus = headersMap["Outreach Status"] ? rowData[headersMap["Outreach Status"] - 1].toString().trim() : "";
        var wasSent = (pipelineStage === "Sent" || pipelineStage === "Follow-up Sent" || outreachStatus === "Email Sent");
        if (!wasSent) continue;
        var cancelled = headersMap["Followup Cancelled"] ? rowData[headersMap["Followup Cancelled"] - 1] : false;
        if (cancelled === true || cancelled.toString().toLowerCase() === "true") continue;
        var account = "";
        if (headersMap["Sent From Account"]) account = rowData[headersMap["Sent From Account"] - 1].toString().trim();
        if (!account && headersMap["Send From Account"]) account = rowData[headersMap["Send From Account"] - 1].toString().trim();
        if (!account) account = getConfigValue(getConfig(), "Default Send Account", "Account A").toString().trim();
        if (!accountsData[account]) continue;
        var f1Due = toDateOrNull_(headersMap["Followup 1 Due Date"] ? rowData[headersMap["Followup 1 Due Date"] - 1] : null);
        var f1Sent = toDateOrNull_(headersMap["Followup 1 Sent Date"] ? rowData[headersMap["Followup 1 Sent Date"] - 1] : null);
        var f2Due = toDateOrNull_(headersMap["Followup 2 Due Date"] ? rowData[headersMap["Followup 2 Due Date"] - 1] : null);
        var f2Sent = toDateOrNull_(headersMap["Followup 2 Sent Date"] ? rowData[headersMap["Followup 2 Sent Date"] - 1] : null);
        if (!f1Sent && f1Due && f1Due.getTime() <= now.getTime()) {
          accountsData[account].fu1Due++;
        } else if (f1Sent && !f2Sent && f2Due && f2Due.getTime() <= now.getTime()) {
          accountsData[account].fu2Due++;
        }
      }
    }
  
    // Clear old data rows (keep header)
    var forecastLastRow = forecastSheet.getLastRow();
    if (forecastLastRow > 1) {
      forecastSheet.getRange(2, 1, forecastLastRow - 1, 8).clearContent();
      forecastSheet.getRange(2, 1, forecastLastRow - 1, 8).setBackground("#f0f4f8");
    }
    var writeData = [];
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    for (var accKey in accountsData) {
      var info = accountsData[accKey];
      var totalDue = info.fu1Due + info.fu2Due;
      var remainingFreshSlots = info.dailyLimit - info.sentToday - totalDue;
      if (remainingFreshSlots < 0) remainingFreshSlots = 0;
      writeData.push([accKey, info.dailyLimit, info.sentToday, info.fu1Due, info.fu2Due, totalDue, remainingFreshSlots, timestamp]);
    }
    if (writeData.length > 0) {
      forecastSheet.getRange(2, 1, writeData.length, 8).setValues(writeData);
      for (var w = 0; w < writeData.length; w++) {
        var rem = writeData[w][6];
        var bg = rem > 10 ? "#c8e6c9" : (rem > 0 ? "#fff9c4" : "#ffcdd2");
        forecastSheet.getRange(w + 2, 7).setBackground(bg);
      }
    }
    Logger.log("updateDailyForecast completed successfully.");
    SpreadsheetApp.getUi().alert("Forecast Updated ✅",
      "Daily Quota Forecast refreshed. Check the 'Daily Quota Forecast' tab.",
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("Error in updateDailyForecast: " + e.toString());
  }
}

// ─────────────────────────────────────────────────────────────
// HOURLY RATE LIMITER
// ─────────────────────────────────────────────────────────────

/**
 * Checks if an account has hit its hourly sending limit for a given type.
 * Automatically resets counters if the 1-hour window has expired.
 * @param {string} account  e.g. "Account A"
 * @param {string} type     "fresh" | "fu1" | "fu2"
 * @return {{ok: boolean, message: string}}
 */
function checkHourlyLimit(account, type) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Hourly Rate Limits");
    if (!sheet) return { ok: true, message: "" };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().trim()] = i; });

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { ok: true, message: "" };

    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][colMap["Account"]].toString().trim().toLowerCase() !== account.toLowerCase()) continue;

      var rowNum = i + 2;
      var now = new Date();
      var maxColName = type === "fresh" ? "Max Emails / Hour" : (type === "fu1" ? "Max FU1 / Hour" : "Max FU2 / Hour");
      var countColName = type === "fresh" ? "Sent This Hour" : (type === "fu1" ? "FU1 This Hour" : "FU2 This Hour");

      var limit = parseInt(data[i][colMap[maxColName]]) || 0;
      if (limit <= 0) return { ok: true, message: "" };

      var current = 0;

      // Reset window if expired (> 1 hour)
      var windowRaw = data[i][colMap["Hour Window Start"]];
      var windowStart = windowRaw ? new Date(windowRaw) : null;
      if (!windowStart || (now.getTime() - windowStart.getTime()) >= 3600000) {
        
        // Rolling Ledger Logic: Append a new block of all accounts
        var latestMap = {};
        for (var j = data.length - 1; j >= 0; j--) {
           var accName = data[j][colMap["Account"]].toString().trim();
           if (accName && !latestMap[accName]) {
             latestMap[accName] = data[j].slice(); // copy the array
           }
        }
        
        var accKeys = Object.keys(latestMap).sort();
        var newRows = [];
        var windowStartStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        
        for (var k = 0; k < accKeys.length; k++) {
           var rData = latestMap[accKeys[k]];
           rData[colMap["Sent This Hour"]] = 0;
           rData[colMap["FU1 This Hour"]] = 0;
           rData[colMap["FU2 This Hour"]] = 0;
           rData[colMap["Hour Window Start"]] = windowStartStr;
           newRows.push(rData);
        }
        
        var nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1, newRows.length, newRows[0].length).setValues(newRows);
        
        // Draw top border if it's a new day
        if (windowStart && now.getDate() !== windowStart.getDate()) {
          sheet.getRange(nextRow, 1, 1, newRows[0].length).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
        
        Logger.log("Hourly window reset. Appended new block for all accounts.");
        current = 0; // count is fresh for this new window
      } else {
        current = parseInt(data[i][colMap[countColName]]) || 0;
      }

      if (current >= limit) {
        var label = type === "fresh" ? "fresh emails" : (type === "fu1" ? "Follow-up 1" : "Follow-up 2");
        return {
          ok: false,
          message: "⏰ Hourly Limit Reached for " + account + "!\n\n" +
                   "You have sent " + current + " " + label + " this hour (limit: " + limit + ").\n\n" +
                   "The counter resets automatically after 1 hour from the first send.\n" +
                   "Go to 'Hourly Rate Limits' tab to adjust your limits."
        };
      }
      return { ok: true, message: "" };
    }
    return { ok: true, message: "" };
  } catch (e) {
    Logger.log("checkHourlyLimit error: " + e.toString());
    return { ok: true, message: "" };
  }
}

/**
 * Increments the hourly send counter for an account/type. Call AFTER a successful send.
 * @param {string} account e.g. "Account A"
 * @param {string} type    "fresh" | "fu1" | "fu2"
 */
function recordHourlySend(account, type) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Hourly Rate Limits");
    if (!sheet) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().trim()] = i; });

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    // Iterate backwards to find the latest row for this account in the rolling ledger
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][colMap["Account"]].toString().trim().toLowerCase() !== account.toLowerCase()) continue;
      var rowNum = i + 2;
      var countColName = type === "fresh" ? "Sent This Hour" : (type === "fu1" ? "FU1 This Hour" : "FU2 This Hour");
      var countSheetCol = colMap[countColName] + 1;
      var windowSheetCol = colMap["Hour Window Start"] + 1;
      if (!data[i][colMap["Hour Window Start"]]) {
        sheet.getRange(rowNum, windowSheetCol).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
      }
      var current = parseInt(sheet.getRange(rowNum, countSheetCol).getValue()) || 0;
      sheet.getRange(rowNum, countSheetCol).setValue(current + 1);
      Logger.log("Hourly counter: " + account + "/" + type + " = " + (current + 1));
      return;
    }
  } catch (e) {
    Logger.log("recordHourlySend error: " + e.toString());
  }
}

// ─────────────────────────────────────────────────────────────
// DAILY 9 AM TRIGGER MANAGEMENT
// ─────────────────────────────────────────────────────────────

/**
 * Creates a daily time-based trigger that runs updateDailyForecast at 9:00 AM.
 * Safe to call multiple times — checks for existing trigger first.
 */
function setupDailyForecastTrigger() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "updateDailyForecast") {
      ui.alert("Already Active",
        "A daily 9 AM forecast trigger is already running.\n\nTo change the time, remove it first.",
        ui.ButtonSet.OK);
      return;
    }
  }
  ScriptApp.newTrigger("updateDailyForecast").timeBased().everyDays(1).atHour(9).create();
  ui.alert("✅ Daily Trigger Set",
    "The Daily Quota Forecast will auto-refresh every morning at 9:00 AM.",
    ui.ButtonSet.OK);
  Logger.log("Daily forecast trigger set for 9 AM.");
}

/**
 * Removes all daily forecast triggers (updateDailyForecast).
 */
function removeDailyForecastTrigger() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "updateDailyForecast") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  ui.alert(removed > 0 ? "Trigger Removed" : "No Trigger Found",
    removed > 0 ? "Daily Quota Forecast trigger removed." : "No active trigger found.",
    ui.ButtonSet.OK);
  Logger.log("removeDailyForecastTrigger: removed " + removed + " trigger(s).");
}
