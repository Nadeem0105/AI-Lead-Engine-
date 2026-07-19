/**
 * AI Lead Generation Engine - Scheduler
 * File: Scheduler.gs
 */

/**
 * 9:00 AM Trigger function.
 * Auto-scores new leads, resets quotas, validates with ZeroBounce, and builds the daily send queue.
 */
function prepareDailyQueue() {
  try {
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leadsSheet = ss.getSheetByName("Leads");
    if (!leadsSheet) return;

    // 0. Refuse to run if any mailbox is still at the untouched bootstrap default quota
    //    (a placeholder must never silently drive real sends).
    if (typeof validateAccountConfigBeforeRun === "function") {
      var quotaCheck = validateAccountConfigBeforeRun();
      if (!quotaCheck.ok) {
        Logger.log("prepareDailyQueue: aborted — account(s) at bootstrap default quota: " +
                   quotaCheck.offenders.join(", ") + ". Set real per-mailbox quotas first.");
        SpreadsheetApp.getUi().alert("Safety Lock Engaged!\n  \n  Account(s) still at the default quota (40): " + quotaCheck.offenders.join(", ") + ".\n  \n  Please change the Daily Quota to a different number in the 'Account Config' sheet to unlock the engine.");
        return;
      }
    }

    // 1. Reset quotas
    if (typeof resetDailyQuotasIfNeeded === "function") {
      resetDailyQuotasIfNeeded();
    }
    
    // 2. Auto-score unscored leads on the sheet (Feature 7)
    // Run the scoring logic headlessly for any new leads that landed on the sheet
    autoScoreAndValidateLeads(leadsSheet);
    
    // 3. Gather leads for the queue
    var spillover = getLeadsMarkedSkippedYesterday(leadsSheet);
    var freshLeads = getLeadsPendingSend(leadsSheet);
    var allLeads = spillover.concat(freshLeads);
    
    if (allLeads.length === 0) {
      Logger.log("No leads to queue for today.");
      SpreadsheetApp.getUi().alert("Zero Leads Found!\n  \n  We looked for leads with Outreach Status = 'Ready for outreach' but found 0. Please make sure you have typed that exactly into the Outreach Status column for your test lead.");
      storeQueueForToday([]);
      return;
    }
    
    // 3.5 Inject explicit email and score properties for ZeroBounceValidator
    var headersMap = getHeadersMap(leadsSheet);
    var emailIdx = (headersMap["Email"] || 0) - 1;
    var scoreIdx = (headersMap["Score"] || 0) - 1;
    for (var k = 0; k < allLeads.length; k++) {
      if (emailIdx >= 0) allLeads[k].email = allLeads[k].data[emailIdx];
      if (scoreIdx >= 0) allLeads[k].score = allLeads[k].data[scoreIdx];
    }
    
    // 4. Validate and assign priority (Feature 1 - handled by ZeroBounceValidator.gs)
    var startTime = new Date().getTime();
    if (typeof validateWithZeroBounce === "function") {
      for (var i = 0; i < allLeads.length; i++) {
        if (new Date().getTime() - startTime > 300000) {
          Logger.log("Timeout Warning: 5 minutes elapsed. Halting queue validation to prevent crash.");
          break;
        }
        validateWithZeroBounce(allLeads[i]);
        assignSendPriority(allLeads[i]);
      }
    }
    
    // 5. Build prioritized queue and save it
    var prioritizedQueue = allLeads;
    if (typeof buildPrioritizedQueue === "function") {
      prioritizedQueue = buildPrioritizedQueue(allLeads);
    }
    
    // 6. Update statuses for leads and stamp ZeroBounce metrics
    var headersMap = getHeadersMap(leadsSheet);
    
    for (var j = 0; j < allLeads.length; j++) {
      var row = allLeads[j].rowNumber;
      
      if (headersMap["ZB Status"] && allLeads[j].zbStatus !== undefined) {
        leadsSheet.getRange(row, headersMap["ZB Status"]).setValue(allLeads[j].zbStatus);
      }
      if (headersMap["ZB Score"] && allLeads[j].zbScore !== undefined) {
        var scoreRange = leadsSheet.getRange(row, headersMap["ZB Score"]);
        scoreRange.setValue(allLeads[j].zbScore);
        scoreRange.setNumberFormat("0"); // Force number format to prevent '09/01/1900' issue
      }
      if (headersMap["Send Priority"] && allLeads[j].sendPriority) {
        leadsSheet.getRange(row, headersMap["Send Priority"]).setValue(allLeads[j].sendPriority);
      }

      if (allLeads[j].sendPriority === "lowPriority" || allLeads[j].sendPriority === "blocked") {
        if (headersMap["Outreach Status"]) {
          leadsSheet.getRange(row, headersMap["Outreach Status"]).setValue("Low Priority");
        }
        if (headersMap["Pipeline Stage"]) {
          leadsSheet.getRange(row, headersMap["Pipeline Stage"]).setValue("Disqualified");
        }
      }
    }
    
    storeQueueForToday(prioritizedQueue);
    Logger.log("Prepared daily queue with " + prioritizedQueue.length + " leads.");
    
    var quotaWarningMsg = "";
    try {
      var accSheet = getAccountConfigSheet_();
      var accData = accSheet.getRange(2, 1, accSheet.getLastRow() - 1, 1).getValues();
      var exhaustedAccounts = [];
      for (var a = 0; a < accData.length; a++) {
        var accName = accData[a][0].toString().trim();
        if (accName && typeof getRemainingQuota === "function") {
          if (getRemainingQuota(accName) <= 0) {
            exhaustedAccounts.push(accName);
          }
        }
      }
      if (exhaustedAccounts.length > 0) {
        quotaWarningMsg = "\n  \n  WARNING: The daily quota for [" + exhaustedAccounts.join(", ") + "] is currently fully reserved for Follow-ups (or already exhausted). No new emails will be drafted for these accounts today!";
      }
    } catch(err) {
      Logger.log("Could not compute quota warnings: " + err);
    }
    
    SpreadsheetApp.getUi().alert("Success! Processed " + allLeads.length + " leads for the daily queue.\n  \n  The ZeroBounce columns should now be populated." + quotaWarningMsg);
    
  } catch (e) {
    Logger.log("Error in prepareDailyQueue: " + e.toString());
  }
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error", e.toString(), ui.ButtonSet.OK);
    Logger.log("Error in prepareDailyQueue: " + e.toString());
  }
}

/**
 * Triggered periodically between 11:00 AM and 1:00 PM.
 * Sends a proportional slice of the daily queue.
 */
function processSendWindow() {
  try {
  
  try {
    var queue = getTodaysQueue();
    if (!queue || queue.length === 0) {
      Logger.log("Send window: queue is empty.");
      SpreadsheetApp.getUi().alert("Queue is Empty!\n  \n  If you just ran 'Prepare Daily Queue', the Engine actively rejected your lead.\n  \n  Look at the 'ZB Status' column in your sheet:\n  - If it says 'Invalid', it's a fake email.\n  - If it says 'Unknown', the Engine blocked it because of the 20% safety ratio rule (you have no 'Valid' leads to balance it out).");
      return;
    }
    
    var now = new Date();
    var windowEnd = new Date();
    windowEnd.setHours(13, 0, 0, 0); // 1:00 PM
    
    var slotsRemaining = countRemainingSendSlots(now, windowEnd);
    if (slotsRemaining <= 0) slotsRemaining = 1;
    
    var batchSize = Math.ceil(queue.length / slotsRemaining);
    var batch = queue.splice(0, batchSize);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leadsSheet = ss.getSheetByName("Leads");
    var headersMap = getHeadersMap(leadsSheet);
    var config = getConfig();
    var outreachMode = getConfigValue(config, "Outreach Mode", "Draft");
    var testRecipient = getConfigValue(config, "Test Email Recipient", "").toString().trim();
    
    for (var i = 0; i < batch.length; i++) {
      var lead = batch[i];
      var r = lead.rowNumber;
      
      // Select account based on score (or existing routing logic)
      var scoreStr = leadsSheet.getRange(r, headersMap["Score"]).getValue().toString().trim();
      var account = null;
      if (typeof selectAccountByScore === "function") {
        account = selectAccountByScore(scoreStr, config);
      }
      if (!account || account === "HOLD") {
        account = getConfigValue(config, "Default Send Account", "Account A").toString().trim();
      }
      
      if (typeof getRemainingQuota === "function" && getRemainingQuota(account) <= 0) {
        markSkippedForTomorrow(leadsSheet, r, headersMap);
        continue;
      }
      
      // Send the email using the existing Outreach engine
      if (typeof processSingleOutreach === "function") {
        var result = processSingleOutreach(leadsSheet, r, headersMap, config, outreachMode, testRecipient, account);
        if (result && result.success) {
          processed++;
          if (typeof recordSend === "function") recordSend(account);
        }
      }
    }
    
    storeQueueForToday(queue);
    Logger.log("Processed send window batch. Remaining in queue: " + queue.length);
    SpreadsheetApp.getUi().alert("Success!\n  \n  Created " + processed + " drafts in Gmail. \n  (Check your Gmail Drafts folder now)");
    
  } catch (e) {
    Logger.log("Error in processSendWindow: " + e.toString());
  }
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error", e.toString(), ui.ButtonSet.OK);
    Logger.log("Error in processSendWindow: " + e.toString());
  }
}

// --- Helper Functions ---

function autoScoreAndValidateLeads(leadsSheet) {
  // Minimal headless version of runScoringOnlyMenu and runFullPipelineMenu
  var headersMap = getHeadersMap(leadsSheet);
  var mapping = resolveColumnMapping(leadsSheet, true);
  var config = getConfig();
  var lastRow = leadsSheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  for (var r = 2; r <= lastRow; r++) {
    var pipelineStage = leadsSheet.getRange(r, headersMap["Pipeline Stage"]).getValue().toString().trim();
    if (pipelineStage.indexOf("Disqualified") !== -1 || pipelineStage === "Sent" || pipelineStage === "Replied") {
      continue;
    }
    
    var scoreStr = leadsSheet.getRange(r, headersMap["Score"]).getValue().toString().trim();
    if (scoreStr === "" && pipelineStage !== "Scored") {
       // Headless score
       var scoreResult = scoreSingleLead(leadsSheet, r, headersMap, config, mapping);
       if (scoreResult && scoreResult.success) {
         leadsSheet.getRange(r, headersMap["Score"]).setValue(scoreResult.score);
         leadsSheet.getRange(r, headersMap["Pipeline Stage"]).setValue("Scored");
       }
    }
  }
}

function getLeadsMarkedSkippedYesterday(leadsSheet) {
  return queryLeadsByStatus(leadsSheet, "skipped_quota_exhausted");
}

function getLeadsPendingSend(leadsSheet) {
  return queryLeadsByStatus(leadsSheet, "Ready for outreach");
}

function queryLeadsByStatus(leadsSheet, targetStatus) {
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  var leads = [];
  if (lastRow <= 1) return leads;
  
  var statusCol = headersMap["Outreach Status"];
  if (!statusCol) return leads;
  
  var target = (targetStatus || "").toString().trim().toLowerCase();
  
  var data = leadsSheet.getRange(2, 1, lastRow - 1, leadsSheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    var cellValue = (data[i][statusCol - 1] || "").toString().trim().toLowerCase();
    if (cellValue === target) {
      leads.push({
        rowNumber: i + 2,
        data: data[i]
      });
    }
  }
  return leads;
}

function storeQueueForToday(queue) {
  var props = PropertiesService.getScriptProperties();
  if (!queue || queue.length === 0) {
    props.deleteProperty("SEND_QUEUE_TODAY");
  } else {
    // Storing as JSON. Only store rowNumbers to save space.
    var simplified = queue.map(function(l) { return { rowNumber: l.rowNumber, score: l.zbScore, priority: l.sendPriority }; });
    props.setProperty("SEND_QUEUE_TODAY", JSON.stringify(simplified));
  }
}

function getTodaysQueue() {
  var props = PropertiesService.getScriptProperties();
  var str = props.getProperty("SEND_QUEUE_TODAY");
  if (!str) return [];
  return JSON.parse(str);
}

function countRemainingSendSlots(now, windowEnd) {
  // Assume a trigger runs every 15 minutes
  var diffMs = windowEnd.getTime() - now.getTime();
  if (diffMs <= 0) return 0;
  return Math.ceil(diffMs / (15 * 60 * 1000));
}

function markSkippedForTomorrow(leadsSheet, rowNumber, headersMap) {
  if (headersMap["Outreach Status"]) {
    leadsSheet.getRange(rowNumber, headersMap["Outreach Status"]).setValue("skipped_quota_exhausted");
  }
}
