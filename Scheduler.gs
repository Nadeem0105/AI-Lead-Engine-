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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var leadsSheet = ss.getSheetByName("Leads");
    if (!leadsSheet) return;
    
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
      storeQueueForToday([]);
      return;
    }
    
    // 4. Validate and assign priority (Feature 1 - handled by ZeroBounceValidator.gs)
    if (typeof validateWithZeroBounce === "function") {
      for (var i = 0; i < allLeads.length; i++) {
        validateWithZeroBounce(allLeads[i]);
        assignSendPriority(allLeads[i]);
      }
    }
    
    // 5. Build prioritized queue and save it
    var prioritizedQueue = allLeads;
    if (typeof buildPrioritizedQueue === "function") {
      prioritizedQueue = buildPrioritizedQueue(allLeads);
    }
    
    storeQueueForToday(prioritizedQueue);
    Logger.log("Prepared daily queue with " + prioritizedQueue.length + " leads.");
    
  } catch (e) {
    Logger.log("Error in prepareDailyQueue: " + e.toString());
  }
}

/**
 * Triggered periodically between 11:00 AM and 1:00 PM.
 * Sends a proportional slice of the daily queue.
 */
function processSendWindow() {
  try {
    var queue = getTodaysQueue();
    if (!queue || queue.length === 0) {
      Logger.log("Send window: queue is empty.");
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
          if (typeof recordSend === "function") recordSend(account);
        }
      }
    }
    
    storeQueueForToday(queue);
    Logger.log("Processed send window batch. Remaining in queue: " + queue.length);
    
  } catch (e) {
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
  
  var data = leadsSheet.getRange(2, 1, lastRow - 1, leadsSheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][statusCol - 1] === targetStatus) {
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
