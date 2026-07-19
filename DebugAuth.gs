function debugDumpDraftRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName("Leads");
  var headersMap = getHeadersMap(leadsSheet);
  var lastRow = leadsSheet.getLastRow();
  
  var result = "Draft Rows Dump:\n\n";
  var found = 0;
  for (var r = 2; r <= lastRow; r++) {
    var status = leadsSheet.getRange(r, headersMap["Outreach Status"]).getValue().toString().trim();
    if (status.indexOf("Draft") !== -1) {
      found++;
      var threadId = headersMap["Thread Id"] ? leadsSheet.getRange(r, headersMap["Thread Id"]).getValue().toString().trim() : "NO COL";
      var account = headersMap["Sent From Account"] ? leadsSheet.getRange(r, headersMap["Sent From Account"]).getValue().toString().trim() : "NO COL";
      var email = leadsSheet.getRange(r, headersMap["Email"]).getValue().toString().trim();
      
      result += "Row " + r + " | " + email + " | " + status + "\n";
      result += "  - Account: " + account + "\n";
      result += "  - Thread Id: " + threadId + "\n\n";
    }
  }
  
  if (found === 0) {
    result += "No rows with 'Draft' in Outreach Status found.";
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.alert("Dump Results", result, ui.ButtonSet.OK);
}
