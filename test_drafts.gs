function debugListDrafts() {
  var ui = SpreadsheetApp.getUi();
  var config = safeGetConfig_();
  var accAEmail = getConfigValue(config, "Account A Email", "").toString().trim();
  var accBEmail = getConfigValue(config, "Account B Email", "").toString().trim();
  
  var result = "Draft Check Results:\n\n";
  
  [accAEmail, accBEmail].forEach(function(email) {
    if (!email) return;
    var service = getGmailService_(email);
    if (!service.hasAccess()) {
      result += email + ": NOT AUTHORIZED\n";
      return;
    }
    
    var url = "https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=10";
    var options = {
      method: "get",
      headers: { Authorization: "Bearer " + service.getAccessToken() },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      if (data.drafts && data.drafts.length > 0) {
        result += email + ": Found " + data.drafts.length + " recent drafts.\n";
        data.drafts.forEach(function(d) {
          result += "  - Draft ID: " + d.id + "\n";
        });
      } else {
        result += email + ": 0 drafts found.\n";
      }
    } else {
      result += email + ": Error fetching drafts - " + response.getContentText() + "\n";
    }
  });
  
  ui.alert("Drafts Check", result, ui.ButtonSet.OK);
}
