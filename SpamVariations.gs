/**
 * Reformats the existing Templates sheet from a vertical layout to a horizontal layout,
 * grouping templates by Type and Account.
 */
function reformatTemplatesHorizontally() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Templates");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Templates sheet not found.");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("Not enough data to format.");
    return;
  }
  
  // Group by (Type, Account)
  var map = {};
  
  for (var i = 1; i < data.length; i++) {
    var type = data[i][0] ? data[i][0].toString().trim() : "";
    
    // In old schema, Account was usually Column D (index 3) but maybe it's already Column B (index 1)?
    // Let's check both just to be safe. If we are reading a vertically-arranged sheet where Account is in col 4:
    var account = data[i][3] ? data[i][3].toString().trim() : (data[i][1] ? data[i][1].toString().trim() : "");
    // Wait, if it's already horizontal, Column 1 is Account. If it's old vertical, Column 3 is Account, and Column 1 is Subject.
    // Let's assume if Column 3 is empty, we check Column 1. Actually, for old vertical, Account is Column 3, Subject is Col 1, Body is Col 2.
    
    var subject, body;
    if (data[i][3] !== undefined && data[i][3] !== "") {
      // Classic vertical layout: Type, Subject, Body, Account
      subject = data[i][1] ? data[i][1].toString().trim() : "";
      body = data[i][2] ? data[i][2].toString().trim() : "";
      account = data[i][3].toString().trim();
    } else {
      // It might already be formatted somewhat or missing account
      subject = data[i][1] ? data[i][1].toString().trim() : "";
      body = data[i][2] ? data[i][2].toString().trim() : "";
      account = "";
    }
    
    // Safety check: if subject is an account name (like Account A), they might have manually changed columns
    if (subject.indexOf("Account") > -1 && body === "") {
      account = subject;
      subject = data[i][2] ? data[i][2].toString().trim() : "";
      body = data[i][3] ? data[i][3].toString().trim() : "";
    }
    
    if (!type && !subject && !body) continue;
    
    var key = type + "|||" + account;
    if (!map[key]) {
      map[key] = { type: type, account: account, variations: [] };
    }
    if (subject || body) {
      map[key].variations.push({ subject: subject, body: body });
    }
  }
  
  // Clear the sheet
  sheet.clear();
  
  // Build new data
  var newData = [];
  var maxVariations = 0;
  
  var keys = Object.keys(map);
  for (var k = 0; k < keys.length; k++) {
    var group = map[keys[k]];
    if (group.variations.length > maxVariations) {
      maxVariations = group.variations.length;
    }
  }
  
  var header = ["Template Type", "Preferred Account"];
  for (var v = 1; v <= maxVariations; v++) {
    header.push("Subject " + v);
    header.push("Body " + v);
  }
  newData.push(header);
  
  for (var k = 0; k < keys.length; k++) {
    var group = map[keys[k]];
    var row = [group.type, group.account];
    for (var v = 0; v < group.variations.length; v++) {
      row.push(group.variations[v].subject);
      row.push(group.variations[v].body);
    }
    // Pad the rest of the row with empty strings
    while (row.length < header.length) {
      row.push("");
    }
    newData.push(row);
  }
  
  sheet.getRange(1, 1, newData.length, header.length).setValues(newData);
  sheet.getRange(1, 1, newData.length, header.length).setWrap(true).setVerticalAlignment("top");
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#f3f3f3");
  
  // Set nice column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  for (var c = 3; c <= header.length; c++) {
    sheet.setColumnWidth(c, c % 2 !== 0 ? 250 : 400); // Subjects narrower, Bodies wider
  }
  
  SpreadsheetApp.getUi().alert("Templates have been arranged horizontally sideways!");
}

/**
 * Appends 16 new highly-distinct Follow-up templates (4 for FU1, 4 for FU2 across Account A and Account B)
 * horizontally to the bottom of the Templates sheet.
 */
function addFollowupVariations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templatesSheet = ss.getSheetByName("Templates");
  
  if (!templatesSheet) {
    SpreadsheetApp.getUi().alert("Templates sheet not found.");
    return;
  }
  
  var variations = [
    // Account A (Ayush) - Follow-up 1
    [
      "Follow-up 1 Template", "Account A",
      "Re: Scaling {Company}'s team", "Hi {First Name},\n\nFollowing up on my previous note. We know finding the right leadership is often the hardest part of scaling.\n\n{AI_INSIGHT}\n\nIf expanding your team is a priority this quarter, I’d love to briefly share how we help platforms like yours secure top executives quickly.\n\nLet me know if you have 10 minutes next week.",
      "Re: Executive hiring for {Company}", "Hi {First Name},\n\nI'm floating this to the top of your inbox since I know how easily things get buried.\n\n{AI_INSIGHT}\n\nWe specialize in taking the heavy lifting out of executive search so founders can focus on execution.\n\nAre you open to a quick introductory chat?",
      "Re: Leadership talent for {Company}", "Hi {First Name},\n\nJust wanted to check if you had a moment to read my last email.\n\n{AI_INSIGHT}\n\nOur team at Butter Search partners exclusively with fast-growing companies to solve their most critical leadership hiring challenges.\n\nWould it make sense to connect briefly this week?",
      "Re: Hiring support for {Company}", "Hi {First Name},\n\nI know you're busy, so I’ll keep this short.\n\n{AI_INSIGHT}\n\nWhen you're ready to scale your senior team, finding talent that can immediately impact growth is essential. We help you do exactly that without the usual friction.\n\nHappy to connect if you’d like to explore this further."
    ],
    // Account A (Ayush) - Follow-up 2
    [
      "Follow-up 2 Template", "Account A",
      "Quick question about {Company}'s roadmap", "Hi {First Name},\n\nI wanted to try one last angle before crossing you off my list.\n\nButter Search has successfully partnered with companies like {Company} to drastically cut down time-to-hire for critical CXO roles, ensuring you never compromise on quality.\n\nIf leadership hiring is on the horizon, I'd love to chat. Otherwise, I'll stop reaching out for now.",
      "Closing the loop on executive hiring", "Hi {First Name},\n\nI haven't heard back, so I assume now might not be the right time to discuss your hiring plans.\n\nJust to leave you with a final thought—our network of elite executives (built by our IIM-C and ex-Naukri alumni founders) is highly curated and ready to deploy into high-growth environments like {Company}.\n\nIf things change and you need support building out your team, please keep us in mind.",
      "Final note regarding leadership talent", "Hi {First Name},\n\nI know how crowded inboxes can get, so this will be my last note.\n\nWe focus entirely on helping platforms like {Company} bypass the talent crunch by directly connecting founders with proven leaders who have successfully scaled similar businesses.\n\nIf you ever find yourself bottlenecked by executive hiring, I hope you'll reach out.",
      "Permission to close your file?", "Hi {First Name},\n\nI'm assuming finding senior leadership isn't an immediate priority for {Company} right now, which is completely understandable.\n\nWe've helped similar platforms reduce their hiring cycles by over 40% while securing top-tier talent. If that becomes a focus for you later this year, feel free to reply to this thread.\n\nWishing you the best with your upcoming milestones!"
    ],
    // Account B (Harshith) - Follow-up 1
    [
      "Follow-up 1 Template", "Account B",
      "Re: Building {Company}'s leadership team", "Hi {First Name},\n\nJust circling back on my previous email. I understand how hectic things can get when you're scaling fast.\n\n{AI_INSIGHT}\n\nWe specialize in giving founders direct access to a highly curated network of senior talent and industry leaders.\n\nWould you have 10 minutes next week for a quick intro?",
      "Re: Strategic hiring at {Company}", "Hi {First Name},\n\nI’m bringing this back to your attention in case it slipped through the cracks.\n\n{AI_INSIGHT}\n\nOur entire focus at Butter Search is ensuring your growth isn't bottlenecked by executive hiring challenges.\n\nCould we schedule a brief call to see if there's a fit?",
      "Re: Executive talent for {Company}", "Hi {First Name},\n\nJust following up on my earlier note.\n\n{AI_INSIGHT}\n\nFinding the right leaders to drive execution is crucial. We handle the heavy lifting of executive search so you can stay focused on building the business.\n\nDoes it make sense to connect briefly about your upcoming hiring plans?",
      "Re: Leadership capacity at {Company}", "Hi {First Name},\n\nI know you've got a lot on your plate, so I'll keep this brief.\n\n{AI_INSIGHT}\n\nWe partner with growing companies like yours to quickly identify and secure the best leadership candidates in the market.\n\nAre you open to a short chat to explore a potential partnership?"
    ],
    // Account B (Harshith) - Follow-up 2
    [
      "Follow-up 2 Template", "Account B",
      "Quick question about {Company}'s hiring plans", "Hi {First Name},\n\nI’ll make this my last follow-up so I don’t clog your inbox.\n\nWe’ve helped high-growth companies like {Company} streamline their executive search process, bringing in elite leaders who can immediately impact scale.\n\nIf leadership hiring becomes a priority later on, I hope you’ll consider Butter Search.",
      "Final thought on senior talent", "Hi {First Name},\n\nSince I haven't heard back, I'm guessing that expanding your senior team isn't a top priority right now.\n\nI just wanted to reiterate that our team (ex-Naukri, PwC, A&M) is completely dedicated to solving the executive hiring bottlenecks that founders often face during rapid expansion.\n\nIf you ever need strategic hiring support, feel free to get in touch.",
      "Closing the loop on leadership recruitment", "Hi {First Name},\n\nI completely understand that timing is everything, so this will be my last email.\n\nOur goal is simply to help platforms like {Company} connect with top-tier executives without the usual friction and delays of traditional search firms.\n\nWishing you continued success, and please keep us in mind for future hiring needs!",
      "One last note for {Company}", "Hi {First Name},\n\nI'm assuming you're fully staffed on the leadership front for now, which is great to hear.\n\nWe take the pain out of executive search for ambitious companies. If you ever find yourself needing experienced leaders to drive execution in the future, don't hesitate to reach out.\n\nBest of luck with your upcoming initiatives!"
    ]
  ];
  
  // Append all variations
  for (var i = 0; i < variations.length; i++) {
    templatesSheet.appendRow(variations[i]);
  }
  
  // Format the new rows
  var lastRow = templatesSheet.getLastRow();
  var addedRange = templatesSheet.getRange(lastRow - variations.length + 1, 1, variations.length, variations[0].length);
  addedRange.setWrap(true);
  addedRange.setVerticalAlignment("top");
  
  SpreadsheetApp.getUi().alert("Success! Added 16 new highly-varied Follow-up templates to the Templates sheet.");
}

