# Inbox Scanner Loop Implementation Task

You are an expert Google Apps Script developer. Please implement the following changes in the `f:\Lead_Generation` repository:

1. **Implement the Inbox Scanner Loop in `ResponseClassifier.gs`**
   - Create a new function called `processInboundResponses()`.
   - This function should read the `Leads` sheet and iterate over all rows where `Thread Id` is present but `Response Status` is empty.
   - For each active thread, fetch it via `GmailApp.getThreadById(threadId)`.
   - Iterate through the messages in the thread (or just look at `messages.length > 1`). Find any message sent AFTER the initial outreach. 
   - **CRITICAL FILTERING:** You must ignore any messages sent by the user's own accounts (e.g. automated follow-ups sent by the FollowupEngine). Only process messages that are from the prospect, or from a mailer-daemon/postmaster (bounces). 
   - You can get the configured sender accounts from `Account Config` or just check if the message sender `getFrom()` does NOT match the `Sent From Account` for that lead (which is stored in the sheet).
   - If a valid inbound message is found, call `classifyReply(message, { row: rowIndex, email: leadEmail })`.

2. **Update the UI Menu in `Main.gs`**
   - In `Main.gs` inside the `onOpen()` function, find the `Email Ops (Daily Workflow)` sub-menu.
   - Insert a new button: `.addItem("3. Scan Inbox for Responses", "processInboundResponses")`.
   - Make sure to update the number of the follow-ups button to `4. Process Follow-ups`.

Make sure there are no logical or technical errors. 
Do not break any existing code.
