# Multi-Mailbox Architecture Refactor

We need to integrate the concepts from `multi-mailbox-gmail-api.gs` into our main Lead Generation project to achieve true multi-account sending and scanning, WITHOUT using Domain-Wide Delegation.

Please perform the following refactoring carefully. Make sure there are no errors and that it aligns with our existing architecture (specifically how we read from the `Config` sheet).

## 1. Setup & OAuth
- Create a new file `GmailAPI.gs` and move the OAuth2 logic (`getService`, `authCallback`, `startAuth`, `checkAuthStatus`) from `multi-mailbox-gmail-api.gs` into it.
- Update `appsscript.json` to include the official Apps Script OAuth2 library: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkPQ9UPbw`.
- In `GmailAPI.gs`, do NOT hardcode `MAILBOX_CONFIG`. Instead, dynamically get the list of active sender emails from the `Config` sheet (e.g. `Account A Email`, `Account B Email`).

## 2. Sending & Drafting (`Outreach.gs`)
- In `processLeadsToDrafts`, replace `GmailApp.createDraft` and `draft.send()` with the `sendFromMailbox` logic. Note that the API uses `https://gmail.googleapis.com/gmail/v1/users/me/messages/send` for sending and `.../drafts` for drafts. Implement a `createDraft` function in `GmailAPI.gs`.

## 3. Follow-ups (`FollowupEngine.gs`)
- Replace the `GmailApp.getThreadById()` and `lastMsg.reply()` logic with `sendFollowupInThread` from the new `GmailAPI.gs`.
- Ensure that `sendFollowup` still supports drafting (by calling a new `draftFollowupInThread` via API if Outreach Mode is Draft).

## 4. Reply Detection (`ResponseClassifier.gs`)
- Update `scanInboxForReplies` to loop through ALL configured sender accounts (using their respective OAuth tokens via `GmailAPI`).
- Use the `checkForReply(mailboxKey, threadId)` logic to determine if a thread has a reply, rather than relying on `GmailApp.search()`.

## 5. Bounce Detection
- Add the `checkForBounce` logic to `ResponseClassifier.gs` or a new file, and schedule it to run daily.

Please execute these changes now using your file editing tools. Ensure all braces and syntax are correct.
