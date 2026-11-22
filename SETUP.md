# Demo 01 setup: what you do by hand

Create and push the project with clasp (see the Quick start in README.md), then:

## 1. Add the API key
In the Apps Script editor, open **Project Settings (gear) → Script Properties → Add script property**.
Name it `ANTHROPIC_API_KEY`, paste your key, and save.

## 2. Authorize and set up
In the editor, choose **setup** in the function dropdown and click **Run**. Approve the consent screen with
**Advanced → Go to Invoice Agent Demo (unsafe) → Allow**. This demo also asks for Google Drive and Google Docs,
for filing PDFs and building the fake invoices. The log prints the sheet and Drive folder links.

## 3. Demo run
1. `seedSampleData()`: about a minute for 10 invoice emails with PDFs. If it reports fewer than 10 labeled,
   run `fixSeedLabels()`.
2. Record the "before": Gmail's `Invoices/Inbox` label with 10 vendor emails, and an empty Drive folder.
3. `processInvoiceInbox()`: run it by hand rather than waiting for the trigger. It takes 1–2 minutes, because Claude
   reads each PDF.
4. Record the "after":
   - The Invoice Log is filled in.
   - Drive has `Invoices/<month>/<vendor>/` with cleanly named PDFs.
   - The **Needs Review** tab shows exactly two invoices: the Brightline duplicate, and the Harbor Print & Sign
     invoice whose total is $520 while the line items sum to $480.
   - The digest email has arrived.
5. `resetDemo()` gets you ready to record again.

## Notes
- The agent only ever emails you. The digest goes to the account owner even if Settings names someone else.
- Every Claude call is in the **AI Log** tab, with tokens and raw JSON.
- Pushing code changes later: run `clasp push` from this folder.
