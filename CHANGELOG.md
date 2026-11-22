# Changelog — Demo 01 Invoice Agent

## 2026-11-20 — Initial release

### Features
- `setup()` creates (or reopens) the `Invoice Agent Demo` sheet with all five tabs. It seeds the Chart of Accounts
  and Settings, creates the three Gmail labels and the `Invoices` Drive root folder (its ID is saved in Settings),
  and replaces the 15-min trigger. You can run it again safely.
- Each thread's label moves to `Invoices/Processed`, or to `Invoices/Error` if any attachment failed, the
  document wasn't an invoice, or the email had no PDF.
- Idempotency: each attachment's key (`<messageId>/<filename>`) is stored in the Invoice Log's
  "Gmail Message ID" column. Already-logged attachments are skipped, so retrying a half-failed thread never
  double-logs.
- Duplicate detection matches on vendor + invoice # (case-insensitive, trimmed) against the Invoice Log. The reason
  names when the original was logged.
- PDFs are filed to `Invoices/<YYYY-MM>/<Vendor>/2026-11-03_Vendor-Name_INV-1042_$1,284.50.pdf`.
- Flagged invoices go to both Invoice Log and Needs Review, with a human-readable Reason and a Resolved checkbox.
- Digest: an HTML email, sent once per run and only if something was processed. It shows the counts, total $, a
  table of flagged invoices with reasons and Drive links, the auto-logged invoices, and errors.
- Settings and Chart of Accounts are read from the sheet, with defaults as fallback.

### Claude API
- Shared Claude helper: API version `2026-10-21`, structured outputs (`gl_category` is an enum
  built from the Chart of Accounts), retries with backoff, and an AI Log row for every call.
- `claude-sonnet-5-5` at `effort: "medium"`.
- **Server-side refusal fallback** is on (`fallbacks: "default"`). If a safety classifier ever declines a
  document, the API retries on a fallback model. Turn it off with `CONFIG.USE_FALLBACKS = false`.
- The prompt tells Claude to copy numbers exactly as printed and never fix arithmetic. Without that,
  the math-mismatch invoice could be silently "corrected" and never flagged.

### Sample data
- Invoice dates are relative to the seed day (10–18 days ago), with due dates from Net terms.
- The duplicate is sent as "Resending: Invoice BOS-20418…", so Gmail keeps it in its own thread instead
  of merging it into the original. This is also how vendors actually resend.
- Seeded threads are found after sending, labeled `Invoices/Inbox`, moved to the inbox, and marked
  unread. `fixSeedLabels()` re-runs that step if Gmail indexing lags.
- The invoice PDFs look more realistic: vendor address, bill-to block, styled line-item table, right-aligned
  amounts, and Net terms.

### Added
- `retryErrors()` moves `Invoices/Error` threads back to `Invoices/Inbox`.
- `resetDemo()` trashes the seeded emails and the filed PDFs, and clears the log tabs. Trash is recoverable for
  30 days.
- Time guard: a run stops after 4.5 min, and the rest is picked up on the next run.
- clasp: `.clasp.json.example` and `.claspignore`.

### Removed
- `onOpen()` menu. This is a standalone script, not bound to the sheet, so the menu could never appear.
  Run functions from the Apps Script editor instead. The `script.container.ui` scope was removed with it.

### Scopes (`appsscript.json`)
- `gmail.modify` and `gmail.send` were replaced with `https://mail.google.com/`, which GmailApp requires.
- Added `userinfo.email`, which `Session.getEffectiveUser().getEmail()` requires.

### Design decisions
- **Digest recipient**: by design the agent never emails anyone but the owner, so the digest
  always goes to the owner. If Settings names someone else, the script logs a warning instead of sending to them.
