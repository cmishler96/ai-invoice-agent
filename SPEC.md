# Demo 01 — Invoice Agent

## Pitch
"Vendors email invoices; the agent reads each PDF, logs it to your books sheet, files it
with a clean name, and only bothers you about the ones that look wrong."

## Demo story (what the viewer sees)
1. Inbox has ~10 vendor emails with PDF invoices, labeled `Invoices/Inbox`.
2. Agent runs. Each PDF is read by Claude (sent directly as a PDF document block).
3. `Invoice Log` sheet fills with vendor, number, dates, amounts, GL category, confidence.
4. PDFs land in Drive as `Invoices/<YYYY-MM>/<Vendor>/2026-11-03_Acme-Supply_INV-1042_$1,284.50.pdf`.
5. Two invoices are flagged in `Needs Review` (a duplicate, and a total that doesn't match line items).
6. A digest email summarizes: processed count, total $, flagged items with reasons.

## Architecture
- Trigger: time-driven, every 15 min → `processInvoiceInbox()`
- Source: Gmail threads with label `Invoices/Inbox`; after processing, swap label to `Invoices/Processed`
  (or `Invoices/Error`).
- AI: Claude Messages API via `UrlFetchApp`; PDF passed as base64 `document` content block.
- Storage: one Google Sheet (tabs below) + Drive folder tree.

## Sheet: `Invoice Agent Demo`
**Invoice Log** — Processed At · Vendor · Invoice # · Invoice Date · Due Date · Subtotal · Tax ·
Total · Currency · GL Category · Confidence · Status (Logged / Needs Review) · Flags · Drive Link ·
Gmail Message ID

**Needs Review** — same columns + Reason · Resolved (checkbox)

**Chart of Accounts** — Category · Description (seeded: Office Supplies, Software & SaaS, Facilities
& Maintenance, Professional Services, Food & Catering, Equipment, Utilities, Travel, Other)

**Settings** — key/value: Review threshold ($), Confidence threshold, Digest recipient, Drive root
folder ID

**AI Log** — Timestamp · Item ID · Model · Input tokens · Output tokens · Raw JSON

## Extraction schema (Claude must return exactly this JSON)
```json
{
  "vendor": "string",
  "invoice_number": "string",
  "invoice_date": "YYYY-MM-DD",
  "due_date": "YYYY-MM-DD or null",
  "currency": "USD",
  "subtotal": 0.0,
  "tax": 0.0,
  "total": 0.0,
  "line_items": [{"description": "string", "qty": 0, "unit_price": 0.0, "amount": 0.0}],
  "gl_category": "one of the Chart of Accounts categories",
  "confidence": 0.0,
  "notes": "anything unusual"
}
```

## Exception rules (computed in code, not by the AI)
- Duplicate: same vendor + invoice number already in Invoice Log
- Math mismatch: |sum(line_items) + tax − total| > $0.02
- Low confidence: confidence < Settings threshold (default 0.8)
- Large amount: total > Settings review threshold (default $5,000)
- Missing fields: vendor, invoice_number, or total is empty
- Not an invoice: model returns `"vendor": null` → label `Invoices/Error`

## Acceptance criteria
- [ ] `setup()` creates sheet tabs, labels, Drive root folder, and the 15-min trigger
- [ ] `seedSampleData()` generates 10 fake invoice PDFs (incl. 1 duplicate, 1 math error) and emails them to the owner with the label applied
- [ ] Running `processInvoiceInbox()` twice never double-logs
- [ ] Every flagged invoice shows a human-readable reason
- [ ] Digest email sends once per run only if something was processed
- [ ] Handles API errors with retry (429/5xx, exponential backoff, max 3) and labels failures `Invoices/Error`
