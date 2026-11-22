/**
 * Fake invoice generator for demo runs. All vendors, addresses, and data are fictional.
 * Builds each invoice as a Google Doc → PDF, emails it to the owner, and applies Invoices/Inbox.
 * Dates are relative to the day you seed, so the demo never looks stale.
 */

/** dayOffset: invoice date relative to today. terms: days until due. */
function sampleInvoices_(lastMonthName) {
  return [
    { vendor: 'Brightline Office Supply', num: 'BOS-20418', dayOffset: -17, terms: 30,
      address: '2210 Commerce Way, Salem, OR 97301',
      items: [['Copy paper, 10 reams', 4, 52.99], ['Toner cartridge HL-L2350', 2, 68.50], ['Dry-erase markers (12pk)', 6, 11.25]] },
    { vendor: 'Cascade Facilities Group', num: 'CFG-7731', dayOffset: -16, terms: 30,
      address: '88 Industrial Pkwy, Tigard, OR 97223',
      items: [['HVAC quarterly service', 1, 840.00], ['Filter replacement', 12, 18.00]] },
    { vendor: 'Northwind Catering Co.', num: 'NWC-1188', dayOffset: -15, terms: 14,
      address: '415 Alder St, Portland, OR 97204',
      items: [['Staff lunch, 45 people', 45, 14.50], ['Delivery fee', 1, 35.00]] },
    { vendor: 'Pixelwise Software', num: 'PW-2026-311', dayOffset: -18, terms: 14,
      address: '1 Market Plaza, Suite 900, San Francisco, CA 94105',
      items: [['Annual license, 25 seats', 25, 96.00]] },
    { vendor: 'Ridgeback IT Services', num: 'RIT-0453', dayOffset: -14, terms: 30,
      address: '730 NW 14th Ave, Portland, OR 97209',
      items: [['Network troubleshooting (hrs)', 6, 125.00], ['Ethernet switch 24-port', 1, 389.00]] },
    { vendor: 'Evergreen Utilities', num: 'EU-88213', dayOffset: -13, terms: 20,
      address: 'PO Box 4400, Portland, OR 97208',
      items: [[`Electric service, ${lastMonthName}`, 1, 1412.37]] },
    { vendor: 'Summit Furniture Outlet', num: 'SFO-5502', dayOffset: -12, terms: 30,
      address: '5600 SE Foster Rd, Portland, OR 97206', taxRate: 0,
      items: [['Student desks', 30, 129.00], ['Teacher chair', 4, 189.00], ['Freight', 1, 250.00]] },
    // Exception: duplicate of the Brightline invoice, re-sent by the vendor.
    { vendor: 'Brightline Office Supply', num: 'BOS-20418', dayOffset: -17, terms: 30, resend: true,
      address: '2210 Commerce Way, Salem, OR 97301',
      items: [['Copy paper, 10 reams', 4, 52.99], ['Toner cartridge HL-L2350', 2, 68.50], ['Dry-erase markers (12pk)', 6, 11.25]] },
    // Exception: the printed total doesn't match the line items ($480.00 vs $520.00).
    { vendor: 'Harbor Print & Sign', num: 'HPS-0902', dayOffset: -11, terms: 14, overrideTotal: 520.00,
      address: '19 Waterfront Dr, Astoria, OR 97103',
      items: [['Vinyl banner 3x8 ft', 2, 145.00], ['Yard signs', 20, 9.50]] },
    { vendor: 'Lumen Professional Services', num: 'LPS-114', dayOffset: -10, terms: 30,
      address: '1200 SW Main St, Suite 410, Portland, OR 97205',
      items: [[`Bookkeeping review, ${lastMonthName}`, 1, 650.00]] },
  ];
}

function seedSampleData() {
  const owner = Session.getEffectiveUser().getEmail();
  const label = getOrCreateLabel_(CONFIG.LABEL_INBOX);
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const lastMonthName = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 1, 1), tz, 'MMMM yyyy');
  const invoices = sampleInvoices_(lastMonthName);

  invoices.forEach(inv => {
    const invDate = addDays_(now, inv.dayOffset);
    inv.date = Utilities.formatDate(invDate, tz, 'MMMM d, yyyy');
    inv.due = Utilities.formatDate(addDays_(invDate, inv.terms), tz, 'MMMM d, yyyy');

    const pdf = buildInvoicePdf_(inv);
    // A different subject keeps the re-sent duplicate in its own Gmail thread, as a real resend would be.
    const subject = (inv.resend ? 'Resending: ' : '') + `Invoice ${inv.num} from ${inv.vendor}`;
    const body = inv.resend
      ? `Hello,\n\nWe haven't seen payment on invoice ${inv.num} yet, so we're re-sending it. ` +
        `Please disregard if already paid.\n\nThank you,\n${inv.vendor} Accounts Receivable`
      : `Hello,\n\nPlease find attached invoice ${inv.num}. Payment is due ${inv.due}.\n\n` +
        `Thank you,\n${inv.vendor} Accounts Receivable`;
    GmailApp.sendEmail(owner, subject, body + '\n\n[demo-seed]', { attachments: [pdf], name: inv.vendor });
    Utilities.sleep(1500);
  });

  // Self-sent mail may land only in Sent (or arrive read). Find each seeded thread and file it like real vendor mail.
  Utilities.sleep(4000);
  const labeled = labelSeededThreads_(label);
  console.log(`Seeded ${invoices.length} invoice emails; ${labeled} labeled ${CONFIG.LABEL_INBOX}.` +
    (labeled < invoices.length ? ' Some may still be indexing — run fixSeedLabels() in a minute.' : ''));
}

/** Re-run if seeding reported fewer labeled threads than it sent (Gmail search can lag a few seconds). */
function fixSeedLabels() {
  const n = labelSeededThreads_(getOrCreateLabel_(CONFIG.LABEL_INBOX));
  console.log(`${n} demo invoice thread(s) labeled ${CONFIG.LABEL_INBOX}.`);
}

function labelSeededThreads_(label) {
  const done = getOrCreateLabel_(CONFIG.LABEL_DONE);
  const error = getOrCreateLabel_(CONFIG.LABEL_ERROR);
  const threads = GmailApp.search('"demo-seed" has:attachment newer_than:1d', 0, 50).filter(t => {
    const names = t.getLabels().map(l => l.getName());
    return names.indexOf(done.getName()) === -1 && names.indexOf(error.getName()) === -1;
  });
  threads.forEach(t => {
    t.addLabel(label);
    t.moveToInbox();
    t.markUnread();
  });
  return threads.length;
}

/** Builds an invoice Doc and returns it as a PDF blob. Trashes the temp Doc. */
function buildInvoicePdf_(inv) {
  const doc = DocumentApp.create(`TMP ${inv.vendor} ${inv.num}`);
  const body = doc.getBody();
  body.setMarginTop(48).setMarginBottom(48);

  body.appendParagraph(inv.vendor).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(inv.address).setForegroundColor('#5f6368');
  body.appendParagraph('INVOICE').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const meta = body.appendTable([
    ['Invoice #', inv.num, 'Bill to', 'Demo Organization'],
    ['Invoice date', inv.date, '', '100 Example St'],
    ['Due date', inv.due, '', 'Portland, OR 97201'],
  ]);
  meta.setBorderWidth(0);
  [0, 2].forEach(col => { for (let r = 0; r < 3; r++) meta.getCell(r, col).editAsText().setBold(true); });

  body.appendParagraph('');
  const rows = [['Description', 'Qty', 'Unit price', 'Amount']];
  let subtotal = 0;
  inv.items.forEach(([desc, qty, price]) => {
    const amt = Math.round(qty * price * 100) / 100;
    subtotal += amt;
    rows.push([desc, String(qty), price.toFixed(2), amt.toFixed(2)]);
  });
  const table = body.appendTable(rows);
  table.setBorderColor('#dadce0');
  for (let c = 0; c < 4; c++) {
    table.getCell(0, c).setBackgroundColor('#f1f3f4').editAsText().setBold(true);
  }
  for (let r = 0; r < rows.length; r++) {
    for (let c = 1; c < 4; c++) {
      table.getCell(r, c).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  }

  const tax = Math.round(subtotal * (inv.taxRate || 0) * 100) / 100;
  const total = inv.overrideTotal != null ? inv.overrideTotal : subtotal + tax;
  [`Subtotal: $${subtotal.toFixed(2)}`, `Tax: $${tax.toFixed(2)}`].forEach(t =>
    body.appendParagraph(t).setAlignment(DocumentApp.HorizontalAlignment.RIGHT));
  body.appendParagraph(`TOTAL DUE: $${total.toFixed(2)}`)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).editAsText().setBold(true).setFontSize(14);

  body.appendParagraph(`Terms: Net ${inv.terms}. Please include the invoice number with your payment.`)
    .setForegroundColor('#5f6368').setFontSize(9);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getAs('application/pdf').setName(`${inv.num}.pdf`);
  file.setTrashed(true);
  return pdf;
}

function addDays_(base, n) {
  return new Date(base.getTime() + n * 24 * 60 * 60 * 1000);
}
