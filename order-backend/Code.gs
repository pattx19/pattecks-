/**
 * PattEcks! pre-order receiver (Google Apps Script).
 * Saves each order as a row in the "Orders" sheet and the payment slip as an image in Google Drive,
 * and emails you for every new order.
 * Setup steps: see README.md in this folder.
 */
const SHEET_NAME = 'Orders';
const SLIP_FOLDER_NAME = 'PattEcks slips';
const STATUSES = ['ใหม่', 'จ่ายแล้ว', 'กำลังผลิต', 'ส่งแล้ว', 'ยกเลิก'];

// Column layout (one row = one order).
const COLS = [
  { title: 'รับเมื่อ', width: 130 },
  { title: 'เลขออเดอร์', width: 140 },
  { title: 'สถานะ', width: 110 },
  { title: 'ชื่อ', width: 160 },
  { title: 'เบอร์โทร', width: 115 },
  { title: 'Instagram', width: 120 },
  { title: 'ที่อยู่จัดส่ง', width: 260 },
  { title: 'รายการ', width: 220 },
  { title: 'ค่าสินค้า', width: 90 },
  { title: 'ค่าส่ง', width: 80 },
  { title: 'ยอดรวม', width: 95 },
  { title: 'สลิป', width: 80 },
  { title: 'หมายเหตุ / เลขพัสดุ', width: 200 },
  { title: 'อีเมล', width: 170 },
  { title: 'ภาษา', width: 60 },
  { title: 'ส่งใบเสร็จแล้ว', width: 130 }
];
const COL = { received: 1, orderId: 2, status: 3, name: 4, phone: 5, ig: 6, address: 7, items: 8,
  itemsTotal: 9, shipping: 10, total: 11, slip: 12, note: 13, email: 14, lang: 15, receiptSent: 16 };
const PAID_STATUS = 'จ่ายแล้ว';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const slipUrl = saveSlip_(data.orderId, data.slip);
    const items = (data.items || []).map(function (i) { return i.product + ' · ' + i.size + ' × ' + i.qty; }).join('\n');
    sheet.appendRow([
      new Date(), data.orderId, STATUSES[0], data.name, "'" + (data.phone || ''), data.instagram || '',
      data.address, items, data.itemsTotal, data.shipping, data.total,
      slipUrl ? '=HYPERLINK("' + slipUrl + '","ดูสลิป")' : '', data.note || '', data.email || '', data.lang || ''
    ]);
    notify_(data, items, slipUrl);
    if (data.email) sendCustomerEmail_('confirm', {
      orderId: data.orderId, name: data.name, email: data.email, lang: data.lang, items: items,
      address: data.address, itemsTotal: data.itemsTotal, shipping: data.shipping, total: data.total, date: new Date()
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true, orderId: data.orderId })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Adds a "PattEcks" menu to the sheet, so you can re-tidy it any time.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PattEcks')
    .addItem('จัดหน้าชีตใหม่', 'formatSheet')
    .addItem('เปิดระบบส่งใบเสร็จอัตโนมัติ (ทำครั้งเดียว)', 'setupReceiptTrigger')
    .addItem('ส่งใบเสร็จของแถวที่เลือกอีกครั้ง', 'resendReceipt')
    .addToUi();
}

/* ---------- Customer emails ---------- */

// Run once (PattEcks menu). After that, changing a row's status to "จ่ายแล้ว" emails that customer a receipt.
function setupReceiptTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onStatusEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  const sheet = getSheet_();
  sheet.getRange(1, COL.receiptSent).setValue(COLS[COL.receiptSent - 1].title)
    .setBackground('#1a1b1d').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setColumnWidth(COL.receiptSent, COLS[COL.receiptSent - 1].width);
  ss.toast('เปิดแล้ว: เปลี่ยนสถานะเป็น "จ่ายแล้ว" ระบบจะส่งใบเสร็จให้ลูกค้าเอง', 'PattEcks', 6);
}

function onStatusEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME || e.range.getColumn() !== COL.status || e.range.getRow() < 2) return;
  if (e.range.getNumRows() !== 1 || String(e.value) !== PAID_STATUS) return;
  sendReceiptForRow_(sheet, e.range.getRow(), false);
}

function resendReceipt() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (sheet.getName() !== SHEET_NAME || row < 2) {
    SpreadsheetApp.getUi().alert('เลือกแถวออเดอร์ในแท็บ Orders ก่อน');
    return;
  }
  const ok = sendReceiptForRow_(sheet, row, true);
  SpreadsheetApp.getActiveSpreadsheet().toast(ok ? 'ส่งใบเสร็จแล้ว' : 'ส่งไม่ได้: ออเดอร์นี้ไม่มีอีเมล', 'PattEcks', 5);
}

function sendReceiptForRow_(sheet, row, force) {
  const v = sheet.getRange(row, 1, 1, COLS.length).getValues()[0];
  const email = String(v[COL.email - 1] || '').trim();
  if (!email) return false;
  if (v[COL.receiptSent - 1] && !force) return false;   // already sent once
  sendCustomerEmail_('receipt', {
    orderId: v[COL.orderId - 1], name: v[COL.name - 1], email: email, lang: v[COL.lang - 1],
    items: String(v[COL.items - 1] || ''), address: v[COL.address - 1],
    itemsTotal: v[COL.itemsTotal - 1], shipping: v[COL.shipping - 1], total: v[COL.total - 1],
    date: v[COL.received - 1] instanceof Date ? v[COL.received - 1] : new Date(), paidOn: new Date()
  });
  sheet.getRange(row, COL.receiptSent).setValue(new Date()).setNumberFormat('d mmm yyyy, HH:mm');
  return true;
}

// kind: 'confirm' (order received, slip being checked) or 'receipt' (payment confirmed).
function sendCustomerEmail_(kind, o) {
  try {
    const th = String(o.lang || '').toLowerCase() !== 'en';
    const T = th ? {
      confirm: 'ได้รับออเดอร์แล้ว', receipt: 'ใบเสร็จรับเงิน',
      hiConfirm: 'ขอบคุณที่สั่งจอง เราได้รับออเดอร์และสลิปของคุณแล้ว กำลังตรวจสอบยอดโอน เมื่อยืนยันแล้วจะส่งใบเสร็จตามมาอีกฉบับ',
      hiReceipt: 'ยืนยันการชำระเงินเรียบร้อย ขอบคุณที่สนับสนุน PattEcks! สินค้าเป็นแบบพรีออเดอร์ เราจะแจ้งเลขพัสดุเมื่อจัดส่ง',
      hello: 'สวัสดีคุณ', order: 'เลขออเดอร์', date: 'วันที่สั่ง', paid: 'วันที่ชำระ', items: 'รายการ',
      sub: 'ค่าสินค้า', ship: 'ค่าส่ง', total: 'ยอดรวม', addr: 'ที่อยู่จัดส่ง', tbc: 'แจ้งภายหลัง',
      foot: 'มีคำถามตอบกลับอีเมลนี้ได้เลย', status: 'สถานะ', stConfirm: 'รอตรวจสลิป', stReceipt: 'ชำระแล้ว'
    } : {
      confirm: 'Order received', receipt: 'Receipt',
      hiConfirm: "Thanks for your pre-order. We've got your order and payment slip and are checking the transfer. You'll get a receipt once it's confirmed.",
      hiReceipt: "Payment confirmed. Thank you for supporting PattEcks! This is a pre-order; we'll send your tracking number when it ships.",
      hello: 'Hi', order: 'Order no.', date: 'Ordered', paid: 'Paid', items: 'Items',
      sub: 'Subtotal', ship: 'Shipping', total: 'Total', addr: 'Ship to', tbc: 'TBC',
      foot: 'Questions? Just reply to this email.', status: 'Status', stConfirm: 'Checking slip', stReceipt: 'Paid'
    };
    const isReceipt = kind === 'receipt';
    const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    const fmtD = function (d) { return Utilities.formatDate(d, tz, 'd MMM yyyy, HH:mm'); };
    const baht = function (n) { return (n === '' || n == null || isNaN(Number(n))) ? T.tbc : '฿' + Number(n).toLocaleString('en-US'); };
    const esc = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
    const lines = String(o.items || '').split('\n').filter(String);
    const row = function (k, val, bold) {
      return '<tr><td style="padding:6px 0;color:#6b7075">' + k + '</td><td style="padding:6px 0;text-align:right;' + (bold ? 'font-weight:700;font-size:18px;color:#1a1b1d' : 'color:#1a1b1d') + '">' + val + '</td></tr>';
    };
    const html =
      '<div style="background:#f4efe6;padding:32px 12px;font-family:Helvetica,Arial,sans-serif">' +
      '<div style="max-width:520px;margin:0 auto;background:#ffffff;border:1px solid #d9d4c7">' +
      '<div style="background:#1a1b1d;color:#ffffff;padding:22px 28px;font-size:22px;font-weight:700;letter-spacing:1px">PattEcks!' +
      '<span style="float:right;font-size:13px;font-weight:400;letter-spacing:' + (th ? '0' : '2px') + ';color:#d9d4c7;padding-top:6px">' + (isReceipt ? T.receipt : T.confirm).toUpperCase() + '</span></div>' +
      '<div style="height:4px;background:#890000"></div>' +
      '<div style="padding:28px">' +
      '<p style="margin:0 0 6px;font-size:16px;color:#1a1b1d">' + T.hello + ' ' + esc(o.name) + '</p>' +
      '<p style="margin:0 0 22px;font-size:14px;line-height:1.6;color:#45484c">' + (isReceipt ? T.hiReceipt : T.hiConfirm) + '</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
      row(T.order, '<b>' + esc(o.orderId) + '</b>') + row(T.date, fmtD(o.date)) +
      (isReceipt ? row(T.paid, fmtD(o.paidOn)) : '') +
      row(T.status, '<span style="color:' + (isReceipt ? '#2e7d32' : '#890000') + ';font-weight:700">' + (isReceipt ? T.stReceipt : T.stConfirm) + '</span>') +
      '</table>' +
      '<div style="border-top:1px solid #e6e1d6;margin:18px 0 10px"></div>' +
      '<div style="font-size:12px;letter-spacing:' + (th ? '0' : '1.5px') + ';color:#6b7075;margin-bottom:6px">' + T.items.toUpperCase() + '</div>' +
      lines.map(function (l) { return '<div style="padding:5px 0;font-size:15px;color:#1a1b1d">' + esc(l) + '</div>'; }).join('') +
      '<div style="border-top:1px solid #e6e1d6;margin:12px 0 6px"></div>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
      row(T.sub, baht(o.itemsTotal)) + row(T.ship, baht(o.shipping)) + row(T.total, baht(o.total), true) + '</table>' +
      '<div style="border-top:1px solid #e6e1d6;margin:16px 0 10px"></div>' +
      '<div style="font-size:12px;letter-spacing:' + (th ? '0' : '1.5px') + ';color:#6b7075;margin-bottom:6px">' + T.addr.toUpperCase() + '</div>' +
      '<div style="font-size:14px;line-height:1.6;color:#1a1b1d;white-space:pre-line">' + esc(o.address) + '</div>' +
      '</div>' +
      '<div style="padding:16px 28px;background:#f4efe6;font-size:12px;color:#6b7075">' + T.foot + '<br><a href="https://www.pattecks.com" style="color:#6b7075">www.pattecks.com</a></div>' +
      '</div></div>';
    const text = [(isReceipt ? T.receipt : T.confirm) + ' — PattEcks!', T.order + ': ' + o.orderId, '', lines.join('\n'), '',
      T.total + ': ' + baht(o.total), T.addr + ': ' + o.address].join('\n');
    // GmailApp sends from your own Gmail (shows in Sent), which spam filters trust more than MailApp.
    const subject = th
      ? (isReceipt ? 'ใบเสร็จออเดอร์ ' : 'ยืนยันออเดอร์ ') + o.orderId + ' จาก PattEcks'
      : (isReceipt ? 'Your receipt for order ' : 'Your PattEcks order ') + o.orderId;
    GmailApp.sendEmail(o.email, subject, text, { htmlBody: html, name: 'PattEcks' });
  } catch (err) {}
}

// Run this once (or from the PattEcks menu) to make the sheet easy to read.
function formatSheet() {
  const sheet = getSheet_();
  const n = COLS.length;
  const maxRows = Math.max(sheet.getMaxRows(), 500);
  if (sheet.getMaxRows() < maxRows) sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());

  // Header: dark band, white bold text, frozen.
  const head = sheet.getRange(1, 1, 1, n);
  head.setValues([COLS.map(function (c) { return c.title; })])
    .setBackground('#1a1b1d').setFontColor('#ffffff').setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  COLS.forEach(function (c, i) { sheet.setColumnWidth(i + 1, c.width); });

  // Body: readable font, top-aligned, wrap long text.
  const body = sheet.getRange(2, 1, maxRows - 1, n);
  body.setFontSize(11).setVerticalAlignment('top').setFontColor('#1a1b1d');
  sheet.getRange(2, 7, maxRows - 1, 2).setWrap(true);   // address, items
  sheet.getRange(2, 13, maxRows - 1, 1).setWrap(true);  // note
  sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat('d mmm yyyy, HH:mm');
  sheet.getRange(2, 5, maxRows - 1, 1).setNumberFormat('@');
  sheet.getRange(2, 9, maxRows - 1, 3).setNumberFormat('฿#,##0;-฿#,##0;"—"');
  sheet.getRange(2, 11, maxRows - 1, 1).setFontWeight('bold');
  sheet.getRange(2, 2, maxRows - 1, 2).setHorizontalAlignment('center');
  sheet.getRange(2, 12, maxRows - 1, 1).setHorizontalAlignment('center');

  // Status: a dropdown with one colour per stage.
  const status = sheet.getRange(2, 3, maxRows - 1, 1);
  status.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  status.setFontWeight('bold');
  const rows = sheet.getRange(2, 1, maxRows - 1, n);
  const colour = { 'ใหม่': '#fff1b8', 'จ่ายแล้ว': '#d4efd4', 'กำลังผลิต': '#d9e6ff', 'ส่งแล้ว': '#ececec', 'ยกเลิก': '#f6d2d2' };
  const rules = STATUSES.map(function (s) {
    const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + s + '"').setBackground(colour[s]).setRanges([rows]);
    if (s === 'ส่งแล้ว') b.setFontColor('#6b7075');
    if (s === 'ยกเลิก') b.setFontColor('#890000').setStrikethrough(true);
    return b.build();
  });
  sheet.setConditionalFormatRules(rules);

  // Thin lines between orders.
  rows.setBorder(null, null, true, null, null, true, '#d9d4c7', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1, 1, maxRows, n).setFontFamily('Sarabun');
}

// Emails you (the Google account that owns this script) for every new order.
function notify_(data, items, slipUrl) {
  try {
    const to = Session.getEffectiveUser().getEmail();
    const body = [
      'ออเดอร์ใหม่ ' + data.orderId,
      '',
      'ชื่อ: ' + data.name,
      'เบอร์: ' + data.phone,
      'Instagram: ' + (data.instagram || '-'),
      'ที่อยู่: ' + data.address,
      '',
      'รายการ:', items,
      'ยอดรวม: ' + (data.total == null ? '(ยังไม่ได้ตั้งราคา)' : data.total + ' บาท'),
      '',
      'สลิป: ' + (slipUrl || '-'),
      'ชีต: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    ].join('\n');
    MailApp.sendEmail(to, 'PattEcks! ออเดอร์ใหม่ ' + data.orderId, body);
  } catch (e) {}
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLS.map(function (c) { return c.title; }));
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveSlip_(orderId, dataUrl) {
  if (!dataUrl) return '';
  const m = String(dataUrl).match(/^data:(image\/[a-z+]+);base64,(.+)$/);
  if (!m) return '';
  const folders = DriveApp.getFoldersByName(SLIP_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(SLIP_FOLDER_NAME);
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], orderId + '.jpg');
  return folder.createFile(blob).getUrl();
}
