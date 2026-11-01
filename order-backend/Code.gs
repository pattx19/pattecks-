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
  { title: 'ภาษา', width: 60 }
];

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
    return ContentService.createTextOutput(JSON.stringify({ ok: true, orderId: data.orderId })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Adds a "PattEcks" menu to the sheet, so you can re-tidy it any time.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PattEcks').addItem('จัดหน้าชีตใหม่', 'formatSheet').addToUi();
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
