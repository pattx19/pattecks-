/**
 * PattEcks! pre-order receiver (Google Apps Script).
 * Saves each order as a row in the "Orders" sheet and the payment slip as an image in Google Drive.
 * Setup steps: see README.md in this folder.
 */
const SHEET_NAME = 'Orders';
const SLIP_FOLDER_NAME = 'PattEcks slips';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const slipUrl = saveSlip_(data.orderId, data.slip);
    const items = (data.items || []).map(function (i) { return i.product + ' ' + i.size + ' x' + i.qty; }).join('\n');
    sheet.appendRow([
      new Date(), data.orderId, data.name, "'" + (data.phone || ''), data.email, data.instagram,
      data.address, items, data.itemsTotal, data.shipping, data.total, slipUrl, data.note, data.lang, 'New'
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, orderId: data.orderId })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Received', 'Order no.', 'Name', 'Phone', 'Email', 'Instagram', 'Address', 'Items', 'Items total', 'Shipping', 'Total', 'Slip', 'Note', 'Language', 'Status']);
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
