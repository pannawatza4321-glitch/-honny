// Google Apps Script: วางใน Extensions > Apps Script ของ Google Sheet
// ตั้งค่า Project Settings > Script properties: KEY = รหัสลับสำหรับหน้า fulfill.html
const SHEET = 'Orders';
const HEADERS = ['timestamp','orderId','name','phone','address','province','postcode','sku','product','qty','price','status'];
const PRICES = {'HNY-LEM-250':159, 'HNY-LGN-250':189}; // ตรวจราคาฝั่งเซิร์ฟเวอร์

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const sh = getSheet_(), now = new Date();
  (d.lines || []).forEach(l => {
    const price = PRICES[l.sku], qty = parseInt(l.qty, 10);
    if (!price || !(qty > 0 && qty < 100)) return;
    sh.appendRow([now, d.orderId, d.name, "'" + d.phone, d.address, d.province, "'" + d.postcode, l.sku, l.name, qty, price, 'new']);
  });
  return ContentService.createTextOutput('ok');
}

function doGet(e) {
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!key || e.parameter.key !== key) return json_({error: 'unauthorized'});
  const v = getSheet_().getDataRange().getValues();
  const head = v.shift();
  return json_(v.map(r => Object.fromEntries(head.map((h, i) => [h, String(r[i])]))));
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
