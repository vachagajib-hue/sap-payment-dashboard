// ===================================================
// Google Apps Script — SAP Payment Dashboard Proxy
// บริษัท รถเจาะไทย จำกัด
// ===================================================
// วิธีใช้:
// 1. เปิด Google Sheets → Extensions → Apps Script
// 2. วางโค้ดนี้ทับโค้ดเดิมทั้งหมด
// 3. กด Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy Web App URL ไปใส่ใน Dashboard HTML

const SHEET_NAMES = {
  invoice:  'ค้างจ่าย-Invoice',
  outgoing: 'ค้างจ่าย-Outgoing',
  ar:       'ค้างรับ'
};

function doGet(e) {
  try {
    const type = e.parameter.type || 'invoice';
    const sheetName = SHEET_NAMES[type];

    if (!sheetName) {
      return makeResponse({ error: 'invalid type: ' + type }, 400);
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return makeResponse({ error: 'ไม่พบ Sheet: ' + sheetName }, 404);
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows    = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // ข้ามแถวที่ว่างทั้งหมด
      if (row.every(cell => cell === '' || cell === null)) continue;
      const obj = {};
      headers.forEach((h, j) => {
        const val = row[j];
        // แปลง Date object เป็น string dd/MM/yyyy
        if (val instanceof Date) {
          obj[h] = Utilities.formatDate(val, 'Asia/Bangkok', 'dd/MM/yyyy');
        } else {
          obj[h] = val !== null && val !== undefined ? String(val) : '';
        }
      });
      rows.push(obj);
    }

    return makeResponse({ type: type, count: rows.length, data: rows });

  } catch (err) {
    return makeResponse({ error: err.message }, 500);
  }
}

function makeResponse(payload, code) {
  const output = ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
