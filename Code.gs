/**
 * ระบบเบิกวัสดุบรรจุภัณฑ์ - Google Apps Script backend
 *
 * วิธีใช้:
 * 1. สร้าง Google Sheet ใหม่ (หรือใช้ที่มีอยู่)
 * 2. เมนู Extensions > Apps Script แล้ววางโค้ดนี้ทับไฟล์ Code.gs
 * 3. กด Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone (หรือ Anyone within [organization] ถ้าต้องการจำกัดเฉพาะในองค์กร)
 * 4. คัดลอก URL ที่ได้ (ลงท้ายด้วย /exec) ไปวางใน index.html ตัวแปร SCRIPT_URL
 *
 * ข้อมูลจะถูกเก็บเป็นแถวในชีตชื่อ "Requests" ภายใน Google Sheet เดียวกับที่ผูก Script นี้ไว้
 * (ไฟล์ Sheet จะอยู่ใน Google Drive ของบัญชีที่ deploy Web App นี้)
 */

const SHEET_NAME = 'Requests';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'ชื่อผู้เบิก', 'แผนก', 'วันที่ต้องการใช้', 'รายการ (JSON)', 'สรุปรายการ', 'หมายเหตุ']);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const name = (data.name || '').toString().trim();
    const department = (data.department || '').toString().trim();
    const dateNeeded = (data.dateNeeded || '').toString().trim();
    const items = Array.isArray(data.items) ? data.items : [];
    const note = (data.note || '').toString().trim();

    if (!name || !department || !dateNeeded || items.length === 0) {
      return jsonResponse_({ ok: false, error: 'missing required fields' });
    }

    const itemsSummary = items.map(i => `${i.name} x${i.qty}`).join(', ');

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      name,
      department,
      dateNeeded,
      JSON.stringify(items),
      itemsSummary,
      note
    ]);

    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'list') {
      const limit = parseInt(e.parameter.limit, 10) || 10;
      const sheet = getSheet_();
      const values = sheet.getDataRange().getValues();
      const header = values.shift(); // remove header row
      const rows = values.slice(-limit).map(row => ({
        timestamp: row[0],
        name: row[1],
        department: row[2],
        dateNeeded: row[3],
        itemsSummary: row[5],
        note: row[6]
      }));
      return jsonResponse_({ ok: true, rows: rows });
    }
    return jsonResponse_({ ok: true, message: 'ระบบเบิกวัสดุบรรจุภัณฑ์ API ทำงานปกติ' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
