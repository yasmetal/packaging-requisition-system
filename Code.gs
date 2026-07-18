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
    if (action === 'summary') {
      const months = parseInt(e.parameter.months, 10) || 6;
      return jsonResponse_({ ok: true, summary: getMonthlySummary_(months) });
    }
    return jsonResponse_({ ok: true, message: 'ระบบเบิกวัสดุบรรจุภัณฑ์ API ทำงานปกติ' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

/**
 * รวมยอดจำนวนการเบิกแต่ละรายการ แยกตามเดือน (ย้อนหลัง N เดือน รวมเดือนปัจจุบัน)
 * คืนค่า { months: ["2026-02", ...], itemNames: [...], series: {itemName: [qty,...]}, totals: {...}, totalRequests, thisMonthRequests }
 */
function getMonthlySummary_(monthsBack) {
  const tz = Session.getScriptTimeZone();
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  values.shift(); // remove header

  // build ordered list of month keys, oldest -> newest, ending at current month
  const now = new Date();
  const monthKeys = [];
  for (let i = monthsBack - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    monthKeys.push(Utilities.formatDate(d, tz, 'yyyy-MM'));
  }

  const itemNames = ['กล่อง', 'เทปใส', 'บับเบิ้ล'];
  const series = {};
  itemNames.forEach(n => { series[n] = monthKeys.map(() => 0); });
  const totals = { 'กล่อง': 0, 'เทปใส': 0, 'บับเบิ้ล': 0 };
  const thisMonthKey = monthKeys[monthKeys.length - 1];
  let totalRequests = 0;
  let thisMonthRequests = 0;

  values.forEach(row => {
    const ts = row[0];
    if (!(ts instanceof Date)) return;
    const monthKey = Utilities.formatDate(ts, tz, 'yyyy-MM');
    const idx = monthKeys.indexOf(monthKey);
    let items = [];
    try { items = JSON.parse(row[4]); } catch (e) { items = []; }

    totalRequests++;
    if (monthKey === thisMonthKey) thisMonthRequests++;

    items.forEach(it => {
      const name = it.name;
      const qty = Number(it.qty) || 0;
      if (!series[name]) { series[name] = monthKeys.map(() => 0); itemNames.push(name); totals[name] = 0; }
      if (idx >= 0) series[name][idx] += qty;
      totals[name] = (totals[name] || 0) + qty;
    });
  });

  return {
    months: monthKeys,
    itemNames: itemNames,
    series: series,
    totals: totals,
    totalRequests: totalRequests,
    thisMonthRequests: thisMonthRequests
  };
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
