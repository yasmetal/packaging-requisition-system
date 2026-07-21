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

function itemLabel_(item) {
  // แต่ละรายการที่มีการระบุ size จะแสดงคำต่อท้ายชื่อให้เหมาะสมกับรายการนั้น
  if (!item.size) return item.name;
  if (item.name === 'ถุงพลาสติก') return `${item.name} เบอร์${item.size}`;
  return `${item.name} ขนาด${item.size}`;
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

    const itemsSummary = items.map(i => `${itemLabel_(i)} x${i.qty}`).join(', ');

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
      // months: จำนวนเดือนย้อนหลัง (ตัวเลข) หรือ "all" เพื่อดูตั้งแต่รายการแรกสุด
      const monthsParam = e.parameter.months === 'all' ? 'all' : (parseInt(e.parameter.months, 10) || 6);
      return jsonResponse_({ ok: true, summary: getMonthlySummary_(monthsParam) });
    }
    return jsonResponse_({ ok: true, message: 'ระบบเบิกวัสดุบรรจุภัณฑ์ API ทำงานปกติ' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

/**
 * รวมยอดจำนวนการเบิกแต่ละรายการ แยกตามเดือน
 * monthsParam: จำนวนเดือนย้อนหลัง (รวมเดือนปัจจุบัน) หรือ 'all' เพื่อไล่ตั้งแต่เดือนของรายการแรกสุดในชีต
 * หมายเหตุ: ถุงพลาสติกทุกเบอร์จะถูกรวมเป็นรายการ "ถุงพลาสติก" รายการเดียว (ไม่แยกตามเบอร์)
 * คืนค่า { months: ["2026-02", ...], itemNames: [...], series: {itemName: [qty,...]}, totals: {...},
 *          totalRequests, thisMonthRequests, requestsByMonth: [n, ...] }
 */
function getMonthlySummary_(monthsParam) {
  const tz = Session.getScriptTimeZone();
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  values.shift(); // remove header

  const now = new Date();
  const monthKeys = [];

  if (monthsParam === 'all') {
    // หาเดือนที่เก่าที่สุดจากข้อมูลจริง แล้วไล่มาจนถึงเดือนปัจจุบัน
    let earliest = now;
    values.forEach(row => {
      const ts = row[0];
      if (ts instanceof Date && ts < earliest) earliest = ts;
    });
    const start = new Date(earliest.getFullYear(), earliest.getMonth(), 1);
    const end = new Date(now.getFullYear(), now.getMonth(), 1);
    let cursor = start;
    while (cursor <= end) {
      monthKeys.push(Utilities.formatDate(cursor, tz, 'yyyy-MM'));
      cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    }
    if (monthKeys.length === 0) {
      monthKeys.push(Utilities.formatDate(now, tz, 'yyyy-MM'));
    }
  } else {
    const monthsBack = monthsParam;
    for (let i = monthsBack - 1; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      monthKeys.push(Utilities.formatDate(d, tz, 'yyyy-MM'));
    }
  }

  const itemNames = ['กล่อง', 'เทปใส', 'บับเบิ้ล', 'ถุงพลาสติก'];
  const series = {};
  itemNames.forEach(n => { series[n] = monthKeys.map(() => 0); });
  const totals = { 'กล่อง': 0, 'เทปใส': 0, 'บับเบิ้ล': 0, 'ถุงพลาสติก': 0 };
  const requestsByMonth = monthKeys.map(() => 0);
  const thisMonthKey = Utilities.formatDate(now, tz, 'yyyy-MM');
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
    if (idx >= 0) requestsByMonth[idx]++;

    items.forEach(it => {
      const name = it.name; // รวมทุกเบอร์ของถุงพลาสติกเป็นรายการเดียว
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
    thisMonthRequests: thisMonthRequests,
    requestsByMonth: requestsByMonth
  };
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
