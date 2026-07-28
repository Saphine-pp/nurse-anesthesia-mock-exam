/**
 * Google Apps Script backend for "มิกซ์ Mock Exam 100 ข้อ" (ทุกชุด)
 *
 * Setup:
 * 1. สร้าง Google Sheet ใหม่ (ว่าง ๆ ก็ได้) เช่นชื่อ "Exam Scores"
 * 2. เปิด Extensions > Apps Script แล้ววางไฟล์นี้ทับ Code.gs ที่มีอยู่
 * 3. Deploy > New deployment > เลือกประเภท "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. คัดลอก Web app URL ที่ได้ แล้วส่งกลับมาวางใน
 *    - mixed_mock_exam_100_tracked.html      (ตัวแปร SUBMIT_URL)
 *    - mixed_mock_exam_100_set2_exam.html    (ตัวแปร SUBMIT_URL)
 *    - dashboard.html                        (ตัวแปร DATA_URL)
 */

const SHEET_NAME = "Responses";
const HEADERS = ["Timestamp", "Name", "Score", "Total", "Vent", "Opioid", "Volatile", "Induction", "NMB", "Set", "Status", "Answered"];

function doPost(e) {
  const sheet = getSheet();
  const data = JSON.parse(e.postData.contents);
  sheet.appendRow([
    new Date(),
    data.name || "",
    data.score,
    data.total,
    data.vent || "",
    data.op || "",
    data.vol || "",
    data.ind || "",
    data.nmb || "",
    data.set || "",
    data.status || "ครบ",
    data.answered != null ? data.answered : data.total
  ]);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // drop header row
  const out = rows.map(function (r) {
    return {
      timestamp: r[0] instanceof Date ? r[0].toISOString() : r[0],
      name: r[1],
      score: r[2],
      total: r[3],
      vent: r[4],
      op: r[5],
      vol: r[6],
      ind: r[7],
      nmb: r[8],
      set: r[9] || "",
      status: r[10] || "ครบ",
      answered: (r[11] !== "" && r[11] != null) ? r[11] : r[3]
    };
  });
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  } else {
    const lastCol = sheet.getLastColumn();
    if (lastCol < HEADERS.length) {
      sheet.getRange(1, lastCol + 1, 1, HEADERS.length - lastCol)
        .setValues([HEADERS.slice(lastCol)]);
    }
  }
  return sheet;
}
