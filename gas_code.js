// ============================================================
//  นร./กสศ.01 — Google Apps Script Backend
//  วางไฟล์นี้ใน Google Apps Script แล้ว Deploy as Web App
//  ตั้งค่า Execute as: Me | Who has access: Anyone
// ============================================================

const SHEET_NAME   = 'ข้อมูลนักเรียน';
const FOLDER_NAME  = 'KSS01_Images';
const SHEET_ID     = ''; // ← ใส่ Spreadsheet ID ของคุณที่นี่

// ─────────────────────────────────────────────────────────────
//  CORS / Router
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'KSS01 API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    let result;

    switch (action) {
      case 'save':        result = saveRecord(payload.data);        break;
      case 'update':      result = updateRecord(payload.data);      break;
      case 'delete':      result = deleteRecord(payload.id);        break;
      case 'getAll':      result = getAllRecords();                  break;
      case 'getByClass':  result = getByClass(payload.classLevel);  break;
      case 'getById':     result = getById(payload.id);             break;
      case 'uploadImage': result = uploadImage(payload);            break;
      default:            result = { success: false, message: 'Unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────────────────────
//  Sheet helpers
// ─────────────────────────────────────────────────────────────
function getSpreadsheet() {
  return SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet() {
  const ss    = getSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    createHeaders(sheet);
  }
  return sheet;
}

function createHeaders(sheet) {
  const headers = [
    'ID', 'วันที่บันทึก', 'โรงเรียน', 'สังกัด',
    'ชื่อ', 'นามสกุล', 'ชั้น', 'เลขประชาชนนักเรียน',
    'สถานภาพครอบครัว', 'อาศัยอยู่กับ',
    'ชื่อผู้ปกครอง', 'นามสกุลผู้ปกครอง', 'ความสัมพันธ์',
    'การศึกษาผู้ปกครอง', 'อาชีพ', 'เบอร์โทร',
    'เลขประชาชนผู้ปกครอง', 'สวัสดิการแห่งรัฐ',
    'จำนวนสมาชิก', 'รายได้รวม', 'รายได้เฉลี่ยต่อคน',
    'ภาระพึ่งพิง', 'การอยู่อาศัย', 'ค่าเช่า',
    'แหล่งไฟฟ้า', 'ที่ดินเกษตร',
    'วิธีเดินทาง', 'ระยะทาง(กม)', 'เวลาเดินทาง', 'ค่าเดินทาง',
    'บ้านเลขที่', 'หมู่', 'ถนน/ซอย', 'ตำบล', 'อำเภอ', 'จังหวัด', 'รหัสไปรษณีย์',
    'URL_รูปนักเรียน', 'URL_รูปนอกบ้าน', 'URL_รูปในบ้าน',
    'ข้อมูล_JSON'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a56db')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(41, 400); // JSON column
}

// ─────────────────────────────────────────────────────────────
//  Drive helpers
// ─────────────────────────────────────────────────────────────
function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

function uploadImage(payload) {
  try {
    const folder    = getOrCreateFolder();
    const blob      = Utilities.newBlob(
      Utilities.base64Decode(payload.base64.split(',')[1] || payload.base64),
      payload.mimeType || 'image/jpeg',
      payload.filename || `img_${Date.now()}.jpg`
    );
    const file      = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
    return { success: true, url, fileId: file.getId() };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  CRUD
// ─────────────────────────────────────────────────────────────
function saveRecord(data) {
  try {
    const sheet = getSheet();
    const id    = Date.now().toString();
    const now   = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

    // Upload images if base64 present
    const imgStudent = data.photos?.prev_student ? uploadImageInline(data.photos.prev_student, `student_${id}.jpg`) : '';
    const imgOut     = data.photos?.prev_out     ? uploadImageInline(data.photos.prev_out,     `out_${id}.jpg`)     : '';
    const imgIn      = data.photos?.prev_in      ? uploadImageInline(data.photos.prev_in,      `in_${id}.jpg`)      : '';

    const row = [
      id, now,
      data.school || '', data.affiliation || '',
      data.std_fname || '', data.std_lname || '', data.std_class || '', data.std_id || '',
      data.family_status || '', data.live_with || '',
      data.grd_fname || '', data.grd_lname || '', data.grd_rel || '',
      data.grd_edu || '', data.grd_job || '', data.grd_phone || '',
      data.grd_id || '', data.welfare ? 'ได้รับ' : 'ไม่ได้รับ',
      (data.members || []).length, data.total_income || 0, data.avg_income || 0,
      data.burden || '', data.housing || '', data.rent || '',
      data.electric || '', data.agri || '',
      data.travel || '', data.travel_km || '', data.travel_time || '', data.travel_cost || '',
      data.addr_no || '', data.addr_moo || '', data.addr_street || '',
      data.addr_tambon || '', data.addr_amphoe || '', data.addr_province || '', data.addr_zip || '',
      imgStudent, imgOut, imgIn,
      JSON.stringify(data)
    ];

    sheet.appendRow(row);
    // Auto-resize columns 1–40
    sheet.autoResizeColumns(1, 40);

    return { success: true, id, message: 'บันทึกสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function uploadImageInline(base64, filename) {
  try {
    if (!base64 || !base64.includes(',')) return '';
    const folder = getOrCreateFolder();
    const blob   = Utilities.newBlob(
      Utilities.base64Decode(base64.split(',')[1]),
      'image/jpeg', filename
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  } catch (e) {
    return '';
  }
}

function updateRecord(data) {
  try {
    const sheet  = getSheet();
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {
        const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

        // Re-upload images only if new base64 provided
        const imgStudent = (data.photos?.prev_student && data.photos.prev_student.startsWith('data:'))
          ? uploadImageInline(data.photos.prev_student, `student_${data.id}.jpg`) : values[i][37];
        const imgOut = (data.photos?.prev_out && data.photos.prev_out.startsWith('data:'))
          ? uploadImageInline(data.photos.prev_out, `out_${data.id}.jpg`) : values[i][38];
        const imgIn = (data.photos?.prev_in && data.photos.prev_in.startsWith('data:'))
          ? uploadImageInline(data.photos.prev_in, `in_${data.id}.jpg`) : values[i][39];

        const row = [
          data.id, now,
          data.school || '', data.affiliation || '',
          data.std_fname || '', data.std_lname || '', data.std_class || '', data.std_id || '',
          data.family_status || '', data.live_with || '',
          data.grd_fname || '', data.grd_lname || '', data.grd_rel || '',
          data.grd_edu || '', data.grd_job || '', data.grd_phone || '',
          data.grd_id || '', data.welfare ? 'ได้รับ' : 'ไม่ได้รับ',
          (data.members || []).length, data.total_income || 0, data.avg_income || 0,
          data.burden || '', data.housing || '', data.rent || '',
          data.electric || '', data.agri || '',
          data.travel || '', data.travel_km || '', data.travel_time || '', data.travel_cost || '',
          data.addr_no || '', data.addr_moo || '', data.addr_street || '',
          data.addr_tambon || '', data.addr_amphoe || '', data.addr_province || '', data.addr_zip || '',
          imgStudent, imgOut, imgIn,
          JSON.stringify(data)
        ];
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return { success: true, message: 'อัปเดตสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteRecord(id) {
  try {
    const sheet  = getSheet();
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAllRecords() {
  try {
    const sheet  = getSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { success: true, data: [] };

    const records = values.slice(1).map(row => {
      try { return JSON.parse(row[40]); } catch (e) {
        // fallback: build object from columns
        return {
          id: row[0], savedAt: row[1],
          school: row[2], affiliation: row[3],
          std_fname: row[4], std_lname: row[5], std_class: row[6], std_id: row[7],
          family_status: row[8], live_with: row[9],
          grd_fname: row[10], grd_lname: row[11], grd_rel: row[12],
          grd_edu: row[13], grd_job: row[14], grd_phone: row[15],
          grd_id: row[16], welfare: row[17] === 'ได้รับ',
          total_income: row[19], avg_income: row[20],
          photos: { prev_student: row[37], prev_out: row[38], prev_in: row[39] }
        };
      }
    });
    return { success: true, data: records };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getByClass(classLevel) {
  try {
    const all = getAllRecords();
    if (!all.success) return all;
    const filtered = classLevel
      ? all.data.filter(r => r.std_class === classLevel)
      : all.data;
    return { success: true, data: filtered };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getById(id) {
  try {
    const all = getAllRecords();
    if (!all.success) return all;
    const record = all.data.find(r => String(r.id) === String(id));
    return record
      ? { success: true, data: record }
      : { success: false, message: 'ไม่พบข้อมูล' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  Utility: test function (run manually in GAS editor)
// ─────────────────────────────────────────────────────────────
function testSetup() {
  const sheet = getSheet();
  const folder = getOrCreateFolder();
  Logger.log('Sheet: ' + sheet.getName());
  Logger.log('Folder: ' + folder.getName());
  Logger.log('Setup OK ✅');
}
