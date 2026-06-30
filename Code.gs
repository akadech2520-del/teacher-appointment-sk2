/**
 * =======================================================
 * ระบบประเมินครูผู้ช่วย ก.ค.ศ. - Backend Google Apps Script
 * =======================================================
 */

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : null;
  var callback = (e && e.parameter) ? e.parameter.callback : null;
  
  // 1. สำหรับการดึงข้อมูล API (เมื่อหน้าเว็บร้องขอ)
  if (action === 'getData') {
    var output = JSON.stringify({
      status: 'success',
      data: {
        teachers: getSheetData("Teachers"),
        directors: getSheetData("Directors"),
        experts: getSheetData("Experts"),
        schools: getSheetData("Schools"),
        appointments: getAppointmentsData(),
        evaluations: getEvaluationsData()
      }
    });
    if (callback) {
      callback = String(callback).replace(/[^\w.$]/g, '');
      return ContentService.createTextOutput(callback + '(' + output + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. สำหรับการเปิดหน้าเว็บผ่านเบราว์เซอร์ปกติ ให้เรียกไฟล์ Index.html
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบประเมินครูผู้ช่วย')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    let resultData = null;

    // ระบบจัดการฐานข้อมูลหลัก (Master Data)
    if (action === 'saveTeacher') saveToSheet("Teachers", data.id, data);
    else if (action === 'deleteTeacher') deleteFromSheet("Teachers", data.id);
    
    else if (action === 'saveSchool') saveToSheet("Schools", data.id, data);
    else if (action === 'deleteSchool') deleteFromSheet("Schools", data.id);
    
    else if (action === 'saveDirector') saveToSheet("Directors", data.id, data);
    else if (action === 'deleteDirector') deleteFromSheet("Directors", data.id);
    
    else if (action === 'saveExpert') saveToSheet("Experts", data.id, data);
    else if (action === 'deleteExpert') deleteFromSheet("Experts", data.id);
    
    // ระบบประเมินและแต่งตั้ง
    else if (action === 'saveAppointment') resultData = saveAppointment(data);
    else if (action === 'saveEval') resultData = saveEval(data.key, data.data);
    
    else throw new Error('ไม่พบ Action ที่ระบุ: ' + action);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: resultData }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ==========================================
 * Helper Functions (ฟังก์ชันจัดการ Google Sheets)
 * ==========================================
 * บันทึกข้อมูลแบบ ID (คอลัมน์ A) และ JSON Data (คอลัมน์ B)
 * ทำให้โครงสร้างยืดหยุ่นมาก ไม่ต้องสร้างคอลัมน์เรียงตามฟิลด์
 */

const SHEET_ID = '1dkP8nhqi9t64jmAxrYTP2PuyLE81J-i0KORhRGwayxU';
const ORDER_FOLDER_NAME = 'คำสั่งแต่งตั้งคณะกรรมการฯ';
const EVAL_FOLDER_NAME = 'ผลคะแนนการประเมินครูผู้ช่วย';
const SIGNATURE_FOLDER_NAME = 'ลายมือชื่อคณะกรรมการประเมินครูผู้ช่วย';

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // ถ้ายังไม่มีชีตนี้ ให้สร้างใหม่และใส่หัวตาราง
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID / Key", "JSON Data"]);
    sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#E6F2FF");
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 600);
  }
  return sheet;
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // มีแค่หัวตารางหรือว่างเปล่า

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if(!row[0]) continue; 
    try {
      result.push(JSON.parse(row[1]));
    } catch(e) {
      // ข้ามแถวที่ JSON ไม่ถูกต้อง
    }
  }
  return result;
}

function saveToSheet(sheetName, id, objData) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const jsonStr = JSON.stringify(objData);

  // ค้นหาว่ามี ID นี้อยู่แล้วหรือไม่ (เพื่ออัปเดต)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 2).setValue(jsonStr);
      return;
    }
  }
  // ถ้าไม่มี ให้เพิ่มแถวใหม่
  sheet.appendRow([id, jsonStr]);
}

function getObjectById(sheetName, id) {
  const items = getSheetData(sheetName);
  for (let i = 0; i < items.length; i++) {
    if (items[i].id == id) return items[i];
  }
  return null;
}

function deleteFromSheet(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// ---- ฟังก์ชันเฉพาะสำหรับ Appointments ----
function getAppointmentsData() {
  const sheet = getSheet("Appointments");
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    if(!data[i][0]) continue;
    try { result[data[i][0]] = JSON.parse(data[i][1]); } catch(e){}
  }
  return result;
}

function saveAppointment(data) {
  const previous = getAppointmentsData()[data.tid] || {};
  data = saveAppointmentSignatures_(data, previous);
  const orderFile = saveAppointmentOrderToDrive(data);
  data.orderFileId = orderFile.id;
  data.orderFileUrl = orderFile.url;
  data.orderFileName = orderFile.name;
  data.orderSavedAt = new Date().toISOString();

  const sheet = getSheet("Appointments");
  const sheetData = sheet.getDataRange().getValues();
  const jsonStr = JSON.stringify(data);
  
  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] == data.tid) {
      sheet.getRange(i + 1, 2).setValue(jsonStr);
      return data;
    }
  }
  sheet.appendRow([data.tid, jsonStr]);
  return data;
}

function saveAppointmentSignatures_(data, previous) {
  if (data.directorSignature && !data.c1Signature) data.c1Signature = data.directorSignature;
  if (data.directorSignatureFileId && !data.c1SignatureFileId) data.c1SignatureFileId = data.directorSignatureFileId;
  if (data.directorSignatureFileUrl && !data.c1SignatureFileUrl) data.c1SignatureFileUrl = data.directorSignatureFileUrl;
  if (data.directorSignatureImageUrl && !data.c1SignatureImageUrl) data.c1SignatureImageUrl = data.directorSignatureImageUrl;
  if (data.directorSignatureName && !data.c1SignatureName) data.c1SignatureName = data.directorSignatureName;

  ['c1', 'c2', 'c3'].forEach(function(key) {
    const rawKey = key + 'Signature';
    const fileIdKey = key + 'SignatureFileId';
    const fileUrlKey = key + 'SignatureFileUrl';
    const imageUrlKey = key + 'SignatureImageUrl';
    const nameKey = key + 'SignatureName';
    const raw = String(data[rawKey] || '');
    const hasFileField = Object.prototype.hasOwnProperty.call(data, fileIdKey);

    if (isDataImage_(raw)) {
      if (data[fileIdKey] || previous[fileIdKey]) trashFileQuietly_(data[fileIdKey] || previous[fileIdKey]);
      const saved = saveSignatureImageToDrive_(raw, data.tid, key, data[nameKey]);
      data[fileIdKey] = saved.id;
      data[fileUrlKey] = saved.fileUrl;
      data[imageUrlKey] = saved.imageUrl;
      data[nameKey] = saved.name;
    } else if (!hasFileField && previous[fileIdKey]) {
      data[fileIdKey] = previous[fileIdKey] || '';
      data[fileUrlKey] = previous[fileUrlKey] || '';
      data[imageUrlKey] = previous[imageUrlKey] || '';
      data[nameKey] = previous[nameKey] || '';
    } else if (hasFileField && !data[fileIdKey] && previous[fileIdKey]) {
      trashFileQuietly_(previous[fileIdKey]);
      data[fileUrlKey] = '';
      data[imageUrlKey] = '';
      data[nameKey] = '';
    }
    data[rawKey] = '';
  });
  data.directorSignatureFileId = data.c1SignatureFileId || '';
  data.directorSignatureFileUrl = data.c1SignatureFileUrl || '';
  data.directorSignatureImageUrl = data.c1SignatureImageUrl || '';
  data.directorSignatureName = data.c1SignatureName || '';
  data.directorSignature = data.c1SignatureImageUrl || previous.c1SignatureImageUrl || previous.directorSignatureImageUrl || previous.directorSignature || '';
  return data;
}

function saveSignatureImageToDrive_(dataUrl, teacherId, key, originalName) {
  const folder = getOrCreateFolder_(SIGNATURE_FOLDER_NAME);
  const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error('รูปแบบไฟล์ลายมือชื่อไม่ถูกต้อง');
  const mimeType = match[1];
  const extension = mimeType.indexOf('png') !== -1 ? 'png' : (mimeType.indexOf('jpeg') !== -1 || mimeType.indexOf('jpg') !== -1 ? 'jpg' : 'img');
  const fileName = sanitizeFileName_('ลายมือชื่อ_' + teacherId + '_' + key + '_' + (originalName || 'signature')) + '.' + extension;
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setDescription('ลายมือชื่ออิเล็กทรอนิกส์สำหรับระบบประเมินครูผู้ช่วย');
  return {
    id: file.getId(),
    fileUrl: file.getUrl(),
    imageUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    name: file.getName()
  };
}

function trashFileQuietly_(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {}
}

function isDataImage_(value) {
  return String(value || '').indexOf('data:image/') === 0;
}

function getSignatureDataUri_(data, key) {
  const raw = String(data[key + 'Signature'] || '');
  if (isDataImage_(raw)) return raw;
  const fileId = data[key + 'SignatureFileId'];
  if (!fileId) return '';
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return '';
  }
}

function saveAppointmentOrderToDrive(data) {
  const teacher = getObjectById("Teachers", data.tid) || {};
  const folder = getOrCreateFolder_(ORDER_FOLDER_NAME);
  const fileName = buildOrderFileName_(data, teacher);

  if (data.orderFileId) {
    try {
      DriveApp.getFileById(data.orderFileId).setTrashed(true);
    } catch (e) {}
  }

  const html = buildAppointmentOrderHtml_(data, teacher);
  const pdfBlob = Utilities
    .newBlob(html, 'text/html', fileName + '.html')
    .getAs(MimeType.PDF)
    .setName(fileName + '.pdf');
  const file = folder.createFile(pdfBlob);
  file.setDescription('คำสั่งแต่งตั้งคณะกรรมการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย');

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}

function getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function buildOrderFileName_(data, teacher) {
  const safeTeacher = sanitizeFileName_(teacher.name || data.tid || 'ครูผู้ช่วย');
  const safeOrder = sanitizeFileName_(data.orderNo || 'ไม่ระบุเลขคำสั่ง');
  return 'คำสั่งแต่งตั้งคณะกรรมการฯ_' + safeTeacher + '_' + safeOrder;
}

function sanitizeFileName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120);
}

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, function(ch) {
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch];
  });
}

function toThaiDigits_(value) {
  const digits = ['๐','๑','๒','๓','๔','๕','๖','๗','๘','๙'];
  return String(value || '').replace(/\d/g, function(n) {
    return digits[Number(n)];
  });
}

function parseThaiDate_(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split('/');
  if (parts.length !== 3) return null;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  let year = parseInt(parts[2], 10);
  if (year > 2400) year -= 543;
  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  return new Date(year, month, day);
}

function addMonths_(date, months) {
  const result = new Date(date);
  result.setMonth(result.getMonth() + months);
  return result;
}

function formatThaiLongDate_(date) {
  const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return 'วันที่ ' + toThaiDigits_(date.getDate()) + ' เดือน' + months[date.getMonth()] + ' พ.ศ. ' + toThaiDigits_(date.getFullYear() + 543);
}

function formatThaiLongDateInput_(dateStr) {
  const parsed = parseThaiDate_(dateStr);
  if (parsed) return formatThaiLongDate_(parsed);
  return toThaiDigits_(dateStr || 'วันที่.......... เดือน.................... พ.ศ. ..........');
}

function getCompletionDateText_(dateStr) {
  const parsed = parseThaiDate_(dateStr);
  if (!parsed) return 'วันที่.......... เดือน.................... พ.ศ. ..........';
  return formatThaiLongDate_(addMonths_(parsed, 24));
}

function getFullSchoolName_(value) {
  const name = String(value || '').trim();
  if (!name) return 'โรงเรียน................................................';
  return name.indexOf('โรงเรียน') === 0 ? name : 'โรงเรียน' + name;
}

function buildAppointmentOrderHtml_(data, teacher) {
  const oNo = escapeHtml_(toThaiDigits_(data.orderNo || '............../2569'));
  const oDateStr = escapeHtml_(formatThaiLongDateInput_(data.orderDate));
  const c1 = escapeHtml_(toThaiDigits_(data.c1 || '....................................'));
  const c2 = escapeHtml_(toThaiDigits_(data.c2 || 'นาย/นาง/นางสาว....................................'));
  const c3 = escapeHtml_(toThaiDigits_(data.c3 || 'ผู้ทรงคุณวุฒิอื่นจากภายนอกสถานศึกษา'));
  const c3Pos = escapeHtml_(toThaiDigits_(data.c3Pos || '.................'));
  const schoolFullName = escapeHtml_(toThaiDigits_(getFullSchoolName_(teacher.school)));
  const teacherName = escapeHtml_(toThaiDigits_(teacher.name || 'นาย/นาง/นางสาว................................................'));
  const startDateText = escapeHtml_(formatThaiLongDateInput_(teacher.startDate));
  const completeDateText = escapeHtml_(getCompletionDateText_(teacher.startDate));
  const c1Signature = getSignatureDataUri_(data, 'c1');
  const c1SignatureHtml = c1Signature ? '<img class="signature-img" src="' + c1Signature + '" alt="ลายมือชื่อ ผอ.รร.">' : '<span class="signature-line">&nbsp;</span>';
  const garudaImageUrl = 'https://img1.pic.in.th/images/images-183d7b3eabccd5225.png';
  const orderStyle = '@page{size:A4;margin:2cm 2cm 2cm 3cm}body{font-family:"TH Sarabun New","Sarabun",Arial,sans-serif;color:#000;margin:0;font-size:15pt;line-height:1.18;letter-spacing:0;word-spacing:0}.order-head{text-align:center;font-weight:700;margin-bottom:10px}.garuda-emblem{display:block;width:3cm;height:auto;margin:0 auto 5px}.order-title{font-size:17pt;font-weight:700}.dash{font-weight:400;margin-top:2px}p{margin:0 0 5px;text-indent:2.5cm;text-align:left;letter-spacing:0;word-spacing:0}.no-indent{text-indent:0}.item{padding-left:22px;text-indent:-22px}.subitem{padding-left:44px;text-indent:-26px}.committee-table{width:100%;border-collapse:collapse;margin:5px 0 9px 0}.committee-table td{border:none;padding:0 4px;vertical-align:top;font-size:15pt;line-height:1.18;letter-spacing:0;word-spacing:0}.committee-table td:first-child{width:22px}.committee-table td:nth-child(2){width:33%}.committee-table td:nth-child(3){width:42%}.committee-table td:nth-child(4){width:18%;white-space:nowrap}.page-break{break-before:page;page-break-before:always;height:0}.signature{width:8.5cm;text-align:center;margin-top:22px;margin-left:auto;line-height:1.18}.signature-date{margin-bottom:8px}.sign-row{height:54px;display:flex;align-items:flex-end;justify-content:center;gap:8px}.sign-prefix{white-space:nowrap}.signature-img{display:block;max-width:155px;max-height:60px;object-fit:contain;margin:0}.signature-line{display:inline-block;width:170px;border-bottom:1px dotted #000;height:30px}.sign-name{margin-top:2px;font-weight:400}';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' + orderStyle + '</style>' +
    '</head><body><main class="order-page">' +
    '<div class="order-head">' +
    '<img class="garuda-emblem" src="' + garudaImageUrl + '" alt="ตราครุฑ">' +
    '<div class="order-title">คำสั่ง' + schoolFullName + '</div>' +
    '<div>ที่&nbsp;&nbsp;' + oNo + '</div>' +
    '<div>เรื่อง&nbsp;&nbsp;แต่งตั้งคณะกรรมการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย</div>' +
    '<div class="dash">.................................................................</div>' +
    '</div>' +
    '<p>อาศัยอำนาจตามความในมาตรา ๕๓ แห่งพระราชบัญญัติระเบียบข้าราชการครูและบุคลากรทางการศึกษา พ.ศ. ๒๕๔๗ และที่แก้ไขเพิ่มเติม ประกอบกับพระราชบัญญัติแก้ไขเพิ่มเติมคำสั่งหัวหน้าคณะรักษาความสงบแห่งชาติ ที่ ๑๙/๒๕๖๐ สั่ง ณ วันที่ ๓ เมษายน พุทธศักราช ๒๕๖๐ พ.ศ. ๒๕๖๕ หนังสือสำนักงาน ก.ค.ศ. ที่ ศธ ๐๒๐๖.๗/ว ๑๙ ลงวันที่ ๑๙ สิงหาคม ๒๕๖๘ เรื่อง หลักเกณฑ์และวิธีการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย และมติ อ.ก.ค.ศ. สำนักงานเขตพื้นที่การศึกษาประถมศึกษาสกลนคร เขต ๒ ครั้งที่ ๑/๒๕๖๙ เมื่อวันที่ ๒๘ มกราคม ๒๕๖๙ เรื่อง การแต่งตั้งบัญชีศึกษานิเทศก์และผู้ทรงคุณวุฒิ เพื่อเป็นกรรมการในการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย สังกัดสำนักงานเขตพื้นที่การศึกษาประถมศึกษาสกลนคร เขต ๒</p>' +
    '<p>เพื่อให้การเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย เป็นเวลา ๒ ปี ก่อนแต่งตั้งให้ดำรงตำแหน่งครู เป็นไปอย่างมีระบบ มีมาตรฐาน เป็นไปตามหลักเกณฑ์และวิธีการที่ ก.ค.ศ. กำหนด และเพื่อพัฒนาครูผู้ช่วยให้มีความรู้ ความสามารถ วินัย คุณธรรม จริยธรรม จรรยาบรรณวิชาชีพ สมรรถนะในการปฏิบัติตนและการปฏิบัติงาน ส่งผลต่อคุณภาพผู้เรียน</p>' +
    '<p>' + schoolFullName + ' จึงแต่งตั้งคณะกรรมการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย ของ ' + teacherName + ' ตำแหน่งครูผู้ช่วย ' + schoolFullName + ' สังกัดสำนักงานเขตพื้นที่การศึกษาประถมศึกษาสกลนคร เขต ๒ ซึ่งได้รับการบรรจุและแต่งตั้งเมื่อ' + startDateText + ' และจะครบกำหนดการเตรียมความพร้อมและพัฒนาอย่างเข้มใน' + completeDateText + ' ดังนี้</p>' +
    '<table class="committee-table">' +
    '<tr><td>๑.</td><td>' + c1 + '</td><td>ผู้อำนวยการ' + schoolFullName + '</td><td>ประธานกรรมการ</td></tr>' +
    '<tr><td>๒.</td><td>' + c2 + '</td><td>ครู (ในสถานศึกษา)</td><td>กรรมการ</td></tr>' +
    '<tr><td>๓.</td><td>' + c3 + '</td><td>ตำแหน่ง ' + c3Pos + '</td><td>กรรมการ</td></tr>' +
    '</table>' +
    '<div class="page-break"></div>' +
    '<p class="no-indent">มีหน้าที่ดังนี้ ให้คณะกรรมการเตรียมความพร้อมและพัฒนาอย่างเข้มดำเนินการ ดังนี้</p>' +
    '<p class="item">๑. ดำเนินการพัฒนาครูผู้ช่วย ด้านวิชาชีพ ด้านสังคม และด้านคุณลักษณะส่วนบุคคล ส่งเสริม สนับสนุน ให้คำปรึกษา สอนงาน ช่วยเหลือ และแนะนำการปฏิบัติหน้าที่ โดยใช้กระบวนการชี้แนะ ดังต่อไปนี้</p>' +
    '<p class="subitem">๑.๑ วางแผนและกำหนดเป้าหมายในการเตรียมความพร้อมและพัฒนาอย่างเข้มร่วมกันกับครูผู้ช่วย</p>' +
    '<p class="subitem">๑.๒ ส่งเสริม สนับสนุน ให้คำปรึกษา สอนงาน ช่วยเหลือ และแนะนำการปฏิบัติหน้าที่ เพื่อให้ครูผู้ช่วยมีทักษะทางวิชาชีพ ทักษะทางสังคม ตลอดจนบุคลิกและลักษณะของความเป็นครูที่ส่งผลต่อการพัฒนาคุณภาพของผู้เรียน</p>' +
    '<p class="subitem">๑.๓ ให้คำปรึกษา แนะนำในการจัดทำแนวทางในการพัฒนาการจัดการเรียนรู้</p>' +
    '<p class="item">๒. ดำเนินการประเมินการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตามแบบบันทึกการประเมิน แบบท้ายหลักเกณฑ์และวิธีการที่ ก.ค.ศ. กำหนด</p>' +
    '<p class="item">๓. รายงานผลการประเมินต่อผู้มีอำนาจตามมาตรา ๕๓ เมื่อประเมินครบแต่ละครั้ง เป็นเวลา ๒ ปี ทั้งนี้ ให้รายงานภายใน ๓๐ วันนับแต่วันที่ประเมินครั้งสุดท้าย และแจ้งผลการประเมิน พร้อมทั้งข้อเสนอแนะ จุดเด่น จุดที่ควรพัฒนา และข้อสังเกตจากการประเมินในแต่ละครั้งให้ครูผู้ช่วยทราบ</p>' +
    '<p class="item">๔. พิจารณาทบทวนผลการประเมินจากข้อโต้แย้งและพยานหลักฐาน ในกรณีที่ครูผู้ช่วยมีผลการประเมินการเตรียมความพร้อมต่ำกว่าเกณฑ์การประเมิน และรายงานผลการพิจารณาทบทวนให้ผู้มีอำนาจตามมาตรา ๕๓ ทราบ</p>' +
    '<p class="item">๕. ปฏิบัติหน้าที่อื่นที่จำเป็นหรือเกี่ยวข้องกับการเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย</p>' +
    '<p>ทั้งนี้ ให้คณะกรรมการที่ได้รับแต่งตั้งปฏิบัติหน้าที่ด้วยความรับผิดชอบ โปร่งใส เป็นธรรม และเป็นไปตามหลักเกณฑ์และวิธีการที่ ก.ค.ศ. กำหนด เพื่อให้การเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย เกิดประสิทธิภาพและประสิทธิผลสูงสุด</p>' +
    '<div class="signature">' +
    '<div class="signature-date">สั่ง ณ ' + oDateStr + '</div>' +
    '<div class="sign-row"><span class="sign-prefix">ลงชื่อ</span>' + c1SignatureHtml + '</div>' +
    '<div class="sign-name">(' + c1 + ')</div>' +
    '<div>ผู้อำนวยการ' + schoolFullName + '</div>' +
    '</div>' +
    '</main></body></html>';
}

// ---- ฟังก์ชันเฉพาะสำหรับ Evaluations ----
function getEvaluationsData() {
  const sheet = getSheet("Evaluations");
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    if(!data[i][0]) continue;
    try { result[data[i][0]] = JSON.parse(data[i][1]); } catch(e){}
  }
  return result;
}

function saveEval(key, evalData) {
  const meta = parseEvalKey_(key);
  const evalFile = saveEvaluationResultToDrive_(key, evalData, meta);
  evalData.evalFileId = evalFile.id;
  evalData.evalFileUrl = evalFile.url;
  evalData.evalFileName = evalFile.name;
  evalData.evalSavedAt = new Date().toISOString();

  if (meta.period === 4) {
    const finalFile = saveFinalEvaluationSummaryToDrive_(key, evalData, meta);
    evalData.finalEvalFileId = finalFile.id;
    evalData.finalEvalFileUrl = finalFile.url;
    evalData.finalEvalFileName = finalFile.name;
    evalData.finalEvalSavedAt = new Date().toISOString();
  }

  const sheet = getSheet("Evaluations");
  const sheetData = sheet.getDataRange().getValues();
  const jsonStr = JSON.stringify(evalData);
  
  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] == key) {
      sheet.getRange(i + 1, 2).setValue(jsonStr);
      return evalData;
    }
  }
  sheet.appendRow([key, jsonStr]);
  return evalData;
}

function parseEvalKey_(key) {
  const parts = String(key || '').split('_');
  const evaluator = parseInt(parts.pop(), 10);
  const period = parseInt(parts.pop(), 10);
  return {
    teacherId: parts.join('_'),
    period: isNaN(period) ? 0 : period,
    evaluator: isNaN(evaluator) ? 0 : evaluator
  };
}

function saveEvaluationResultToDrive_(key, evalData, meta) {
  const teacher = getObjectById("Teachers", meta.teacherId) || {};
  const folder = getOrCreateFolder_(EVAL_FOLDER_NAME);
  const fileName = buildEvalFileName_(teacher, meta);

  if (evalData.evalFileId) {
    try {
      DriveApp.getFileById(evalData.evalFileId).setTrashed(true);
    } catch (e) {}
  }

  const html = buildEvaluationResultHtml_(evalData, teacher, meta);
  const pdfBlob = Utilities
    .newBlob(html, 'text/html', fileName + '.html')
    .getAs(MimeType.PDF)
    .setName(fileName + '.pdf');
  const file = folder.createFile(pdfBlob);
  file.setDescription('ผลคะแนนการประเมินครูผู้ช่วย รายรอบ');

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}

function saveFinalEvaluationSummaryToDrive_(key, evalData, meta) {
  const teacher = getObjectById("Teachers", meta.teacherId) || {};
  const folder = getOrCreateFolder_(EVAL_FOLDER_NAME);
  const fileName = buildFinalEvalFileName_(teacher);

  if (evalData.finalEvalFileId) {
    try {
      DriveApp.getFileById(evalData.finalEvalFileId).setTrashed(true);
    } catch (e) {}
  }

  const allEvaluations = getEvaluationsData();
  allEvaluations[key] = evalData;
  const html = buildFinalEvaluationSummaryHtml_(allEvaluations, teacher, meta.teacherId);
  const pdfBlob = Utilities
    .newBlob(html, 'text/html', fileName + '.html')
    .getAs(MimeType.PDF)
    .setName(fileName + '.pdf');
  const file = folder.createFile(pdfBlob);
  file.setDescription('สรุปผลคะแนนการประเมินครูผู้ช่วย รอบสุดท้าย');

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}

function buildEvalFileName_(teacher, meta) {
  return 'ผลประเมิน_' +
    sanitizeFileName_(teacher.name || meta.teacherId || 'ครูผู้ช่วย') +
    '_ครั้งที่_' + meta.period +
    '_กรรมการ_' + meta.evaluator;
}

function buildFinalEvalFileName_(teacher) {
  return 'สรุปผลประเมินรอบสุดท้าย_' + sanitizeFileName_(teacher.name || 'ครูผู้ช่วย');
}

function getEvalCriteria_(period) {
  if (period === 1 || period === 2) return { pro: 9, soc: 0, per: 7, proTotal: 15, socTotal: 0, perTotal: 11 };
  return { pro: 11, soc: 3, per: 11, proTotal: 15, socTotal: 3, perTotal: 16 };
}

function countPassed_(obj) {
  let count = 0;
  obj = obj || {};
  Object.keys(obj).forEach(function(key) {
    if (obj[key]) count++;
  });
  return count;
}

function getEvalScore_(evalData, period) {
  const cr = getEvalCriteria_(period);
  const proCount = countPassed_(evalData.pro);
  const socCount = countPassed_(evalData.soc);
  const perCount = countPassed_(evalData.per);
  const passed = (proCount >= cr.pro) && (perCount >= cr.per) && (period < 3 || socCount >= cr.soc);
  return {
    proCount: proCount,
    socCount: socCount,
    perCount: perCount,
    passed: passed,
    label: passed ? 'ผ่าน' : 'ไม่ผ่าน',
    criteria: cr
  };
}

function buildEvaluationResultHtml_(evalData, teacher, meta) {
  const score = getEvalScore_(evalData, meta.period);
  const notes = evalData.notes || {};
  const evaluatorName = getEvaluatorName_(teacher.id || meta.teacherId, meta.evaluator);
  const signerName = escapeHtml_(getEvaluatorPersonName_(teacher.id || meta.teacherId, meta.evaluator));
  const signature = isDataImage_(evalData.signature) ? evalData.signature : getEvaluatorSignature_(teacher.id || meta.teacherId, meta.evaluator);
  const signatureHtml = signature ? '<img class="signature-img" src="' + signature + '" alt="ลายมือชื่อกรรมการ">' : '<br><br>';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>@page{size:A4;margin:2cm}body{font-family:"Sarabun","TH Sarabun New",Arial,sans-serif;font-size:14pt;color:#000;line-height:1.45}h1,h2{margin:0;text-align:center}h1{font-size:18pt}h2{font-size:15pt;margin-bottom:18px}.info{margin:18px 0}.info div{margin:4px 0}.score-table{width:100%;border-collapse:collapse;margin:14px 0}.score-table th,.score-table td{border:1px solid #000;padding:6px 8px;text-align:center}.score-table th{font-weight:700}.left{text-align:left!important}.result{font-size:18pt;font-weight:700;text-align:center;margin:18px 0}.notes{margin-top:12px}.notes div{margin-bottom:8px}.signature{margin-top:42px;text-align:center;margin-left:55%;line-height:1.45}.signature-img{display:block;max-width:150px;max-height:70px;object-fit:contain;margin:6px auto 2px}</style>' +
    '</head><body>' +
    '<h1>ผลคะแนนการประเมินการเตรียมความพร้อมและพัฒนาอย่างเข้ม</h1>' +
    '<h2>ตำแหน่งครูผู้ช่วย ครั้งที่ ' + meta.period + ' / กรรมการคนที่ ' + meta.evaluator + '</h2>' +
    '<div class="info">' +
    '<div><strong>ชื่อ-สกุล:</strong> ' + escapeHtml_(teacher.name || '-') + '</div>' +
    '<div><strong>สถานศึกษา:</strong> ' + escapeHtml_(teacher.school || '-') + '</div>' +
    '<div><strong>ผู้ประเมิน:</strong> ' + escapeHtml_(evaluatorName) + '</div>' +
    '<div><strong>วันที่บันทึก:</strong> ' + formatThaiLongDate_(new Date()) + '</div>' +
    '</div>' +
    '<table class="score-table"><thead><tr><th>ด้านการประเมิน</th><th>คะแนน/รายการที่ผ่าน</th><th>เกณฑ์ผ่าน</th><th>ผล</th></tr></thead><tbody>' +
    '<tr><td class="left">ด้านวิชาชีพ</td><td>' + score.proCount + ' / ' + score.criteria.proTotal + '</td><td>' + score.criteria.pro + '</td><td>' + (score.proCount >= score.criteria.pro ? 'ผ่าน' : 'ไม่ผ่าน') + '</td></tr>' +
    '<tr><td class="left">ด้านสังคม</td><td>' + score.socCount + ' / ' + score.criteria.socTotal + '</td><td>' + score.criteria.soc + '</td><td>' + (meta.period < 3 ? '-' : (score.socCount >= score.criteria.soc ? 'ผ่าน' : 'ไม่ผ่าน')) + '</td></tr>' +
    '<tr><td class="left">ด้านคุณลักษณะส่วนบุคคล</td><td>' + score.perCount + ' / ' + score.criteria.perTotal + '</td><td>' + score.criteria.per + '</td><td>' + (score.perCount >= score.criteria.per ? 'ผ่าน' : 'ไม่ผ่าน') + '</td></tr>' +
    '</tbody></table>' +
    '<div class="result">ผลการประเมิน: ' + score.label + '</div>' +
    '<div class="notes"><div><strong>จุดเด่น:</strong> ' + escapeHtml_(notes.strength || '-') + '</div>' +
    '<div><strong>จุดที่ควรพัฒนา:</strong> ' + escapeHtml_(notes.improve || '-') + '</div>' +
    '<div><strong>ข้อเสนอแนะ:</strong> ' + escapeHtml_(notes.suggest || '-') + '</div></div>' +
    '<div class="signature">' + signatureHtml + '(' + signerName + ')<br>กรรมการผู้ประเมิน</div>' +
    '</body></html>';
}

function getEvaluatorName_(teacherId, evaluator) {
  const ap = getAppointmentsData()[teacherId] || {};
  if (evaluator === 1 && ap.c1) return ap.c1 + ' (ประธานกรรมการ)';
  if (evaluator === 2 && ap.c2) return ap.c2 + ' (กรรมการ)';
  if (evaluator === 3 && ap.c3) return ap.c3 + ' (กรรมการ)';
  return 'กรรมการคนที่ ' + evaluator;
}

function getEvaluatorPersonName_(teacherId, evaluator) {
  const ap = getAppointmentsData()[teacherId] || {};
  if (evaluator === 1 && ap.c1) return ap.c1;
  if (evaluator === 2 && ap.c2) return ap.c2;
  if (evaluator === 3 && ap.c3) return ap.c3;
  return 'กรรมการคนที่ ' + evaluator;
}

function getEvaluatorSignature_(teacherId, evaluator) {
  const ap = getAppointmentsData()[teacherId] || {};
  if (evaluator === 1) return getSignatureDataUri_(ap, 'c1');
  if (evaluator === 2) return getSignatureDataUri_(ap, 'c2');
  if (evaluator === 3) return getSignatureDataUri_(ap, 'c3');
  return '';
}

function getSummaryForPeriod_(allEvaluations, teacherId, period) {
  let passCount = 0;
  let rows = [];
  for (let evaluator = 1; evaluator <= 3; evaluator++) {
    const key = teacherId + '_' + period + '_' + evaluator;
    const ev = allEvaluations[key];
    if (!ev) {
      rows.push({ evaluator: evaluator, label: 'รอประเมิน', proCount: '-', socCount: '-', perCount: '-' });
      continue;
    }
    const score = getEvalScore_(ev, period);
    if (score.passed) passCount++;
    rows.push({
      evaluator: evaluator,
      label: score.label,
      proCount: score.proCount,
      socCount: score.socCount,
      perCount: score.perCount
    });
  }
  return {
    rows: rows,
    final: passCount >= 2 ? 'ผ่าน' : (rows.some(function(row) { return row.label === 'รอประเมิน'; }) ? 'รอผล' : 'ไม่ผ่าน')
  };
}

function buildFinalEvaluationSummaryHtml_(allEvaluations, teacher, teacherId) {
  let bodyRows = '';
  for (let period = 1; period <= 4; period++) {
    const summary = getSummaryForPeriod_(allEvaluations, teacherId, period);
    summary.rows.forEach(function(row) {
      bodyRows += '<tr><td>ครั้งที่ ' + period + '</td><td>กรรมการคนที่ ' + row.evaluator + '</td><td>' + row.proCount + '</td><td>' + row.socCount + '</td><td>' + row.perCount + '</td><td>' + row.label + '</td><td>' + summary.final + '</td></tr>';
    });
  }
  const finalRound = getSummaryForPeriod_(allEvaluations, teacherId, 4);
  const ap = getAppointmentsData()[teacherId] || {};
  const signerName = escapeHtml_(ap.c1 || 'ผู้อำนวยการโรงเรียน');
  const signature = getSignatureDataUri_(ap, 'c1');
  const signatureHtml = signature ? '<img class="signature-img" src="' + signature + '" alt="ลายมือชื่อ ผอ.รร.">' : '<br><br>';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>@page{size:A4;margin:2cm}body{font-family:"Sarabun","TH Sarabun New",Arial,sans-serif;font-size:13pt;color:#000;line-height:1.4}h1,h2{text-align:center;margin:0}h1{font-size:18pt}h2{font-size:15pt;margin-bottom:16px}.info{margin:16px 0}.info div{margin:4px 0}table{width:100%;border-collapse:collapse;margin-top:12px}th,td{border:1px solid #000;padding:5px;text-align:center}th{font-weight:700}.result{font-size:18pt;font-weight:700;text-align:center;margin-top:18px}.signature{margin-top:40px;text-align:center;margin-left:55%;line-height:1.45}.signature-img{display:block;max-width:150px;max-height:70px;object-fit:contain;margin:6px auto 2px}</style>' +
    '</head><body>' +
    '<h1>สรุปผลคะแนนการประเมินรอบสุดท้าย</h1>' +
    '<h2>การเตรียมความพร้อมและพัฒนาอย่างเข้ม ตำแหน่งครูผู้ช่วย</h2>' +
    '<div class="info"><div><strong>ชื่อ-สกุล:</strong> ' + escapeHtml_(teacher.name || '-') + '</div>' +
    '<div><strong>สถานศึกษา:</strong> ' + escapeHtml_(teacher.school || '-') + '</div>' +
    '<div><strong>วันที่บันทึก:</strong> ' + formatThaiLongDate_(new Date()) + '</div></div>' +
    '<table><thead><tr><th>รอบ</th><th>ผู้ประเมิน</th><th>วิชาชีพ</th><th>สังคม</th><th>คุณลักษณะ</th><th>ผลผู้ประเมิน</th><th>ผลรอบ</th></tr></thead><tbody>' + bodyRows + '</tbody></table>' +
    '<div class="result">ผลสรุปรอบสุดท้าย: ' + finalRound.final + '</div>' +
    '<div class="signature">' + signatureHtml + '(' + signerName + ')<br>ผู้รับรองผลการประเมิน</div>' +
    '</body></html>';
}
