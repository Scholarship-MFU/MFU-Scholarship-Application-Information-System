function doGet(e) {
  try {
    // ใช้ HtmlOutput เพื่อแก้ปัญหาเปิดใน Facebook แล้วขึ้น undefined
    var html = HtmlService.createHtmlOutputFromFile('Index');
    
    html.setTitle('ระบบข้อมูลการรับสมัครทุนการศึกษา มหาวิทยาลัยแม่ฟ้าหลวง');
    html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
    return html;
  } catch (error) {
    return HtmlService.createHtmlOutput("Error: ไม่สามารถโหลดหน้าเว็บได้ - " + error.toString());
  }
}

function getScholarshipData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Scholarships');
  
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // ดึง Timezone ของ Script (เช่น Asia/Bangkok) เพื่อใช้อ้างอิง
  const timeZone = Session.getScriptTimeZone();

  const scholarshipList = rows.map((row) => {
    let item = {};
    headers.forEach((header, index) => {
      if (header === 'years' || header === 'tags') {
         item[header] = row[index] ? row[index].toString().split(',').map(s => s.trim()) : [];
      } else if (header === 'deadline') {
         // --- แก้ไขตรงนี้ ---
         // ใช้ Utilities.formatDate เพื่อคงวันที่ตาม Timezone เดิม ไม่แปลงเป็น UTC
         if (row[index] instanceof Date) {
           item[header] = Utilities.formatDate(row[index], timeZone, "yyyy-MM-dd");
         } else {
           item[header] = row[index] ? String(row[index]) : '';
         }
      } else {
         item[header] = row[index];
      }
    });
    return item;
  });

  return scholarshipList;
}

// ฟังก์ชันสำหรับนับและดึงยอดผู้เข้าชม
function incrementAndGetVisitorCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');
  
  // ถ้ายังไม่มีชีต 'Settings' ให้สร้างใหม่และตั้งค่าเริ่มต้น
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.getRange('A1').setValue('Visitor Count');
    sheet.getRange('B1').setValue(0); // เริ่มต้นที่ 0
  }
  
  const countRange = sheet.getRange('B1');
  let currentCount = countRange.getValue();
  
  // ตรวจสอบความถูกต้องของตัวเลข
  if (isNaN(currentCount) || currentCount === '') {
    currentCount = 0;
  }
  
  const newCount = currentCount + 1; // บวกเพิ่ม 1 ทุกครั้งที่มีการรันฟังก์ชัน
  countRange.setValue(newCount); // บันทึกค่าใหม่ลง Sheet
  
  return newCount;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
