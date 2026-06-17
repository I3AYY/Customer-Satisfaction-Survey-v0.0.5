/**
 * ============================================================================
 * เว็บแอปพลิเคชันระบบประเมินความพึงพอใจผู้รับบริการ (Customer Satisfaction Survey)
 * เวอร์ชัน: 0.0.5 AI Core (MedTech Luminous Theme & Form Comment Integration)
 * กลุ่มงานพยาธิวิทยาคลินิกและเทคนิคการแพทย์ โรงพยาบาลมะเร็งสุราษฎร์ธานี
 * ============================================================================
 */

function doGet(e) {
  setupSheets();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบประเมินความพึงพอใจ - กลุ่มงานพยาธิวิทยาคลินิกและเทคนิคการแพทย์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ติดตั้งและตรวจสอบฐานข้อมูล (Google Sheets) อัตโนมัติ
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ชีตเก็บคำถาม (SurveyQuestions)
  let sqSheet = ss.getSheetByName('SurveyQuestions');
  if (!sqSheet) {
    sqSheet = ss.insertSheet('SurveyQuestions');
    sqSheet.appendRow(['ID', 'QuestionText', 'ImageURL', 'Status']);
    sqSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    
    const sampleQuestions = [
      [1, 'ห้องสะอาด เป็นระเบียบ', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Clean+Room', true],
      [2, 'รอคิวไม่นาน', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Fast+Queue', true],
      [3, 'ยืนยันตัวตนก่อนเจาะเลือด', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Identify+Patient', true],
      [4, 'แจ้งเวลารอผลตรวจชัดเจน', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Clear+TAT', true],
      [5, 'เจ้าหน้าที่พูดจาสุภาพ ให้คำแนะนำชัดเจน', 'https://placehold.co/800x500/F0F9FF/1E40AF?text=Polite+Staff', true]
    ];
    sampleQuestions.forEach(q => sqSheet.appendRow(q));
    sqSheet.setFrozenRows(1);
    sqSheet.autoResizeColumns(1, 4);
  }

  // 2. ชีตเก็บการตั้งค่าและโลโก้ (Logo)
  let logoSheet = ss.getSheetByName('Logo');
  if (!logoSheet) {
    logoSheet = ss.insertSheet('Logo');
    logoSheet.getRange('A1:B1').setValues([['Setting', 'Value']]).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    logoSheet.getRange('A2:B3').setValues([
      // ใช้ Thumbnail API เพื่อแก้ปัญหาภาพไม่แสดง
      ['Logo URL', 'https://drive.google.com/thumbnail?id=1NvjW7zbszUCWPhBcH4th6uVsYqjACdED&sz=w800'],
      ['QR Code URL', 'https://placehold.co/200x200/ffffff/1e3a8a?text=Scan+to+Comment']
    ]);
    logoSheet.setFrozenRows(1);
    logoSheet.autoResizeColumns(1, 2);
  } else {
    const data = logoSheet.getDataRange().getValues();
    let hasQR = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'QR Code URL') hasQR = true;
    }
    if (!hasQR) {
      logoSheet.appendRow(['QR Code URL', 'https://placehold.co/200x200/ffffff/1e3a8a?text=Scan+to+Comment']);
    }
  }

  // 3. ชีตเก็บคำตอบ (Responses)
  let respSheet = ss.getSheetByName('Responses');
  if (!respSheet) {
    respSheet = ss.insertSheet('Responses');
    const headers = ['Timestamp', 'Year', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5'];
    respSheet.appendRow(headers);
    respSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    respSheet.setFrozenRows(1);
    
    const range = respSheet.getRange("C2:G2000");
    const ruleSatisfied = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('พอใจ').setBackground('#d1fae5').setFontColor('#065f46').setRanges([range]).build();
    const ruleDissatisfied = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ไม่พอใจ').setBackground('#ffe4e6').setFontColor('#9f1239').setRanges([range]).build();

    const rules = respSheet.getConditionalFormatRules();
    rules.push(ruleSatisfied, ruleDissatisfied);
    respSheet.setConditionalFormatRules(rules);
    respSheet.autoResizeColumns(1, 7);
  }
}

/**
 * ดึงข้อมูลเริ่มต้น (Questions, Settings, Years)
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ดึงคำถาม
  const sqSheet = ss.getSheetByName('SurveyQuestions');
  const sqData = sqSheet.getDataRange().getValues();
  const activeQuestions = [];
  for (let i = 1; i < sqData.length; i++) {
    if (sqData[i][3] === true || sqData[i][3].toString().toUpperCase() === 'TRUE') {
      activeQuestions.push({ id: sqData[i][0], text: sqData[i][1], imageUrl: sqData[i][2] });
    }
  }

  // 2. ดึงการตั้งค่า Logo และ QR
  const logoSheet = ss.getSheetByName('Logo');
  const logoData = logoSheet.getDataRange().getValues();
  let logoUrl = '';
  let qrUrl = '';
  for (let i = 1; i < logoData.length; i++) {
    if (logoData[i][0] === 'Logo URL') logoUrl = logoData[i][1].toString().trim();
    if (logoData[i][0] === 'QR Code URL') qrUrl = logoData[i][1].toString().trim();
  }

  // 3. ดึงปีที่มีคนตอบ 
  const yearSet = new Set();
  const respSheet = ss.getSheetByName('Responses');
  if (respSheet && respSheet.getLastRow() > 1) {
    respSheet.getRange(2, 2, respSheet.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) yearSet.add(r[0].toString()); });
  }

  // รองรับการค้นหาชีตของ Google Form (ชื่อ "การตอบแบบฟอร์ม..." หรือ "Form Responses...")
  let formSheet = null;
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sName = sheets[i].getName();
    if (sName.includes('การตอบแบบฟอร์ม') || sName.includes('Form Responses') || sName === 'AdditionalComments') {
      formSheet = sheets[i];
      break;
    }
  }

  if (formSheet && formSheet.getLastRow() > 1) {
    const isGoogleForm = formSheet.getName().includes('การตอบแบบฟอร์ม') || formSheet.getName().includes('Form Responses');
    if (isGoogleForm) {
      // แปลง Timestamp คอลัมน์แรกให้เป็นปี พ.ศ. อัตโนมัติ
      formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 1).getValues().forEach(r => {
        if (r[0]) {
           try {
             const d = new Date(r[0]);
             if(!isNaN(d.getTime())) yearSet.add((d.getFullYear() + 543).toString());
           } catch(e) {}
        }
      });
    } else {
      // โครงสร้างตาราง AdditionalComments เดิมที่มีคอลัมน์ Year ไว้แล้ว
      formSheet.getRange(2, 2, formSheet.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) yearSet.add(r[0].toString()); });
    }
  }
  
  const availableYears = Array.from(yearSet).sort((a, b) => b - a);
  const currentYearBE = (new Date().getFullYear() + 543).toString();
  
  return { questions: activeQuestions, logoUrl: logoUrl, qrUrl: qrUrl, availableYears: availableYears, currentYear: currentYearBE };
}

/**
 * บันทึกคำตอบแบบประเมินลงตาราง Responses
 */
function submitSurveyResponse(answers) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error('ระบบกำลังประมวลผลหนาแน่น กรุณาลองกดส่งใหม่อีกครั้ง'); }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Responses');
    const now = new Date();
    const timestamp = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    const yearBE = (now.getFullYear() + 543).toString();
    
    const rowData = [timestamp, yearBE];
    for (let i = 0; i < 5; i++) rowData.push(answers[i] ? answers[i] : "");
    sheet.appendRow(rowData);
    
    lock.releaseLock();
    return { success: true };
  } catch (error) {
    lock.releaseLock();
    throw new Error("เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + error.toString());
  }
}

/**
 * ดึงข้อมูลสถิติและคอมเมนต์สำหรับหน้า Dashboard
 */
function getDashboardStats(filterYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const stats = {
    totalRespondents: 0,
    overall: { satisfied: 0, dissatisfied: 0 },
    questions: {
      Q1: { satisfied: 0, dissatisfied: 0 }, Q2: { satisfied: 0, dissatisfied: 0 },
      Q3: { satisfied: 0, dissatisfied: 0 }, Q4: { satisfied: 0, dissatisfied: 0 },
      Q5: { satisfied: 0, dissatisfied: 0 }
    },
    comments: []
  };
  
  // 1. ประมวลผลคำตอบหลัก
  const sheet = ss.getSheetByName('Responses');
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    data.forEach(row => {
      const rowYear = row[1] ? row[1].toString() : '';
      if (filterYear === 'ทั้งหมด' || filterYear === rowYear) {
        stats.totalRespondents++;
        
        const qKeys = ['Q1', 'Q2', 'Q3', 'Q4', 'Q5'];
        for(let i=0; i<5; i++){
           const ans = row[i+2]; // คอลัมน์ C ถึง G
           if (ans === 'พอใจ') {
             stats.questions[qKeys[i]].satisfied++;
             stats.overall.satisfied++;
           } else if (ans === 'ไม่พอใจ') {
             stats.questions[qKeys[i]].dissatisfied++;
             stats.overall.dissatisfied++;
           }
        }
      }
    });
  }
  
  // 2. ดึงคอมเมนต์ (รองรับชีตอัตโนมัติจาก Google Form)
  let commSheet = null;
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sName = sheets[i].getName();
    if (sName.includes('การตอบแบบฟอร์ม') || sName.includes('Form Responses') || sName === 'AdditionalComments') {
      commSheet = sheets[i];
      break;
    }
  }
  
  if (commSheet) {
    const cLastRow = commSheet.getLastRow();
    const isGoogleForm = commSheet.getName().includes('การตอบแบบฟอร์ม') || commSheet.getName().includes('Form Responses');

    if (cLastRow > 1) {
      // ดึงข้อมูล 3 คอลัมน์แรก (เผื่อเคสฟอร์มมีคำถามเดียว จะได้ดึงข้อความจากคอลัมน์ B)
      const numCols = isGoogleForm ? commSheet.getLastColumn() : 3;
      const cData = commSheet.getRange(2, 1, cLastRow - 1, Math.max(numCols, 2)).getValues();
      
      cData.forEach(r => {
        const timestamp = r[0];
        let cYear = '';
        let commentText = '';

        if (isGoogleForm) {
           // คำนวณปีจาก Timestamp ของฟอร์ม
           try {
             const dateObj = new Date(timestamp);
             if(!isNaN(dateObj.getTime())) {
               cYear = (dateObj.getFullYear() + 543).toString();
             }
           } catch(e) {}
           commentText = r[1] ? r[1].toString() : ''; // คำตอบแบบฟอร์มมักอยู่คอลัมน์ B
        } else {
           cYear = r[1] ? r[1].toString() : '';
           commentText = r[2] ? r[2].toString() : '';
        }

        if (filterYear === 'ทั้งหมด' || filterYear === cYear) {
           if (commentText) { // ตรวจสอบว่าไม่เป็นช่องว่าง
             let formattedTime = "";
             try {
               formattedTime = Utilities.formatDate(new Date(timestamp), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
             } catch(e) { formattedTime = timestamp.toString(); }
             
             stats.comments.push({ timestamp: formattedTime, text: commentText });
           }
        }
      });
    }
  }
  
  // จัดเรียงคอมเมนต์ล่าสุดขึ้นก่อนเสมอ
  stats.comments.reverse();
  
  return stats;
}
