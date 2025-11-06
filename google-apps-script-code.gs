/**
 * ระบบบันทึกน้ำหนัก-ส่วนสูงออนไลน์
 * Google Apps Script Backend
 * By ครูเปิงมางfc
 *
 * คำแนะนำการติดตั้ง:
 * 1. เปิด Google Sheets ของคุณ
 * 2. ไปที่ Extensions > Apps Script
 * 3. คัดลอกโค้ดทั้งหมดนี้ไปวางในไฟล์ Code.gs
 * 4. กด Deploy > New deployment
 * 5. เลือก Web app
 * 6. Execute as: Me
 * 7. Who has access: Anyone
 * 8. กด Deploy และคัดลอก URL ที่ได้
 * 9. นำ URL ไปใส่ในตัวแปร GOOGLE_APPS_SCRIPT_URL ในไฟล์ HTML
 */

// ชื่อ Sheet ที่จะใช้เก็บข้อมูล
const SHEET_NAME = 'ข้อมูลนักเรียน';

/**
 * ฟังก์ชันหลักสำหรับรับคำขอ GET
 */
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'load') {
      return loadStudents();
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Invalid action'
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.toString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันหลักสำหรับรับคำขอ POST
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'save') {
      return saveStudents(data.data);
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Invalid action'
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.toString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ดึงข้อมูลนักเรียนทั้งหมดจาก Google Sheets
 */
function loadStudents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // ถ้ายังไม่มี Sheet ให้สร้างใหม่
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      initializeSheet(sheet);

      return ContentService.createTextOutput(
        JSON.stringify({
          success: true,
          students: [],
          message: 'Sheet created'
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();

    // ถ้ามีแค่ header หรือไม่มีข้อมูล
    if (data.length <= 1) {
      return ContentService.createTextOutput(
        JSON.stringify({
          success: true,
          students: []
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // แปลงข้อมูลเป็น JSON
    const headers = data[0];
    const students = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // ข้ามแถวที่ว่าง
      if (!row[0]) continue;

      students.push({
        studentId: row[0].toString(),
        fullName: row[1],
        grade: row[2],
        gender: row[3],
        weight: parseFloat(row[4]),
        height: parseFloat(row[5]),
        age: parseInt(row[6]),
        bmi: parseFloat(row[7]),
        status: row[8]
      });
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        success: true,
        students: students,
        count: students.length
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.toString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * บันทึกข้อมูลนักเรียนลง Google Sheets
 */
function saveStudents(students) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // ถ้ายังไม่มี Sheet ให้สร้างใหม่
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      initializeSheet(sheet);
    } else {
      // ล้างข้อมูลเก่า (เก็บไว้แค่ header)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
    }

    // เตรียมข้อมูลสำหรับเขียนลง Sheet
    const dataToWrite = students.map(student => [
      student.studentId,
      student.fullName,
      student.grade,
      student.gender,
      student.weight,
      student.height,
      student.age,
      student.bmi,
      student.status,
      new Date().toLocaleString('th-TH')
    ]);

    // เขียนข้อมูลลง Sheet
    if (dataToWrite.length > 0) {
      sheet.getRange(2, 1, dataToWrite.length, 10).setValues(dataToWrite);

      // Format ตาราง
      formatSheet(sheet, dataToWrite.length + 1);
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        success: true,
        message: 'Data saved successfully',
        count: dataToWrite.length
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: error.toString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * สร้าง Header สำหรับ Sheet
 */
function initializeSheet(sheet) {
  const headers = [
    'รหัสนักเรียน',
    'ชื่อ-สกุล',
    'ชั้น',
    'เพศ',
    'น้ำหนัก (kg)',
    'ส่วนสูง (cm)',
    'อายุ (ปี)',
    'BMI',
    'สถานะ',
    'วันที่บันทึก'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#64B5F6');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // ตั้งค่าความกว้างคอลัมน์
  sheet.setColumnWidth(1, 120);  // รหัสนักเรียน
  sheet.setColumnWidth(2, 200);  // ชื่อ-สกุล
  sheet.setColumnWidth(3, 80);   // ชั้น
  sheet.setColumnWidth(4, 60);   // เพศ
  sheet.setColumnWidth(5, 100);  // น้ำหนัก
  sheet.setColumnWidth(6, 100);  // ส่วนสูง
  sheet.setColumnWidth(7, 80);   // อายุ
  sheet.setColumnWidth(8, 80);   // BMI
  sheet.setColumnWidth(9, 120);  // สถานะ
  sheet.setColumnWidth(10, 180); // วันที่บันทึก

  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Format ตาราง
 */
function formatSheet(sheet, lastRow) {
  // Format all data cells
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
  dataRange.setHorizontalAlignment('center');
  dataRange.setVerticalAlignment('middle');

  // Alternate row colors
  for (let i = 2; i <= lastRow; i++) {
    const rowRange = sheet.getRange(i, 1, 1, 10);
    if (i % 2 === 0) {
      rowRange.setBackground('#F5F5F5');
    } else {
      rowRange.setBackground('#FFFFFF');
    }
  }

  // Format status column with colors
  for (let i = 2; i <= lastRow; i++) {
    const statusCell = sheet.getRange(i, 9);
    const status = statusCell.getValue();

    switch (status) {
      case 'ปกติ':
        statusCell.setBackground('#C8E6C9');
        statusCell.setFontColor('#2E7D32');
        break;
      case 'ต่ำกว่าเกณฑ์':
        statusCell.setBackground('#FFECB3');
        statusCell.setFontColor('#F57F17');
        break;
      case 'น้ำหนักเกิน':
        statusCell.setBackground('#FFCCBC');
        statusCell.setFontColor('#D84315');
        break;
      case 'อ้วน':
        statusCell.setBackground('#FFCDD2');
        statusCell.setFontColor('#C62828');
        break;
    }
    statusCell.setFontWeight('bold');
  }

  // Add borders
  const allRange = sheet.getRange(1, 1, lastRow, 10);
  allRange.setBorder(true, true, true, true, true, true);
}

/**
 * ฟังก์ชันสำหรับทดสอบ (สามารถรันได้จาก Apps Script Editor)
 */
function testInitialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  initializeSheet(sheet);
  Logger.log('Sheet initialized successfully');
}

/**
 * ฟังก์ชันสำหรับสร้างข้อมูลทดสอบ
 */
function createSampleData() {
  const sampleStudents = [
    {
      studentId: '60001',
      fullName: 'สมชาย ใจดี',
      grade: 'ม.1/1',
      gender: 'ชาย',
      weight: 45.5,
      height: 165,
      age: 13,
      bmi: '16.71',
      status: 'ต่ำกว่าเกณฑ์'
    },
    {
      studentId: '60002',
      fullName: 'สมหญิง สวยงาม',
      grade: 'ม.1/1',
      gender: 'หญิง',
      weight: 50,
      height: 160,
      age: 13,
      bmi: '19.53',
      status: 'ปกติ'
    },
    {
      studentId: '60003',
      fullName: 'สมศักดิ์ แข็งแรง',
      grade: 'ม.2/1',
      gender: 'ชาย',
      weight: 70,
      height: 170,
      age: 14,
      bmi: '24.22',
      status: 'น้ำหนักเกิน'
    }
  ];

  saveStudents(sampleStudents);
  Logger.log('Sample data created successfully');
}

/**
 * ฟังก์ชันสำหรับแสดง URL ของ Web App
 */
function getWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  Logger.log('Web App URL: ' + url);
  return url;
}
