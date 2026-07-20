// ==========================================
// Google Apps Script (Code.gs)
// สำหรับระบบเช็คชื่อนักเรียนเข้าแถว
// ==========================================

// กำหนด Google Sheet ID ของคุณที่นี่ (เพื่อแก้ปัญหา getActiveSpreadsheet หาไฟล์ไม่เจอตอนรันเป็น Web App)
const SHEET_ID = '1vXHZsIts2xUVkpGswOFgVA475QBnCjOXmRxx0fdFhRg';

function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Create sheets if they don't exist
  const sheets = ['Classrooms', 'Students', 'Attendance'];
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      const sheet = ss.insertSheet(name);
      // Setup headers
      if (name === 'Classrooms') {
        sheet.appendRow(['id', 'name', 'level', 'academicYear']);
      } else if (name === 'Students') {
        sheet.appendRow(['id', 'studentId', 'title', 'firstName', 'lastName', 'classroomId', 'status']);
      } else if (name === 'Attendance') {
        sheet.appendRow(['id', 'date', 'classroomId', 'studentId', 'status', 'recordedBy', 'recordedAt']);
      }
    }
  });
}

// ------------------------------------------
// 1. HTTP GET - ดึงข้อมูลจาก Sheets ไปยัง React (Pull)
// ------------------------------------------
function doGet(e) {
  try {
    // ใช้ openById แทน getActiveSpreadsheet
    const targetId = (e && e.parameter && e.parameter.sheetId) ? e.parameter.sheetId : SHEET_ID;
    const ss = SpreadsheetApp.openById(targetId);
    
    const data = {
      classrooms: getSheetDataAsObjects(ss, 'Classrooms'),
      students: getSheetDataAsObjects(ss, 'Students'),
      attendance: getSheetDataAsObjects(ss, 'Attendance')
    };

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: data
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------
// 2. HTTP POST - บันทึกข้อมูลจาก React ลง Sheets (Push)
// ------------------------------------------
function doPost(e) {
  try {
    // Parse incoming JSON
    const body = JSON.parse(e.postData.contents);
    
    if (body.action === 'syncAll') {
      const targetId = body.sheetId ? body.sheetId : SHEET_ID;
      const ss = SpreadsheetApp.openById(targetId);
      const payload = body.payload;
      
      if (payload.classrooms) {
        writeDataToSheet(ss, 'Classrooms', payload.classrooms, ['id', 'name', 'level', 'academicYear']);
      }
      if (payload.students) {
        writeDataToSheet(ss, 'Students', payload.students, ['id', 'studentId', 'title', 'firstName', 'lastName', 'classroomId', 'status']);
      }
      if (payload.attendance) {
        writeDataToSheet(ss, 'Attendance', payload.attendance, ['id', 'date', 'classroomId', 'studentId', 'status', 'recordedBy', 'recordedAt']);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Data synchronized successfully'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    throw new Error('Unknown action: ' + body.action);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------
// Helper Functions
// ------------------------------------------

// อ่านข้อมูลจาก Sheet แปลงเป็น Array of Objects
function getSheetDataAsObjects(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Empty or only headers
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] !== undefined ? row[index] : '';
    });
    return obj;
  });
}

// เขียนข้อมูล Array of Objects ลง Sheet (ล้างข้อมูลเก่าออกหมด แล้วเขียนใหม่)
function writeDataToSheet(ss, sheetName, dataArray, headers) {
  let sheet = ss.getSheetByName(sheetName);
  
  // ถ้ายังไม่มี Sheet ให้สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // ล้างข้อมูลเก่า
  sheet.clear();
  
  // พิมพ์หัวคอลัมน์
  sheet.appendRow(headers);
  
  if (!dataArray || dataArray.length === 0) return;
  
  // สร้าง Array 2 มิติ สำหรับข้อมูลแต่ละแถว
  const rows = dataArray.map(obj => {
    return headers.map(header => {
      // ป้องกันเรื่อง Object ถูกแปลงเป็น string '[object Object]' ให้เก็บเป็น string เปล่าถ้าไม่มีค่า
      return obj[header] !== undefined && obj[header] !== null ? obj[header] : '';
    });
  });
  
  // เขียนลง Sheet รวดเดียว (ทำงานเร็วมาก)
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}
