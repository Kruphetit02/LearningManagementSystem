// ====================================
// อ้างอิง : https://www.youtube.com/@Pormodtanoy/videos
// ====================================
const SPREADSHEET_ID = '1VInU80stIVBeRx-B4oqBo50YMVJPI0VK65cJgcAaZuA'; // แก้ไขจุดที่ 1 : ใส่ ID spreadsheet ของท่าน 
const FOLDER_ID = '1k82WN6jYWBH1qfNYW1J4LSA6Gf-gOoSB'; // แก้ไขจุดที่ 2 :  ใส่ ID Folder  สำหรับเก็บไฟล์

// Sheet Names
const SHEETS = {
  SYLLABUS: 'แผนการจัดการเรียนรู้',
  USERS: 'ผู้ใช้งาน',
  LOGS: 'ประวัติการใช้งาน'
};

// ====================================
// Main Functions
// ====================================
function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ตรวจสอบว่าสามารถเข้าถึง spreadsheet ได้หรือไม่
    const sheets = ss.getSheets();
    if (sheets.length === 0) {
      Utilities.sleep(1000); // รอ 1 วินาที
      SpreadsheetApp.flush();
      throw new Error('ไม่สามารถอ่าน sheet ได้ กรุณาตรวจสอบสิทธิ์');
    }
    
    if (!ss.getSheetByName(SHEETS.USERS)) {
      initializeSheets();
    }
    
    const output = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบบริหารจัดการแผนการเรียนรู้ออนไลน์')
      .setFaviconUrl('https://img1.pic.in.th/images/logoc577678dae0d285f.png') // แก้ไขจุดที่ 3 : เปลี่ยนภาพ favicon ของท่านโดยฝากภาพโลโกไว้บนเว็บฝากภาพฟรี https://pic.in.th/?lang=th
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
    return output;
    
  } catch (error) {
    Logger.log('doGet Error: ' + error.toString());
    Logger.log('Error details: ' + JSON.stringify(error));
    
    // สร้างหน้าแจ้งเตือนข้อผิดพลาด
    return HtmlService.createHtmlOutput(`
      <div style="display: flex; height: 100vh; align-items: center; justify-content: center; text-align: center; font-family: 'Sarabun', Arial, sans-serif;">
        <div>
          <h2>ไม่สามารถเชื่อมต่อกับระบบได้</h2>
          <p>เกิดข้อผิดพลาดในการเชื่อมต่อ</p>
          <p>กรุณาติดต่อผู้ดูแลระบบ</p>
          <p style="font-size: 0.8em; color: #666;">Error: ${error.message || 'Unknown error'}</p>
        </div>
      </div>
    `);
  }
}

// ====================================
// Initialize Spreadsheet
// ====================================
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Create Syllabus Sheet
    let syllabusSheet = ss.getSheetByName(SHEETS.SYLLABUS);
    if (!syllabusSheet) {
      syllabusSheet = ss.insertSheet(SHEETS.SYLLABUS);
      syllabusSheet.getRange('A1:Q1').setValues([[
        'ID', 'รหัสวิชา', 'ชื่อวิชา', 'ครูผู้สอน', 'ภาคเรียน',
        'ปีการศึกษา', 'กลุ่มสาระ', 'สถานะ', 'วันที่ส่ง',
        'หมายเหตุ', 'เวอร์ชัน', 'ไฟล์ ID', 'ไฟล์ URL',
        'หน่วยที่', 'ทั้งหมดกี่หน่วย', 'แผนที่', 'ทั้งหมดกี่แผน' // คอลัมน์ที่เพิ่มใหม่
      ]]);
      syllabusSheet.getRange('A1:Q1').setFontWeight('bold')
        .setBackground('#4a90e2').setFontColor('#ffffff');
      syllabusSheet.setFrozenRows(1);
    }

    // บังคับคอลัมน์ F (ปีการศึกษา) ทั้งคอลัมน์เป็น plain text
    // ป้องกัน Sheets แปลงค่าเป็น Date อัตโนมัติ
    syllabusSheet.getRange('F:F').setNumberFormat('@STRING@');

    // Create Users Sheet
    let usersSheet = ss.getSheetByName(SHEETS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(SHEETS.USERS);
      usersSheet.getRange('A1:H1').setValues([[
        'Username', 'Password', 'ชื่อ-นามสกุล', 'อีเมล', 'กลุ่มสาระ', 'Role', 'Active', 'วันที่สมัคร'
      ]]);
      usersSheet.getRange('A1:H1').setFontWeight('bold')
        .setBackground('#6b8e23').setFontColor('#ffffff');
      usersSheet.setFrozenRows(1);
    }

    // Create Logs Sheet
    let logsSheet = ss.getSheetByName(SHEETS.LOGS);
    if (!logsSheet) {
      logsSheet = ss.insertSheet(SHEETS.LOGS);
      logsSheet.getRange('A1:E1').setValues([[
        'Timestamp', 'User', 'Action', 'Details', 'IP Address'
      ]]);
      logsSheet.getRange('A1:E1').setFontWeight('bold')
        .setBackground('#9b59b6').setFontColor('#ffffff');
      logsSheet.setFrozenRows(1);
    }

    return { success: true, message: 'Sheets initialized successfully' };
  } catch (error) {
    Logger.log('initializeSheets Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


function fixAcademicYearColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('แผนการสอน');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('No data to fix');
    return;
  }

  // Set format คอลัมน์ F เป็น plain text ก่อน
  sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('@STRING@');

  const values = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  const fixed = values.map(row => {
    const val = row[0];
    if (val instanceof Date) {
      return [Utilities.formatDate(val, 'GMT+7', 'yyyy')];
    }
    return [val ? val.toString().trim() : ''];
  });

  sheet.getRange(2, 6, lastRow - 1, 1).setValues(fixed);
  Logger.log('Fixed ' + fixed.length + ' rows in academic_year column');
}


// ==================================== 
// 1. แก้ไขฟังก์ชัน getAllSyllabus (ฝั่ง server)
// ====================================

function getAllSyllabus(userRole, userGroup) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues(); 

    // ฟังก์ชันช่วยจัดการกรณีเป็นวันที่ให้กลับมาเป็นข้อความ 1-5
    const cleanValue = (val) => {
      if (val instanceof Date) {
        return Utilities.formatDate(val, "GMT+7", "d-M");
      }
      return val ? val.toString().trim() : '';
    };

    const result = data
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => {
        return {
          id:             row[0]  ? row[0].toString()        : '',
          course_code:    row[1]  ? row[1].toString()        : '',
          course_name:    row[2]  ? row[2].toString()        : '',
          instructor:     row[3]  ? row[3].toString()        : '',
          semester:       row[4]  ? row[4].toString()        : '',
          academic_year:  cleanValue(row[5]),
          group:          row[6]  ? row[6].toString()        : '',
          status:         row[7]  ? row[7].toString()        : 'pending',
          submitted_date: row[8]  ? row[8].toString()        : '',
          reviewer_notes: row[9]  ? row[9].toString()        : '',
          version:        row[10] ? row[10].toString()       : '1',
          file_id:        row[11] ? row[11].toString()       : '',
          file_url:       row[12] ? row[12].toString()       : '',
          unit_number:    cleanValue(row[13]),
          total_units:    cleanValue(row[14]),
          lesson_number:  cleanValue(row[15]),
          total_lessons:  cleanValue(row[16])
        };
      });

    if (userRole === 'admin') return result;
    else if (userGroup) return result.filter(item => item.group === userGroup);
    else return result;

  } catch (error) {
    return [];
  }
}

// เพิ่มฟังก์ชันนี้เพื่อ debug
function testGetAllSyllabus() {
  const result = getAllSyllabus();
  Logger.log('Test result count: ' + result.length);
  Logger.log('Test result: ' + JSON.stringify(result));
  return result;
}

// ==================================== 
// 2. ฟังก์ชันบันทึกข้อมูลใหม่ (ใส่ ' นำหน้าป้องกัน Sheet แปลงเป็นวันที่)
// ====================================
function createSyllabus(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    if (!sheet) throw new Error('ไม่พบชีตแผนการจัดการเรียนรู้');
    
    const id = data.id || Utilities.getUuid();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([
      id,
      data.course_code || '',
      data.course_name || '',
      data.instructor || '',
      data.semester || '',
      "'" + (data.academic_year || ''), // บังคับเป็นข้อความ
      data.group || '',
      data.status || 'pending',
      timestamp,
      data.reviewer_notes || '',
      data.version || 1,
      data.file_id || '',
      data.file_url || '',
      "'" + (data.unit_number || ''),   // บังคับเป็นข้อความ
      "'" + (data.total_units || ''),
      "'" + (data.lesson_number || ''), // บังคับเป็นข้อความ (1-5 จะไม่เพี้ยน)
      "'" + (data.total_lessons || '')
    ]);
    
    // แจ้งเตือน Telegram
    try {
      if ((data.status || 'pending') === 'pending') {
        sendTelegramNotification(data);
      }
    } catch (e) {}

    return { success: true, data: { id: id, course_code: data.course_code } };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ====================================
// updateSyllabus (ปรับแล้ว — Telegram ส่งเฉพาะยืนยันส่ง)
// ====================================

// ==================================== 
// 3. ฟังก์ชันอัปเดตข้อมูล (รักษาเครื่องหมาย ' ไว้)
// ====================================
function updateSyllabus(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    const searchId = data.id ? data.id.toString().trim() : '';
    
    let rowIndex = -1, existingData = null;
    for (let i = 0; i < dataRange.length; i++) {
      if (dataRange[i][0]?.toString().trim() === searchId) {
        rowIndex = i;
        existingData = dataRange[i];
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: 'ไม่พบข้อมูล' };
    
    const actualRow = rowIndex + 2;
    const formattedDate = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    const newStatus = data.status || existingData[7] || 'pending';
    
    const updateValues = [[
      data.course_code || existingData[1],
      data.course_name || existingData[2],
      data.instructor || existingData[3],
      data.semester || existingData[4],
      "'" + (data.academic_year || existingData[5]),
      data.group || existingData[6],
      newStatus,
      formattedDate,
      data.reviewer_notes || existingData[9],
      data.version || existingData[10],
      data.file_id || existingData[11],
      data.file_url || existingData[12],
      "'" + (data.unit_number || existingData[13]),
      "'" + (data.total_units || existingData[14]),
      "'" + (data.lesson_number || existingData[15]),
      "'" + (data.total_lessons || existingData[16])
    ]];
    
    sheet.getRange(actualRow, 2, 1, 16).setValues(updateValues);

    // แจ้งเตือน Email & Telegram (ถ้ามีการส่งงาน)
    try {
      if (data.originalStatus === 'not_submitted' && newStatus === 'pending') {
        sendTelegramNotification(data);
      }
      const userEmail = findUserEmail(data.instructor);
      if (userEmail) sendEmailNotification(data, userEmail, 'STATUS_CHANGE');
    } catch (e) {}

    return { success: true, message: 'อัปเดตข้อมูลสำเร็จ' };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}


function deleteSyllabus(id) {
  try {
    Logger.log('deleteSyllabus called with ID: ' + id);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    
    if (!sheet) {
      return { success: false, message: 'ไม่พบชีตแผนการจัดการเรียนรู้' };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: false, message: 'ไม่มีข้อมูลในชีต' };
    }
    
    // ดึงข้อมูล ID ทั้งหมด
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    Logger.log('Total rows to search: ' + dataRange.length);
    
    // แปลง ID ให้เป็น string และหา row ที่ตรงกัน
    const searchId = id ? id.toString().trim() : '';
    Logger.log('Searching for ID: "' + searchId + '"');
    
    let rowIndex = -1;
    for (let i = 0; i < dataRange.length; i++) {
      const cellId = dataRange[i][0] ? dataRange[i][0].toString().trim() : '';
      if (cellId === searchId) {
        rowIndex = i;
        Logger.log('Found match at index: ' + i + ' (row ' + (i + 2) + ')');
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log('ID not found. Available IDs: ' + dataRange.map(r => r[0]).join(', '));
      return { success: false, message: 'ไม่พบข้อมูล ID: ' + searchId };
    }
    
    const actualRow = rowIndex + 2;
    Logger.log('Deleting row: ' + actualRow);
    
    // ดึงข้อมูลก่อนลบเพื่อบันทึก log
    const rowData = sheet.getRange(actualRow, 1, 1, 13).getValues()[0];
    const courseCode = rowData[1] || 'Unknown';
    
    // ลบแถว
    sheet.deleteRow(actualRow);
    
    Logger.log('Delete successful for row: ' + actualRow);
    
    try {
      logActivity('DELETE', 'Deleted syllabus: ' + courseCode + ' (ID: ' + id + ')');
    } catch (logError) {
      Logger.log('Non-fatal logging error: ' + logError.toString());
    }
    
    return { success: true, message: 'ลบข้อมูลสำเร็จ' };
    
  } catch (error) {
    Logger.log('deleteSyllabus Error: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return { 
      success: false, 
      message: 'เกิดข้อผิดพลาด: ' + error.toString() 
    };
  }
}

// ====================================
// User Management with Role and Active Status
// ====================================
function registerUser(userData) {
  try {
    Logger.log('[registerUser] Starting registration...');
    Logger.log('[registerUser] Received data: ' + JSON.stringify(userData));
    
    // ตรวจสอบข้อมูลที่จำเป็น
    if (!userData || !userData.username || !userData.password || 
        !userData.fullname || !userData.email || !userData.group) {
      Logger.log('[registerUser] Missing required fields');
      return { 
        success: false, 
        message: 'ข้อมูลไม่ครบถ้วน กรุณากรอกข้อมูลให้ครบทุกช่อง' 
      };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.USERS);
    
    if (!sheet) {
      Logger.log('[registerUser] Users sheet not found');
      return { 
        success: false, 
        message: 'ไม่พบตารางผู้ใช้งาน กรุณาติดต่อผู้ดูแลระบบ' 
      };
    }

    const lastRow = sheet.getLastRow();
    Logger.log('[registerUser] Last row: ' + lastRow);
    
    // ตรวจสอบ username ซ้ำ
    if (lastRow >= 2) {
      const usernames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      const emails = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
      
      const usernameExists = usernames.some(row => 
        row[0] && row[0].toString().trim().toLowerCase() === userData.username.trim().toLowerCase()
      );
      
      if (usernameExists) {
        Logger.log('[registerUser] Username already exists');
        return { 
          success: false, 
          message: 'ชื่อผู้ใช้นี้มีผู้อื่นใช้งานอยู่แล้ว' 
        };
      }
      
      const emailExists = emails.some(row => 
        row[0] && row[0].toString().trim().toLowerCase() === userData.email.trim().toLowerCase()
      );
      
      if (emailExists) {
        Logger.log('[registerUser] Email already exists');
        return { 
          success: false, 
          message: 'อีเมลนี้ลงทะเบียนไว้แล้ว' 
        };
      }
    }

    // สร้าง timestamp
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    Logger.log('[registerUser] Adding user to sheet...');
    
    // เพิ่มข้อมูลลงในชีต
    sheet.appendRow([
      userData.username.trim(),
      userData.password,
      userData.fullname.trim(),
      userData.email.trim(),
      userData.group,
      'user',      // role
      'active',    // status
      timestamp    // registration date
    ]);
    
    Logger.log('[registerUser] User registered successfully: ' + userData.username);

    // บันทึก log
    try {
      logActivity('REGISTER', `สมัครสมาชิกใหม่: ${userData.username} (${userData.fullname})`);
    } catch (logError) {
      Logger.log('[registerUser] Non-fatal logging error: ' + logError.toString());
    }

    return { 
      success: true, 
      message: 'สมัครสมาชิกสำเร็จ',
      username: userData.username 
    };

  } catch (error) {
    Logger.log('[registerUser] FATAL ERROR: ' + error.toString());
    Logger.log('[registerUser] Stack trace: ' + error.stack);
    return { 
      success: false, 
      message: `เกิดข้อผิดพลาด: ${error.message || error.toString()}` 
    };
  }
}


function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.USERS);
    
    if (sheet.getLastRow() <= 1) {
      return { success: false, message: 'ไม่พบผู้ใช้ในระบบ' };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    
    const user = data.find(row => row[0] === username && row[1] === password);
    
    if (user) {
      // ตรวจสอบสถานะ active
      const activeStatus = user[6] ? user[6].toString().toLowerCase() : 'active';
      
      if (activeStatus !== 'active') {
        return { 
          success: false, 
          message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ' 
        };
      }
      
      logActivity('LOGIN', 'User logged in: ' + username);
      return {
        success: true,
        user: {
          username: user[0],
          fullname: user[2],
          email: user[3],
          group: user[4],
          role: user[5] || 'user' // ถ้าไม่มีให้เป็น user
        }
      };
    }
    
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (error) {
    Logger.log('loginUser Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


// ====================================
// File Upload
// ====================================
function uploadFile(fileData, fileName, mimeType) {
  try {
    Logger.log('uploadFile called with fileName: ' + fileName);
    
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file = folder.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    Logger.log('File uploaded successfully with ID: ' + file.getId());
    
    return {
      success: true,
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: file.getName()
    };
  } catch (error) {
    Logger.log('uploadFile Error: ' + error.toString());
    return { 
      success: false, 
      message: 'ไม่สามารถอัปโหลดไฟล์ได้: ' + error.toString() 
    };
  }
}


// ====================================
// PDF Generation — Fixed & Simplified
// ====================================
function getSyllabusById(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.SYLLABUS);

  const data = sheet.getDataRange().getValues();
  const header = data.shift();

  const found = data.find(r => r[0].toString() === id.toString());

  if (!found) {
    return { success: false, message: "ไม่พบข้อมูล" };
  }

  const item = {
    id: found[0],
    course_code: found[1],
    course_name: found[2],
    instructor: found[3],
    semester: found[4],
    academic_year: found[5],
    group: found[6],
    status: found[7],
    submitted_date: found[8],
    reviewer_notes: found[9]
  };

  return { success: true, item };
}


// ====================================
// Statistics
// ====================================
function getStatistics() {
  try {
    const data = getAllSyllabus();
    
    const stats = {
      total: data.length,
      pending: data.filter(s => s.status === 'pending').length,
      approved: data.filter(s => s.status === 'approved').length,
      revision: data.filter(s => s.status === 'revision').length,
      not_submitted: data.filter(s => s.status === 'not_submitted').length
    };
    
    const semesterStats = {};
    data.forEach(item => {
      const key = `${item.semester}/${item.academic_year}`;
      if (!semesterStats[key]) {
        semesterStats[key] = { pending: 0, approved: 0, revision: 0, not_submitted: 0 };
      }
      semesterStats[key][item.status] = (semesterStats[key][item.status] || 0) + 1;
    });
    
    const groupStats = {};
    data.forEach(item => {
      if (!groupStats[item.group]) {
        groupStats[item.group] = { pending: 0, approved: 0, revision: 0, not_submitted: 0 };
      }
      groupStats[item.group][item.status] = (groupStats[item.group][item.status] || 0) + 1;
    });
    
    return {
      success: true,
      stats: stats,
      semesterStats: semesterStats,
      groupStats: groupStats
    };
  } catch (error) {
    Logger.log('getStatistics Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ====================================
// Activity Logging
// ====================================
function logActivity(action, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LOGS);
    
    if (sheet) {
      sheet.appendRow([
        Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
        Session.getActiveUser().getEmail() || 'Anonymous',
        action,
        details,
        ''
      ]);
    }
  } catch (error) {
    Logger.log('logActivity Error: ' + error.toString());
  }
}

// ====================================
// Utility Functions
// ====================================
function testConnection() {
  return {
    success: true,
    message: 'Connection successful',
    timestamp: new Date().toISOString()
  };
}

function getSpreadsheetUrl() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getUrl();
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

function getFolderUrl() {
  try {
    return DriveApp.getFolderById(FOLDER_ID).getUrl();
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

// ====================================
// Email Notification (เหมือนเดิม — แจ้งทุกการเปลี่ยนแปลง)
// ====================================
function sendEmailNotification(syllabusData, userEmail, action) {
  try {
    if (!userEmail || !userEmail.trim() || userEmail === 'undefined') {
      return { success: false, message: 'ไม่มีอีเมลผู้รับ' };
    }

    const emailStr = userEmail.toString().trim();
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(emailStr)) {
      return { success: false, message: 'รูปแบบอีเมลไม่ถูกต้อง' };
    }

    const statusLabels = {
      'not_submitted': 'ยังไม่ส่ง',
      'pending': 'รอตรวจ',
      'approved': 'ตรวจแล้ว',
      'revision': 'ส่งกลับแก้ไข'
    };
    const statusThai = statusLabels[syllabusData.status] || syllabusData.status;

    let subject = '', body = '';

    if (action === 'EDIT') {
      subject = '📝 แจ้งเตือน: ข้อมูลแผนการจัดการเรียนรู้ถูกแก้ไขแล้ว';
      body = `
        <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9;">
          <div style="background: linear-gradient(135deg, #4a90e2, #357abd); padding: 30px; border-radius: 10px 10px 0 0; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">📝 แก้ไขข้อมูลสำเร็จ</h1>
          </div>
          
          <div style="background-color: white; padding: 30px; border-radius: 0 0 10px 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <p style="font-size: 16px; color: #333; margin-bottom: 20px;">
              เรียน คุณ<strong>${syllabusData.instructor}</strong>
            </p>
            
            <p style="font-size: 16px; color: #333; margin-bottom: 20px;">
              ข้อมูลแผนการจัดการเรียนรู้ของคุณได้รับการแก้ไขเรียบร้อยแล้ว
            </p>
            
            <div style="background-color: #f0f7ff; padding: 20px; border-radius: 8px; border-left: 4px solid #4a90e2; margin: 20px 0;">
              <h3 style="color: #4a90e2; margin-top: 0; margin-bottom: 15px; font-size: 18px;">📋 รายละเอียด</h3>
              <table style="width: 100%; border-collapse: collapse;">
                <tr><td style="padding: 8px 0; color: #666;">รหัสวิชา:</td><td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_code}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">ชื่อวิชา:</td><td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_name}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">ครูผู้สอน:</td><td style="padding: 8px 0; color: #333;">${syllabusData.instructor}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">ภาคเรียน/ปีการศึกษา:</td><td style="padding: 8px 0; color: #333;">${syllabusData.semester}/${syllabusData.academic_year}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">กลุ่มสาระการเรียนรู้:</td><td style="padding: 8px 0; color: #333;">${syllabusData.group}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">สถานะ:</td><td style="padding: 8px 0;"><span style="background-color: #28a745; color: white; padding: 4px 12px; border-radius: 12px;">${statusThai}</span></td></tr>
                <tr><td style="padding: 8px 0; color: #666; vertical-align: top;">หมายเหตุ:</td><td style="padding: 8px 0; color: #333;">${syllabusData.reviewer_notes || '-'}</td></tr>
              </table>
            </div>
            
            <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center;">
              <p style="color: #666; font-size: 14px;">
                วันที่แก้ไข: <strong>${Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm น.')}</strong>
              </p>
            </div>
            
            <div style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border-radius: 8px; border-left: 4px solid #ffc107;">
              <p style="margin: 0; color: #856404; font-size: 14px;">
                💡 <strong>หมายเหตุ:</strong> อีเมลนี้ส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ
              </p>
            </div>
          </div>
          
          <div style="text-align: center; margin-top: 20px; padding: 15px; color: #666; font-size: 12px;">
            <p>พัฒนาโดย <strong>นายชัยรัตน์ มานิช โรงเรียนวิชูทิศ สำนักงานเขตดินแดง</strong> @canva AI</p>
            <p>© ${new Date().getFullYear()} ระบบบริหารจัดการแผนการเรียนรู้ออนไลน์</p>
          </div>
        </div>
      `;
    } else if (action === 'STATUS_CHANGE') {
      subject = '🔔 แจ้งเตือน: สถานะแผนการจัดการเรียนรู้เปลี่ยนแปลง';
      
      let statusColor = '#28a745', statusIcon = '✅';
      if (syllabusData.status === 'pending') { statusColor = '#ffc107'; statusIcon = '⏳'; }
      else if (syllabusData.status === 'revision') { statusColor = '#fd7e14'; statusIcon = '🔄'; }
      else if (syllabusData.status === 'not_submitted') { statusColor = '#dc3545'; statusIcon = '❌'; }
      
      body = `
        <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9;">
          <div style="background: linear-gradient(135deg, #6b8e23, #556b2f); padding: 30px; border-radius: 10px 10px 0 0; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 24px;">🔔 สถานะเปลี่ยนแปลง</h1>
          </div>
          
          <div style="background-color: white; padding: 30px; border-radius: 0 0 10px 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <p style="font-size: 16px; color: #333; margin-bottom: 20px;">
              เรียน คุณ<strong>${syllabusData.instructor}</strong>
            </p>
            
            <p style="font-size: 16px; color: #333; margin-bottom: 20px;">
              สถานะแผนการจัดการเรียนรู้ของคุณได้รับการเปลี่ยนแปลง
            </p>
            
            <div style="background-color: #f0f7ff; padding: 20px; border-radius: 8px; border-left: 4px solid ${statusColor}; margin: 20px 0;">
              <h3 style="color: ${statusColor}; margin-top: 0; margin-bottom: 15px; font-size: 18px;">📋 รายละเอียด</h3>
              <table style="width: 100%; border-collapse: collapse;">
                <tr><td style="padding: 8px 0; color: #666;">รหัสวิชา:</td><td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_code}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">ชื่อวิชา:</td><td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_name}</td></tr>
                <tr><td style="padding: 8px 0; color: #666;">สถานะใหม่:</td><td style="padding: 8px 0;"><span style="background-color: ${statusColor}; color: white; padding: 6px 16px; border-radius: 12px; font-weight: 600;">${statusIcon} ${statusThai}</span></td></tr>
                ${syllabusData.reviewer_notes ? `<tr><td style="padding: 8px 0; color: #666; vertical-align: top;">หมายเหตุ:</td><td style="padding: 8px 0; color: #333;">${syllabusData.reviewer_notes}</td></tr>` : ''}
              </table>
            </div>
            
            <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center;">
              <p style="color: #666; font-size: 14px;">
                วันที่อัปเดต: <strong>${Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm น.')}</strong>
              </p>
            </div>
            
            <div style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border-radius: 8px; border-left: 4px solid #ffc107;">
              <p style="margin: 0; color: #856404; font-size: 14px;">
                💡 <strong>หมายเหตุ:</strong> อีเมลนี้ส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ
              </p>
            </div>
          </div>
          
          <div style="text-align: center; margin-top: 20px; padding: 15px; color: #666; font-size: 12px;">
            <p>พัฒนาโดย <strong>นายชัยรัตน์ มานิช โรงเรียนวิชูทิศ สำนักงานเขตดินแดง</strong> @กรุงเทพมหานคร</p>
            <p>© ${new Date().getFullYear()} ระบบบริหารจัดการแผนการเรียนรู้ออนไลน์</p>
          </div>
        </div>
      `;
    }
    
    MailApp.sendEmail({ to: emailStr, subject, htmlBody: body, noReply: true });
    Logger.log('✅ อีเมลส่งถึง:', emailStr);
    return { success: true };

  } catch (error) {
    Logger.log('❌ sendEmailNotification Error:', error.toString());
    return { success: false, message: error.toString() };
  }
}


// ฟังก์ชันค้นหาอีเมลจากชื่อครู
function findUserEmail(instructorName) {
  try {
    Logger.log('findUserEmail called with: ' + instructorName + ' (' + typeof instructorName + ')');
    
    // ตรวจสอบและปฏิเสธค่าปัญหา
    if (!instructorName || 
        instructorName.toString().trim() === '' || 
        instructorName === 'null' || 
        instructorName === 'undefined') {
      Logger.log('findUserEmail: Invalid or empty instructor name');
      return null;
    }
    
    // แปลงเป็น string ที่ปลอดภัย
    let nameStr = instructorName.toString().trim();
    if (nameStr.length === 0) {
      Logger.log('findUserEmail: Empty string after toString()');
      return null;
    }
    
    // เปิดไฟล์
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.USERS);
    
    if (!sheet) {
      Logger.log('findUserEmail Error: Users sheet not found');
      return null;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No users found in sheet');
      return null;
    }
    
    // อ่าน USER ข้อมูลทั้งแถว (คอลัมน์ A:H)
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // ค้นหาโดยใช้ exact match (ชื่อเต็ม)
    for (let i = 0; i < data.length; i++) {
      const fullname = data[i][2] ? data[i][2].toString().trim().toLowerCase() : '';
      const email = data[i][3] ? data[i][3].toString().trim() : '';
      
      // เปรียบเทียบชื่อเต็ม (exact match ไม่สนตัวพิมพ์)
      if (fullname === nameStr.toLowerCase()) {
        Logger.log('Found exact match for ' + nameStr + ': ' + email);
        return email;
      }
    }
    
    // หากไม่พบให้คืนค่า null
    Logger.log('No email found for instructor: ' + nameStr);
    return null;
    
  } catch (error) {
    const errorMsg = error.toString();
    Logger.log('findUserEmail Error: ' + errorMsg);
    return null;
  }
}

// ====================================
// Test Functions
// ====================================
function testEmail() {
  const testEmail = 'modtanoy.tanong@gmail.com';
  
  const testData = {
    id: '12345',
    course_code: 'TEST101',
    course_name: 'วิชาทดสอบ',
    instructor: 'ดร.ทดสอบ ระบบ',
    semester: '1',
    academic_year: '2567',
    group: 'สาขาสุขภาพดิจิทัล',
    status: 'approved',
    submitted_date: '2024-11-28 10:30:00',
    reviewer_notes: 'ทดสอบการส่งอีเมล'
  };
  
  const result = sendEmailNotification(testData, testEmail, 'EDIT');
  Logger.log('Test email result: ' + JSON.stringify(result));
  
  return result;
}

function testStatusChangeEmail() {
  const testEmail = 'pormodtanoy@gmail.com';
  
  const testData = {
    id: '12345',
    course_code: 'TEST101',
    course_name: 'วิชาทดสอบ',
    instructor: 'ดร.ทดสอบ ระบบ',
    semester: '1',
    academic_year: '2567',
    group: 'สาขาสุขภาพดิจิทัล',
    status: 'revision',
    submitted_date: '2024-11-28 10:30:00',
    reviewer_notes: 'กรุณาแก้ไขรูปแบบการเขียนใหม่'
  };
  
  const result = sendEmailNotification(testData, testEmail, 'STATUS_CHANGE');
  Logger.log('Test status change email result: ' + JSON.stringify(result));
  
  return result;
}


// ====================================
// Telegram Notification (แจ้งเตือนเมื่อส่งแผนการจัดการเรียนรู้)
// ====================================
const TELEGRAM_BOT_TOKEN = '8888180069:AAH8ggyaBBDrodmKN1mCwqWono5MCY6LkVk';   // ← แก้ไชจุดที่4 แก้เป็น Token จริงจาก BotFather
const TELEGRAM_CHAT_ID  = '7411061347';      // ← แก้ไขจุดที่ 5 แก้เป็น chat_id (เช่น -100xxxxxxxxxx)

function sendTelegramNotification(syllabusData) {
  try {
    if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID ||
        TELEGRAM_BOT_TOKEN === 'YOUR_BOT_TOKEN_HERE' ||
        TELEGRAM_CHAT_ID === 'YOUR_CHAT_ID_HERE') {
      Logger.log('⚠️ Telegram config not set. Skipping.');
      return { success: false, message: 'Config missing' };
    }

    const message =
      `📤 *มีการส่งแผนการจัดการเรียนรู้ใหม่*\n\n` +
      `📌 *รหัสวิชา:* ${syllabusData.course_code || '-'}\n` +
      `📚 *ชื่อวิชา:* ${syllabusData.course_name || '-'}\n` +
      `👨‍🏫 *ครูผู้สอน:* ${syllabusData.instructor || '-'}\n` +
      `📆 *ภาคเรียน:* ${syllabusData.semester || '-'} / ${syllabusData.academic_year || '-'}\n` +
      `👥 *กลุ่มสาระการเรียนรู้:* ${syllabusData.group || '-'}\n` +
      `✅ *สถานะ:* รอตรวจ (ผู้สอนได้กดยืนยันส่งแล้ว)\n` +
      `\n🕒 ${Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm น.')}`;

    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
    const payload = {
      chat_id: TELEGRAM_CHAT_ID,
      text: message,
      parse_mode: 'Markdown',
      disable_web_page_preview: true
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    Logger.log(`📡 Telegram API Response [${code}]: ${text.substring(0, 300)}...`);

    if (code === 200) {
      const result = JSON.parse(text);
      if (result.ok) {
        Logger.log('✅ Telegram ส่งสำเร็จ');
        return { success: true };
      } else {
        throw new Error(`Telegram error: ${result.description || 'Unknown'}`);
      }
    } else if (code === 400) {
      throw new Error('Bad Request — ตรวจสอบ chat_id หรือ token');
    } else if (code === 401) {
      throw new Error('Invalid BOT_TOKEN — ตรวจสอบ token อีกครั้ง');
    } else if (code === 403) {
      throw new Error('Chat not found / bot not added — ตรวจสอบ chat_id และว่า bot อยู่ในกลุ่มแล้วหรือไม่');
    } else {
      throw new Error(`HTTP ${code}: ${text}`);
    }

  } catch (e) {
    Logger.log('❌ Telegram send failed: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ฟังก์ชันทดสอบส่ง Telegram ด้วยข้อมูลตัวอย่าง
function testTelegram() {
  const testData = {
    course_code: 'TEST101',
    course_name: 'ระบบสารสนเทศ',
    instructor: 'ทดสอบ ระบบ',
    semester: '1',
    academic_year: '2568',
    group: 'IT',
    status: 'pending'
  };
  Logger.log('🚀 เริ่มทดสอบ Telegram...');
  const result = sendTelegramNotification(testData);
  Logger.log('✅ ผลการทดสอบ:', JSON.stringify(result));
  return result;
}

