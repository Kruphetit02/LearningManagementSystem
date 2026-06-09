// ====================================
// อ้างอิง : https://www.youtube.com/@Pormodtanoy/videos
// ====================================
const SPREADSHEET_ID = '1QMwunrikf3ZTrvdNBvtbmMaikwsjM5psDAHxP6tzfSs'; // เปลี่ยนเป็น ไอดีสเปรดชีต
const FOLDER_ID = '1NX3dZLuX1mwwzsKVWB1W-wJCCVRA7OFx'; // สร้างโฟลเดอร์ แล้าเอาไอดีโฟลเดอร์มมาใส่

// Sheet Names
const SHEETS = {
  SYLLABUS: 'ประมวลการสอน',
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
      .setFaviconUrl('https://img1.pic.in.th/images/logoc577678dae0d285f.png')
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
      syllabusSheet.getRange('A1:M1').setValues([[
        'ID', 'รหัสวิชา', 'ชื่อวิชา', 'อาจารย์ผู้สอน', 'ภาคเรียน', 
        'ปีการศึกษา', 'กลุ่มเรียน', 'สถานะ', 'วันที่ส่ง', 
        'หมายเหตุ', 'เวอร์ชัน', 'ไฟล์ ID', 'ไฟล์ URL'
      ]]);
      syllabusSheet.getRange('A1:M1').setFontWeight('bold')
        .setBackground('#4a90e2').setFontColor('#ffffff');
      syllabusSheet.setFrozenRows(1);
    }
    
    // Create Users Sheet with Role and Active columns
    let usersSheet = ss.getSheetByName(SHEETS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(SHEETS.USERS);
      usersSheet.getRange('A1:H1').setValues([[
        'Username', 'Password', 'ชื่อ-นามสกุล', 'อีเมล', 'กลุ่มเรียน', 'Role', 'Active', 'วันที่สมัคร'
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

// ==================================== 
// 1. แก้ไขฟังก์ชัน getAllSyllabus (ฝั่ง server)
// ====================================

function getAllSyllabus(userRole, userGroup) {
  try {
    Logger.log('getAllSyllabus: Starting...');
    Logger.log(`User role: ${userRole}, User group: ${userGroup}`);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      Logger.log('getAllSyllabus: Cannot access spreadsheet');
      return [];
    }
    
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    if (!sheet) {
      Logger.log('getAllSyllabus: Sheet not found');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log('getAllSyllabus: Last row = ' + lastRow);
    
    if (lastRow <= 1) {
      Logger.log('getAllSyllabus: No data rows');
      return [];
    }
    
    // อ่านข้อมูลทั้งหมด (คอลัมน์ A-M)
    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    Logger.log('getAllSyllabus: Raw data count = ' + data.length);
    
    const result = data
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => ({
        id: row[0] ? row[0].toString() : '',
        course_code: row[1] ? row[1].toString() : '',
        course_name: row[2] ? row[2].toString() : '',
        instructor: row[3] ? row[3].toString() : '',
        semester: row[4] ? row[4].toString() : '',
        academic_year: row[5] ? row[5].toString() : '',
        group: row[6] ? row[6].toString() : '',
        status: row[7] ? row[7].toString() : 'pending',
        submitted_date: row[8] ? row[8].toString() : '',
        reviewer_notes: row[9] ? row[9].toString() : '',
        version: row[10] ? row[10].toString() : '1',
        file_id: row[11] ? row[11].toString() : '',
        file_url: row[12] ? row[12].toString() : ''
      }));
    
    Logger.log('getAllSyllabus: Filtered result count = ' + result.length);
    
    // กรองข้อมูลตามสิทธิ์ผู้ใช้
    if (userRole === 'admin') {
      Logger.log('getAllSyllabus: Admin accessing, returning all data');
      return result;
    } else if (userGroup) {
      // แปลงกลุ่มให้เป็นตัวพิมพ์เล็กเพื่อเปรียบเทียบแบบไม่สนตัวพิมพ์
      const userGroupLower = userGroup.toLowerCase().trim();
      const filtered = result.filter(item => 
        item.group && item.group.toLowerCase().trim() === userGroupLower
      );
      Logger.log(`getAllSyllabus: User group filter (${userGroup}), returning ${filtered.length} items`);
      return filtered;
    } else {
      Logger.log('getAllSyllabus: No role/group, returning all data');
      return result;
    }
    
  } catch (error) {
    Logger.log('getAllSyllabus Error: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
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

function createSyllabus(data) {
  try {
    Logger.log('createSyllabus called with data: ' + JSON.stringify(data));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    
    if (!sheet) {
      throw new Error('ไม่พบชีตประมวลการสอน');
    }
    
    const id = data.id || Utilities.getUuid();
    
    // ✅ สร้างวันที่ในรูปแบบที่ถูกต้อง
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    Logger.log('Generated timestamp: ' + timestamp);
    
    // เพิ่มข้อมูลลงในชีต
    sheet.appendRow([
      id,
      data.course_code || '',
      data.course_name || '',
      data.instructor || '',
      data.semester || '',
      data.academic_year || '',
      data.group || '',
      data.status || 'pending',
      timestamp, // ✅ ใช้ timestamp ที่ format แล้ว
      data.reviewer_notes || '',
      data.version || 1,
      data.file_id || '',
      data.file_url || ''
    ]);
    
    Logger.log('Data saved successfully with ID: ' + id);
    
    try {
      logActivity('CREATE', 'Created syllabus: ' + data.course_code);
    } catch (logError) {
      Logger.log('Non-fatal logging error: ' + logError.toString());
    }
    
    return { 
      success: true, 
      data: { 
        id: id,
        course_code: data.course_code,
        course_name: data.course_name
      } 
    };
    
  } catch (error) {
    Logger.log('createSyllabus Error: ' + error.toString());
    return { 
      success: false, 
      message: 'เกิดข้อผิดพลาด: ' + error.toString() 
    };
  }
}

function updateSyllabus(data) {
  try {
    Logger.log('updateSyllabus called with data: ' + JSON.stringify(data));
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    
    if (!sheet) {
      return { success: false, message: 'ไม่พบชีตประมวลการสอน' };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: false, message: 'ไม่มีข้อมูลในชีต' };
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    Logger.log('Total rows to search: ' + dataRange.length);
    
    const searchId = data.id ? data.id.toString().trim() : '';
    Logger.log('Searching for ID: "' + searchId + '"');
    
    let rowIndex = -1;
    let existingData = null;
    
    for (let i = 0; i < dataRange.length; i++) {
      const cellId = dataRange[i][0] ? dataRange[i][0].toString().trim() : '';
      if (cellId === searchId) {
        rowIndex = i;
        existingData = dataRange[i];
        Logger.log('Found match at index: ' + i + ' (row ' + (i + 2) + ')');
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log('ID not found. Available IDs: ' + dataRange.map(r => r[0]).join(', '));
      return { success: false, message: 'ไม่พบข้อมูล ID: ' + searchId };
    }
    
    const actualRow = rowIndex + 2;
    Logger.log('Updating row: ' + actualRow);
    
    // ใช้ชื่ออาจารย์จากข้อมูลเดิมถ้าไม่มีใน data ที่ส่งมา
    const instructorName = data.instructor || existingData[3] || '';
    Logger.log('Instructor name (from data or existing): "' + instructorName + '"');
    
    const formattedDate = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    const updateValues = [[
      data.course_code || existingData[1] || '',
      data.course_name || existingData[2] || '',
      instructorName,
      data.semester || existingData[4] || '',
      data.academic_year || existingData[5] || '',
      data.group || existingData[6] || '',
      data.status || existingData[7] || 'pending',
      formattedDate,
      data.reviewer_notes || existingData[9] || '',
      data.version || existingData[10] || 1,
      data.file_id || existingData[11] || '',
      data.file_url || existingData[12] || ''
    ]];
    
    sheet.getRange(actualRow, 2, 1, 12).setValues(updateValues);
    
    Logger.log('Update successful for row: ' + actualRow);
    
    // ส่งอีเมลแจ้งเตือน
    try {
      Logger.log('=== EMAIL NOTIFICATION START ===');
      Logger.log('Instructor name to lookup: "' + instructorName + '"');
      
      if (instructorName && instructorName.trim() !== '') {
        const userEmail = findUserEmail(instructorName.trim());
        Logger.log('Email lookup result: ' + (userEmail || 'NOT FOUND'));
        
        if (userEmail) {
          const emailData = {
            id: data.id || existingData[0],
            course_code: data.course_code || existingData[1],
            course_name: data.course_name || existingData[2],
            instructor: instructorName,
            semester: data.semester || existingData[4],
            academic_year: data.academic_year || existingData[5],
            group: data.group || existingData[6],
            status: data.status || existingData[7],
            submitted_date: formattedDate,
            reviewer_notes: data.reviewer_notes || existingData[9] || ''
          };
          
          // กำหนด action: ถ้ามี originalStatus และต่างจาก status ปัจจุบัน = เปลี่ยนสถานะ
          const hasStatusChanged = data.originalStatus && data.originalStatus !== data.status;
          const action = hasStatusChanged ? 'STATUS_CHANGE' : 'EDIT';
          
          Logger.log('Email action: ' + action);
          Logger.log('Original status: ' + data.originalStatus + ', New status: ' + data.status);
          Logger.log('Sending email to: ' + userEmail);
          
          const emailResult = sendEmailNotification(emailData, userEmail, action);
          Logger.log('Email sent result: ' + JSON.stringify(emailResult));
        } else {
          Logger.log('WARNING: No email found for instructor: ' + instructorName);
          Logger.log('Please check if this name matches any user in the Users sheet');
        }
      } else {
        Logger.log('ERROR: Instructor name is empty or invalid');
      }
      Logger.log('=== EMAIL NOTIFICATION END ===');
    } catch (emailError) {
      Logger.log('Email notification error (non-fatal): ' + emailError.toString());
      Logger.log('Email error stack: ' + emailError.stack);
    }
    
    try {
      logActivity('UPDATE', 'Updated syllabus: ' + (data.course_code || 'Unknown') + ' (ID: ' + data.id + ')');
    } catch (logError) {
      Logger.log('Non-fatal logging error: ' + logError.toString());
    }
    
    return { success: true, message: 'อัปเดตข้อมูลสำเร็จ' };
    
  } catch (error) {
    Logger.log('updateSyllabus Error: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return { 
      success: false, 
      message: 'เกิดข้อผิดพลาด: ' + error.toString() 
    };
  }
}



function deleteSyllabus(id) {
  try {
    Logger.log('deleteSyllabus called with ID: ' + id);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SYLLABUS);
    
    if (!sheet) {
      return { success: false, message: 'ไม่พบชีตประมวลการสอน' };
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
// Email Notification
// ====================================
function sendEmailNotification(syllabusData, userEmail, action, attempts = 1) {
  try {
    Logger.log('Sending email to: ' + (userEmail || 'undefined'));
    
    // เพิ่มการตรวจสอบหลายระดับ
    if (!userEmail || userEmail.toString().trim() === '' || userEmail === 'undefined') {
      Logger.log('No email provided, skipping notification');
      return { success: false, message: 'ไม่มีอีเมลผู้รับ' };
    }
    
    // แปลง userEmail เป็น string และตรวจสอบค่า
    const emailStr = userEmail.toString().trim();
    if (!emailStr || emailStr === '') {
      Logger.log('Invalid email format: ' + userEmail);
      return { success: false, message: 'รูปแบบอีเมลไม่ถูกต้อง' };
    }
    
    // ตรวจสอบรูปแบบอีเมลพื้นฐาน
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(emailStr)) {
      Logger.log('Invalid email format detected: ' + emailStr);
      return { success: false, message: 'รูปแบบอีเมลไม่ถูกต้อง' };
    }
    
    // กำหนดสถานะเป็นภาษาไทย
    const statusLabels = {
      'not_submitted': 'ยังไม่ส่ง',
      'pending': 'รอตรวจ',
      'approved': 'ตรวจแล้ว',
      'revision': 'ส่งกลับแก้ไข'
    };
    
    const statusThai = statusLabels[syllabusData.status] || syllabusData.status;
    
    // กำหนด subject และ body ตาม action
    let subject = '';
    let body = '';
    
    if (action === 'EDIT') {
      subject = '📝 แจ้งเตือน: ข้อมูลประมวลการสอนถูกแก้ไขแล้ว';
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
              ข้อมูลประมวลการสอนของคุณได้รับการแก้ไขเรียบร้อยแล้ว
            </p>
            
            <div style="background-color: #f0f7ff; padding: 20px; border-radius: 8px; border-left: 4px solid #4a90e2; margin: 20px 0;">
              <h3 style="color: #4a90e2; margin-top: 0; margin-bottom: 15px; font-size: 18px;">📋 รายละเอียด</h3>
              
              <table style="width: 100%; border-collapse: collapse;">
                <tr>
                  <td style="padding: 8px 0; color: #666; width: 40%;">รหัสวิชา:</td>
                  <td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_code}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">ชื่อวิชา:</td>
                  <td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_name}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">อาจารย์ผู้สอน:</td>
                  <td style="padding: 8px 0; color: #333;">${syllabusData.instructor}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">ภาคเรียน/ปีการศึกษา:</td>
                  <td style="padding: 8px 0; color: #333;">${syllabusData.semester}/${syllabusData.academic_year}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">กลุ่มเรียน:</td>
                  <td style="padding: 8px 0; color: #333;">${syllabusData.group}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">สถานะ:</td>
                  <td style="padding: 8px 0;">
                    <span style="background-color: #28a745; color: white; padding: 4px 12px; border-radius: 12px; font-size: 14px; font-weight: 500;">
                      ${statusThai}
                    </span>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666; vertical-align: top;">หมายเหตุ:</td>
                  <td style="padding: 8px 0; color: #333;">${syllabusData.reviewer_notes || '-'}</td>
                </tr>
              </table>
            </div>
            
            <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center;">
              <p style="color: #666; font-size: 14px; margin-bottom: 10px;">
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
            <p style="margin: 5px 0;">พัฒนาโดย <strong>พ่อมดตะนอย</strong> @canva AI</p>
            <p style="margin: 5px 0;">© ${new Date().getFullYear()} ระบบบริหารจัดการแผนการเรียนรู้ออนไลน์</p>
          </div>
        </div>
      `;
    } else if (action === 'STATUS_CHANGE') {
      subject = '🔔 แจ้งเตือน: สถานะประมวลการสอนเปลี่ยนแปลง';
      
      // กำหนดสีตามสถานะ
      let statusColor = '#28a745';
      let statusIcon = '✅';
      
      if (syllabusData.status === 'pending') {
        statusColor = '#ffc107';
        statusIcon = '⏳';
      } else if (syllabusData.status === 'revision') {
        statusColor = '#fd7e14';
        statusIcon = '🔄';
      } else if (syllabusData.status === 'not_submitted') {
        statusColor = '#dc3545';
        statusIcon = '❌';
      }
      
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
              สถานะประมวลการสอนของคุณได้รับการเปลี่ยนแปลง
            </p>
            
            <div style="background-color: #f0f7ff; padding: 20px; border-radius: 8px; border-left: 4px solid ${statusColor}; margin: 20px 0;">
              <h3 style="color: ${statusColor}; margin-top: 0; margin-bottom: 15px; font-size: 18px;">📋 รายละเอียด</h3>
              
              <table style="width: 100%; border-collapse: collapse;">
                <tr>
                  <td style="padding: 8px 0; color: #666; width: 40%;">รหัสวิชา:</td>
                  <td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_code}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">ชื่อวิชา:</td>
                  <td style="padding: 8px 0; color: #333; font-weight: 600;">${syllabusData.course_name}</td>
                </tr>
                <tr>
                  <td style="padding: 8px 0; color: #666;">สถานะใหม่:</td>
                  <td style="padding: 8px 0;">
                    <span style="background-color: ${statusColor}; color: white; padding: 6px 16px; border-radius: 12px; font-size: 16px; font-weight: 600;">
                      ${statusIcon} ${statusThai}
                    </span>
                  </td>
                </tr>
                ${syllabusData.reviewer_notes ? `
                <tr>
                  <td style="padding: 8px 0; color: #666; vertical-align: top;">หมายเหตุ:</td>
                  <td style="padding: 8px 0; color: #333;">${syllabusData.reviewer_notes}</td>
                </tr>
                ` : ''}
              </table>
            </div>
            
            <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center;">
              <p style="color: #666; font-size: 14px; margin-bottom: 10px;">
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
            <p style="margin: 5px 0;">พัฒนาโดย <strong>พ่อมดตะนอย</strong> @canva AI</p>
            <p style="margin: 5px 0;">© ${new Date().getFullYear()} ระบบจัดการประมวลการสอน</p>
          </div>
        </div>
      `;
    }
    
    // ส่งอีเมล
     try {
      MailApp.sendEmail({
        to: emailStr,
        subject: subject,
        htmlBody: body,
        noReply: true
      });
      
      Logger.log('Email sent successfully to: ' + emailStr);
      return { success: true, message: 'ส่งอีเมลสำเร็จ' };
    } catch (sendError) {
      Logger.log('Failed to send email: ' + sendError.toString());
      
      // พยายามใหม่อีกครั้งหากยังไม่เกินจำนวนครั้งที่กำหนด
      if (attempts < 3) {
        Logger.log('Retrying email send... Attempt ' + (attempts + 1));
        Utilities.sleep(1000); // รอ 1 วินาทีก่อนลองใหม่
        return sendEmailNotification(syllabusData, userEmail, action, attempts + 1);
      }
      
      return { success: false, message: 'ไม่สามารถส่งอีเมลได้: ' + sendError.toString() };
    }
    
  } catch (error) {
    Logger.log('sendEmailNotification Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ฟังก์ชันค้นหาอีเมลจากชื่ออาจารย์
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

