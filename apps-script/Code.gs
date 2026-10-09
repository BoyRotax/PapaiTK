/**
 * ระบบบริหารจัดการศูนย์พัฒนาเด็กเล็ก - ฉบับสมบูรณ์ (Ultimate Edition)
 *
 * การแก้ไขจากฉบับเดิม (ดู apps-script/README.md):
 *  1. เพิ่ม action getAttendanceReport (ใช้พิมพ์ใบเช็คชื่อรายเดือน attendance_report.html)
 *  2. แก้ saveRoutineData ให้หาแถวเดิมเจอ (ตัด ' หน้าวันที่ก่อนเทียบ) ไม่เพิ่มแถวซ้ำ
 */

const SHEET_STUDENTS  = "Students";
const SHEET_USERS     = "Users";
const SHEET_TEACHERS  = "Teachers";
const SHEET_ATTEND    = "Attendance";
const SHEET_GROWTH    = "Growth";
const SHEET_HOLIDAYS  = "Holidays";
const SHEET_SPECIAL   = "SpecialStudents";
const SHEET_NEWS      = "News";
const SHEET_ROUTINE   = "DailyRoutine";
const DRIVE_FOLDER_ID = "1kY6uz84wD9oKF3Zg_4sPotO7719BdJ4P"; 

function forceDriveAuth() {
  DriveApp.createFile("dummy_test.txt", "Auth Test");
  SpreadsheetApp.getActive();
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = e.parameter.action;
    
    if (action === "login") return handleLogin(ss, e.parameter.username, e.parameter.password);
    if (action === "getStudents") return getAllData(ss, SHEET_STUDENTS);
    if (action === "getTeachers") return getAllData(ss, SHEET_TEACHERS);
    if (action === "searchStudent") return searchData(ss, SHEET_STUDENTS, 1, e.parameter.id);
    if (action === "checkHoliday") return createResponse(checkHoliday(e.parameter.date));
    if (action === "getKitchenData") return createResponse(getKitchenSummary());
    if (action === "getSpecialStudents") return getSpecialStudentsList();
    if (action === "getRoutineReport") return getRoutineReport();
    if (action === "getPresentStudents") return getPresentStudents(e.parameter.date);
    if (action === "getAttendanceReport") return getAttendanceReport();   // 🌟 ใหม่
    if (action === "getAllHolidays") return getAllHolidays();
    if (action === "getNews") return getNewsData(3);
    if (action === "getAllNews") return getNewsData(null);

    // 🌟 ระบบจัดการตั้งค่าและผู้ใช้
    if (action === "getUsers") return getAllUsers();
    if (action === "getSettings") return getSettings();

    return createResponse({ "status": "error", "msg": "Invalid Action in doGet" });
  } catch (error) { 
    return createResponse({ "status": "error", "msg": error.message });
  }
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = "";
    var requestData = {};

    if (e.postData && e.postData.contents) {
      try {
        requestData = JSON.parse(e.postData.contents);
        action = requestData.action;
      } catch(err) {
        action = e.parameter.action;
        requestData = e.parameter;
      }
    } else {
      action = e.parameter.action;
      requestData = e.parameter;
    }

    if (!action) return createResponse({ "status": "error", "msg": "No action specified" });
    
    if (action === "deleteStudent") return deleteData(ss, SHEET_STUDENTS, 1, requestData.studentId);
    if (action === "updateStudent" || action === "insertStudent") return saveStudent(ss, requestData, action);
    if (action === "deleteTeacher") return deleteData(ss, SHEET_TEACHERS, 1, requestData.empId); 
    if (action === "updateTeacher" || action === "insertTeacher") return saveTeacher(ss, requestData, action);
    if (action === "saveAttendance") return saveAttendanceList(requestData.data || e.parameter.data);
    if (action === "saveGrowth") return saveGrowthData(requestData.data || e.parameter.data);
    if (action === "getParentPortalData") return getParentPortalData(requestData.studentId || e.parameter.studentId);
    if (action === "saveRoutine") return saveRoutineData(requestData.data || e.parameter.data);
    
    if (action === "saveSpecialStudent") {
      var payload = typeof requestData.data === 'string' ? JSON.parse(requestData.data) : (requestData.data || requestData);
      return saveSpecialStudent(payload);
    }
    if (action === "updateSpecialPayment") {
      var payload = typeof requestData.data === 'string' ? JSON.parse(requestData.data) : (requestData.data || requestData);
      return updateSpecialPayment(payload);
    }
    if (action === "saveNews") {
      var payload = typeof requestData.data === 'string' ? JSON.parse(requestData.data) : (requestData.data || requestData);
      return saveNewsData(payload);
    }
    if (action === "deleteNews") {
      var payload = typeof requestData.data === 'string' ? JSON.parse(requestData.data) : (requestData.data || requestData);
      return deleteNewsData(payload.newsId);
    }

    // 🌟 ระบบจัดการตั้งค่าและผู้ใช้
    if (action === "saveUser") return saveUserData(requestData.data || e.parameter.data);
    if (action === "deleteUser") return deleteUserData(requestData.username || e.parameter.username);
    if (action === "saveSettings") return saveSettingsData(requestData.data || e.parameter.data);

    return createResponse({ "status": "error", "msg": "Invalid Action in doPost: " + action });
  } catch (error) { 
    // Fallback แบบเก่าสำหรับ Form URL Encoded
    if (e.parameter && e.parameter.action) {
      if (e.parameter.action === "saveSpecialStudent") return saveSpecialStudent(JSON.parse(e.parameter.data));
      if (e.parameter.action === "updateSpecialPayment") return updateSpecialPayment(JSON.parse(e.parameter.data));
      if (e.parameter.action === "saveRoutine") return saveRoutineData(JSON.parse(e.parameter.data));
      if (e.parameter.action === "saveUser") return saveUserData(JSON.parse(e.parameter.data));
      if (e.parameter.action === "saveSettings") return saveSettingsData(JSON.parse(e.parameter.data));
    }
    return createResponse({ "status": "error", "msg": "doPost Error: " + error.message });
  }
}

// ==========================================
// 🧑‍💻 ระบบจัดการผู้ใช้งาน (Users)
// ==========================================
function getAllUsers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return createResponse([]);
    var data = sheet.getDataRange().getDisplayValues();
    var users = [];
    for (var i = 1; i < data.length; i++) {
      users.push({
        Username: data[i][1], Password: data[i][2], Name: data[i][3],
        Role: data[i][4], ClassLevel: data[i][5], Room: data[i][6], Status: data[i][7]
      });
    }
    return createResponse({status: "success", data: users});
  } catch (e) { return createResponse({status: "error", msg: e.message}); }
}

function saveUserData(dataParam) {
  try {
    var dataObj = typeof dataParam === 'string' ? JSON.parse(dataParam) : dataParam;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return createResponse({status: "error", msg: "ไม่พบชีต Users"});
    
    var data = sheet.getDataRange().getValues();
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(dataObj.Username).toLowerCase()) {
        rowIndex = i + 1; break;
      }
    }

    var rowData = [
      timestamp, dataObj.Username, dataObj.Password, dataObj.Name, 
      dataObj.Role, dataObj.ClassLevel || "", dataObj.Room || "", dataObj.Status || "เปิดใช้งาน"
    ];

    if (dataObj.isEdit === true) {
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 1, 1, 8).setValues([rowData]);
        return createResponse({status: "success", msg: "อัปเดตผู้ใช้งานเรียบร้อย"});
      } else {
        return createResponse({status: "error", msg: "ไม่พบผู้ใช้งานนี้ในระบบ"});
      }
    } else {
      if (rowIndex > -1) {
        return createResponse({status: "error", msg: "Username นี้มีอยู่ในระบบแล้ว กรุณาใช้ชื่ออื่น"});
      } else {
        sheet.appendRow(rowData);
        return createResponse({status: "success", msg: "เพิ่มผู้ใช้งานใหม่เรียบร้อย"});
      }
    }
  } catch (e) { return createResponse({status: "error", msg: e.message}); }
}

function deleteUserData(username) {
  try {
    if(String(username).toLowerCase() === 'admin') return createResponse({status: "error", msg: "ไม่สามารถลบบัญชีผู้ดูแลระบบหลักได้"});
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(username).toLowerCase()) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex);
      return createResponse({status: "success", msg: "ลบผู้ใช้งานเรียบร้อย"});
    }
    return createResponse({status: "error", msg: "ไม่พบผู้ใช้งาน"});
  } catch (e) { return createResponse({status: "error", msg: e.message}); }
}

// ==========================================
// ⚙️ ระบบตั้งค่า (Settings)
// ==========================================
function getSettings() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (!sheet) return createResponse({});
    var data = sheet.getDataRange().getDisplayValues();
    var settings = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) settings[data[i][0]] = data[i][1];
    }
    return createResponse({status: "success", data: settings});
  } catch (e) { return createResponse({status: "error", msg: e.message}); }
}

function saveSettingsData(dataParam) {
  try {
    var dataObj = typeof dataParam === 'string' ? JSON.parse(dataParam) : dataParam;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) {
        sheet = ss.insertSheet("Settings");
        sheet.appendRow(["Key", "Value"]);
    }

    var keys = Object.keys(dataObj);
    for (var k = 0; k < keys.length; k++) {
      var keyName = keys[k];
      var val = dataObj[keyName];

      if (keyName === "SchoolLogo" && String(val).startsWith("data:image")) {
         var url = uploadImageToDrive(val, "LOGO_" + new Date().getTime() + ".png");
         if (!url.startsWith("ERROR")) val = url; 
      }

      var data = sheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === keyName) {
           sheet.getRange(i + 1, 2).setValue("'" + val); 
           found = true; break;
        }
      }
      if (!found) sheet.appendRow([keyName, "'" + val]);
    }
    return createResponse({status: "success", msg: "บันทึกการตั้งค่าเรียบร้อย"});
  } catch (e) { return createResponse({status: "error", msg: e.message}); }
}

// ==========================================
// 📰 ระบบข่าวสารและกิจกรรม (News)
// ==========================================
function getNewsData(limit) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NEWS);
    if (!sheet) return createResponse([]);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse([]);
    
    var headers = data[0];
    var result = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
      result.push(obj);
      if (limit && result.length >= limit) break;
    }
    return createResponse({ status: "success", data: result });
  } catch(e) { return createResponse({ status: "error", msg: e.message }); }
}

function saveNewsData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NEWS) || ss.insertSheet(SHEET_NEWS);
    if (sheet.getLastRow() === 0) sheet.appendRow(["Timestamp", "NewsDate", "Title", "Content", "ImageUrl"]);

    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    if (data.editId) {
       var sData = sheet.getDataRange().getValues();
       var rowIndex = -1;
       for(var i=1; i<sData.length; i++) {
         if(String(sData[i][0]) === String(data.editId)) { rowIndex = i + 1; break; }
       }
       if (rowIndex > -1) {
          var imageUrl = sData[rowIndex-1][4];
          if (data.base64Image) {
             imageUrl = uploadImageToDrive(data.base64Image, "NEWS_" + new Date().getTime() + ".jpg");
          }
          sheet.getRange(rowIndex, 2).setValue(data.newsDate);
          sheet.getRange(rowIndex, 3).setValue(data.title);
          sheet.getRange(rowIndex, 4).setValue(data.content);
          sheet.getRange(rowIndex, 5).setValue(imageUrl);
          return createResponse({ status: "success", msg: "อัปเดตข่าวสารเรียบร้อยแล้ว!" });
       }
    }

    var imageUrl = "https://images.unsplash.com/photo-1588072432836-e10032774350?auto=format&fit=crop&w=800&q=80";
    if (data.base64Image) {
      imageUrl = uploadImageToDrive(data.base64Image, "NEWS_" + new Date().getTime() + ".jpg");
      if (imageUrl.startsWith("ERROR")) throw new Error("อัปโหลดรูปไม่สำเร็จ: " + imageUrl);
    }

    sheet.appendRow([timestamp, data.newsDate, data.title, data.content, imageUrl]);
    return createResponse({ status: "success", msg: "ลงข่าวสารเรียบร้อยแล้ว!" });
  } catch (e) { return createResponse({ status: "error", msg: e.message }); }
}

function deleteNewsData(newsId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NEWS);
    if (!sheet) return createResponse({status: "error", msg: "ไม่พบฐานข้อมูลข่าวสาร"});
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(newsId)) {
        sheet.deleteRow(i + 1);
        return createResponse({ status: "success", msg: "ลบข่าวสารเรียบร้อย" });
      }
    }
    return createResponse({ status: "error", msg: "ไม่พบข้อมูลที่ต้องการลบ" });
  } catch (e) { return createResponse({ status: "error", msg: e.message }); }
}

// ==========================================
// ส่วนภาคปกติของระบบ
// ==========================================
function getSpecialStudentsList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SPECIAL);
  if (!sheet) return createResponse([]);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return createResponse([]);
  var headers = data[0], result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
    result.push(obj);
  }
  return createResponse(result);
}

function saveSpecialStudent(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_SPECIAL);
    if (!sheet) { 
        sheet = ss.insertSheet(SHEET_SPECIAL); 
        sheet.appendRow(["Timestamp", "StudentId", "FullName", "NickName", "ClassLevel", "Room", "StudentType", "Tel", "TimeSlot", "PaymentStatus", "SlipUrl", "PaymentMonth", "PaymentDate"]);
    } else { 
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (headers.indexOf("PaymentStatus") === -1) { 
            sheet.getRange(1, sheet.getLastColumn() + 1).setValue("PaymentStatus"); 
            sheet.getRange(1, sheet.getLastColumn() + 2).setValue("SlipUrl");
        } 
        if (headers.indexOf("PaymentMonth") === -1) { 
            sheet.getRange(1, sheet.getLastColumn() + 1).setValue("PaymentMonth");
            sheet.getRange(1, sheet.getLastColumn() + 2).setValue("PaymentDate"); 
        } 
    }
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var sId = data.studentId || "";
    if (data.studentType === "เด็กภายนอก") {
       var lastRow = sheet.getLastRow();
       var exCount = 1;
       if(lastRow > 1) { 
           var allData = sheet.getDataRange().getValues();
           for(var i = 1; i < allData.length; i++) { 
               if(String(allData[i][1]).indexOf("EX") > -1) exCount++;
           } 
       }
       sId = "EX" + String(exCount).padStart(4, '0');
    }
    
    if (sheet.getLastRow() > 1) {
      var existingData = sheet.getDataRange().getValues();
      for (var i = 1; i < existingData.length; i++) {
        if (sId !== "" && existingData[i][1] == sId && existingData[i][8] == data.timeSlot) return createResponse({ status: "error", msg: "นักเรียนรหัสนี้ ลงทะเบียนเรียนรอบเวลานี้ไปแล้วครับ" });
        if (data.studentType === "เด็กภายนอก" && existingData[i][2] == data.fullName && existingData[i][8] == data.timeSlot) return createResponse({ status: "error", msg: "มีชื่อนักเรียนคนนี้ ลงทะเบียนรอบเวลานี้ไปแล้วครับ" });
      }
    }
    
    var slipUrl = "";
    if (data.base64Slip) { 
        slipUrl = uploadImageToDrive(data.base64Slip, "SLIP_" + sId + "_" + new Date().getTime() + ".jpg");
        if (slipUrl.startsWith("ERROR")) throw new Error("อัปโหลดสลิปไม่สำเร็จ: " + slipUrl); 
    }
    
    sheet.appendRow([timestamp, sId, data.fullName || "", data.nickName || "", data.classLevel || "", data.room || "-", data.studentType || "เด็กใน ศพด.", data.tel || "-", data.timeSlot || "-", data.paymentStatus || "ยังไม่ชำระ", slipUrl]);
    var newRowIndex = sheet.getLastRow();
    var updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMonth = updatedHeaders.indexOf("PaymentMonth") + 1;
    var colDate = updatedHeaders.indexOf("PaymentDate") + 1;
    if(colMonth > 0) sheet.getRange(newRowIndex, colMonth).setValue("'" + (data.paymentMonth || "-"));
    if(colDate > 0) sheet.getRange(newRowIndex, colDate).setValue(data.paymentStatus === "ชำระแล้ว" ? timestamp : "-");
    return createResponse({ status: "success", msg: "ลงทะเบียนเรียนพิเศษสำเร็จ!" });
  } catch (e) { return createResponse({ status: "error", msg: e.message }); }
}

function updateSpecialPayment(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName(SHEET_SPECIAL);
    if(!sheet) return createResponse({ status: "error", msg: "ไม่พบฐานข้อมูลเรียนพิเศษ" });
    var sData = sheet.getDataRange().getValues(); var headers = sData[0];
    var colMonth = headers.indexOf("PaymentMonth") + 1; var colDate = headers.indexOf("PaymentDate") + 1; var colStatus = headers.indexOf("PaymentStatus") + 1;
    var colSlip = headers.indexOf("SlipUrl") + 1;
    if(colMonth === 0) { colMonth = sheet.getLastColumn() + 1; sheet.getRange(1, colMonth).setValue("PaymentMonth"); }
    if(colDate === 0) { colDate = sheet.getLastColumn() + 1; sheet.getRange(1, colDate).setValue("PaymentDate"); }
    
    var rowIndex = -1;
    for(var i=1; i<sData.length; i++) { 
        if(sData[i][1] == data.studentId && sData[i][8] == data.timeSlot) { rowIndex = i + 1; break; } 
    }
    
    if(rowIndex === -1) return createResponse({ status: "error", msg: "ไม่พบข้อมูลรหัสนักเรียนในรอบเวลานี้" });
    if(data.isDropOut === true) { sheet.deleteRow(rowIndex); return createResponse({ status: "success", msg: "ยกเลิกการเรียนพิเศษและลบรายชื่อเรียบร้อย" }); }
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var slipUrl = sData[rowIndex-1][colSlip-1];
    if (data.base64Slip) { slipUrl = uploadImageToDrive(data.base64Slip, "SLIP_" + data.studentId + "_" + new Date().getTime() + ".jpg"); }
    
    sheet.getRange(rowIndex, colMonth).setValue("'" + data.paymentMonth);
    sheet.getRange(rowIndex, colStatus).setValue(data.paymentStatus); 
    sheet.getRange(rowIndex, colSlip).setValue(slipUrl);
    
    if(data.paymentStatus === "ชำระแล้ว") { sheet.getRange(rowIndex, colDate).setValue(timestamp); } 
    else { sheet.getRange(rowIndex, colDate).setValue("-"); }
    
    return createResponse({ status: "success", msg: "อัปเดตข้อมูลเรียบร้อย" });
  } catch (e) { return createResponse({ status: "error", msg: e.message }); }
}

function saveStudent(ss, data, action) {
  var sheet = ss.getSheetByName(SHEET_STUDENTS) || ss.insertSheet(SHEET_STUDENTS); var timestamp = new Date();
  if(data.base64Image) { 
      var newUrl = uploadImageToDrive(data.base64Image, "STU_" + data.studentId + "_" + new Date().getTime() + ".jpg");
      if(newUrl !== "" && !newUrl.startsWith("ERROR")) data.photoUrl = newUrl; 
  }
  
  var rowData = [timestamp, data.studentId, data.classLevel || "", data.room || "", data.fullName || "", data.nickName || "", data.birthDate || "", data.gender || "", data.bloodType || "", data.province || "", data.amphoe || "", data.tambon || "", data.addressDetail || "", data.fatherName || "", data.fatherTel || "", data.motherName || "", data.motherTel || "", data.emergencyName || "", data.emergencyTel || "", data.medicalCondition || "-", data.allergies || "-", data.photoUrl || ""];
  
  var sheetData = sheet.getDataRange().getValues(); var rowIndex = -1;
  for (var i = 1; i < sheetData.length; i++) { if (sheetData[i][1] == data.studentId) { rowIndex = i + 1; break; } }
  if (action === "updateStudent") { 
      if (rowIndex > -1) { 
          sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
          return createResponse({ "status": "success", "msg": "อัปเดตข้อมูลสำเร็จ" }); 
      } 
      return createResponse({ "status": "not_found", "msg": "ไม่พบข้อมูลนักเรียน" });
  } 
  if (rowIndex > -1) return createResponse({ "status": "error", "msg": "มีรหัสนักเรียนนี้แล้ว" });
  sheet.appendRow(rowData);
  return createResponse({ "status": "success", "msg": "บันทึกข้อมูลใหม่เรียบร้อย" });
}

function saveTeacher(ss, data, action) {
  var sheet = ss.getSheetByName(SHEET_TEACHERS) || ss.insertSheet(SHEET_TEACHERS);
  var timestamp = new Date();
  if(data.base64Image) { 
      var newUrl = uploadImageToDrive(data.base64Image, "TCH_" + data.empId + "_" + new Date().getTime() + ".jpg");
      if(newUrl !== "" && !newUrl.startsWith("ERROR")) data.photoUrl = newUrl; 
  }
  var rowData = [timestamp, data.empId || "", data.title || "", data.name || "", data.position || "", data.assignedClass || "", data.assignedRoom || "", data.tel || "", data.startDate || "", data.status || "ปัจจุบัน", data.photoUrl || ""];
  var sheetData = sheet.getDataRange().getValues(); var rowIndex = -1;
  for (var i = 1; i < sheetData.length; i++) { if (sheetData[i][1] == data.empId) { rowIndex = i + 1; break; } }
  if (action === "updateTeacher") { 
      if (rowIndex > -1) { 
          sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
          return createResponse({ "status": "success", "msg": "อัปเดตข้อมูลครูเรียบร้อย" }); 
      } 
      return createResponse({ "status": "not_found", "msg": "ไม่พบข้อมูลครู" });
  } 
  if (rowIndex > -1) return createResponse({ "status": "error", "msg": "มีรหัสพนักงานนี้แล้ว" });
  sheet.appendRow(rowData);
  return createResponse({ "status": "success", "msg": "บันทึกข้อมูลครูใหม่เรียบร้อย" });
}

function saveAttendanceList(dataParam) {
  try { 
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTEND);
      if (!sheet) return createResponse({ status: "error", message: "หาแผ่นงาน Attendance ไม่เจอ" }); 
      var records = typeof dataParam === 'string' ? JSON.parse(dataParam) : dataParam; 
      var timestamp = new Date(); 
      records.forEach(function(rec) { sheet.appendRow([timestamp, rec.date, rec.studentId, rec.classLevel, rec.room, rec.status, rec.note]); });
      return createResponse({ status: "success", message: "บันทึกการเช็คชื่อเรียบร้อยแล้ว!" }); 
  } catch (e) { return createResponse({ status: "error", message: e.message }); }
}

function saveGrowthData(dataParam) {
  try { 
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROWTH);
      if (!sheet) return createResponse({ status: "error", message: "หาแผ่นงาน Growth ไม่เจอ" }); 
      var records = typeof dataParam === 'string' ? JSON.parse(dataParam) : dataParam; 
      var timestamp = new Date(); 
      records.forEach(function(rec) { if (rec.weight || rec.height) sheet.appendRow([timestamp, rec.monthYear, rec.studentId, rec.weight, rec.height]); });
      return createResponse({ status: "success", message: "บันทึกพัฒนาการเรียบร้อยแล้ว!" }); 
  } catch (e) { return createResponse({ status: "error", message: e.message }); }
}

function getKitchenSummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var studentSheet = ss.getSheetByName(SHEET_STUDENTS);
    var attendSheet = ss.getSheetByName(SHEET_ATTEND);
    var todayFmt = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"); var parts = todayFmt.split("/"); var yyyy = parseInt(parts[2]);
    if (yyyy < 2500) yyyy += 543; var todayStr = parts[0] + "/" + parts[1] + "/" + yyyy;
    var sData = studentSheet.getDataRange().getValues(); var roomCounter = {}; 
    for (var i = 1; i < sData.length; i++) { 
        var cName = String(sData[i][2]).trim();
        var rName = String(sData[i][3]).trim(); 
        if (!cName) continue; 
        if (!roomCounter[cName]) roomCounter[cName] = new Set(); 
        if (rName) roomCounter[cName].add(rName);
    }
    var classMap = {}; var finalSummary = {};
    for (var i = 1; i < sData.length; i++) { 
        var cName = String(sData[i][2]).trim(); var rName = String(sData[i][3]).trim();
        if (!cName) continue; 
        var internalKey = cName + "||" + rName; 
        if (!classMap[internalKey]) { 
            var displayName = cName;
            if (roomCounter[cName].size > 1 && rName) displayName = cName + " (ห้อง " + rName + ")";
            classMap[internalKey] = { displayName: displayName, total: 0, present: 0, checked: false }; 
        } 
        classMap[internalKey].total++;
    }
    if (attendSheet) {
      var aData = attendSheet.getDataRange().getValues(); var seenStudentsToday = {};
      for (var j = 1; j < aData.length; j++) {
        var recordDateRaw = aData[j][1];
        var recordDateStr = "";
        if (Object.prototype.toString.call(recordDateRaw) === '[object Date]') { 
            var rFmt = Utilities.formatDate(recordDateRaw, "GMT+7", "dd/MM/yyyy"); var rParts = rFmt.split("/");
            var rYyyy = parseInt(rParts[2]); if (rYyyy < 2500) rYyyy += 543;
            recordDateStr = rParts[0] + "/" + rParts[1] + "/" + rYyyy; 
        } else { recordDateStr = String(recordDateRaw).trim(); }
        var studentId = String(aData[j][2]).trim(); var cName = String(aData[j][3]).trim(); var rName = String(aData[j][4]).trim();
        var status = String(aData[j][5]).trim(); var internalKey = cName + "||" + rName;
        if (recordDateStr === todayStr && classMap[internalKey]) { 
            classMap[internalKey].checked = true; 
            if (status === "มา" && !seenStudentsToday[studentId]) { 
                classMap[internalKey].present++;
                seenStudentsToday[studentId] = true; 
            } 
        }
      }
    }
    for (var key in classMap) { 
        var info = classMap[key];
        finalSummary[info.displayName] = { total: info.total, present: info.present, checked: info.checked };
    }
    return { status: "success", date: todayStr, summary: finalSummary };
  } catch (e) { return { status: "error", message: e.message }; }
}

function getParentPortalData(studentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var studentSheet = ss.getSheetByName(SHEET_STUDENTS);
    if (!studentSheet) return createResponse({ status: "error", message: "ไม่พบแผ่นงานนักเรียน" });
    var studentData = studentSheet.getDataRange().getValues(); var headers = studentData[0];
    var studentInfo = null; var roomsInClass = new Set(); var studentClass = "";
    for (var i = 1; i < studentData.length; i++) { 
        if (studentData[i][1] == studentId) { 
            studentInfo = {};
            headers.forEach(function(h, idx) { studentInfo[h] = studentData[i][idx]; }); 
            studentInfo['Age'] = calculateAge(studentInfo['BirthDate']); 
            studentClass = studentInfo['ClassLevel'];
        } 
    }
    if (!studentInfo) return createResponse({ status: "not_found", message: "ไม่พบรหัสนักเรียนในระบบ" });
    for (var i = 1; i < studentData.length; i++) { 
        if (studentData[i][2] == studentClass) roomsInClass.add(studentData[i][3]);
    }
    studentInfo['ShowRoom'] = roomsInClass.size > 1;
    var teacherSheet = ss.getSheetByName(SHEET_TEACHERS); studentInfo['Teacher'] = "ไม่ระบุ";
    if (teacherSheet) { 
        var tData = teacherSheet.getDataRange().getValues(); 
        for(var t = 1; t < tData.length; t++) { 
            if(tData[t][5] == studentInfo['ClassLevel'] && (tData[t][6] == studentInfo['Room'] || tData[t][6] == "")) { 
                studentInfo['Teacher'] = (tData[t][2] || "") + (tData[t][3] || "");
                break; 
            } 
        } 
    }
    var attSheet = ss.getSheetByName(SHEET_ATTEND); var attendanceHistory = [];
    if (attSheet) { 
        var attData = attSheet.getDataRange().getValues(); 
        for (var i = attData.length - 1; i >= 1; i--) { 
            if (attData[i][2] == studentId) { 
                var dDate = new Date(attData[i][1]);
                var dateString = isNaN(dDate) ? attData[i][1] : Utilities.formatDate(dDate, "GMT+7", "dd/MM/yyyy"); 
                attendanceHistory.push({ date: dateString, status: attData[i][5], note: attData[i][6] || "-" });
                if (attendanceHistory.length >= 30) break; 
            } 
        } 
    }
    var growthSheet = ss.getSheetByName(SHEET_GROWTH); var latestGrowth = null;
    if (growthSheet) { 
        var growthData = growthSheet.getDataRange().getValues(); 
        for (var i = growthData.length - 1; i >= 1; i--) { 
            if (growthData[i][2] == studentId) { 
                latestGrowth = { monthYear: growthData[i][1], weight: growthData[i][3], height: growthData[i][4] };
                break; 
            } 
        } 
    }
    return createResponse({ status: "success", student: studentInfo, attendance: attendanceHistory, growth: latestGrowth });
  } catch (e) { return createResponse({ status: "error", message: e.message }); }
}

function checkHoliday(dateStr) {
  try { 
      var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName(SHEET_HOLIDAYS); var d = new Date(dateStr);
      var day = d.getDay(); if (day === 0 || day === 6) return { isHoliday: true, name: "วันหยุดเสาร์-อาทิตย์" };
      if (!sheet) return { isHoliday: false }; var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) { 
          var holidayDate = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "yyyy-MM-dd");
          if (holidayDate === dateStr) return { isHoliday: true, name: data[i][1] }; 
      } 
      return { isHoliday: false };
  } catch (e) { return { isHoliday: false, error: e.message }; }
}

function uploadImageToDrive(base64Str, fileName) {
  if(!base64Str) return "";
  try { 
      var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID); var contentType = "image/jpeg"; var data = base64Str;
      if (base64Str.indexOf('data:') === 0) { 
          contentType = base64Str.split(';')[0].substring(5); data = base64Str.split(',')[1]; 
      } 
      var blob = Utilities.newBlob(Utilities.base64Decode(data), contentType, fileName);
      var file = folder.createFile(blob); 
      return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500"; 
  } catch(e) { return "ERROR: " + e.message; }
}

function calculateAge(birthDateString) {
  if (!birthDateString) return "-"; var birthDate = new Date(birthDateString); if (isNaN(birthDate.getTime())) return "-";
  var bYear = birthDate.getFullYear(); if (bYear > 2400) bYear -= 543; var today = new Date();
  var years = today.getFullYear() - bYear; var months = today.getMonth() - birthDate.getMonth(); if (today.getDate() < birthDate.getDate()) months--;
  if (months < 0) { years--; months += 12; } var ageStr = "";
  if (years > 0) ageStr += years + " ปี ";
  if (months > 0) ageStr += months + " เดือน"; return ageStr.trim() || "ไม่ถึง 1 เดือน";
}

function handleLogin(ss, username, password) {
  var sheet = ss.getSheetByName(SHEET_USERS); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { 
      if (data[i][1] == username && data[i][2] == password) { 
          return createResponse({ "status": "success", "name": data[i][3], "role": data[i][4], "classLevel": data[i][5] || "", "room": data[i][6] || "" });
      } 
  }
  return createResponse({ "status": "error", "msg": "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" });
}

function getAllData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return createResponse([]); var data = sheet.getDataRange().getValues(); if (data.length <= 1) return createResponse([]);
  var headers = data[0], result = []; 
  for (var i = 1; i < data.length; i++) { 
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; }); 
      if (sheetName === SHEET_STUDENTS && obj['BirthDate']) obj['Age'] = calculateAge(obj['BirthDate']); 
      result.push(obj); 
  } 
  return createResponse(result);
}

function searchData(ss, sheetName, searchColIndex, searchValue) {
  var sheet = ss.getSheetByName(sheetName); var data = sheet.getDataRange().getValues(); var headers = data[0];
  for (var i = 1; i < data.length; i++) { 
      if (data[i][searchColIndex] == searchValue) { 
          var obj = {};
          headers.forEach(function(h, idx) { obj[h] = data[i][idx]; }); 
          if (sheetName === SHEET_STUDENTS && obj['BirthDate']) obj['Age'] = calculateAge(obj['BirthDate']);
          return createResponse({ "status": "found", "result": obj }); 
      } 
  } 
  return createResponse({ "status": "not_found", "msg": "ไม่พบข้อมูล" });
}

function deleteData(ss, sheetName, searchColIndex, searchValue) {
  var sheet = ss.getSheetByName(sheetName); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { 
      if (data[i][searchColIndex] == searchValue) { 
          sheet.deleteRow(i + 1);
          return createResponse({ "status": "success", "msg": "ลบข้อมูลเรียบร้อย" }); 
      } 
  } 
  return createResponse({ "status": "not_found", "msg": "ไม่พบข้อมูล" });
}

// ==========================================
// 🦷🥛 ระบบแปรงฟัน ดื่มนม (Daily Routine)
// ==========================================
function saveRoutineData(dataParam) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROUTINE);
    if (!sheet) return createResponse({ status: "error", message: "หาแผ่นงาน DailyRoutine ไม่เจอ" });

    var records = typeof dataParam === 'string' ? JSON.parse(dataParam) : dataParam;
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    // 🌟 แก้: ใช้ค่าที่แสดงจริง (ไม่มี ' นำหน้า) เพื่อเทียบกับวันที่ที่ตัด ' ออกแล้ว
    var sheetData = sheet.getDataRange().getDisplayValues();
    
    records.forEach(function(rec) {
      var recDate = String(rec.date).replace(/'/g, "");
      var rowIndex = -1;
      for (var i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][1]).trim() === recDate && String(sheetData[i][2]).trim() === String(rec.studentId).trim()) {
          rowIndex = i + 1; break;
        }
      }

      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 1).setValue(timestamp);
        sheet.getRange(rowIndex, 6).setValue(rec.brushed ? "✅ แปรงฟัน" : "❌ ไม่ได้แปรง");
        sheet.getRange(rowIndex, 7).setValue(rec.milk ? "✅ ดื่ม" : "❌ ไม่ได้ดื่ม");
      } else {
        // ใส่ ' นำหน้าวันที่ เพื่อให้ชีตเก็บเป็นข้อความ ไม่แปลงเป็นวันที่เอง
        sheet.appendRow([timestamp, "'" + recDate, rec.studentId, rec.classLevel, rec.room, (rec.brushed ? "✅ แปรงฟัน" : "❌ ไม่ได้แปรง"), (rec.milk ? "✅ ดื่ม" : "❌ ไม่ได้ดื่ม")]);
        // เพิ่มแถวใหม่ลงในข้อมูลที่เทียบ เพื่อไม่ให้ซ้ำถ้ามีรายการเดียวกันซ้ำใน batch เดียวกัน
        sheetData.push([timestamp, recDate, String(rec.studentId)]);
      }
    });
    return createResponse({ status: "success", message: "บันทึกข้อมูลแปรงฟัน-ดื่มนมเรียบร้อยแล้ว!" });
  } catch (e) { return createResponse({ status: "error", message: e.message }); }
}

function getRoutineReport() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROUTINE);
    if (!sheet) return createResponse([]);
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({
        Date: data[i][1], StudentId: data[i][2], ClassLevel: data[i][3],
        Room: data[i][4], BrushedTeeth: data[i][5], DrankMilk: data[i][6]
      });
    }
    return createResponse({ status: "success", data: result });
  } catch (e) { return createResponse({ status: "error", message: e.message }); }
}

function getPresentStudents(dateStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTEND);
    if (!sheet) return createResponse([]);
    var data = sheet.getDataRange().getDisplayValues();
    var presentIds = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === dateStr && String(data[i][5]).indexOf("มา") > -1) {
        presentIds.push(String(data[i][2]));
      }
    }
    return createResponse({status: "success", data: presentIds});
  } catch(e) { return createResponse({status: "error", message: e.message}); }
}

// 🌟 ใหม่: ข้อมูลเช็คชื่อทั้งหมด (ใช้พิมพ์ใบเช็คชื่อรายเดือน)
// คอลัมน์ของชีต Attendance: [0]=เวลาบันทึก [1]=วันที่ [2]=รหัสนักเรียน [3]=ชั้น [4]=ห้อง [5]=สถานะ [6]=หมายเหตุ
function getAttendanceReport() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTEND);
    if (!sheet) return createResponse({ status: "success", data: [] });
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({ Date: data[i][1], StudentId: data[i][2], Status: data[i][5] });
    }
    return createResponse({ status: "success", data: result });
  } catch (e) { return createResponse({ status: "error", msg: e.message }); }
}

function getAllHolidays() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
    if (!sheet) return createResponse([]);
    var data = sheet.getDataRange().getDisplayValues();
    var h = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) h.push({ Date: data[i][0], Name: data[i][1] });
    }
    return createResponse({status: "success", data: h});
  } catch (e) { return createResponse({status: "error", message: e.message}); }
}

function createResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
