var SPREADSHEET_ID = '1VyqSbID2Qi8GCpWnOnBbomyhSBTQkwAX__wVkWJYBBw'; // หากสร้างแบบ Standalone script ให้นำ ID ของ Google Sheets มาใส่ในนี้ครับ

// ถ้า SPREADSHEET_ID ว่าง ระบบจะพยายามดึงจาก Active Spreadsheet
function getDB() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  try {
    setupAuthDatabase();
  } catch (err) {
    Logger.log("Auto-setup error on doGet: " + err.toString());
  }
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Maintenance Management System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ------------------------------------------------------------------
// 1. Setup Database
// ------------------------------------------------------------------
function setupDatabase() {
  var ss = getDB();
  var sheets = [
    { name: "Users", headers: ["username", "password_hash", "role", "status", "created_at"] },
    { name: "UsersLogBook", headers: ["Log_ID", "Timestamp", "Username", "Role", "Action", "Remarks", "Device_Browser", "IP_Address"] },
    { name: "Equipments", headers: ["Serial Number", "TYPE", "Controller", "Fac", "Line", "MODEL", "Vendor", "QR Code코드", "QR Image코드이미지", "FAC2", "Remark"] },
    { name: "Repair History", headers: ["Date", "Serial Number", "Ticket_ID", "TYPE", "FAC2", "LINE", "Task Classification", "Before Symptoms", "Before Problem", "After Repair Completed", "PART", "Repair Started", "Repair Ended", "Total Time (Minutes)", "Worker작업자", "Before1", "Before2", "Before3", "After1", "After2", "After3", "Before1 Preview", "Before2 Preview", "Before3 Preview", "After1 Preview", "After2 Preview", "After3 Preview"] },
    { name: "Machines", headers: ["Machine_ID", "Machine_Name", "PM_Type", "Last_PM_Date", "Next_PM_Date", "Location", "QR_Link", "History_Link"] },
    { name: "Tickets", headers: ["Ticket_ID", "Timestamp", "Machine_ID", "Issue", "Reporter", "Status", "Urgency"] },
    { name: "History", headers: ["Ticket_ID", "Timestamp", "Old_Status", "New_Status", "Update_By"] },
    { name: "Repairs", headers: ["Ticket_ID", "Timestamp", "Repairer", "Repair_Details", "Before_Image", "After_Image", "Repair_Status"] },
    { name: "SpareParts", headers: ["Part_ID", "Ticket_ID", "Part_Name", "Qty", "Status"] }
  ];

  sheets.forEach(function(s) {
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
    }
    sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers]).setFontWeight("bold");
    
    // Copy "Machine List" from template if it doesn't exist
    var machineListSheet = ss.getSheetByName("Machine List");
    if (machineListSheet && machineListSheet.getLastRow() <= 1) {
      ss.deleteSheet(machineListSheet);
      machineListSheet = null;
    }
    
    if (!machineListSheet) {
      try {
        var templateSS = SpreadsheetApp.openById("1QbW-NTqkljS6SwE6vyzRi6dZAzKz4MsGbXPy84Sedqw");
        var templateSheet = templateSS.getSheetByName("Machine List") || templateSS.getSheets()[0];
        templateSheet.copyTo(ss).setName("Machine List");
      } catch (e) {
        Logger.log("Error copying Machine List: " + e.toString());
      }
    }
    
    // Add default user if not exists
    if (s.name === "Users" && sheet.getLastRow() === 1) {
      sheet.appendRow(["admin", hashPassword("1234"), "Admin", "Active", new Date()]);
    }
    // Add dummy machine if not exists
    if (s.name === "Machines" && sheet.getLastRow() === 1) {
      sheet.appendRow(["M001", "CNC Machine 1", 30, new Date(), new Date(new Date().setDate(new Date().getDate() + 30)), "Zone A", "", ""]);
    }
  });

  setupAuthDatabase();
  return "Database Setup Complete!";
}

function setupAuthDatabase() {
  var ss = getDB();
  
  // 1. Users sheet
  var userHeaders = ["username", "password_hash", "role", "status", "created_at"];
  var userSheet = ss.getSheetByName("Users");
  if (!userSheet) {
    userSheet = ss.insertSheet("Users");
    userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]).setFontWeight("bold");
    userSheet.appendRow(["admin", hashPassword("1234"), "Admin", "Active", new Date()]);
  } else {
    var curHeaders = userSheet.getRange(1, 1, 1, Math.max(userSheet.getLastColumn(), 1)).getValues()[0];
    if (curHeaders.indexOf("password_hash") === -1) {
      var lastRow = userSheet.getLastRow();
      var migratedUsers = [];
      if (lastRow > 1) {
        var usersData = userSheet.getRange(2, 1, lastRow - 1, curHeaders.length).getValues();
        usersData.forEach(function(row) {
          var u = (row[0] || "").toString().trim();
          if (!u) return;
          var rawPwd = (row[1] || "").toString().trim();
          var role = (row[2] && row[2].toString().toLowerCase() === 'admin') ? 'Admin' : 'User';
          var hashed = rawPwd ? hashPassword(rawPwd) : hashPassword("1234");
          migratedUsers.push([u, hashed, role, "Active", new Date()]);
        });
      }
      userSheet.clear();
      userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]).setFontWeight("bold");
      migratedUsers.forEach(function(uRow) {
        userSheet.appendRow(uRow);
      });
    }
  }

  // Always ensure 'admin' exists in Users
  var checkData = userSheet.getDataRange().getValues();
  var hasAdmin = false;
  for (var i = 1; i < checkData.length; i++) {
    if ((checkData[i][0] || "").toString().trim().toLowerCase() === "admin") {
      hasAdmin = true;
      break;
    }
  }
  if (!hasAdmin) {
    userSheet.appendRow(["admin", hashPassword("1234"), "Admin", "Active", new Date()]);
  }

  // 2. UsersLogBook sheet
  var logHeaders = ["Log_ID", "Timestamp", "Username", "Role", "Action", "Remarks", "Device_Browser", "IP_Address"];
  var logSheet = ss.getSheetByName("UsersLogBook");
  if (!logSheet) {
    logSheet = ss.insertSheet("UsersLogBook");
    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight("bold");
  }

  return "Auth Database Setup Complete!";
}

// ------------------------------------------------------------------
// 2. Authentication & User Management
// ------------------------------------------------------------------
function hashPassword(password) {
  if (!password) return "";
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password.toString(), Utilities.Charset.UTF_8);
  var hash = "";
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i] < 0 ? rawHash[i] + 256 : rawHash[i];
    var byteStr = byteVal.toString(16);
    hash += (byteStr.length === 1 ? "0" : "") + byteStr;
  }
  return hash;
}

function loginUser(username, password, clientInfo) {
  var ss = getDB();
  var sheet = ss.getSheetByName("Users");
  var logSheet = ss.getSheetByName("UsersLogBook");
  
  // Auto-heal / Auto-setup if missing or old headers
  if (!sheet || !logSheet) {
    setupAuthDatabase();
    sheet = ss.getSheetByName("Users");
  } else {
    var curHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    if (curHeaders.indexOf("password_hash") === -1) {
      setupAuthDatabase();
      sheet = ss.getSheetByName("Users");
    }
  }
  
  var data = sheet.getDataRange().getValues();
  var uClean = (username || "").toString().trim().toLowerCase();
  var hashed = hashPassword((password || "").toString().trim());
  var rawPwd = (password || "").toString().trim();
  
  for (var i = 1; i < data.length; i++) {
    var rowUser = (data[i][0] || "").toString().trim().toLowerCase();
    var rowHash = (data[i][1] || "").toString().trim();
    var rowRole = data[i][2] || "User";
    var rowStatus = data[i][3] || "Active";
    
    // Support matching both hashed password and legacy plain-text password
    var isMatch = (rowHash === hashed || rowHash === rawPwd);
    
    if (rowUser === uClean && isMatch) {
      // Auto-upgrade plain-text password to hash
      if (rowHash !== hashed) {
        sheet.getRange(i + 1, 2).setValue(hashed);
      }
      
      if (rowStatus === "Inactive") {
        return { success: false, message: "บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ" };
      }
      
      // Log to UsersLogBook ONLY on Login Success
      try {
        var lSheet = ss.getSheetByName("UsersLogBook");
        if (!lSheet) {
          setupAuthDatabase();
          lSheet = ss.getSheetByName("UsersLogBook");
        }
        var timestamp = new Date();
        var logId = "LOG-" + Utilities.formatDate(timestamp, "Asia/Bangkok", "yyMMdd-HHmmss") + "-" + Math.floor(100 + Math.random() * 900);
        var device = (clientInfo && clientInfo.device) ? clientInfo.device : "";
        var ip = (clientInfo && clientInfo.ip) ? clientInfo.ip : "";
        lSheet.appendRow([
          logId,
          timestamp,
          data[i][0],
          rowRole,
          "LOGIN_SUCCESS",
          "เข้าสู่ระบบสำเร็จ",
          device,
          ip
        ]);
      } catch (logErr) {
        Logger.log("Error logging user login: " + logErr.toString());
      }
      
      return {
        success: true,
        username: data[i][0],
        name: data[i][0],
        role: rowRole,
        status: rowStatus
      };
    }
  }
  return { success: false, message: "Username หรือ Password ไม่ถูกต้อง" };
}

function login(username, password, clientInfo) {
  return loginUser(username, password, clientInfo);
}

function registerUser(username, password) {
  var ss = getDB();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) {
    setupAuthDatabase();
    sheet = ss.getSheetByName("Users");
  }
  
  var uClean = (username || "").toString().trim().toLowerCase();
  var uOrig = (username || "").toString().trim();
  
  if (!/^[A-Za-z0-9_]+$/.test(uOrig)) {
    return { success: false, message: "กรุณาตั้งชื่อผู้ใช้เป็นภาษาอังกฤษและตัวเลขเท่านั้น" };
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var existingUser = (data[i][0] || "").toString().trim().toLowerCase();
    if (existingUser === uClean) {
      return { success: false, message: "มีผู้ใช้งานชื่อนี้แล้ว กรุณาใช้ชื่ออื่น" };
    }
  }
  
  var hashed = hashPassword((password || "").toString().trim());
  var role = (uClean === "admin") ? "Admin" : "User";
  sheet.appendRow([uOrig, hashed, role, "Active", new Date()]);
  return { success: true, message: "ลงทะเบียนสำเร็จ กรุณาเข้าสู่ระบบ" };
}

function getUsersList(adminUsername) {
  var ss = getDB();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Users sheet not found" };
  
  var data = sheet.getDataRange().getValues();
  var isAdmin = false;
  var checkUser = (adminUsername || "").toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === checkUser) {
      if ((data[i][2] || "").toString().toLowerCase() === "admin") {
        isAdmin = true;
      }
      break;
    }
  }
  if (!isAdmin) {
    return { success: false, message: "Permission Denied: เฉพาะผู้ดูแลระบบ (Admin) เท่านั้น" };
  }
  
  var list = [];
  for (var j = 1; j < data.length; j++) {
    list.push({
      username: data[j][0],
      role: data[j][2] || "User",
      status: data[j][3] || "Active",
      createdAt: data[j][4] ? Utilities.formatDate(new Date(data[j][4]), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "-"
    });
  }
  return { success: true, users: list };
}

function adminResetPassword(adminUsername, targetUsername, newPassword) {
  var ss = getDB();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  var isAdmin = false;
  var checkUser = (adminUsername || "").toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === checkUser && (data[i][2] || "").toString().toLowerCase() === "admin") {
      isAdmin = true;
      break;
    }
  }
  if (!isAdmin) return { success: false, message: "Permission Denied" };
  
  var target = (targetUsername || "").toString().trim().toLowerCase();
  var newHashed = hashPassword((newPassword || "").toString().trim());
  for (var j = 1; j < data.length; j++) {
    if ((data[j][0] || "").toString().trim().toLowerCase() === target) {
      sheet.getRange(j + 1, 2).setValue(newHashed);
      return { success: true, message: "รีเซ็ตรหัสผ่านของผู้ใช้ " + targetUsername + " เรียบร้อยแล้ว" };
    }
  }
  return { success: false, message: "ไม่พบผู้ใช้นี้ในระบบ" };
}

function adminToggleStatus(adminUsername, targetUsername, newStatus) {
  var ss = getDB();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  var isAdmin = false;
  var checkUser = (adminUsername || "").toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === checkUser && (data[i][2] || "").toString().toLowerCase() === "admin") {
      isAdmin = true;
      break;
    }
  }
  if (!isAdmin) return { success: false, message: "Permission Denied" };
  
  var target = (targetUsername || "").toString().trim().toLowerCase();
  for (var j = 1; j < data.length; j++) {
    if ((data[j][0] || "").toString().trim().toLowerCase() === target) {
      sheet.getRange(j + 1, 4).setValue(newStatus);
      return { success: true, message: "เปลี่ยนสถานะผู้ใช้เป็น " + newStatus + " เรียบร้อยแล้ว" };
    }
  }
  return { success: false, message: "ไม่พบผู้ใช้นี้ในระบบ" };
}

function adminUpdateRole(adminUsername, targetUsername, newRole) {
  var ss = getDB();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  var isAdmin = false;
  var checkUser = (adminUsername || "").toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === checkUser && (data[i][2] || "").toString().toLowerCase() === "admin") {
      isAdmin = true;
      break;
    }
  }
  if (!isAdmin) return { success: false, message: "Permission Denied" };
  
  var target = (targetUsername || "").toString().trim().toLowerCase();
  for (var j = 1; j < data.length; j++) {
    if ((data[j][0] || "").toString().trim().toLowerCase() === target) {
      sheet.getRange(j + 1, 3).setValue(newRole);
      return { success: true, message: "เปลี่ยนสิทธิ์ผู้ใช้เป็น " + newRole + " เรียบร้อยแล้ว" };
    }
  }
  return { success: false, message: "ไม่พบผู้ใช้นี้ในระบบ" };
}

// ------------------------------------------------------------------
// 3. API Data Fetching
// ------------------------------------------------------------------
function getMachines() {
  var sheet = getDB().getSheetByName("Machine List");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var machines = [];
  if (data.length < 2) return [];
  
  var headers = data[0];
  var lineIndex = headers.indexOf("Line");
  if (lineIndex === -1) lineIndex = 2; // Default to index 2
  
  var vendorIdx = headers.indexOf("Vender");
  if (vendorIdx === -1) vendorIdx = headers.indexOf("Vendor"); // Fallback
  
  var mcModelIdx = headers.indexOf("M/C Model");
  var serNoIdx = headers.indexOf("Ser.No.");
  var mfgDateIdx = headers.indexOf("MFG.Date");
  var powerSupplyIdx = headers.indexOf("Power Supply");
  var optionIdx = headers.indexOf("Option");
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][lineIndex]) {
      var mfgDate = mfgDateIdx > -1 ? data[i][mfgDateIdx] : "";
      if (mfgDate instanceof Date) {
         mfgDate = Utilities.formatDate(mfgDate, "Asia/Bangkok", "dd/MM/yyyy");
      }
      
      machines.push({
        id: data[i][lineIndex],
        name: data[i][lineIndex],
        details: {
          vendor: vendorIdx > -1 ? data[i][vendorIdx] : "",
          mcModel: mcModelIdx > -1 ? data[i][mcModelIdx] : "",
          serNo: serNoIdx > -1 ? data[i][serNoIdx] : "",
          mfgDate: mfgDate,
          powerSupply: powerSupplyIdx > -1 ? data[i][powerSupplyIdx] : "",
          option: optionIdx > -1 ? data[i][optionIdx] : ""
        }
      });
    }
  }
  return machines;
}

function getTickets() {
  var sheet = getDB().getSheetByName("Tickets");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  var headers = data[0];
  var tickets = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[headers[j]] = (val !== undefined && val !== null) ? val : "";
    }
    tickets.push(obj);
  }
  return tickets.reverse(); // Newest first
}

function getTicketDetails(ticketId) {
  var db = getDB();
  
  // Get Ticket Info
  var tSheet = db.getSheetByName("Tickets");
  var ticketInfo = null;
  if (tSheet) {
    var tData = tSheet.getDataRange().getValues();
    if (tData.length > 1) {
      var tHeaders = tData[0];
      for (var i = 1; i < tData.length; i++) {
        if (tData[i][0] == ticketId) {
          ticketInfo = {};
          for (var j = 0; j < tHeaders.length; j++) {
            var val = tData[i][j];
            if (val instanceof Date) {
              val = Utilities.formatDate(val, "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss");
            }
            ticketInfo[tHeaders[j]] = (val !== undefined && val !== null) ? val : "";
          }
          break;
        }
      }
    }
  }

  // Get History
  var hSheet = db.getSheetByName("History");
  var history = [];
  if (hSheet) {
    var hData = hSheet.getDataRange().getValues();
    for (var i = 1; i < hData.length; i++) {
      if (hData[i][0] == ticketId) {
        var hTime = hData[i][1];
        if (hTime instanceof Date) {
          hTime = Utilities.formatDate(hTime, "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss");
        }
        history.push({
          Timestamp: (hTime !== undefined && hTime !== null) ? hTime : "",
          Old_Status: hData[i][2] || "",
          New_Status: hData[i][3] || "",
          Update_By: hData[i][4] || ""
        });
      }
    }
  }

  // Get Repairs
  var rSheet = db.getSheetByName("Repairs");
  var repairs = [];
  if (rSheet) {
    var rData = rSheet.getDataRange().getValues();
    for (var i = 1; i < rData.length; i++) {
      if (rData[i][0] == ticketId) {
        var rTime = rData[i][1];
        if (rTime instanceof Date) {
          rTime = Utilities.formatDate(rTime, "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm:ss");
        }
        repairs.push({
          Timestamp: (rTime !== undefined && rTime !== null) ? rTime : "",
          Repairer: rData[i][2] || "",
          Repair_Details: rData[i][3] || "",
          Before_Image: rData[i][4] || "",
          After_Image: rData[i][5] || "",
          Repair_Status: rData[i][6] || ""
        });
      }
    }
  }

  // Get Spare Parts
  var spSheet = db.getSheetByName("SpareParts");
  var spareParts = [];
  if (spSheet) {
    var spData = spSheet.getDataRange().getValues();
    for (var i = 1; i < spData.length; i++) {
      if (spData[i][1] == ticketId) {
        spareParts.push({
          Part_ID: spData[i][0] || "",
          Part_Name: spData[i][2] || "",
          Qty: spData[i][3] || "",
          Status: spData[i][4] || ""
        });
      }
    }
  }

  return {
    ticket: ticketInfo,
    history: history.reverse(),
    repairs: repairs.reverse(),
    spareParts: spareParts
  };
}

// ------------------------------------------------------------------
// 4. API Operations
// ------------------------------------------------------------------
function createTicket(data) {
  var sheet = getDB().getSheetByName("Tickets");
  var timestamp = new Date();
  
  var dateStr = Utilities.formatDate(timestamp, "Asia/Bangkok", "yyMMdd");
  var lastRow = sheet.getLastRow();
  var seq = 1;
  if (lastRow > 1) {
    var lastTicketId = sheet.getRange(lastRow, 1).getValue();
    if (lastTicketId.toString().indexOf(dateStr) !== -1) {
      var parts = lastTicketId.split('-');
      if(parts.length > 1) {
        seq = parseInt(parts[1]) + 1;
      }
    }
  }
  var ticketId = "TK" + dateStr + "-" + Utilities.formatString("%03d", seq);

  var status = "รอดำเนินการ (Pending / 대기중)";
  sheet.appendRow([
    ticketId, 
    timestamp, 
    data.machineId, 
    data.issue, 
    data.reporter, 
    status,
    data.urgency || "ปกติ (Normal / 보통)"
  ]);

  var hSheet = getDB().getSheetByName("History");
  hSheet.appendRow([ticketId, timestamp, "New", status, data.reporter]);

  // ซิงค์เข้า Repair History
  try {
    var machineListSheet = getDB().getSheetByName("Machine List");
    var mData = machineListSheet ? machineListSheet.getDataRange().getValues() : [];
    var serialNum = "", type = "", fac2 = "";
    
    // ค้นหาข้อมูลเครื่องจักร (บรรทัดที่ 0 คือ header)
    if (mData.length > 0) {
      var mHeaders = mData[0];
      var lineIdx = mHeaders.indexOf("Line");
      if (lineIdx === -1) lineIdx = 2;
      var snIdx = mHeaders.indexOf("Serial Number");
      var typeIdx = mHeaders.indexOf("Type");
      var facIdx = mHeaders.indexOf("FAC 2");

      for (var i = 1; i < mData.length; i++) {
        if (mData[i][lineIdx] == data.machineId) {
          if (snIdx !== -1) serialNum = mData[i][snIdx];
          if (typeIdx !== -1) type = mData[i][typeIdx];
          if (facIdx !== -1) fac2 = mData[i][facIdx];
          break;
        }
      }
    }

    var rSheet = getDB().getSheetByName("Repair History");
    if (rSheet) {
      var repairRow = [];
      repairRow[0] = timestamp;
      repairRow[1] = serialNum;
      repairRow[2] = ticketId;
      repairRow[3] = type;
      repairRow[4] = fac2;
      repairRow[5] = data.machineId;
      repairRow[6] = ""; // Task Classification เว้นว่างไว้
      repairRow[7] = data.issue; // Before Symptoms
      
      // เติมช่องที่เหลือให้ครบ 27 คอลัมน์
      while(repairRow.length < 27) repairRow.push("");
      rSheet.appendRow(repairRow);
    }
  } catch (e) {
    Logger.log("Error syncing to Repair History: " + e.toString());
  }

  return { success: true, ticketId: ticketId };
}

function updateTicket(data) {
  var db = getDB();
  var timestamp = new Date();
  
  var tSheet = db.getSheetByName("Tickets");
  var tData = tSheet.getDataRange().getValues();
  var oldStatus = "";
  var rowIndex = -1;
  
  for (var i = 1; i < tData.length; i++) {
    if (tData[i][0] == data.ticketId) {
      oldStatus = tData[i][5]; // F (index 5) = Status
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: "ไม่พบใบงาน" };
  }

  if (data.newStatus && data.newStatus !== oldStatus) {
    tSheet.getRange(rowIndex, 6).setValue(data.newStatus);
    var hSheet = db.getSheetByName("History");
    hSheet.appendRow([data.ticketId, timestamp, oldStatus, data.newStatus, data.updateBy]);
  }

  var finalStatus = data.newStatus || oldStatus;

  var beforeImageUrl = "";
  var afterImageUrl = "";
  if (data.beforeImageBase64) {
    beforeImageUrl = uploadImageToDrive(data.beforeImageBase64, data.ticketId + "_Before");
  }
  if (data.afterImageBase64) {
    afterImageUrl = uploadImageToDrive(data.afterImageBase64, data.ticketId + "_After");
  }

  if (data.repairDetails || beforeImageUrl || afterImageUrl) {
    var rSheet = db.getSheetByName("Repairs");
    rSheet.appendRow([
      data.ticketId, 
      timestamp, 
      data.updateBy, 
      data.repairDetails || "", 
      beforeImageUrl, 
      afterImageUrl, 
      finalStatus
    ]);
  }

  return { success: true };
}

function addSparePart(data) {
  var sheet = getDB().getSheetByName("SpareParts");
  var partId = "SP" + new Date().getTime();
  sheet.appendRow([partId, data.ticketId, data.partName, data.qty, data.status]);
  return { success: true };
}

function uploadImageToDrive(base64Data, filename) {
  try {
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(byteCharacters, type, filename + ".jpg");
    
    var folderName = "Maintenance_Images";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log(e.toString());
    return "";
  }
}

function checkPMDaily() {
  var sheet = getDB().getSheetByName("Machines");
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0,0,0,0);
  
  var alertMessages = [];
  
  for (var i = 1; i < data.length; i++) {
    var machineId = data[i][0];
    var machineName = data[i][1];
    var nextPmDate = new Date(data[i][4]);
    nextPmDate.setHours(0,0,0,0);
    
    var timeDiff = nextPmDate.getTime() - today.getTime();
    var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24));
    
    if (diffDays <= 7 && diffDays > 0) {
      alertMessages.push("⚠️ ใกล้ถึงกำหนด PM ภายใน " + diffDays + " วัน: " + machineName + " (" + machineId + ")");
    } else if (diffDays === 0) {
      alertMessages.push("🔴 ถึงกำหนด PM วันนี้: " + machineName + " (" + machineId + ")");
    } else if (diffDays < 0) {
      alertMessages.push("🚨 เลยกำหนด PM มาแล้ว " + Math.abs(diffDays) + " วัน: " + machineName + " (" + machineId + ")");
    }
  }
  
  // Note: Add your LINE notify implementation or email here if needed.
}

// ------------------------------------------------------------------
// 7. OnEdit Trigger for AppSheet Images in Repair History
// ------------------------------------------------------------------
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Repair History") return;
  
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row === 1) return; // Skip header
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colName = headers[col - 1];
  
  // ตรวจสอบว่าแก้ไขในคอลัมน์รูปภาพหรือไม่
  var imageColumns = ["Before1", "Before2", "Before3", "After1", "After2", "After3"];
  if (imageColumns.indexOf(colName) !== -1) {
    var fileName = e.value;
    var previewColName = colName + " Preview";
    var previewColIndex = headers.indexOf(previewColName) + 1;
    
    if (previewColIndex > 0) {
      if (!fileName) {
        sheet.getRange(row, previewColIndex).clearContent();
        return;
      }
      
      try {
        // หาไฟล์ภาพใน Google Drive จากชื่อไฟล์ (AppSheet มักจะเก็บเป็นชื่อไฟล์ที่มีนามสกุล)
        // ตัด path ออกให้เหลือแค่ชื่อไฟล์
        var actualName = fileName.toString().split('/').pop();
        var files = DriveApp.getFilesByName(actualName);
        
        if (files.hasNext()) {
          var file = files.next();
          // เปิดสิทธิ์ให้เข้าถึงได้เพื่อให้ IMAGE() ทำงาน
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          var url = "https://drive.google.com/uc?id=" + file.getId();
          sheet.getRange(row, previewColIndex).setFormula('=IMAGE("' + url + '")');
        } else {
          sheet.getRange(row, previewColIndex).setValue("ไม่พบไฟล์ใน Drive");
        }
      } catch (err) {
        Logger.log(err);
      }
    }
  }
}

// ------------------------------------------------------------------
// 8. Translation Helper
// ------------------------------------------------------------------
function translateToEnglish(text) {
  try {
    return LanguageApp.translate(text, '', 'en');
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// ------------------------------------------------------------------
// 9. Auto-Sync from Repair History (AppSheet)
// ------------------------------------------------------------------
function syncStatusFromRepairHistory() {
  var ss = getDB();
  var rhSheet = ss.getSheetByName("Repair History");
  var tSheet = ss.getSheetByName("Tickets");
  var hSheet = ss.getSheetByName("History");
  
  if (!rhSheet || !tSheet || !hSheet) return;
  
  var rhData = rhSheet.getDataRange().getValues();
  if (rhData.length < 2) return;
  
  var tData = tSheet.getDataRange().getValues();
  
  // Create ticket lookup
  var ticketsMap = {};
  for (var i = 1; i < tData.length; i++) {
    ticketsMap[tData[i][0]] = {
      rowIndex: i + 1,
      status: tData[i][5],
      reporter: tData[i][4]
    };
  }
  
  var rhHeaders = rhData[0];
  var rhTicketIdIdx = rhHeaders.indexOf("Ticket_ID");
  var rhSerialIdx = rhHeaders.indexOf("Serial Number");
  var rhAfterRepairIdx = rhHeaders.indexOf("After Repair Completed");
  
  // รองรับทั้งชื่อคอลัมน์เก่าและใหม่
  var rhWorkerIdx = rhHeaders.indexOf("Worker작업자");
  if (rhWorkerIdx === -1) rhWorkerIdx = rhHeaders.indexOf("Worker");
  
  if (rhTicketIdIdx === -1) return;
  
  var timestamp = new Date();
  
  for (var j = 1; j < rhData.length; j++) {
    var row = rhData[j];
    var ticketId = row[rhTicketIdIdx];
    if (!ticketId) continue;
    
    var serial = rhSerialIdx !== -1 ? row[rhSerialIdx] : "";
    var afterRepair = rhAfterRepairIdx !== -1 ? row[rhAfterRepairIdx] : "";
    var workerName = (rhWorkerIdx !== -1 && row[rhWorkerIdx]) ? row[rhWorkerIdx] : "AppSheet System";
    
    var expectedStatus = "";
    var isSimultaneous = false;
    
    if (serial !== "" && afterRepair !== "") {
      expectedStatus = "เสร็จสิ้น (Completed / 완료)";
      isSimultaneous = true;
    } else if (afterRepair !== "") {
      expectedStatus = "เสร็จสิ้น (Completed / 완료)";
    } else if (serial !== "") {
      expectedStatus = "กำลังซ่อม (In Progress / 진행중)";
    }
    
    if (expectedStatus !== "" && ticketsMap[ticketId]) {
      var currentStatus = ticketsMap[ticketId].status;
      if (currentStatus !== expectedStatus && !currentStatus.includes("เสร็จสิ้น")) {
        // Update Tickets sheet
        tSheet.getRange(ticketsMap[ticketId].rowIndex, 6).setValue(expectedStatus);
        
        var oldStatusToLog = currentStatus;
        if (isSimultaneous) {
           oldStatusToLog = "กำลังซ่อม (In Progress / 진행중)";
        }
        
        // Append to History
        hSheet.appendRow([ticketId, timestamp, oldStatusToLog, expectedStatus, workerName]);
        
        // Update local map to prevent duplicate updates if script runs again
        ticketsMap[ticketId].status = expectedStatus;
      }
    }
  }
}

// ฟังก์ชันสำหรับติดตั้ง Trigger ให้ทำงานอัตโนมัติ
function setupAutoSyncTrigger() {
  var ss = getDB();
  var triggers = ScriptApp.getUserTriggers(ss);
  
  // ลบ Trigger เดิมที่มีชื่อซ้ำกันเพื่อป้องกันการซ้ำซ้อน
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncStatusFromRepairHistory") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // สร้าง Trigger ใหม่: ทำงานทุกครั้งที่มีการเปลี่ยนข้อมูลใน Spreadsheet (รวมถึงจาก AppSheet)
  ScriptApp.newTrigger("syncStatusFromRepairHistory")
    .forSpreadsheet(ss)
    .onChange()
    .create();
    
  return "ติดตั้งระบบ Auto-Sync สถานะจาก Repair History เสร็จสมบูรณ์!";
}
