var SPREADSHEET_ID = '1VyqSbID2Qi8GCpWnOnBbomyhSBTQkwAX__wVkWJYBBw'; // หากสร้างแบบ Standalone script ให้นำ ID ของ Google Sheets มาใส่ในนี้ครับ

// ถ้า SPREADSHEET_ID ว่าง ระบบจะพยายามดึงจาก Active Spreadsheet
function getDB() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
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
    { name: "Users", headers: ["username", "password", "role", "redirect"] },
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
    // ถ้ามีชีทอยู่แล้วแต่มีแค่บรรทัดเดียว (แค่หัวข้อ) ให้ลบทิ้งเพื่อดึงข้อมูลใหม่
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
      sheet.appendRow(["admin", "1234", "admin", ""]);
    }
    // Add dummy machine if not exists
    if (s.name === "Machines" && sheet.getLastRow() === 1) {
      sheet.appendRow(["M001", "CNC Machine 1", 30, new Date(), new Date(new Date().setDate(new Date().getDate() + 30)), "Zone A", "", ""]);
    }
  });
  return "Database Setup Complete!";
}

// ------------------------------------------------------------------
// 2. Authentication
// ------------------------------------------------------------------
function login(username, password) {
  var sheet = getDB().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { success: true, name: data[i][0], role: data[i][2], redirect: data[i][3], username: data[i][0] };
    }
  }
  return { success: false, message: "Username หรือ Password ไม่ถูกต้อง" };
}

function registerUser(username, password) {
  var sheet = getDB().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == username) {
      return { success: false, message: "มีผู้ใช้งานชื่อนี้แล้ว" };
    }
  }
  const redirectUrl = 'https://sites.google.com/view/pavaritport/homepage';
  sheet.appendRow([username, password, "member", redirectUrl]);
  return { success: true, message: "สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ" };
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
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][lineIndex]) {
      machines.push({
        id: data[i][lineIndex],
        name: data[i][lineIndex]
      });
    }
  }
  return machines;
}

function getTickets() {
  var sheet = getDB().getSheetByName("Tickets");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var tickets = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    tickets.push(obj);
  }
  return tickets.reverse(); // Newest first
}

function getTicketDetails(ticketId) {
  var db = getDB();
  
  // Get Ticket Info
  var tSheet = db.getSheetByName("Tickets");
  var tData = tSheet.getDataRange().getValues();
  var tHeaders = tData[0];
  var ticketInfo = null;
  for (var i = 1; i < tData.length; i++) {
    if (tData[i][0] == ticketId) {
      ticketInfo = {};
      for (var j = 0; j < tHeaders.length; j++) {
        ticketInfo[tHeaders[j]] = tData[i][j];
      }
      break;
    }
  }

  // Get History
  var hSheet = db.getSheetByName("History");
  var hData = hSheet.getDataRange().getValues();
  var history = [];
  for (var i = 1; i < hData.length; i++) {
    if (hData[i][0] == ticketId) {
      history.push({
        Timestamp: hData[i][1],
        Old_Status: hData[i][2],
        New_Status: hData[i][3],
        Update_By: hData[i][4]
      });
    }
  }

  // Get Repairs
  var rSheet = db.getSheetByName("Repairs");
  var rData = rSheet.getDataRange().getValues();
  var repairs = [];
  for (var i = 1; i < rData.length; i++) {
    if (rData[i][0] == ticketId) {
      repairs.push({
        Timestamp: rData[i][1],
        Repairer: rData[i][2],
        Repair_Details: rData[i][3],
        Before_Image: rData[i][4],
        After_Image: rData[i][5],
        Repair_Status: rData[i][6]
      });
    }
  }

  // Get Spare Parts
  var spSheet = db.getSheetByName("SpareParts");
  var spData = spSheet.getDataRange().getValues();
  var spareParts = [];
  for (var i = 1; i < spData.length; i++) {
    if (spData[i][1] == ticketId) {
      spareParts.push({
        Part_ID: spData[i][0],
        Part_Name: spData[i][2],
        Qty: spData[i][3],
        Status: spData[i][4]
      });
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
