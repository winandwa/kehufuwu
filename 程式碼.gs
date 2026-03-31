const CONFIG = {
  MASTER_SHEET_NAME: "品牌客訴總表",
  DISPATCH_SHEET_ID: "1hbkxLEKKXeOQl2VQQYKo1bA4WVMjZjVNeEZGa89wsAs", 
  IMAGE_FOLDER_ID: "1yKe1SGfCKkiQI1yO0Lfe_6I9zl9aVPyt" 
};

function doGet(e) {
  const brand = e.parameter.brand || "ALL";
  const template = HtmlService.createTemplateFromFile('Index');
  template.brand = brand;
  return template.evaluate()
    .setTitle("客訴管理系統 - " + brand)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
    const lastId = generateID(sheet, todayStr);
    
    let imageUrls = [];
    if (formData.fileData && formData.fileData.length > 0) {
      const folder = DriveApp.getFolderById(CONFIG.IMAGE_FOLDER_ID);
      formData.fileData.forEach((fileObj, index) => {
        const contentType = fileObj.base64String.split(",")[0].split(":")[1].split(";")[0];
        const bytes = Utilities.base64Decode(fileObj.base64String.split(",")[1]);
        const blob = Utilities.newBlob(bytes, contentType, `${lastId}_${index + 1}`);
        const newFile = folder.createFile(blob);
        imageUrls.push(newFile.getUrl());
      });
    }

    const complaintDate = new Date(formData.date);
    const rowData = [
      formData.brand, lastId, complaintDate, formData.customerName, formData.productNo,
      formData.platform, formData.buyPlatform, formData.orderNo, formData.content,
      formData.category, formData.solution, formData.comment, imageUrls.join(", "),
      "未解決", new Date(), ""
    ];
    sheet.appendRow(rowData);

    if (formData.solution.includes("汶和")) { syncToDispatchSystem(lastId, formData); }
    return { status: "success", id: lastId };
  } catch (error) { return { status: "error", message: error.toString() }; }
}

function generateID(sheet, todayStr) {
  const data = sheet.getDataRange().getValues();
  let count = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] && data[i][1].toString().startsWith(todayStr)) {
      const parts = data[i][1].toString().split('WA');
      if (parts.length > 1) { count = parseInt(parts[1]) + 1; break; }
    }
  }
  return todayStr + "WA" + count.toString().padStart(3, '0');
}

function getBrandRecords(brand) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data
    .filter(row => row[0] && row[1] && (brand === "ALL" || row[0].toString().trim() === brand))
    .reverse() 
    .map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        if (row[i] instanceof Date) { obj[header] = Utilities.formatDate(row[i], "GMT+8", "yyyy-MM-dd"); }
        else { obj[header] = row[i] || ""; }
      });
      return obj;
    });
}

function getDashboardData(brand) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); 
    
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    const monthStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM");
    
    let dayStats = {}, monthStats = {}, unresolved = 0;
    
    data.forEach(row => {
      if (!row[0] || !row[2]) return;
      if (brand !== "ALL" && row[0].toString().trim() !== brand) return;
      
      let d = "";
      if (row[2] instanceof Date) {
        d = Utilities.formatDate(row[2], "GMT+8", "yyyy-MM-dd");
      } else {
        d = row[2].toString();
      }
      
      const m = d.slice(0, 7);
      const cat = row[9] || "未分類";
      const status = row[13] ? row[13].toString().trim() : "未解決";

      if (d === todayStr) dayStats[cat] = (dayStats[cat] || 0) + 1;
      if (m === monthStr) monthStats[cat] = (monthStats[cat] || 0) + 1;
      if (status !== "已結案") unresolved++;
    });
    
    return {
      day: Object.keys(dayStats).map(k => ({cat: k, count: dayStats[k]})),
      month: Object.keys(monthStats).map(k => ({cat: k, count: monthStats[k]})),
      unresolved: unresolved
    };
  } catch (e) { return { day: [], month: [], unresolved: 0, error: e.toString() }; }
}

function updateComplaintProgress(complaintId, newProgress) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf("編碼");
  const progIdx = data[0].indexOf("後續進度記錄");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === complaintId) {
      const time = Utilities.formatDate(new Date(), "GMT+8", "MM/dd HH:mm");
      const old = data[i][progIdx] || "";
      sheet.getRange(i + 1, progIdx + 1).setValue(old + (old ? "\n" : "") + `[${time}] ${newProgress}`);
      return { status: "success" };
    }
  }
}

function syncToDispatchSystem(id, formData) {
  const dispatchSS = SpreadsheetApp.openById(CONFIG.DISPATCH_SHEET_ID);
  const dispatchSheet = dispatchSS.getSheets()[0];
  dispatchSheet.appendRow([`[客訴派工] ${id}`, `客戶：${formData.customerName}\n內容：${formData.content}`, "品保", "", new Date(), "待處理"]);
}