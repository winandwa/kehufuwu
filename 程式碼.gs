/**
 * 系統參數設定
 */
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
    
    // 1. 生成編碼
    const lastId = generateID(sheet, todayStr);
    
    // 2. 處理圖片上傳
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

    // 3. 寫入總表 (確保欄位數量 16 欄與順序正確)
    const rowData = [
      formData.brand,
      lastId,
      formData.date,
      formData.customerName,
      formData.productNo,
      formData.platform,      
      formData.buyPlatform,   
      formData.orderNo,       
      formData.content,
      formData.category,
      formData.solution,
      formData.comment,       
      imageUrls.join(", "),
      "未解決",
      new Date(),             // 建立時間
      ""                      // 後續進度記錄 (預設空白)
    ];
    sheet.appendRow(rowData);

    // 4. 派工連動
    if (formData.solution.includes("汶和")) {
      syncToDispatchSystem(lastId, formData);
    }

    return { status: "success", id: lastId };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function generateID(sheet, todayStr) {
  const data = sheet.getDataRange().getValues();
  let count = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] && data[i][1].toString().startsWith(todayStr)) {
      const lastIdStr = data[i][1].toString();
      const lastNum = parseInt(lastIdStr.split('WA')[1]);
      if (!isNaN(lastNum)) {
        count = lastNum + 1;
        break;
      }
    }
  }
  return todayStr + "WA" + count.toString().padStart(3, '0');
}

function syncToDispatchSystem(id, formData) {
  const dispatchSS = SpreadsheetApp.openById(CONFIG.DISPATCH_SHEET_ID);
  const dispatchSheet = dispatchSS.getSheets()[0];
  dispatchSheet.appendRow([
    `[客訴派工] ${id}`, 
    `客戶：${formData.customerName}\n產品：${formData.productNo}\n內容：${formData.content}\n說明：${formData.comment}`,
    "產品技術-品保",
    "", 
    new Date(),
    "待處理"
  ]);
}

function getBrandRecords(brand) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data
    .filter(row => brand === "ALL" || row[0] === brand)
    .reverse() 
    .map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        if (row[i] instanceof Date) {
          obj[header] = Utilities.formatDate(row[i], "GMT+8", "yyyy-MM-dd");
        } else {
          obj[header] = row[i];
        }
      });
      return obj;
    });
}

function updateComplaintProgress(complaintId, newProgress) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idIdx = headers.indexOf("編碼");
    const progressIdx = headers.indexOf("後續進度記錄");
    
    if (progressIdx === -1) return { status: "error", message: "請在試算表標題列補上『後續進度記錄』" };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === complaintId) {
        const timestamp = Utilities.formatDate(new Date(), "GMT+8", "MM/dd HH:mm");
        const oldProgress = data[i][progressIdx] || "";
        const updatedText = oldProgress + (oldProgress ? "\n" : "") + `[${timestamp}] ${newProgress}`;
        
        sheet.getRange(i + 1, progressIdx + 1).setValue(updatedText);
        return { status: "success" };
      }
    }
    return { status: "error", message: "找不到案件編碼" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getDashboardData(brand) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
  const monthStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM");
  
  let dayStats = {}, monthStats = {}, unresolved = 0;
  
  data.forEach(row => {
    if (brand !== "ALL" && row[0] !== brand) return;
    
    // 確保日期是字串格式
    const d = row[2] instanceof Date ? Utilities.formatDate(row[2], "GMT+8", "yyyy-MM-dd") : row[2].toString();
    const m = d.slice(0, 7);
    const cat = row[9]; 
    const status = row[13]; 
    
    if (d === todayStr) dayStats[cat] = (dayStats[cat] || 0) + 1;
    if (m === monthStr) monthStats[cat] = (monthStats[cat] || 0) + 1;
    if (status !== "已結案") unresolved++;
  });
  
  return {
    day: Object.keys(dayStats).map(k => ({cat: k, count: dayStats[k]})),
    month: Object.keys(monthStats).map(k => ({cat: k, count: monthStats[k]})),
    unresolved: unresolved
  };
}