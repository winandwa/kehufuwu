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
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { status: "error", message: "系統忙碌，請稍後再試" };
  }

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

    // [修正] 客訴日期改存純文字 "yyyy-MM-dd"，避免 UTC 時區造成日期偏移一天
    // formData.date 已是 "2026-03-31" 格式，直接存入即可
    const rowData = [
      formData.brand, lastId, formData.date, formData.customerName, formData.productNo,
      formData.platform, formData.buyPlatform, formData.orderNo, formData.content,
      formData.category, formData.solution, formData.comment, imageUrls.join(", "),
      "未解決", Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"), ""
    ];
    sheet.appendRow(rowData);

    const dispatchSolutions = ["請汶和解答", "請汶和延伸供應商"];
    if (dispatchSolutions.includes(formData.solution)) {
      try {
        syncToDispatchSystem(lastId, formData);
      } catch (syncError) {
        console.warn("派工系統寫入失敗：" + syncError.toString());
      }
    }

    return { status: "success", id: lastId };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
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

// [新增] 日期統一轉換函式：不管是 Date 物件還是字串都能正確輸出 "yyyy-MM-dd"
function toDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
  }
  return val.toString().slice(0, 10);
}

function getBrandRecords(brand) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  return data
    .filter(row => {
      if (!row[0] || !row[1]) return false;
      if (brand === "ALL") return true;
      // [修正] 移除全形/半形空格再比對，避免隱藏字元造成比對失敗
      const rowBrand = row[0].toString().replace(/[\s\u3000]/g, "");
      const targetBrand = brand.replace(/[\s\u3000]/g, "");
      return rowBrand === targetBrand;
    })
    .reverse()
    .map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        // [修正] 日期欄位用 toDateStr 統一處理
        if (row[i] instanceof Date) {
          obj[header] = toDateStr(row[i]);
        } else {
          obj[header] = row[i] !== undefined && row[i] !== null ? row[i].toString() : "";
        }
      });
      return obj;
    });
}

function getDashboardData(brand) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    data.shift(); // 移除 header

    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    const monthStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM");

    let dayStats = {}, monthStats = {}, unresolved = 0;

    data.forEach(row => {
      if (!row[0] || !row[2]) return;

      // [修正] 品牌比對移除空白字元
      if (brand !== "ALL") {
        const rowBrand = row[0].toString().replace(/[\s\u3000]/g, "");
        const targetBrand = brand.replace(/[\s\u3000]/g, "");
        if (rowBrand !== targetBrand) return;
      }

      // [修正] 用 toDateStr 統一處理日期，不管存的是 Date 物件還是字串
      const d = toDateStr(row[2]);
      if (!d) return;

      const m = d.slice(0, 7);
      const cat = row[9] ? row[9].toString().trim() : "未分類";
      const status = row[13] ? row[13].toString().trim() : "未解決";

      if (d === todayStr) dayStats[cat] = (dayStats[cat] || 0) + 1;
      if (m === monthStr) monthStats[cat] = (monthStats[cat] || 0) + 1;
      if (status !== "已結案") unresolved++;
    });

    return {
      day: Object.keys(dayStats).map(k => ({ cat: k, count: dayStats[k] })),
      month: Object.keys(monthStats).map(k => ({ cat: k, count: monthStats[k] })),
      unresolved: unresolved
    };
  } catch (e) {
    return { day: [], month: [], unresolved: 0, error: e.toString() };
  }
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
  return { status: "error", message: "找不到案件編碼：" + complaintId };
}

function syncToDispatchSystem(id, formData) {
  const dispatchSS = SpreadsheetApp.openById(CONFIG.DISPATCH_SHEET_ID);
  const dispatchSheet = dispatchSS.getSheets()[0];
  dispatchSheet.appendRow([
    `[客訴派工] ${id}`,
    `客戶：${formData.customerName}\n內容：${formData.content}`,
    "品保", "", new Date(), "待處理"
  ]);
}
function testDebug() {
  const brand = "TLV"; // 你的品牌名稱
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("品牌客訴總表");
  console.log("工作表：", sheet ? "找到了" : "找不到！");
  
  const data = sheet.getDataRange().getValues();
  console.log("總行數（含header）：", data.length);
  console.log("Header：", JSON.stringify(data[0]));
  
  if (data.length > 1) {
    const rawBrand = data[1][0];
    console.log("第一筆品牌原始值：", JSON.stringify(rawBrand));
    console.log("品牌比對結果：", rawBrand.toString().replace(/[\s\u3000]/g, "") === brand);
  }
  
  const dashResult = getDashboardData(brand);
  console.log("getDashboardData 結果：", JSON.stringify(dashResult));
  
  const recordResult = getBrandRecords(brand);
  console.log("getBrandRecords 筆數：", recordResult.length);
}