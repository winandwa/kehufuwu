const CONFIG = {
  MASTER_SHEET_NAME: "品牌客訴總表",
  MENU_SHEET_NAME: "選單設定表",
  DISPATCH_SHEET_ID: "1hbkxLEKKXeOQl2VQQYKo1bA4WVMjZjVNeEZGa89wsAs",
  IMAGE_FOLDER_ID: "1yKe1SGfCKkiQI1yO0Lfe_6I9zl9aVPyt"
};

function doGet(e) {
  const brand = e.parameter.brand || "ALL";
  const template = HtmlService.createTemplateFromFile('Index');
  template.brand = brand;
  // brand=ALL 才是管理者，看得到進階分析；各品牌窗口看不到
  template.isAdmin = (brand === "ALL");
  return template.evaluate()
    .setTitle("客訴管理系統 - " + brand)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGetAdmin(e) {
  const brand = e.parameter.brand || "ALL";
  const template = HtmlService.createTemplateFromFile('Index');
  template.brand = brand;
  template.isAdmin = true;
  return template.evaluate()
    .setTitle("客訴管理系統（後台）- " + brand)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// [新增] 讀取對應品牌的料號清單（選單設定表 A欄=品牌, B欄=料號）
function getProductList(brand) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MENU_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const products = [];
    data.forEach(function(row) {
      if (!row[0] || !row[1]) return;
      const rowBrand = row[0].toString().replace(/[\s\u3000]/g, "");
      const targetBrand = brand.replace(/[\s\u3000]/g, "");
      if (brand === "ALL" || rowBrand === targetBrand) {
        products.push(row[1].toString().trim());
      }
    });
    return products;
  } catch (e) {
    return [];
  }
}

function processForm(formData) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { status: "error", message: "系統忙碌，請稍後再試" }; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
    const lastId = generateID(sheet, todayStr);
    let imageUrls = [];
    if (formData.fileData && formData.fileData.length > 0) {
      const folder = DriveApp.getFolderById(CONFIG.IMAGE_FOLDER_ID);
      formData.fileData.forEach(function(fileObj, index) {
        const contentType = fileObj.base64String.split(",")[0].split(":")[1].split(";")[0];
        const bytes = Utilities.base64Decode(fileObj.base64String.split(",")[1]);
        const blob = Utilities.newBlob(bytes, contentType, lastId + "_" + (index + 1));
        imageUrls.push(folder.createFile(blob).getUrl());
      });
    }
    sheet.appendRow([
      formData.brand, lastId, formData.date, formData.customerName,
      formData.productNo, formData.batchNo,
      formData.platform, formData.buyPlatform, formData.orderNo,
      formData.content, formData.category, formData.solution,
      formData.comment, imageUrls.join(", "),
      "未解決", Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"), ""
    ]);
    const dispatchSolutions = ["請汶和解答", "請汶和延伸供應商"];
    if (dispatchSolutions.includes(formData.solution)) {
      try { syncToDispatchSystem(lastId, formData); }
      catch (syncError) { console.warn("派工失敗：" + syncError); }
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

function toDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
  return val.toString().slice(0, 10);
}

function getBrandRecords(brand) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data
    .filter(function(row) {
      if (!row[0] || !row[1]) return false;
      if (brand === "ALL") return true;
      return row[0].toString().replace(/[\s\u3000]/g, "") === brand.replace(/[\s\u3000]/g, "");
    })
    .reverse()
    .map(function(row) {
      var obj = {};
      headers.forEach(function(header, i) {
        obj[header] = (row[i] instanceof Date) ? toDateStr(row[i]) : (row[i] != null ? row[i].toString() : "");
      });
      return obj;
    });
}

function getDashboardData(brand) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    data.shift();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    const monthStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM");
    let dayStats = {}, monthStats = {}, unresolved = 0;
    let productStats = {}, dateStats = {};

    data.forEach(function(row) {
      if (!row[0] || !row[2]) return;
      if (brand !== "ALL" && row[0].toString().replace(/[\s\u3000]/g, "") !== brand.replace(/[\s\u3000]/g, "")) return;
      const d = toDateStr(row[2]);
      if (!d) return;
      const m = d.slice(0, 7);
      const cat = row[10] ? row[10].toString().trim() : "未分類";
      const status = row[14] ? row[14].toString().trim() : "未解決";
      const product = row[4] ? row[4].toString().trim() : "未知料號";
      if (d === todayStr) dayStats[cat] = (dayStats[cat] || 0) + 1;
      if (m === monthStr) monthStats[cat] = (monthStats[cat] || 0) + 1;
      if (status !== "已結案") unresolved++;
      if (!productStats[product]) productStats[product] = { total: 0, unresolved: 0 };
      productStats[product].total++;
      if (status !== "已結案") productStats[product].unresolved++;
      const daysAgo = Math.floor((now - new Date(d)) / 86400000);
      if (daysAgo >= 0 && daysAgo < 30) dateStats[d] = (dateStats[d] || 0) + 1;
    });

    const dateList = [];
    for (let i = 29; i >= 0; i--) {
      const d = new Date(now.getTime() - i * 86400000);
      const dStr = Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd");
      dateList.push({ date: dStr, count: dateStats[dStr] || 0 });
    }

    return {
      day: Object.keys(dayStats).map(function(k) { return { cat: k, count: dayStats[k] }; }),
      month: Object.keys(monthStats).map(function(k) { return { cat: k, count: monthStats[k] }; }),
      unresolved: unresolved,
      products: Object.keys(productStats).map(function(k) { return { product: k, total: productStats[k].total, unresolved: productStats[k].unresolved }; }).sort(function(a,b){return b.total-a.total;}),
      dates: dateList
    };
  } catch (e) {
    return { day: [], month: [], unresolved: 0, products: [], dates: [], error: e.toString() };
  }
}

function updateComplaintProgress(complaintId, newProgress) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf("編碼");
  const progIdx = data[0].indexOf("後續進度記錄");
  const statusIdx = data[0].indexOf("處理狀態");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === complaintId) {
      if (data[i][statusIdx] && data[i][statusIdx].toString().trim() === "已結案") {
        return { status: "error", message: "此案件已結案，無法再更新進度" };
      }
      const time = Utilities.formatDate(new Date(), "GMT+8", "MM/dd HH:mm");
      const old = data[i][progIdx] || "";
      sheet.getRange(i + 1, progIdx + 1).setValue(old + (old ? "\n" : "") + "[" + time + "] " + newProgress);
      return { status: "success" };
    }
  }
  return { status: "error", message: "找不到案件編碼：" + complaintId };
}

function setComplaintStatus(complaintId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf("編碼");
  const statusIdx = data[0].indexOf("處理狀態");
  const progIdx = data[0].indexOf("後續進度記錄");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === complaintId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      const time = Utilities.formatDate(new Date(), "GMT+8", "MM/dd HH:mm");
      const old = data[i][progIdx] || "";
      const log = newStatus === "已結案"
        ? "[" + time + "] ✅ 品牌客服結案"
        : "[" + time + "] 🔄 品牌客服重新開啟案件";
      sheet.getRange(i + 1, progIdx + 1).setValue(old + (old ? "\n" : "") + log);
      return { status: "success" };
    }
  }
  return { status: "error", message: "找不到案件編碼：" + complaintId };
}

// [修正] 派工欄位對應正確版本
function syncToDispatchSystem(id, formData) {
  const dispatchSS = SpreadsheetApp.openById(CONFIG.DISPATCH_SHEET_ID);
  const dispatchSheet = dispatchSS.getSheetByName("派工任務") || dispatchSS.getSheets()[0];

  const now = new Date();
  const deadline = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000); // 3天後

  const solutionLabel = formData.solution === "請汶和解答" ? "品質問題解答" : "延伸供應商協助";
  const description =
    "【客訴編碼】" + id + "\n" +
    "【品牌】" + formData.brand + "\n" +
    "【客人】" + formData.customerName + "\n" +
    "【料號】" + formData.productNo + "\n" +
    "【批號】" + (formData.batchNo || "-") + "\n" +
    "【客訴內容】" + formData.content + "\n" +
    "【解決方向】" + formData.solution;

  // 欄位順序：任務編號/建立時間/申請人姓名/申請人帳號/派工部門/派工對象/派工對象帳號/派工事項/派工說明/期待完成時間/任務狀態
  dispatchSheet.appendRow([
    id,                                                                          // 任務編號
    Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"),                 // 建立時間
    "客服系統",                                                                  // 申請人姓名
    "",                                                                          // 申請人帳號（系統自動，留空）
    "產品技術-品保",                                                             // 派工部門
    "詹雅筑",                                                                    // 派工對象
    "fanny.ppppp@gmail.com",                                                    // 派工對象帳號
    "品牌客訴回覆",                                                              // 派工事項
    description,                                                                 // 派工說明
    Utilities.formatDate(deadline, "GMT+8", "yyyy-MM-dd"),                      // 期待完成時間
    "待接受"                                                                     // 任務狀態
  ]);
}