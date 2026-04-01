const CONFIG = {
  MASTER_SHEET_NAME: "品牌客訴總表",
  MENU_SHEET_NAME: "選單設定表",
  DISPATCH_SHEET_ID: "1hbkxLEKKXeOQl2VQQYKo1bA4WVMjZjVNeEZGa89wsAs",
  IMAGE_FOLDER_ID: "1yKe1SGfCKkiQI1yO0Lfe_6I9zl9aVPyt"
};

function doGet(e) {
  const brand = e.parameter.brand || "ALL";
  const page = e.parameter.page || "index";
  if (page === "report") {
    const template = HtmlService.createTemplateFromFile('Report');
    return template.evaluate()
      .setTitle("客訴分析報表")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.brand = brand;
  return template.evaluate()
    .setTitle("客服記錄系統 - " + brand)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== 選單讀取（跳過第一列 header）=====
// 選單設定表：A=品牌, B=料號, C=客訴分類, D=解決方向, E=子欄位定義, F=客訴平台, G=購買平台
function getProductList(brand) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MENU_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    data.shift(); // 跳過第一列 header
    const products = [];
    data.forEach(function(row) {
      if (!row[0] || !row[1]) return;
      const rowBrand = row[0].toString().replace(/[\s\u3000]/g, "");
      const targetBrand = brand.replace(/[\s\u3000]/g, "");
      if (brand === "ALL" || rowBrand === targetBrand) products.push(row[1].toString().trim());
    });
    return products;
  } catch (e) { return []; }
}

function getMenuOptions() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MENU_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    data.shift(); // 跳過第一列 header
    const categories = [], solutions = [], solutionFields = {}, platforms = [], buyPlatforms = [];
    data.forEach(function(row) {
      if (row[2] && row[2].toString().trim()) categories.push(row[2].toString().trim());
      if (row[3] && row[3].toString().trim()) {
        const sol = row[3].toString().trim();
        solutions.push(sol);
        // E欄子欄位定義：用 | 分隔
        if (row[4] && row[4].toString().trim()) {
          solutionFields[sol] = row[4].toString().trim().split('|').map(function(s){ return s.trim(); }).filter(Boolean);
        } else {
          solutionFields[sol] = [];
        }
      }
      if (row[5] && row[5].toString().trim()) platforms.push(row[5].toString().trim());
      if (row[6] && row[6].toString().trim()) buyPlatforms.push(row[6].toString().trim());
    });
    return {
      categories: [...new Set(categories)],
      solutions: [...new Set(solutions)],
      solutionFields: solutionFields,
      platforms: [...new Set(platforms)],
      buyPlatforms: [...new Set(buyPlatforms)]
    };
  } catch (e) { return { categories: [], solutions: [], solutionFields: {}, platforms: [], buyPlatforms: [] }; }
}

// ===== 主要送出 =====
function processForm(formData) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { status: "error", message: "系統忙碌，請稍後再試" }; }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
    const lastId = generateID(sheet, todayStr, formData.brand);

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

    // A~AE 共31欄
    sheet.appendRow([
      formData.brand,           // A 品牌
      lastId,                   // B 編碼
      formData.caseType || "客訴處理", // C 處理類型
      formData.date,            // D 建立日期
      formData.customerName,    // E 客人名稱
      formData.orderNo || "",   // F 訂單編號
      formData.platform || "",  // G 客訴平台
      formData.buyPlatform || "", // H 購買平台
      formData.productNo || "", // I 客訴產品
      formData.batchNo || "",   // J 產品批號
      formData.category || "",  // K 客訴分類
      formData.subType || "",   // L 子類型/原因
      formData.solution || "",  // M 解決方向
      formData.content || "",   // N 客訴內容
      formData.comment || "",   // O 第一階段補充說明
      formData.deliveryInfo || "", // P 收件資訊（保留相容）
      imageUrls.join(", "),     // Q 圖片網址清單
      formData.videoLinks || "", // R 影片連結
      formData.needDispatch ? "是" : "否", // S 派工汶和
      "未解決",                  // T 處理狀態
      Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"), // U 建立時間
      "",                       // V 後續進度記錄
      formData.recipientName || "",    // W 收件人
      formData.recipientAddress || "", // X 收件地址
      formData.recipientTime || "",    // Y 方便收件時間
      formData.exchangeItem || "",     // Z 換貨品項
      formData.resendItem || "",       // AA 補寄品項
      formData.quantity || "",         // AB 數量
      formData.compensationNote || "", // AC 補償說明
      formData.noteContent || "",      // AD 備註
      formData.discountNote || ""      // AE 折扣說明
    ]);

    if (formData.needDispatch === true) {
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

// [修改] 編碼加品牌：20260401WA-TLV-001
function generateID(sheet, todayStr, brand) {
  const brandCode = (brand || "WA").replace(/[\s\u3000]/g, "");
  const prefix = todayStr + "WA-" + brandCode + "-";
  const data = sheet.getDataRange().getValues();
  let count = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] && data[i][1].toString().startsWith(prefix)) {
      const parts = data[i][1].toString().split("-");
      const last = parseInt(parts[parts.length - 1]);
      if (!isNaN(last)) { count = last + 1; break; }
    }
  }
  return prefix + count.toString().padStart(3, '0');
}

function toDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
  return val.toString().slice(0, 10);
}

// ===== 查詢 =====
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
    let dayStats = {}, monthStats = {}, unresolved = 0, productStats = {}, dateStats = {};

    data.forEach(function(row) {
      if (!row[0] || !row[3]) return;
      if (brand !== "ALL" && row[0].toString().replace(/[\s\u3000]/g, "") !== brand.replace(/[\s\u3000]/g, "")) return;
      const d = toDateStr(row[3]);
      if (!d) return;
      const m = d.slice(0, 7);
      const cat = row[10] ? row[10].toString().trim() : "未分類"; // K 客訴分類
      const status = row[19] ? row[19].toString().trim() : "未解決"; // T 處理狀態
      const product = row[8] ? row[8].toString().trim() : "未知料號"; // I 客訴產品

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

// ===== 更新進度 =====
function updateComplaintProgress(complaintId, newProgress) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf("編碼");
  const progIdx = data[0].indexOf("後續進度記錄");
  const statusIdx = data[0].indexOf("處理狀態");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === complaintId) {
      if (data[i][statusIdx] && data[i][statusIdx].toString().trim() === "已結案") {
        return { status: "error", message: "此案件已結案" };
      }
      const time = Utilities.formatDate(new Date(), "GMT+8", "MM/dd HH:mm");
      const old = data[i][progIdx] || "";
      sheet.getRange(i + 1, progIdx + 1).setValue(old + (old ? "\n" : "") + "[" + time + "] " + newProgress);
      return { status: "success" };
    }
  }
  return { status: "error", message: "找不到案件：" + complaintId };
}

function setComplaintStatus(complaintId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idIdx = data[0].indexOf("編碼");
  const statusIdx = data[0].indexOf("處理狀態");
  const progIdx = data[0].indexOf("後續進度記錄");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === complaintId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      const time = Utilities.formatDate(new Date(), "GMT+8", "MM/dd HH:mm");
      const old = data[i][progIdx] || "";
      const log = newStatus === "已結案" ? "[" + time + "] ✅ 結案" : "[" + time + "] 🔄 重新開啟";
      sheet.getRange(i + 1, progIdx + 1).setValue(old + (old ? "\n" : "") + log);
      return { status: "success" };
    }
  }
  return { status: "error", message: "找不到案件：" + complaintId };
}

// ===== 派工 =====
function syncToDispatchSystem(id, formData) {
  const dispatchSS = SpreadsheetApp.openById(CONFIG.DISPATCH_SHEET_ID);
  const dispatchSheet = dispatchSS.getSheetByName("派工任務") || dispatchSS.getSheets()[0];
  const now = new Date();
  const deadline = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
  const description =
    "【客訴編碼】" + id + "\n【品牌】" + formData.brand +
    "\n【客人】" + formData.customerName +
    "\n【料號】" + (formData.productNo || "-") +
    "\n【批號】" + (formData.batchNo || "-") +
    "\n【內容】" + (formData.content || "-");
  dispatchSheet.appendRow([
    id,
    Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"),
    "客服系統", "",
    "產品技術-品保", "詹雅筑", "fanny.ppppp@gmail.com",
    "品牌客訴回覆", description,
    Utilities.formatDate(deadline, "GMT+8", "yyyy-MM-dd"),
    "待接受"
  ]);
}