/**
 * 智光商工 資處一仁 點餐系統 - Google Apps Script 雲端後端程式
 * 
 * 運作說明：
 * 1. doGet: 當網頁請求讀取時，自動從試算表讀取所有已點餐學生資料回傳。
 * 2. doPost: 當學生點餐、取消點餐或老師重置時，自動寫入或清除試算表資料。
 */

const SHEET_NAME = "點餐明細";

// 讀取全班最新訂餐資料 (GET)
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const orders = {};
    
    // 第一行為標題列，從第二行 (index 1) 開始讀取
    for (let i = 1; i < data.length; i++) {
      const seat = String(data[i][0]).padStart(2, '0');
      const itemName = data[i][2];
      if (seat && itemName) {
        orders[seat] = {
          seat: seat,
          name: data[i][1] || "",
          item: itemName,
          price: parseInt(data[i][3]) || 0,
          time: data[i][4] || ""
        };
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      count: Object.keys(orders).length,
      orders: orders,
      timestamp: new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" })
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 寫入、取消或重置訂餐資料 (POST)
function doPost(e) {
  try {
    let postData = {};
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    }
    
    const action = postData.action;
    const sheet = getOrCreateSheet();
    
    // 1. 學生點餐或修改點餐
    if (action === "order") {
      const seat = String(postData.seat).padStart(2, '0');
      const name = postData.name || "";
      const item = postData.item || "";
      const price = parseInt(postData.price) || 0;
      const time = postData.time || new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" });
      
      const data = sheet.getDataRange().getValues();
      let foundRow = -1;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).padStart(2, '0') === seat) {
          foundRow = i + 1; // 試算表列號從 1 開始
          break;
        }
      }
      
      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 5).setValues([[seat, name, item, price, time]]);
      } else {
        sheet.appendRow([seat, name, item, price, time]);
      }
      
      // 自動依照座號排序（由小到大）
      sortSheetBySeat(sheet);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: `座號 ${seat} 訂餐已記錄`
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. 取消點餐
    else if (action === "cancel") {
      const seat = String(postData.seat).padStart(2, '0');
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).padStart(2, '0') === seat) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: `座號 ${seat} 訂餐已取消`
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 3. 重置所有訂單
    else if (action === "reset") {
      sheet.clearContents();
      sheet.appendRow(["座號", "姓名", "訂購餐點", "金額", "訂餐時間"]);
      formatHeaderRow(sheet);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "全班訂餐紀錄已重置清空"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "未知的 action 操作指令"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 取得或建立「點餐明細」分頁
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["座號", "姓名", "訂購餐點", "金額", "訂餐時間"]);
    formatHeaderRow(sheet);
  }
  return sheet;
}

// 格式化標題列
function formatHeaderRow(sheet) {
  const header = sheet.getRange(1, 1, 1, 5);
  header.setBackground("#ff85a2");
  header.setFontColor("#ffffff");
  header.setFontWeight("bold");
  header.setHorizontalAlignment("center");
}

// 自動按座號排序
function sortSheetBySeat(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, 5);
    range.sort({ column: 1, ascending: true });
  }
}
