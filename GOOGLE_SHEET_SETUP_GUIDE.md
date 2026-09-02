# 智光商工 資處一仁 點餐系統 - Google 試算表雲端同步教學 (3 分鐘完成)

## 📌 為什麼先前老師端看到 0 人點餐？
- **原因**：GitHub Pages 是靜態網頁空間，先前點餐資料只儲存在各個學生手機的「瀏覽器本地儲存區 (LocalStorage)」內，沒有上傳至雲端資料庫。所以每位學生只能看到自己的手機紀錄，老師的電腦上自然是 0 人。
- **解決方案**：透過 **Google 試算表 (Google Apps Script)** 免費建立專屬的班級雲端後端，31 位學生手機與老師電腦秒級雙向同步！

---

## 🚀 3 步驟設定 Google 試算表雲端同步

### 第一步：建立 Google 試算表
1. 用您的 Google 帳號開啟 [Google 試算表 (點此新建空白試算表)](https://sheets.new)。
2. 將試算表命名為：`資處一仁點餐紀錄表`。

---

### 第二步：貼上 Apps Script 程式碼
1. 點選試算表上方選單的 **「擴充功能」 ➔ 「Apps Script」**。
2. 將編輯器內原本的程式碼全部刪除，**完整貼上**以下程式碼：

```javascript
const SHEET_NAME = "點餐明細";

function doGet(e) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const orders = {};
  for (let i = 1; i < data.length; i++) {
    const seat = String(data[i][0]).padStart(2, '0');
    if (seat && data[i][2]) {
      orders[seat] = {
        name: data[i][1],
        item: data[i][2],
        price: data[i][3],
        time: data[i][4]
      };
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: "success", orders: orders }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const sheet = getOrCreateSheet();
    
    if (action === "order") {
      const seat = String(postData.seat).padStart(2, '0');
      const name = postData.name;
      const item = postData.item;
      const price = postData.price;
      const time = postData.time;
      
      const data = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).padStart(2, '0') === seat) {
          foundRow = i + 1;
          break;
        }
      }
      
      if (foundRow > 0) {
        sheet.getRange(foundRow, 1, 1, 5).setValues([[seat, name, item, price, time]]);
      } else {
        sheet.appendRow([seat, name, item, price, time]);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "訂單已儲存" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "cancel") {
      const seat = String(postData.seat).padStart(2, '0');
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).padStart(2, '0') === seat) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "訂單已取消" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "reset") {
      sheet.clearContents();
      sheet.appendRow(["座號", "姓名", "訂購餐點", "金額", "訂餐時間"]);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "全部重置完成" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["座號", "姓名", "訂購餐點", "金額", "訂餐時間"]);
  }
  return sheet;
}
```

3. 點擊上方的 **「儲存 💾」** 圖示。

---

### 第三步：發布部署為網頁應用程式
1. 點擊右上角的 **「部署」 ➔ 「新增部署」**。
2. 點選左側齒輪圖示 ➔ 選擇 **「網頁應用程式 (Web App)」**。
3. 說明填寫：`資處一仁點餐同步API`。
4. **【最關鍵步驟】**：
   - 執行身分：**我 (您的 Google 帳號)**
   - 誰可以存取：**所有人 (Anyone)**
5. 點擊 **「部署」** ➔ 如跳出授權提示，點擊「審查權限」➔ 選擇您的 Google 帳號 ➔ 點「Advanced (進階)」➔ 點「Go to (前往專案)」➔ 點「Allow (允許)」。
6. 複製彈出視窗中的 **「網頁應用程式網址 (Web App URL)」** (通常為 `https://script.google.com/macros/s/AKfycb.../exec`)。

---

### 第四步：貼上網址開始使用
1. 開啟點餐網頁 `seat.html`。
2. 點擊上方工具列的 **「☁️ 雲端同步設定」**。
3. 將剛才複製的網址貼上，點擊 **「💾 儲存並測試連線」**。
4. 工具列狀態變為 **「🟢 雲端已同步」** 即代表連線成功！
5. 現在無論哪位學生在手機點餐，老師的電腦與所有學生手機都會在數秒內自動同步更新，且老師的 Google 試算表也會即時自動記錄！
