/**
 * 📈 正二 ETF 折溢價紀錄
 * 把今天的折溢價(%)寫進「Beta紀錄」分頁的 R 欄（對應 M~Q 欄那 4 檔正二）
 * 資料來源：證交所公開的 ETF 即時資料（all_etf.txt）
 * 折溢價(%) = (成交價 ÷ 預估淨值 − 1) × 100
 * 失敗時只記 Log，不會影響原本的 Beta 紀錄
 */
function appendPremiumLog_(ss) {
  var sheet = ss.getSheetByName("Beta紀錄");
  if (!sheet) return;

  var premiums = fetchEtfPremiums_(ETF2X_WATCH);

  if (String(sheet.getRange("R1").getValue()).trim() === "") {
    sheet.getRange("R1").setValue("折溢價(%)");
  }

  var last = sheet.getLastRow();
  if (last < 2) return;

  var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var keys = sheet.getRange(2, 13, last - 1, 2).getValues();   // M 日期、N 標的
  var out = sheet.getRange(2, 18, last - 1, 1).getValues();    // R 折溢價
  var changed = false;

  for (var i = 0; i < keys.length; i++) {
    var v = keys[i][0];
    if (v === "" || v === null) continue;
    var ds = (v instanceof Date) ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : String(v).trim();
    if (ds !== today) continue;
    var code = String(keys[i][1]).trim().toUpperCase();
    if (premiums[code] !== undefined) {
      out[i][0] = premiums[code];
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(2, 18, last - 1, 1).setNumberFormat("0.00");
    sheet.getRange(2, 18, last - 1, 1).setValues(out);
  }
}

/**
 * 從證交所抓 ETF 資料，回傳 { "00631L": 0.15, ... }（單位 %）
 */
function fetchEtfPremiums_(codes) {
  var result = {};
  try {
    var res = UrlFetchApp.fetch("https://mis.twse.com.tw/stock/data/all_etf.txt", {
      method: "get",
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0", "Accept": "application/json" }
    });
    if (res.getResponseCode() !== 200) {
      Logger.log("⚠️ 折溢價來源回應異常：" + res.getResponseCode());
      return result;
    }
    var json = JSON.parse(res.getContentText());

    function walk(node) {
      if (Array.isArray(node)) {
        for (var n = 0; n < node.length; n++) walk(node[n]);
      } else if (node && typeof node === "object") {
        if (node.a !== undefined && node.e !== undefined && node.f !== undefined) {
          handle(node);
        } else {
          for (var key in node) walk(node[key]);
        }
      }
    }

    function handle(item) {
      var code = String(item.a).trim().toUpperCase();
      if (codes.indexOf(code) === -1) return;
      var price = Number(item.e);
      var nav = Number(item.f);
      var p;
      if (price > 0 && nav > 0) {
        p = (price / nav - 1) * 100;
      } else {
        p = Number(item.g);   // 退而求其次：直接用來源給的折溢價欄位
      }
      if (!isNaN(p)) result[code] = Math.round(p * 100) / 100;
      Logger.log("折溢價原始資料 " + code + "：" + JSON.stringify(item));
    }

    walk(json);
  } catch (e) {
    Logger.log("⚠️ 折溢價抓取失敗：" + e.message);
  }
  return result;
}

/**
 * 🧪 手動測試用：在 Apps Script 編輯器選這個函式按「執行」，
 * 再看「執行記錄」有沒有 4 檔正二的折溢價數字
 */
function testPremium() {
  Logger.log(JSON.stringify(fetchEtfPremiums_(ETF2X_WATCH)));
}
