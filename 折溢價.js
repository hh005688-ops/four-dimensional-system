/**
 * 📈 正二 ETF 折溢價紀錄
 * 把今天的折溢價(%)寫進「Beta紀錄」分頁的 R 欄（對應 M~Q 欄那 4 檔正二）
 * 資料來源：鉅亨網「淨值與折溢價」頁面
 * 折溢價(%) = (市價 ÷ 淨值 − 1) × 100
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
 * 從鉅亨網「淨值與折溢價」頁面抓 ETF 資料，回傳 { "00631L": 0.45, ... }（單位 %）
 * 表格每列的順序：代碼、名稱、市價、漲跌、淨值、漲跌、折溢價(元)、折溢價(%)
 * 折溢價(%) = (市價 ÷ 淨值 − 1) × 100，正數 = 溢價、負數 = 折價
 */
function fetchEtfPremiums_(codes) {
  var result = {};
  try {
    var res = UrlFetchApp.fetch("https://www.cnyes.com/twstock/etf/detail/discount-premium", {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
        "Accept-Language": "zh-TW,zh;q=0.9"
      }
    });
    if (res.getResponseCode() !== 200) {
      Logger.log("⚠️ 折溢價來源回應異常：" + res.getResponseCode());
      return result;
    }
    var html = res.getContentText();

    for (var c = 0; c < codes.length; c++) {
      var code = String(codes[c]).toUpperCase();
      var at = html.indexOf(">" + code + "<");
      if (at === -1) {
        Logger.log("⚠️ 折溢價頁面找不到 " + code);
        continue;
      }
      var end = html.indexOf("</tr>", at);
      var rowHtml = html.substring(at, end === -1 ? at + 6000 : end);

      // 取出這一列所有的文字格：[代碼, 名稱, 市價, 漲跌, 淨值, 漲跌, 折溢價(元), 折溢價(%)...]
      var cells = [];
      var re = />([^<>]+)</g, m;
      while ((m = re.exec(rowHtml)) !== null) {
        var t = m[1].trim();
        if (t !== "") cells.push(t);
      }
      var price = Number(String(cells[2]).replace(/,/g, ""));
      var nav = Number(String(cells[4]).replace(/,/g, ""));
      var p;
      if (price > 0 && nav > 0) {
        p = (price / nav - 1) * 100;
      } else {
        p = Number(String(cells[6]).replace(/[%+,]/g, ""));   // 退而求其次：直接讀頁面上的 %
      }
      if (!isNaN(p)) result[code] = Math.round(p * 100) / 100;
      Logger.log("折溢價原始資料 " + code + "：" + JSON.stringify(cells.slice(0, 8)));
    }
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
