/**
 * 📈 正二 ETF 折溢價紀錄（正式淨值版）
 *
 * 把折溢價(%)寫進「Beta紀錄」分頁的 R 欄（對應 M~Q 欄那 4 檔正二）
 * 資料來源：MoneyDJ「近30日淨值」頁面（每天一筆：淨值、市價、折溢價）
 * 折溢價(%) = (市價 ÷ 淨值 − 1) × 100，正數 = 溢價、負數 = 折價
 *
 * 注意：每天的正式淨值要收盤後、晚上才會公布，所以「當天」的折溢價要等隔天才有。
 * 程式每次執行都會把 R 欄裡「日期有對上、而且 MoneyDJ 已經有資料」的列補上，
 * 所以不用特別改排程時間，晚一天的資料會自動補進去。
 * 當天那一列在正式淨值公布前會保持空白（不會放不準的數字）。
 * 失敗時只記 Log，不會影響原本的 Beta 紀錄。
 */
function appendPremiumLog_(ss) {
  var sheet = ss.getSheetByName("Beta紀錄");
  if (!sheet) return;

  if (String(sheet.getRange("R1").getValue()).trim() === "") {
    sheet.getRange("R1").setValue("折溢價(%)");
  }

  var last = sheet.getLastRow();
  if (last < 2) return;

  // 每檔 ETF 抓一次：{ "00631L": { "2026-11-03": 0.45, ... }, ... }
  var official = {};
  for (var c = 0; c < ETF2X_WATCH.length; c++) {
    official[ETF2X_WATCH[c]] = fetchOfficialPremiums_(ETF2X_WATCH[c]);
  }

  var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var keys = sheet.getRange(2, 13, last - 1, 2).getValues();   // M 日期、N 標的
  var out = sheet.getRange(2, 18, last - 1, 1).getValues();    // R 折溢價
  var changed = false;

  for (var i = 0; i < keys.length; i++) {
    var v = keys[i][0];
    if (v === "" || v === null) continue;
    var ds = (v instanceof Date) ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : String(v).trim();
    var code = String(keys[i][1]).trim().toUpperCase();
    var table = official[code];
    if (!table) continue;

    if (table[ds] !== undefined) {
      if (out[i][0] !== table[ds]) {
        out[i][0] = table[ds];
        changed = true;
      }
    } else if (ds === today && out[i][0] !== "") {
      out[i][0] = "";   // 當天正式淨值還沒公布，清掉舊的不準數字
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(2, 18, last - 1, 1).setNumberFormat("0.00");
    sheet.getRange(2, 18, last - 1, 1).setValues(out);
  }
}

/**
 * 抓玩股網單一檔 ETF 的折溢價表，回傳 { "2026-11-03": 0.45, ... }（單位 %）
 */
function fetchOfficialPremiums_(code) {
  var result = {};
  try {
    var url = "https://www.moneydj.com/etf/x/basic/basic0003.xdjhtm?etfid=" + String(code).toLowerCase() + ".tw";
    var res = UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
        "Accept-Language": "zh-TW,zh;q=0.9"
      }
    });
    if (res.getResponseCode() !== 200) {
      Logger.log("⚠️ 折溢價來源回應異常 " + code + "：" + res.getResponseCode());
      return result;
    }
    var html = res.getContentText();
    result = parseOfficialPremiums_(html);
    if (Object.keys(result).length === 0) {
      var plain = html.replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<[^>]+>/g, "|").replace(/\s+/g, " ");
      var at = Math.max(0, plain.indexOf("折溢價"));
      Logger.log("⚠️ 折溢價頁面讀不到表格 " + code + "，頁面片段：" + plain.substring(at, at + 600));
    }
  } catch (e) {
    Logger.log("⚠️ 折溢價抓取失敗 " + code + "：" + e.message);
  }
  return result;
}

/**
 * 從頁面文字找出每一天的資料。
 * 每列的順序：日期、淨值、市價、折溢價(%)
 * 不依賴固定的網頁標籤，只看「日期 後面跟著 3 個數字」，
 * 並且檢查「折溢價 = (市價 ÷ 淨值 − 1) × 100」有對上才收，讀錯的列會自動略過。
 */
function parseOfficialPremiums_(html) {
  var out = {};
  var text = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, "|")
    .replace(/&nbsp;/g, " ")
    .replace(/−/g, "-");   // 把特殊的減號換成一般減號

  var num = "([+-]?[\\d,]*\\.?\\d+)";
  var sep = "[\\s|]+";
  var re = new RegExp(
    "(\\d{4}[-/]\\d{2}[-/]\\d{2})" + sep + num + sep + num + sep + num + "\\s*%?", "g");

  var m;
  while ((m = re.exec(text)) !== null) {
    var date = m[1].replace(/\//g, "-");
    if (out[date] !== undefined) continue;   // 同一天只取第一筆
    var a = Number(m[2].replace(/,/g, ""));
    var b = Number(m[3].replace(/,/g, ""));
    var pct = Number(m[4].replace(/,/g, ""));
    if (!(a > 0 && b > 0) || isNaN(pct)) continue;
    // 驗算：折溢價要等於 (市價÷淨值−1)×100（欄位順序對調也能容許）
    var ok = Math.abs((b / a - 1) * 100 - pct) <= 0.06 ||
             Math.abs((a / b - 1) * 100 - pct) <= 0.06;
    if (!ok) continue;
    out[date] = Math.round(pct * 100) / 100;
  }
  return out;
}

/**
 * 🧩 手動補填：在 Apps Script 編輯器選這個函式按「執行」，
 * 會把 R 欄所有已經有正式淨值的日期補上（不會動到 Beta 的數字）
 */
function backfillPremium() {
  appendPremiumLog_(SpreadsheetApp.getActiveSpreadsheet());
}

/**
 * 🧪 手動測試用：選這個函式按「執行」，看「執行記錄」有沒有 4 檔正二、各個日期的折溢價
 */
function testPremium() {
  for (var i = 0; i < ETF2X_WATCH.length; i++) {
    Logger.log(ETF2X_WATCH[i] + "：" + JSON.stringify(fetchOfficialPremiums_(ETF2X_WATCH[i])));
  }
}
