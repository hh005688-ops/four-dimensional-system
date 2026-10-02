/**
 * =========================================================================
 * 🏛️ 四維系統 - 歷史數據、公式注入與 Beta/Alpha 核心計算 (v5.5.0 動態全自動版)
 * =========================================================================
 */

var CONFIG = {
  MAIN_SHEET: "總覽",
  LOG_SHEET: "每日紀錄",
  DATA_SHEET: "後台數據",
  INDEX_TICKER: "TPE:IX0001",       // 台股加權指數 (B、C 欄)
  NDX_TICKER: "INDEXNASDAQ:NDX",     // 那斯達克 100 指數 (D、E 欄)
  START_COL: 6                       // 自訂個股/ETF 由 F 欄起算 (第 6 欄)
};

function runDailyRecord() {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
  if (!accessCheck.success) {
    SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 這裡可以放你本來想在 runDailyRecord 裡執行的主邏輯（例如呼叫 recordDailySnapshot 等）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  syncOneYearHistory(ss);
}

/**
 * 歷史價格與公式注入（主動動態掃描台美股持倉，全自動重建後台欄位與 Beta）
 */
function syncOneYearHistory(ss) {
  if (typeof resolveSpreadsheet === 'function') {
    ss = resolveSpreadsheet(ss);
  }
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  if (!ss) {
    Logger.log("❌ 找不到可用的試算表，終止執行 syncOneYearHistory");
    return { success: false, message: "❌ 找不到可用的試算表" };
  }

  var dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  var twSheet = ss.getSheetByName("台股區");
  var usSheet = ss.getSheetByName("美股區");

  if (!dataSheet) {
    Logger.log("❌ 找不到工作表：" + CONFIG.DATA_SHEET);
    return { success: false, message: "❌ 找不到工作表：" + CONFIG.DATA_SHEET };
  }

  // 1. 🌟 主動動態從「台股區」與「美股區」抓取最新所有持倉代碼
  var stocks = [];
  
  // 抓取台股 (加上 TPE: 前綴)
  if (twSheet) {
    var twCodes = twSheet.getRange("A2:A7").getValues();
    for (var t = 0; t < twCodes.length; t++) {
      var codeTW = String(twCodes[t][0] || "").trim().toUpperCase();
      if (codeTW !== "") {
        stocks.push(codeTW.startsWith("TPE:") ? codeTW : "TPE:" + codeTW);
      }
    }
  }

  // 抓取美股 (直接使用代號)
  if (usSheet) {
    var usCodes = usSheet.getRange("A2:A7").getValues();
    for (var u = 0; u < usCodes.length; u++) {
      var codeUS = String(usCodes[u][0] || "").trim().toUpperCase();
      if (codeUS !== "") {
        stocks.push(codeUS);
      }
    }
  }

  Logger.log("🔍 動態偵測到 " + stocks.length + " 支最新持倉標的：" + stocks.join(", "));

  var maxCols = dataSheet.getMaxColumns();
  var maxRows = dataSheet.getMaxRows();
  var totalNeededCols = CONFIG.START_COL + (stocks.length * 2) + 2;

  // 自動擴充欄位數（防溢出）
  if (maxCols < totalNeededCols) {
    dataSheet.insertColumnsAfter(maxCols, totalNeededCols - maxCols);
    maxCols = dataSheet.getMaxColumns();
  }

  // 2. 清空整張工作表內容（保留首列格式）
  dataSheet.getRange(1, 1, 1, maxCols).clearContent();
  if (maxRows > 1) {
    dataSheet.getRange(2, 1, maxRows - 1, maxCols).clear();
  }

  // 3. 建立固定的雙大盤標題 (A ~ E 欄)
  dataSheet.getRange(1, 1).setValue("日期");
  dataSheet.getRange(1, 2).setValue("TPE:IX0001");
  dataSheet.getRange(1, 3).setValue("漲跌幅");
  dataSheet.getRange(1, 4).setValue("NDX");
  dataSheet.getRange(1, 5).setValue("NASDAQ 100 報酬率");

  // 4. 動態寫入自訂標的標題
  var returnStartCol = CONFIG.START_COL + stocks.length;
  for (var s = 0; s < stocks.length; s++) {
    dataSheet.getRange(1, CONFIG.START_COL + s).setValue(stocks[s]);
    dataSheet.getRange(1, returnStartCol + s).setValue(stocks[s] + " 報酬率");
  }

  SpreadsheetApp.flush();
  Logger.log("⏳ 正在注入 GoogleFinance 公式與對齊日期序列...");

  // 5. A 欄：動態日期序列 (滾動 1 年 366 天)
  dataSheet.getRange("A2").setFormula('=SEQUENCE(366, 1, TODAY()-365)');
  
  // 6. B、C 欄：台股大盤價格與報酬率
  var taiexFormula = '=ARRAYFORMULA(IFERROR(VLOOKUP(INT(A2:A367), {INT(INDEX(GOOGLEFINANCE("' + CONFIG.INDEX_TICKER + '", "price", TODAY()-375, TODAY()+1),,1)), INDEX(GOOGLEFINANCE("' + CONFIG.INDEX_TICKER + '", "price", TODAY()-375, TODAY()+1),,2)}, 2, FALSE), IFERROR(IF(A2:A367=TODAY(), GOOGLEFINANCE("' + CONFIG.INDEX_TICKER + '", "price"), ""))))';
  dataSheet.getRange("B2").setFormula(taiexFormula);
  var taiexReturn = '=ARRAYFORMULA(IF(B3:B367="", "", IFERROR(B3:B367 / LOOKUP(ROW(B3:B367)-1, ROW(B2:B367)/(B2:B367<>""), B2:B367) - 1, "")))';
  dataSheet.getRange(3, 3).setFormula(taiexReturn);

  // 7. D、E 欄：那斯達克 100 指數價格與報酬率
  var ndxFormula = '=ARRAYFORMULA(IFERROR(VLOOKUP(INT(A2:A367), {INT(INDEX(GOOGLEFINANCE("' + CONFIG.NDX_TICKER + '", "price", TODAY()-375, TODAY()+1),,1)), INDEX(GOOGLEFINANCE("' + CONFIG.NDX_TICKER + '", "price", TODAY()-375, TODAY()+1),,2)}, 2, FALSE), IFERROR(IF(A2:A367=TODAY(), GOOGLEFINANCE("' + CONFIG.NDX_TICKER + '", "price"), ""))))';
  dataSheet.getRange("D2").setFormula(ndxFormula);
  var ndxReturn = '=ARRAYFORMULA(IF(D3:D367="", "", IFERROR(D3:D367 / LOOKUP(ROW(D3:D367)-1, ROW(D2:D367)/(D2:D367<>""), D2:D367) - 1, "")))';
  dataSheet.getRange(3, 5).setFormula(ndxReturn);

  // 8. F 欄開始：自訂標的價格與報酬率注入
  for (var i = 0; i < stocks.length; i++) {
    var stockTicker = stocks[i];
    var colPrice = CONFIG.START_COL + i; 
    var colReturn = returnStartCol + i; 
    var priceColLetter = getColumnLetter(colPrice);

    // 價格公式
    var priceFormula = '=ARRAYFORMULA(IFERROR(VLOOKUP(INT(A2:A367), {INT(INDEX(GOOGLEFINANCE("' + stockTicker + '", "price", TODAY()-375, TODAY()+1),,1)), INDEX(GOOGLEFINANCE("' + stockTicker + '", "price", TODAY()-375, TODAY()+1),,2)}, 2, FALSE), IFERROR(IF(A2:A367=TODAY(), GOOGLEFINANCE("' + stockTicker + '", "price"), ""))))';
    dataSheet.getRange(2, colPrice).setFormula(priceFormula);
    
    // 報酬率公式
    var returnFormula = '=ARRAYFORMULA(IF(' + priceColLetter + '3:' + priceColLetter + '367="", "", IFERROR(' + priceColLetter + '3:' + priceColLetter + '367 / LOOKUP(ROW(' + priceColLetter + '3:' + priceColLetter + '367)-1, ROW(' + priceColLetter + '2:' + priceColLetter + '367)/(' + priceColLetter + '2:' + priceColLetter + '367<>""), ' + priceColLetter + '2:' + priceColLetter + '367) - 1, "")))';
    dataSheet.getRange(3, colReturn).setFormula(returnFormula);
  }

  // 9. 精準格式化設定
  dataSheet.getRange("A2:A367").setNumberFormat("yyyy/MM/dd");
  dataSheet.getRange("B2:B367").setNumberFormat("#,##0.00");
  dataSheet.getRange("C3:C367").setNumberFormat("0.00%");
  dataSheet.getRange("D2:D367").setNumberFormat("#,##0.00");
  dataSheet.getRange("E3:E367").setNumberFormat("0.00%");
  
  if (stocks.length > 0) {
    dataSheet.getRange(2, CONFIG.START_COL, 366, stocks.length).setNumberFormat("#,##0.00");
    dataSheet.getRange(3, returnStartCol, 365, stocks.length).setNumberFormat("0.00%");
  }
  
  SpreadsheetApp.flush();
  Utilities.sleep(1500); // 等待 Google Finance 公式取得數據

  // =========================================================================
  // ⚡ 10. 後台純數值計算：批次寫入台股與美股的 Beta、整體加權 Beta 與 Alpha
  // =========================================================================
  calculateAndWriteBetaAlpha(ss, dataSheet);

  try {
    ss.toast("🔄 歷史數據、Beta 與 Alpha 已全自動刷新完成！", "四維決策中心", 4);
  } catch(e) {}
  Logger.log("✅ 歷史數據與 Beta/Alpha 刷新完成！");

  return { success: true, message: "⚡ 台美股歷史數據與 Beta/Alpha 已全數刷新計算完成！" };
}

/**
 * ⚡ 內部核心函式：純數值 Slope (Beta) 與 Alpha 記憶體批次運算
 */
function calculateAndWriteBetaAlpha(ss, dataSheet) {
  var lastRow = dataSheet.getLastRow();
  var lastCol = dataSheet.getLastColumn();
  if (lastRow < 3) return;

  var rawHeader = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rawData = dataSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  var twBenchIdx = 2; // C 欄 (Index 2): 台股加權報酬率
  var usBenchIdx = 4; // E 欄 (Index 4): NASDAQ 100 報酬率

  var colMap = {};
  for (var c = 0; c < rawHeader.length; c++) {
    var h = String(rawHeader[c] || "").trim();
    if (h) colMap[h] = c;
  }

  function calculateSlope(stockIdx, benchIdx) {
    if (stockIdx === undefined || benchIdx === undefined) return 0;
    var sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, count = 0;

    for (var r = 0; r < rawData.length; r++) {
      var y = Number(rawData[r][stockIdx]);
      var x = Number(rawData[r][benchIdx]);

      if (!isNaN(x) && !isNaN(y) && rawData[r][stockIdx] !== "" && rawData[r][benchIdx] !== "") {
        sumX += x;
        sumY += y;
        sumXY += x * y;
        sumX2 += x * x;
        count++;
      }
    }
    if (count < 2) return 0;
    var denominator = (count * sumX2 - sumX * sumX);
    if (denominator === 0) return 0;
    var slope = (count * sumXY - sumX * sumY) / denominator;
    return Math.round(slope * 10) / 10;
  }

  // 🇹🇼 台股區計算回寫
  var twSheet = ss.getSheetByName("台股區");
  if (twSheet) {
    var twStockCodes = twSheet.getRange("A2:A7").getValues();
    var twMarketVals = twSheet.getRange("F2:F7").getValues();
    var twBetaWrites = [];
    var totalTwMarketVal = 0;
    var weightedTwBetaSum = 0;

    for (var i = 0; i < twStockCodes.length; i++) {
      var code = String(twStockCodes[i][0] || "").trim().toUpperCase();
      var mVal = Number(twMarketVals[i][0]) || 0;
      totalTwMarketVal += mVal;

      if (!code) {
        twBetaWrites.push([""]);
        continue;
      }

      var cIdx = colMap["TPE:" + code + " 報酬率"];
      if (cIdx === undefined) cIdx = colMap[code + " 報酬率"];

      var beta = (cIdx !== undefined) ? calculateSlope(cIdx, twBenchIdx) : 0;
      twBetaWrites.push([beta]);
      weightedTwBetaSum += beta * mVal;
    }

    twSheet.getRange("L2:L7").setValues(twBetaWrites);

    var overallTwBeta = totalTwMarketVal > 0 ? Math.round((weightedTwBetaSum / totalTwMarketVal) * 10) / 10 : 0;
    twSheet.getRange("C10").setValue(overallTwBeta);

    var twTodayRet = Number(twSheet.getRange("E10").getValue()) || 0;
    var twBenchTodayRet = Number(twSheet.getRange("A12").getValue()) || 0;
    var twAlpha = Math.round((twTodayRet - (overallTwBeta * twBenchTodayRet)) * 100) / 100;
    twSheet.getRange("G10").setValue(twAlpha);
  }

  // 🇺🇸 美股區計算回寫
  var usSheet = ss.getSheetByName("美股區");
  if (usSheet) {
    var usStockCodes = usSheet.getRange("A2:A7").getValues();
    var usMarketVals = usSheet.getRange("F2:F7").getValues();
    var usBetaWrites = [];
    var totalUsMarketVal = 0;
    var weightedUsBetaSum = 0;

    for (var j = 0; j < usStockCodes.length; j++) {
      var uCode = String(usStockCodes[j][0] || "").trim().toUpperCase();
      var uVal = Number(usMarketVals[j][0]) || 0;
      totalUsMarketVal += uVal;

      if (!uCode) {
        usBetaWrites.push([""]);
        continue;
      }

      var uIdx = colMap[uCode + " 報酬率"];
      var uBeta = (uIdx !== undefined) ? calculateSlope(uIdx, usBenchIdx) : 0;
      usBetaWrites.push([uBeta]);
      weightedUsBetaSum += uBeta * uVal;
    }

    usSheet.getRange("L2:L7").setValues(usBetaWrites);

    var overallUsBeta = totalUsMarketVal > 0 ? Math.round((weightedUsBetaSum / totalUsMarketVal) * 10) / 10 : 0;
    usSheet.getRange("C10").setValue(overallUsBeta);

    var usTodayRet = Number(usSheet.getRange("E10").getValue()) || 0;
    var usBenchTodayRet = Number(rawData[rawData.length - 1][usBenchIdx]) || 0;
    var usAlpha = Math.round((usTodayRet - (overallUsBeta * usBenchTodayRet)) * 100) / 100;
    usSheet.getRange("G10").setValue(usAlpha);
  }
}

/**
 * 欄位數字轉英文代號輔助函式
 */
function getColumnLetter(col) {
  var temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter;
}