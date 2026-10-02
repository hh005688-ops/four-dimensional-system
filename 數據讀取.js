/**
 * 🛠️ 共用安全數字解析防呆
 */
function parseNum_(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  var clean = String(val).replace(/[^0-9.-]+/g, "");
  var num = Number(clean);
  return isNaN(num) ? 0 : num;
}

/**
 * 🏛️ 四維系統：首頁儀表板與持倉數據讀取模組
 */
function getInitDataUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表，請檢查 ID 或授權。" };

  try {
    SpreadsheetApp.flush();
    var mainSheet = ss.getSheetByName("總覽");
    var currentCash = 0, totalAssets = 0, totalLiabilities = 0, netWorth = 0;
    var pledgeMaintenance = 0, realLeverage = 0, assetBeta = 0, todayProfit = 0;

    if (mainSheet) {
      var ov = mainSheet.getRange("A1:E15").getValues();
      function cell_(rows, r, c) {
        return (rows && rows[r] && rows[r][c] !== undefined) ? rows[r][c] : 0;
      }
      currentCash         = Math.round(parseNum_(cell_(ov, 4, 1)));   // B5
      totalAssets         = Math.round(parseNum_(cell_(ov, 5, 1)));   // B6
      totalLiabilities    = Math.round(parseNum_(cell_(ov, 11, 1)));  // B12
      netWorth            = Math.round(parseNum_(cell_(ov, 14, 0)));  // A15
      var rawMaint        = parseNum_(cell_(ov, 14, 2));              // C15
      pledgeMaintenance   = rawMaint <= 10 ? +(rawMaint * 100).toFixed(2) : +rawMaint.toFixed(2);
      realLeverage        = +(parseNum_(cell_(ov, 3, 4))).toFixed(2); // E4
      assetBeta           = +(parseNum_(cell_(ov, 5, 4))).toFixed(2); // E6
      todayProfit         = Math.round(parseNum_(cell_(ov, 14, 1)));  // B15
    }

    var stocks = [];
    var twSheet = ss.getSheetByName("台股區");
    if (twSheet) {
      // 🎯 擴大範圍至 K 欄 (A2:K7) 以讀取目標佔比
      var twData = twSheet.getRange("A2:K7").getValues();
      for (var i = 0; i < twData.length; i++) {
        if (isStockAreaBoundary_(twData[i][0])) break;
        var code = String(twData[i][0]).trim();
        var name = String(twData[i][1]).trim();
        var qty = parseNum_(twData[i][4]);
        var mVal = parseNum_(twData[i][5]);
        var pQty = parseNum_(twData[i][8]);
        var tRatio = parseNum_(twData[i][10]) * 100; // K 欄 (第11欄，索引10) 目標佔比

        if (isInventoryHoldingRow_(twData[i][0], qty)) {
          stocks.push({
            market: "TW",
            code: code,
            name: name,
            totalQty: qty,
            pledgedQty: pQty,
            unpledgedQty: Math.max(0, qty - pQty),
            marketVal: mVal,
            targetRatio: tRatio, // 🎯 傳遞 K 欄數值給前端
            row: 2 + i
          });
        }
      }
    }

    var usSheet = ss.getSheetByName("美股區");
    if (usSheet) {
      // 🎯 擴大範圍至 K 欄 (A2:K7) 以讀取目標佔比
      var usData = usSheet.getRange("A2:K7").getValues();
      for (var j = 0; j < usData.length; j++) {
        if (isStockAreaBoundary_(usData[j][0])) break;
        var uCode = String(usData[j][0]).trim();
        var uName = String(usData[j][1]).trim();
        var uQty = parseNum_(usData[j][4]);
        var uTwdVal = parseNum_(usData[j][8]);
        var uTRatio = parseNum_(usData[j][10]) * 100; // K 欄 (第11欄，索引10) 目標佔比

        if (isInventoryHoldingRow_(usData[j][0], uQty)) {
          stocks.push({
            market: "US",
            code: uCode,
            name: uName,
            totalQty: uQty,
            pledgedQty: 0,
            unpledgedQty: uQty,
            marketVal: uTwdVal,
            targetRatio: uTRatio, // 🎯 傳遞 K 欄數值給前端
            row: 2 + j
          });
        }
      }
    }

    var pledges = [];
    var loans = [];
    var debtSheet = ss.getSheetByName("負債/資金");
    if (debtSheet) {
      var pData = debtSheet.getRange("A5:G7").getValues();
      for (var p = 0; p < pData.length; p++) {
        var pName = String(pData[p][0]).trim();
        var pBal = Math.round(parseNum_(pData[p][6]));
        var pRate = parseNum_(pData[p][2]);
        if (pName !== "" && pBal > 0) {
          pledges.push({ row: 5 + p, name: pName + " 質押", balance: pBal, rate: (pRate <= 1 ? +(pRate * 100).toFixed(2) : pRate) });
        }
      }

      var lData = debtSheet.getRange("A9:G11").getValues();
      for (var l = 0; l < lData.length; l++) {
        var lName = String(lData[l][0]).trim();
        var lBal = Math.round(parseNum_(lData[l][6]));
        var lPay = parseNum_(lData[l][3]);
        var lPeriods = parseNum_(lData[l][5]);
        if (lName !== "" && lBal > 0) {
          loans.push({ row: 9 + l, name: lName + " 信貸", balance: lBal, defaultPay: lPay, periods: lPeriods });
        }
      }

      var oName = String(debtSheet.getRange("A17").getValue()).trim();
      var oBal = Math.round(parseNum_(debtSheet.getRange("G17").getValue()));
      if (oName !== "" && oBal > 0) {
        loans.push({ row: 17, name: oName, balance: oBal, defaultPay: 0, periods: 0 });
      }
    }

    var futures = {
      twMargin: 100000,
      twEquity: 100000,
      twPosition: "持倉觀望 (0口)",
      usMargin: 100000,
      usEquityTwd: 3188830,
      usPosition: "持倉觀望 (0口)",
      usRate: 31.89
    };

    var fSheet = ss.getSheetByName("期貨區");
    if (fSheet) {
      var rawTwMargin = parseNum_(fSheet.getRange("B11").getValue());
      futures.twMargin = !isNaN(rawTwMargin) ? rawTwMargin : 100000;

      var rawTwEquity = parseNum_(fSheet.getRange("D11").getValue());
      futures.twEquity = !isNaN(rawTwEquity) ? rawTwEquity : 100000;

      var twFStatus = String(fSheet.getRange("I2").getValue()).trim();
      if (twFStatus !== "已平倉" && twFStatus !== "") {
        futures.twPosition = String(fSheet.getRange("B2").getValue()) + " " + fSheet.getRange("E2").getValue() + "口";
      }

      var rawRate = parseNum_(fSheet.getRange("C14").getValue());
      futures.usRate = !isNaN(rawRate) && rawRate > 0 ? +rawRate.toFixed(2) : 31.89;

      var rawUsMargin = parseNum_(fSheet.getRange("B25").getValue());
      futures.usMargin = !isNaN(rawUsMargin) ? rawUsMargin : 100000;

      var rawUsEquity = parseNum_(fSheet.getRange("E25").getValue());
      futures.usEquityTwd = Math.round(!isNaN(rawUsEquity) ? rawUsEquity : 3188830);

      var usFStatus = String(fSheet.getRange("I16").getValue()).trim();
      if (usFStatus !== "已平倉" && usFStatus !== "") {
        futures.usPosition = String(fSheet.getRange("B16").getValue()) + " " + fSheet.getRange("E16").getValue() + "口";
      }
    }

    return {
      success: true,
      sheetName: ss.getName(),
      netWorth: netWorth,
      totalAssets: totalAssets,
      totalLiabilities: totalLiabilities,
      currentCash: currentCash,
      pledgeMaintenance: pledgeMaintenance,
      realLeverage: realLeverage,
      assetBeta: assetBeta,
      coverageRatio: realLeverage,
      todayProfit: todayProfit,
      stocks: stocks,
      pledges: pledges,
      loans: loans,
      futures: futures
    };
  } catch (err) {
    return { success: false, message: "❌ 讀取數據失敗: " + err.message };
  }
}

/**
 * 🏛️ 四維系統：讀取「每日紀錄」完整歷史趨勢數據
 */
function getDailyLogHistoryUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    var logSheet = ss.getSheetByName("每日紀錄") || ss.getSheetByName("每日記錄");
    if (!logSheet) {
      return { success: false, message: "❌ 找不到『每日紀錄』分頁" };
    }

    var lastRow = logSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, dates: [], assets: [], netWorths: [], marketDDs: [], netDDs: [] };
    }

    var rowCount = lastRow - 1;
    var rawData = logSheet.getRange(2, 1, rowCount, 12).getValues();
    var dates = [], assets = [], netWorths = [], marketDDs = [], netDDs = [];
    var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";

    for (var i = 0; i < rawData.length; i++) {
      var d = rawData[i][0];
      if (!d) continue;
      var dStr = (d instanceof Date) ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : String(d).trim();
      if (!dStr) continue;

      dates.push(dStr);
      assets.push(parseNum_(rawData[i][1]));
      netWorths.push(parseNum_(rawData[i][2]));
      
      var mDD = parseNum_(rawData[i][9]);
      marketDDs.push(+(mDD <= 1 && mDD >= -1 ? mDD * 100 : mDD).toFixed(2));

      var nDD = parseNum_(rawData[i][11]);
      netDDs.push(+(nDD <= 1 && nDD >= -1 ? nDD * 100 : nDD).toFixed(2));
    }

    return {
      success: true,
      dates: dates,
      assets: assets,
      netWorths: netWorths,
      marketDDs: marketDDs,
      netDDs: netDDs
    };
  } catch (err) {
    return { success: false, message: "❌ 讀取總覽歷史失敗: " + err.message };
  }
}

/**
 * 🏛️ 四維系統：讀取台美股歷史市值資料（美股自動以美股區 A10 匯率折合為台幣）
 */
function getStockLogHistoryUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";

    // 🌟 從美股區 A10 取得最新匯率（預設為 31.89 作為防呆）
    var currentRate = 31.89;
    var usSheet = ss.getSheetByName("美股區") || ss.getSheetByName("美股");
    if (usSheet) {
      var rVal = parseNum_(usSheet.getRange("A10").getValue());
      if (rVal > 0) currentRate = rVal;
    }

    // 1. 讀取台股區 (第15列開始，單位：台幣)
    var twSheet = ss.getSheetByName("台股區") || ss.getSheetByName("台股");
    var twDates = [], twValues = [];
    if (twSheet) {
      var lastRowTw = twSheet.getLastRow();
      if (lastRowTw >= 15) {
        var twData = twSheet.getRange(15, 1, lastRowTw - 14, 5).getValues();
        for (var i = 0; i < twData.length; i++) {
          var d = twData[i][0];
          var val = parseNum_(twData[i][1]);
          if (!val && twData[i][2]) val = parseNum_(twData[i][2]);
          
          if (!d) continue;
          var dStr = (d instanceof Date) ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : String(d).trim();
          if (dStr && val > 0) {
            twDates.push(dStr);
            twValues.push(val);
          }
        }
      }
    }

    // 2. 讀取美股區 (第15列開始，原本為 USD，乘以 A10 匯率折合為 TWD)
    var usDates = [], usValues = [];
    if (usSheet) {
      var lastRowUs = usSheet.getLastRow();
      if (lastRowUs >= 15) {
        var usData = usSheet.getRange(15, 1, lastRowUs - 14, 5).getValues();
        for (var j = 0; j < usData.length; j++) {
          var d2 = usData[j][0];
          var val2Usd = parseNum_(usData[j][1]);
          if (!val2Usd && usData[j][2]) val2Usd = parseNum_(usData[j][2]);

          if (!d2) continue;
          var dStr2 = (d2 instanceof Date) ? Utilities.formatDate(d2, tz, "yyyy-MM-dd") : String(d2).trim();
          if (dStr2 && val2Usd > 0) {
            usDates.push(dStr2);
            // 🌟 美金市值 * 美股區 A10 匯率折合台幣
            usValues.push(Math.round(val2Usd * currentRate));
          }
        }
      }
    }

    return {
      success: true,
      exchangeRateUsed: currentRate,
      tw: { dates: twDates, values: twValues },
      us: { dates: usDates, values: usValues, unit: "TWD" }
    };
  } catch (err) {
    return { success: false, message: "❌ 讀取台美股歷史失敗: " + err.message };
  }
}