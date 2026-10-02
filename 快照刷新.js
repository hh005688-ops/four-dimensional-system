/**
 * 🏛️ 四維系統：每日快照打卡與行情指標刷新模組
 */

// 1. 台股收盤快照 (14:00)
function triggerTaiwanStockSnapshotUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    if (typeof recordTaiwanStock === 'function') {
      var res = recordTaiwanStock(ss);
      return (typeof res === 'object') ? res : { success: true, message: "✅ 台股收盤快照已完成！" };
    }
    return { success: false, message: "❌ 未找到 recordTaiwanStock 函式" };
  } catch (err) {
    return { success: false, message: "❌ 台股快照失敗: " + err.message };
  }
}

// 2. 美股收盤快照 (06:15)
function triggerUSStockSnapshotUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    if (typeof recordUSStock === 'function') {
      var res = recordUSStock(ss);
      return (typeof res === 'object') ? res : { success: true, message: "✅ 美股收盤快照已完成！" };
    }
    return { success: false, message: "❌ 未找到 recordUSStock 函式" };
  } catch (err) {
    return { success: false, message: "❌ 美股快照失敗: " + err.message };
  }
}

// 3. 全系統每日總覽快照 (14:15 精準版)
function triggerDailySnapshotUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    SpreadsheetApp.flush();
    var mainSheet = ss.getSheetByName("總覽");
    var logSheet = ss.getSheetByName("每日紀錄");
    var twSheet = ss.getSheetByName("台股區");

    if (!mainSheet || !logSheet) {
      return { success: false, message: "❌ 找不到『總覽』或『每日紀錄』分頁" };
    }

    var now = new Date();
    var dayOfWeek = now.getDay();
    var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
    var todayStr = Utilities.formatDate(now, tz, "yyyy/MM/dd");

    if (dayOfWeek === 0 || dayOfWeek === 6) {
      return { success: true, message: "☕ 今日為週末非開盤日，已跳過快照。" };
    }

    var totalAsset  = Number(mainSheet.getRange("B6").getValue()) || 0;
    var netAsset    = Number(mainSheet.getRange("A15").getValue()) || 0;
    var todayProfit = Number(mainSheet.getRange("B15").getValue()) || 0;
    var totalAlpha  = Number(mainSheet.getRange("E8").getValue()) || 0;
    var todayReturn = Number(mainSheet.getRange("E10").getValue()) || 0;

    var benchReturn = 0;
    if (twSheet) {
      benchReturn = Number(twSheet.getRange("A12").getValue()) || 0;
    }

    var lastRow = logSheet.getLastRow();
    var prevCumProfit = 0;
    var maxMarketVal = totalAsset;
    var maxNetVal = netAsset;

    if (lastRow >= 2) {
      var historyData = logSheet.getRange(2, 1, lastRow - 1, 12).getValues();
      for (var r = 0; r < historyData.length; r++) {
        var hDate = historyData[r][0];
        var hDateStr = (hDate instanceof Date) 
          ? Utilities.formatDate(hDate, tz, "yyyy/MM/dd")
          : String(hDate).replace(/-/g, "/").trim();

        if (hDateStr !== todayStr) {
          var mVal = Number(historyData[r][1]) || 0;
          var nVal = Number(historyData[r][2]) || 0;
          if (mVal > maxMarketVal) maxMarketVal = mVal;
          if (nVal > maxNetVal) maxNetVal = nVal;

          if (r === historyData.length - 1 || (r === historyData.length - 2 && String(historyData[historyData.length - 1][0]).includes(todayStr))) {
            prevCumProfit = Number(historyData[r][6]) || 0;
          }
        }
      }
    }

    var curMaxMarket = Math.max(maxMarketVal, totalAsset);
    var curMaxNet    = Math.max(maxNetVal, netAsset);
    var marketDD     = curMaxMarket > 0 ? (totalAsset - curMaxMarket) / curMaxMarket : 0;
    var netDD        = curMaxNet > 0 ? (netAsset - curMaxNet) / curMaxNet : 0;
    var cumProfit    = prevCumProfit + todayProfit;

    var snapshotRow = [
      todayStr, totalAsset, netAsset, todayProfit, totalAlpha, todayReturn,
      cumProfit, benchReturn, curMaxMarket, marketDD, curMaxNet, netDD
    ];

    var targetRowIdx = lastRow + 1;
    if (lastRow >= 2) {
      var checkDateVal = logSheet.getRange(lastRow, 1).getValue();
      var checkDateStr = (checkDateVal instanceof Date)
        ? Utilities.formatDate(checkDateVal, tz, "yyyy/MM/dd")
        : String(checkDateVal).replace(/-/g, "/").trim();

      if (checkDateStr === todayStr) {
        targetRowIdx = lastRow;
      }
    }

    logSheet.getRange(targetRowIdx, 1, 1, 12).setValues([snapshotRow]);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: "✅ 每日總覽快照成功！已寫入第 " + targetRowIdx + " 列 (累積獲利: $" + Math.round(cumProfit).toLocaleString() + ")"
    };
  } catch (err) {
    return { success: false, message: "❌ 每日總覽快照失敗: " + err.message };
  }
}

// 4. 刷新台美股 Beta / Alpha
function refreshBetaAlphaUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    if (typeof syncOneYearHistory === 'function') {
      return syncOneYearHistory(ss);
    }
    SpreadsheetApp.flush();
    return { success: true, message: "⚡ 指標參數已強制重新計算並同步！" };
  } catch (err) {
    return { success: false, message: "❌ 刷新 Beta/Alpha 失敗: " + err.message };
  }
}

// 5. 刷新期貨點位與美元匯率
function refreshMarketDataUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    if (typeof getMarketData === 'function') {
      getMarketData(ss);
      return { success: true, message: "📈 期貨點位與美元匯率已成功刷新！" };
    }
    SpreadsheetApp.flush();
    return { success: true, message: "📈 市場行情數據已強制刷新同步！" };
  } catch (err) {
    return { success: false, message: "❌ 刷新市場行情失敗: " + err.message };
  }
}

/**
 * 📋 讀取「每日紀錄」、「台股區」、「美股區」的歷史數據供前端明細表格顯示
 */
function getSnapshotHistoryDataUniversal(sheetIdentifier) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";

  // 1. 讀取「每日紀錄」 (總覽)
  var logSheet = ss.getSheetByName("每日紀錄") || ss.getSheetByName("每日記錄");
  var generalData = [];
  if (logSheet) {
    var lastRow = logSheet.getLastRow();
    if (lastRow >= 2) {
      var values = logSheet.getRange(2, 1, lastRow - 1, 12).getValues();
      generalData = values.map(function(row) {
        var d = row[0];
        if (!d) return null;
        var dateStr = (d instanceof Date) ? Utilities.formatDate(d, tz, "yyyy/MM/dd") : String(d).replace(/-/g, "/").trim();
        return {
          date: dateStr,
          totalAsset: parseNum_(row[1]),
          netAsset: parseNum_(row[2]),
          todayProfit: parseNum_(row[3]),
          totalAlpha: parseNum_(row[4]),
          todayReturn: parseNum_(row[5]),
          cumProfit: parseNum_(row[6]),
          marketDD: parseNum_(row[9])
        };
      }).filter(function(x){ return x !== null && x.date !== ""; });
    }
  }

  // 2. 讀取「台股區」歷史明細 (第15列開始)
  var twSheet = ss.getSheetByName("台股區") || ss.getSheetByName("台股");
  var taiwanData = [];
  if (twSheet) {
    var lastRowTw = twSheet.getLastRow();
    if (lastRowTw >= 15) {
      var twValues = twSheet.getRange(15, 1, lastRowTw - 14, 7).getValues(); // A~G 欄
      taiwanData = twValues.map(function(row) {
        var d = row[0];
        if (!d) return null;
        var dateStr = (d instanceof Date) ? Utilities.formatDate(d, tz, "yyyy/MM/dd") : String(d).replace(/-/g, "/").trim();
        return {
          date: dateStr,
          marketVal: parseNum_(row[1]),
          profit: parseNum_(row[2]),
          returnRate: parseNum_(row[3]),
          benchReturn: parseNum_(row[4]),
          alpha: parseNum_(row[5])
        };
      }).filter(function(x){ return x !== null && x.date !== ""; });
    }
  }

  // 3. 讀取「美股區」歷史明細 (第15列開始)
  var usSheet = ss.getSheetByName("美股區") || ss.getSheetByName("美股");
  var usData = [];
  if (usSheet) {
    var lastRowUs = usSheet.getLastRow();
    if (lastRowUs >= 15) {
      var usValues = usSheet.getRange(15, 1, lastRowUs - 14, 7).getValues(); // A~G 欄
      usData = usValues.map(function(row) {
        var d = row[0];
        if (!d) return null;
        var dateStr = (d instanceof Date) ? Utilities.formatDate(d, tz, "yyyy/MM/dd") : String(d).replace(/-/g, "/").trim();
        return {
          date: dateStr,
          marketVal: parseNum_(row[1]),
          profit: parseNum_(row[2]),
          returnRate: parseNum_(row[3]),
          fxProfit: parseNum_(row[4]),
          alpha: parseNum_(row[5]),
          nasdaqReturn: parseNum_(row[6])
        };
      }).filter(function(x){ return x !== null && x.date !== ""; });
    }
  }

  return { 
    success: true, 
    general: generalData, 
    taiwan: taiwanData, 
    us: usData 
  };
}
