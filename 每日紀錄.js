/**
 * ==========================================
 * 四維系統：每日紀錄 雙圖動態渲染核心 (跨分頁快取版)
 * ==========================================
 */

/**
 * 1. 每日紀錄專屬圖表渲染器（包含：市值雙折線圖 + 回撤雙面積圖）
 */
function runDailyRecord() {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
  if (!accessCheck.success) {
    SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 這裡可以放你本來想在 runDailyRecord 裡執行的主邏輯（例如呼叫 recordDailySnapshot 等）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  recordDailySnapshot(ss);
}

function updateDailyLogCharts(e) {
  var ss = resolveSpreadsheet(e);
  if (!ss) return;

  var logSheet = ss.getSheetByName("每日紀錄") || ss.getSheetByName("每日記錄");
  if (!logSheet) return;

  var startRow = 2; // 資料從第 2 列開始
  var lastRow = logSheet.getLastRow();
  if (lastRow < startRow) return;

  // 1. 取得或建立隱藏快取分頁
  var cacheSheetName = "_CHART_CACHE";
  var cacheSheet = ss.getSheetByName(cacheSheetName);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(cacheSheetName);
    cacheSheet.hideSheet();
  }

  // 2. 批次讀取 A 到 L 欄完整數據（不受篩選器影響）
  var totalRows = lastRow - startRow + 1;
  var rawData = logSheet.getRange(startRow, 1, totalRows, 12).getValues();

  var validData = rawData.filter(function(row) {
    var dateVal = row[0];
    var marketVal = parseNumber(row[1]);
    return dateVal !== "" && !isNaN(marketVal) && marketVal > 0;
  });

  if (validData.length === 0) return;

  // 3. 讀取 N1 下拉選單天數
  var rangeSelect = String(logSheet.getRange("N1").getValue() || "近30天");
  var days = 30;
  var upper = rangeSelect.toUpperCase();
  if (upper.indexOf("ALL") !== -1 || rangeSelect.indexOf("全") !== -1) {
    days = validData.length;
  } else {
    var numMatch = rangeSelect.match(/\d+/);
    if (numMatch) days = parseInt(numMatch[0], 10);
  }

  var takeCount = Math.min(validData.length, days);
  var targetData = validData.slice(validData.length - takeCount);

  // 4. 將繪圖所需資料批次寫入快取分頁（每日紀錄專用 G~K 欄，即第 7~11 欄）
  var cacheStartCol = 7;
  cacheSheet.getRange(1, cacheStartCol, 1500, 5).clearContent();

  var writeArray = targetData.map(function(row) {
    var dateVal = (row[0] instanceof Date) 
      ? Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone() || "Asia/Taipei", "yyyy/MM/dd") 
      : String(row[0]).replace(/-/g, "/").trim();
    var marketVal = parseNumber(row[1]);   // B欄: 總資產市值
    var netVal = parseNumber(row[2]);      // C欄: 淨資產
    var ddMarket = parseNumber(row[9]);    // J欄: 動態回撤率
    var ddNet = parseNumber(row[11]);      // L欄: 淨值回撤率
    return [dateVal, marketVal, netVal, ddMarket, ddNet];
  });

  cacheSheet.getRange(1, cacheStartCol, writeArray.length, 5).setValues(writeArray);

  // 5. 計算市值圖 Y 軸自適應邊界
  var allAssets = writeArray.flatMap(function(r) { return [r[1], r[2]]; }).filter(function(v){ return v > 0; });
  var assetBound = minMaxNums_(allAssets);
  var minAsset = assetBound.min;
  var maxAsset = assetBound.max;
  var padAsset = (maxAsset - minAsset) * 0.05 || 10000;
  var assetMin = Math.floor(Math.max(0, minAsset - padAsset));
  var assetMax = Math.ceil(maxAsset + padAsset);

  // 6. 計算回撤圖 Y 軸自適應邊界 (百分比，負數向下)
  var allDD = writeArray.flatMap(function(r) { return [r[3], r[4]]; });
  var minDD = minMaxNums_(allDD).min;
  var ddFloor = (minDD < 0) ? Math.floor(minDD * 10) / 10 - 0.05 : -0.1;

  // 7. 清理舊圖表
  var charts = logSheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    logSheet.removeChart(charts[i]);
  }

  var dateRange = cacheSheet.getRange(1, 7, writeArray.length, 1);
  var marketRange = cacheSheet.getRange(1, 8, writeArray.length, 1);
  var netRange = cacheSheet.getRange(1, 9, writeArray.length, 1);
  var ddMarketRange = cacheSheet.getRange(1, 10, writeArray.length, 1);
  var ddNetRange = cacheSheet.getRange(1, 11, writeArray.length, 1);

  // ==========================================
  // 📈 圖表 1：總資產市值和淨資產 (雙折線圖)
  // ==========================================
  var chart1 = logSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .setPosition(2, 13, 5, 5) // 放置於 M2
    .clearRanges()
    .addRange(dateRange)
    .addRange(marketRange)
    .addRange(netRange)
    .setOption('title', '總資產市值和淨資產 (' + rangeSelect + ')')
    .setOption('width', 580)
    .setOption('height', 280)
    .setOption('colors', ['#1e3a8a', '#dc2626']) // 總市值：深藍，淨資產：紅
    .setOption('lineWidth', 3)
    .setOption('pointSize', 3)
    .setOption('curveType', 'function')
    .setOption('legend.position', 'top')
    .setOption('chartArea.width', '80%')
    .setOption('chartArea.left', '15%')
    .setOption('vAxis.viewWindowMode', 'explicit')
    .setOption('vAxis.viewWindow.min', assetMin)
    .setOption('vAxis.viewWindow.max', assetMax)
    .setOption('vAxis.gridlines.color', '#e2e8f0')
    .setOption('vAxis.textStyle.bold', true)
    .build();

  // ==========================================
  // 📉 圖表 2：總(淨)資產回撤率 (雙面積圖)
  // ==========================================
  var chart2 = logSheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .setPosition(15, 13, 5, 5) // 放置於 M15
    .clearRanges()
    .addRange(dateRange)
    .addRange(ddMarketRange)
    .addRange(ddNetRange)
    .setOption('title', '總(淨)資產回撤率 (' + rangeSelect + ')')
    .setOption('width', 580)
    .setOption('height', 280)
    .setOption('colors', ['#dc2626', '#0f172a']) // 總回撤：紅，淨值回撤：深黑
    .setOption('areaOpacity', 0.35)
    .setOption('lineWidth', 2)
    .setOption('pointSize', 0)
    .setOption('legend.position', 'top')
    .setOption('chartArea.width', '80%')
    .setOption('chartArea.left', '15%')
    .setOption('vAxis.viewWindowMode', 'explicit')
    .setOption('vAxis.viewWindow.min', ddFloor)
    .setOption('vAxis.viewWindow.max', 0.02)
    .setOption('vAxis.format', '0.00%')
    .setOption('vAxis.gridlines.color', '#e2e8f0')
    .setOption('vAxis.textStyle.bold', true)
    .build();

  logSheet.insertChart(chart1);
  logSheet.insertChart(chart2);
}

/**
 * 2. 編輯觸發器聯動 (加入每日紀錄 N1 的監聽)
 */
function handleFrontendEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var a1 = e.range.getA1Notation();

  // 監聽 台股區 I12
  if (sheetName === "台股區" && a1 === "I12") {
    renderFrontendAreaChart({
      sheet: sheet,
      rangeSelect: e.value || "近15天",
      titlePrefix: "台股總市值趨勢",
      dataStartRow: 15,
      xCol: 1,
      yCol: 2,
      anchorRow: 14,
      anchorCol: 9
    });
  }
  // 監聽 美股區 I12
  else if (sheetName === "美股區" && a1 === "I12") {
    renderFrontendAreaChart({
      sheet: sheet,
      rangeSelect: e.value || "近15天",
      titlePrefix: "美股總市值趨勢",
      dataStartRow: 15,
      xCol: 1,
      yCol: 2,
      anchorRow: 14,
      anchorCol: 9
    });
  }
  // 🚀 監聽 每日紀錄 N1
  else if ((sheetName === "每日紀錄" || sheetName === "每日記錄") && a1 === "N1") {
    updateDailyLogCharts(e);
  }
}

/**
 * 📷 4D-MS 總覽專屬：每日資產快照核心函式 (v5.3.0 精準修復版)
 */
function recordDailySnapshot(ss) {
  if (recordDailySnapshot._busy) {
    return { success: false, message: "⏸️ 每日快照執行中，已阻擋重複／遞迴呼叫" };
  }
  recordDailySnapshot._busy = true;
  try {
    return recordDailySnapshotImpl_(ss);
  } finally {
    recordDailySnapshot._busy = false;
  }
}

function recordDailySnapshotImpl_(ss) {
  if (typeof resolveSpreadsheet === 'function') {
    ss = resolveSpreadsheet(ss);
  }
  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch(e) {}
  }
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  var mainSheet = ss.getSheetByName("總覽");
  var logSheet = ss.getSheetByName("每日紀錄");
  var twSheet = ss.getSheetByName("台股區");

  if (!mainSheet || !logSheet) {
    Logger.log("❌ 找不到【總覽】或【每日紀錄】分頁");
    return { success: false, message: "❌ 找不到【總覽】或【每日紀錄】分頁" };
  }

  var now = new Date();
  var dayOfWeek = now.getDay(); // 0 = 週日, 6 = 週六
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
  var todayStr = Utilities.formatDate(now, tz, "yyyy/MM/dd");

  // ==========================================
  // 🛡️ 防呆機制 1：週末直接攔截 (週六、週日不記錄)
  // ==========================================
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    try {
      ss.toast("☕ 今日為週末非開盤日，系統已自動跳過快照紀錄！", "4D-MS 風控防呆", 4);
    } catch(e) {}
    Logger.log("🛡️ 今日為週末 (" + todayStr + ")，終止寫入快照。");
    return { success: true, message: "☕ 今日為週末非開盤日，已跳過快照。" };
  }

  // ==========================================
  // 🛡️ 防呆機制 2：讀取核心數據與休市判斷
  // ==========================================
  var totalAsset  = Number(mainSheet.getRange("B6").getValue()) || 0;   // 總資產市值
  var netAsset    = Number(mainSheet.getRange("A15").getValue()) || 0;  // 淨資產
  var todayProfit = Number(mainSheet.getRange("B15").getValue()) || 0;  // 今日獲利
  var totalAlpha  = Number(mainSheet.getRange("E8").getValue()) || 0;   // 資產總 Alpha (E7)
  var todayReturn = Number(mainSheet.getRange("E10").getValue()) || 0;  // 今日總體報酬率

  // 大盤漲跌幅 (從台股區 A12 即時取得)
  var benchReturn = 0;
  if (twSheet) {
    benchReturn = Number(twSheet.getRange("A12").getValue()) || 0;
  }

  if (totalAsset <= 0) {
    try {
      ss.toast("⚠️ 總資產數值異常 (<= 0)，終止寫入！", "4D-MS 錯誤防呆", 4);
    } catch(e) {}
    return { success: false, message: "⚠️ 總資產數值異常 (<= 0)，終止寫入！" };
  }

  // ==========================================
  // 📊 歷史高點、回撤與累積獲利計算
  // ==========================================
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

        // 取最後一筆非今日的累積獲利
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
    todayStr,           // A: 日期
    totalAsset,        // B: 總資產市值
    netAsset,          // C: 淨資產
    todayProfit,       // D: 今日獲利
    totalAlpha,        // E: 資產總 Alpha
    todayReturn,       // F: 今日報酬率
    cumProfit,         // G: 累積獲利
    benchReturn,       // H: 今日大盤損益
    curMaxMarket,      // I: 歷史最高市值
    marketDD,          // J: 動態回撤率
    curMaxNet,         // K: 淨值歷史高點
    netDD              // L: 淨值回撤率
  ];

  // ==========================================
  // ✍️ 判斷覆蓋或新增資料列 (批次寫入)
  // ==========================================
  var targetRowIdx = lastRow + 1;

  if (lastRow >= 2) {
    var checkDateVal = logSheet.getRange(lastRow, 1).getValue();
    var checkDateStr = (checkDateVal instanceof Date)
      ? Utilities.formatDate(checkDateVal, tz, "yyyy/MM/dd")
      : String(checkDateVal).replace(/-/g, "/").trim();

    if (checkDateStr === todayStr) {
      targetRowIdx = lastRow; // 當日覆蓋更新
    }
  }

  // 1 次 IO 批次寫入純數值
  logSheet.getRange(targetRowIdx, 1, 1, 12).setValues([snapshotRow]);
  SpreadsheetApp.flush();

  // 聯動刷新圖表 (若有定義)
  if (typeof updateDailyLogCharts === 'function') {
    try { updateDailyLogCharts(ss); } catch(e) {}
  }

  try {
    ss.toast("📷 今日開盤快照 (" + todayStr + ") 已寫入第 " + targetRowIdx + " 列！", "4D-MS 決策中心", 4);
  } catch(e) {}

  return {
    success: true,
    message: "✅ 每日總覽快照成功！已寫入第 " + targetRowIdx + " 列 (累積獲利: $" + Math.round(cumProfit).toLocaleString() + ")"
  };
}