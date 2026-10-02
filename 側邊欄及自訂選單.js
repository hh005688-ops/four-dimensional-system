// ============================================================================
// 🏛️ 四維系統 - 頂部選單與前端 UI / 圖表聯動引擎 (MenuAndUI.gs)
// ============================================================================

/**
 * 1. 自動觸發器：開啟試算表時建立頂部完整選單（含解鎖與各項功能）
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🏛️ 四維決策中心');
  
  menu.addItem('🔑 啟動/解鎖系統', 'showActivationPrompt')
      .addItem('⚡ 開啟四維資產即時監控', 'showFourDimensionsSidebar')
      .addItem('🧠 開啟四維智腦特助 (AI)', 'OPEN_BRAIN_SIDEBAR')
      .addItem('🌐 開啟系統 Web App', 'openWebAppDirectly')
      .addSeparator()
      .addItem('🚀 舊版/基礎版一鍵無痛轉移', 'migrateFromOldVersion')
      .addSeparator()
      .addItem('⚡ 一鍵更新全部期貨行情 (國內+海期)', 'updateAllFutures')
      .addItem('🇹🇼 僅更新國內期貨行情', 'updateFuturesQuotes')
      .addItem('🌐 僅更新海外期貨行情', 'updateOverseasFuturesQuotes')
      .addSeparator()
      .addItem('📝 執行每日紀錄', 'recordDailySnapshot')
      .addItem('🇹🇼 執行台股紀錄', 'recordTaiwanStock')
      .addItem('🇺🇸 執行美股紀錄', 'recordUSStock')
      .addSeparator()
      .addItem('📊 執行月結紀錄 ', 'recordMonthlySummary')
      .addSeparator()
      .addItem('📡 執行每日盤後速報 (LINE)', 'sendDailyBriefing')
      .addItem('📊 週末深度週報 (LINE)', 'sendWeeklyDeepReport')
      .addItem('🏆 月度資產大報告 (LINE)', 'sendMonthlyDeepReport')
      .addSeparator()
      .addItem('⏰ 啟用/更新每日自動排程與圖表引擎 (必點)', 'setupDailyTrigger')
      .addItem('🔄 重置/同步 1 年歷史數據公式', 'syncOneYearHistory')
      .addToUi();
}

// ==========================================
// ⚡ 四維資產即時監控 - 側邊欄開啟
// ==========================================
function showFourDimensionsSidebar() {
  runWithAccessCheck_(function() {
    var html = HtmlService.createHtmlOutputFromFile('側邊欄操作')
      .setTitle('⚡ 四維資產即時監控')
      .setWidth(360);
    SpreadsheetApp.getUi().showSidebar(html);
  });
}

// ==========================================
// 🧠 四維智腦特助 (AI 側邊欄 - 本地獨立版)
// ==========================================
function OPEN_BRAIN_SIDEBAR() {
  runWithAccessCheck_(function() {
    try {
      var html = HtmlService.createHtmlOutputFromFile('側邊欄ai')
                  .setTitle('🤖 四維智腦特助')
                  .setWidth(360);
      SpreadsheetApp.getUi().showSidebar(html);
    } catch(e) {
      SpreadsheetApp.getActiveSpreadsheet().toast("❌ 找不到智腦側邊欄介面檔案 (Brain.html)", "系統提示", 5);
    }
  });
}

function GET_SELECTED_DATA_FULL() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return "未選取任何區域";
  const values = range.getValues();
  return values.map(row => row.filter(String).join(" | ")).join("\n") || "選取區域為空白";
}

function CALL_BRAIN_API(prompt) {
  try {
    // 💡 這裡您可以直接在本地端實作呼叫 AI API（例如呼叫您的 DeepSeek 核心）
    // 如果原本是依賴 D4MS 的後端，請將對應的 API 呼叫邏輯直接寫在此處或呼叫本地函式
    
    if (typeof callDeepSeekAgent === 'function') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      return callDeepSeekAgent("deepseek-chat", "你是一個高冷、說話幽默犀利、極度重視數據邏輯的四維系統首席分析師。", prompt, ss);
    }
    
    return "❌ 本地端尚未定義 AI 核心呼叫函式";
  } catch(e) {
    return "❌ 執行智腦 API 失敗：" + e.toString();
  }
}

// ============================================================================
// 📈 前端圖表即時聯動監聽器 (由安裝型觸發器呼叫)
// ============================================================================

function handleDropdownEdit(e) {
  if (handleDropdownEdit._busy) return;
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (sheetName === "_CHART_CACHE") return;

  var a1 = e.range.getA1Notation();
  handleDropdownEdit._busy = true;
  try {
    if (a1 === "I12" || a1 === "I13") {
      if (sheetName === "台股區") {
        updateTaiwanStockChartFrontend(sheet);
      } else if (sheetName === "美股區") {
        updateUSStockChartFrontend(sheet);
      }
    } else if ((sheetName === "每日紀錄" || sheetName === "每日記錄") && a1 === "N1") {
      updateDailyLogChartsFrontend(sheet);
    }
  } finally {
    handleDropdownEdit._busy = false;
  }
}

function onEdit(e) {
  handleDropdownEdit(e);
}

/**
 * 台股圖表渲染引擎
 */
function updateTaiwanStockChartFrontend(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("台股區");
  }
  if (!sheet) return;
  var val12 = String(sheet.getRange("I12").getValue()).trim();
  var val13 = String(sheet.getRange("I13").getValue()).trim();
  var rangeSelect = val12 || val13 || "近15天";

  renderFrontendAreaChart({
    sheet: sheet,
    titlePrefix: "台股總市值趨勢",
    rangeSelect: rangeSelect,
    anchorRow: 14,
    anchorCol: 9,
    dataStartRow: 15,
    xCol: 1,
    yCol: 2
  });
}

/**
 * 美股圖表渲染引擎
 */
function updateUSStockChartFrontend(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("美股區");
  }
  if (!sheet) return;
  var val12 = String(sheet.getRange("I12").getValue()).trim();
  var val13 = String(sheet.getRange("I13").getValue()).trim();
  var rangeSelect = val12 || val13 || "近15天";

  renderFrontendAreaChart({
    sheet: sheet,
    titlePrefix: "美股總市值趨勢 (USD)",
    rangeSelect: rangeSelect,
    anchorRow: 14,
    anchorCol: 9,
    dataStartRow: 15,
    xCol: 1,
    yCol: 2
  });
}

/**
 * 每日紀錄專屬圖表渲染引擎
 */
function updateDailyLogChartsFrontend(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!sheet) {
    sheet = ss.getSheetByName("每日紀錄") || ss.getSheetByName("每日記錄");
  }
  if (!sheet) return;

  var startRow = 2;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  var cacheSheetName = "_CHART_CACHE";
  var cacheSheet = ss.getSheetByName(cacheSheetName);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(cacheSheetName);
    cacheSheet.hideSheet();
  }

  var totalRows = lastRow - startRow + 1;
  var rawData = sheet.getRange(startRow, 1, totalRows, 12).getValues();

  var validData = rawData.filter(function(row) {
    var dateVal = row[0];
    var marketVal = parseFloat(String(row[1]).replace(/[\$,]/g, '').trim()) || 0;
    return dateVal !== "" && !isNaN(marketVal) && marketVal > 0;
  });

  if (validData.length === 0) return;

  var rangeSelect = String(sheet.getRange("N1").getValue() || "近30天");
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

  var cacheStartCol = 7;
  cacheSheet.getRange(1, cacheStartCol, 1500, 5).clearContent();

  var timeZone = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
  var writeArray = targetData.map(function(row) {
    var dateVal = (row[0] instanceof Date) 
      ? Utilities.formatDate(row[0], timeZone, "yyyy/MM/dd") 
      : String(row[0]).replace(/-/g, "/").trim();
    var marketVal = parseFloat(String(row[1]).replace(/[\$,]/g, '').trim()) || 0;
    var netVal    = parseFloat(String(row[2]).replace(/[\$,]/g, '').trim()) || 0;
    var ddMarket  = parseFloat(String(row[9]).replace(/[\$,%]/g, '').trim()) / (String(row[9]).indexOf('%') !== -1 ? 100 : 1) || 0;
    var ddNet     = parseFloat(String(row[11]).replace(/[\$,%]/g, '').trim()) / (String(row[11]).indexOf('%') !== -1 ? 100 : 1) || 0;
    return [dateVal, marketVal, netVal, ddMarket, ddNet];
  });

  cacheSheet.getRange(1, cacheStartCol, writeArray.length, 5).setValues(writeArray);

  var allAssets = writeArray.flatMap(function(r) { return [r[1], r[2]]; }).filter(function(v){ return v > 0; });
  var assetBound = minMaxNums_(allAssets);
  var minAsset = assetBound.min;
  var maxAsset = assetBound.max;
  var padAsset = (maxAsset - minAsset) * 0.05 || 10000;
  var assetMin = Math.floor(Math.max(0, minAsset - padAsset));
  var assetMax = Math.ceil(maxAsset + padAsset);

  var allDD = writeArray.flatMap(function(r) { return [r[3], r[4]]; });
  var minDD = minMaxNums_(allDD).min;
  var ddFloor = (minDD < 0) ? Math.floor(minDD * 10) / 10 - 0.05 : -0.1;

  var charts = sheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    sheet.removeChart(charts[i]);
  }

  var dateRange     = cacheSheet.getRange(1, 7, writeArray.length, 1);
  var marketRange   = cacheSheet.getRange(1, 8, writeArray.length, 1);
  var netRange      = cacheSheet.getRange(1, 9, writeArray.length, 1);
  var ddMarketRange = cacheSheet.getRange(1, 10, writeArray.length, 1);
  var ddNetRange    = cacheSheet.getRange(1, 11, writeArray.length, 1);

  var chart1 = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .setPosition(2, 13, 5, 5)
    .clearRanges()
    .addRange(dateRange)
    .addRange(marketRange)
    .addRange(netRange)
    .setOption('title', '總資產市值和淨資產 (' + rangeSelect + ')')
    .setOption('width', 580)
    .setOption('height', 280)
    .setOption('colors', ['#1e3a8a', '#dc2626'])
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

  var chart2 = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .setPosition(15, 13, 5, 5)
    .clearRanges()
    .addRange(dateRange)
    .addRange(ddMarketRange)
    .addRange(ddNetRange)
    .setOption('title', '總(淨)資產回撤率 (' + rangeSelect + ')')
    .setOption('width', 580)
    .setOption('height', 280)
    .setOption('colors', ['#dc2626', '#0f172a'])
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

  sheet.insertChart(chart1);
  sheet.insertChart(chart2);
}

/**
 * 通用深色科技風圖表繪製核心
 */
function renderFrontendAreaChart(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = config.sheet;
  var startRow = config.dataStartRow || 15;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  var cacheSheetName = "_CHART_CACHE";
  var cacheSheet = ss.getSheetByName(cacheSheetName);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(cacheSheetName);
    cacheSheet.hideSheet();
  }

  var totalRows = lastRow - startRow + 1;
  var rawData = sheet.getRange(startRow, config.xCol, totalRows, (config.yCol - config.xCol + 1)).getValues();
  
  var validData = rawData.filter(function(row) {
    var dateVal = row[0];
    var numVal = parseFloat(String(row[row.length - 1]).replace(/[\$,]/g, '').trim()) || 0;
    return dateVal !== "" && !isNaN(numVal) && numVal > 0;
  });

  if (validData.length === 0) return;

  var rangeSelect = config.rangeSelect || "近15天";
  var days = 15;
  var upper = rangeSelect.toUpperCase();
  if (upper.indexOf("ALL") !== -1 || rangeSelect.indexOf("全") !== -1) {
    days = validData.length;
  } else {
    var numMatch = rangeSelect.match(/\d+/);
    if (numMatch) days = parseInt(numMatch[0], 10);
  }

  var takeCount = Math.min(validData.length, days);
  var targetData = validData.slice(validData.length - takeCount);

  var sheetName = sheet.getName();
  var cacheCol = (sheetName === "台股區") ? 1 : 4; 
  
  cacheSheet.getRange(1, cacheCol, 1000, 2).clearContent();

  var writeArray = targetData.map(function(item) {
    var val = parseFloat(String(item[item.length - 1]).replace(/[\$,]/g, '').trim()) || 0;
    return [item[0], val];
  });
  
  cacheSheet.getRange(1, cacheCol, writeArray.length, 2).setValues(writeArray);

  var yValues = writeArray.map(function(r) { return r[1]; });
  var yBound = minMaxNums_(yValues);
  var minVal = yBound.min;
  var maxVal = yBound.max;
  var padding = (maxVal - minVal) * 0.05 || 1000;
  var chartMin = Math.floor(Math.max(0, minVal - padding));
  var chartMax = Math.ceil(maxVal + padding);

  if (chartMin === chartMax) {
    chartMin = Math.max(0, chartMin - 1000);
    chartMax = chartMax + 1000;
  }

  var charts = sheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    sheet.removeChart(charts[i]);
  }

  var xRange = cacheSheet.getRange(1, cacheCol, writeArray.length, 1);
  var yRange = cacheSheet.getRange(1, cacheCol + 1, writeArray.length, 1);
  var displayTitle = config.titlePrefix + " (" + rangeSelect + ")";

  var finalChart = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .setPosition(config.anchorRow, config.anchorCol, 5, 5)
    .clearRanges()
    .addRange(xRange)
    .addRange(yRange)
    .setOption('title', displayTitle)
    .setOption('width', 580)
    .setOption('height', 280)
    .setOption('colors', ['#0f172a'])
    .setOption('areaOpacity', 0.35)
    .setOption('lineWidth', 3)
    .setOption('pointSize', 4)
    .setOption('curveType', 'function')
    .setOption('legend.position', 'none')
    .setOption('chartArea.width', '75%')
    .setOption('chartArea.left', '18%')
    .setOption('vAxis.viewWindowMode', 'explicit')
    .setOption('vAxis.viewWindow.min', chartMin)
    .setOption('vAxis.viewWindow.max', chartMax)
    .setOption('vAxis.gridlines.color', '#e2e8f0')
    .setOption('vAxis.textStyle.color', '#475569')
    .setOption('vAxis.textStyle.bold', true)
    .build();

  sheet.insertChart(finalChart);
}