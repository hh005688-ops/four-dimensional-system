/**
 * ⚡ 四維系統：全自動台美股資產紀錄與圖表更新腳本 (v4.7.0 Web App 通用版)
 */

// ==========================================
// --- 1. 台股紀錄核心 (支援 Web App 遠端呼叫) ---
// ==========================================
function runDailyRecord() {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
  if (!accessCheck.success) {
    SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 取得目前試算表物件，並傳入 recordTaiwanStock 中
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  recordTaiwanStock(ss);
}

function recordTaiwanStock(sheetIdentifier) {
  if (recordTaiwanStock._busy) {
    return { success: false, message: "⏸️ 台股快照執行中，已阻擋重複／遞迴呼叫" };
  }
  recordTaiwanStock._busy = true;
  try {
    return recordTaiwanStockImpl_(sheetIdentifier);
  } finally {
    recordTaiwanStock._busy = false;
  }
}

function recordTaiwanStockImpl_(sheetIdentifier) {
  // 🛡️ 直接在函式內解析，不受任何外部同名函式干擾
  var ss = null;
  if (sheetIdentifier && typeof sheetIdentifier.getSheetByName === 'function') {
    ss = sheetIdentifier;
  } else if (typeof sheetIdentifier === 'string' && sheetIdentifier.trim() !== '') {
    var str = sheetIdentifier.trim();
    var match = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
    var id = (match && match[1]) ? match[1] : str;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch(err) {
      Logger.log("openById 錯誤: " + err.message);
    }
  } else {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }

  if (!ss) {
    return { success: false, message: "❌ 無法獲取試算表物件，請檢查網址或權限。" };
  }

  var sheet = ss.getSheetByName("台股區"); 
  if (!sheet) {
    return { success: false, message: "❌ 找不到『台股區』分頁，請檢查分頁名稱！" };
  }

  // 🛡️ 週末休市過濾
  var dayOfWeek = new Date().getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return { success: false, message: "⏸️ 今天是週末（休市），未執行台股快照。" };
  }
  
  // ⭐️ 強制先更新公式，避免讀到載入中的 0
  SpreadsheetApp.flush();

  // 🛡️ 數值抓取與防呆
  var totalValue = parseNumber(sheet.getRange("B10").getValue());
  if (totalValue === 0) {
    return { success: false, message: "⚠️ 台股總市值為 0（或報價載入中），已取消寫入。" };
  }

  var timeZone = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
  var date = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  
  var todayProfit = parseNumber(sheet.getRange("D10").getValue()); // 今日總損益
  var changePct   = parseNumber(sheet.getRange("E10").getValue()); // 今日漲跌幅
  var beta        = parseNumber(sheet.getRange("C10").getValue()); // 整體Beta值 (C10)
  var marketPct   = parseNumber(sheet.getRange("A12").getValue()); // 加權指數今日漲跌 (A12)

  // 📐 動態計算台股 Alpha 值：E10 - (C10 * A12)
  var alphaValue  = changePct - (beta * marketPct);

  var startRow = 15;

  // 🚀 尋找最新寫入列
  var targetRow = getNextAvailableRow(sheet, startRow, 1);

  // 📝 批次寫入 (A 到 F 欄)
  sheet.getRange(targetRow, 1, 1, 6).setValues([[
    date, totalValue, todayProfit, changePct, marketPct, alphaValue
  ]]);

  SpreadsheetApp.flush();

  // 📊 更新圖表
  try {
    updateSystemChart(sheet, "台股總市值趨勢", 14, 9, targetRow, 1, 2, startRow);
  } catch(chartErr) {
    Logger.log("圖表更新略過: " + chartErr.message);
  }

  return { 
    success: true, 
    message: "✅ 台股紀錄與圖表已更新（第 " + targetRow + " 列）",
    row: targetRow 
  };
}

// ==========================================
// --- 2. 美股紀錄核心 (支援 Web App 遠端呼叫) ---
// ==========================================
function runDailyRecord() {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
  if (!accessCheck.success) {
    SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 取得目前試算表物件，並傳入 recordUSStock 中
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  recordUSStock(ss);
}

function recordUSStock(sheetIdentifier) {
  if (recordUSStock._busy) {
    return { success: false, message: "⏸️ 美股快照執行中，已阻擋重複／遞迴呼叫" };
  }
  recordUSStock._busy = true;
  try {
    return recordUSStockImpl_(sheetIdentifier);
  } finally {
    recordUSStock._busy = false;
  }
}

function recordUSStockImpl_(sheetIdentifier) {
  var ss = resolveSpreadsheet(sheetIdentifier);
  if (!ss) {
    return { success: false, message: "❌ 無法獲取試算表物件，請檢查網址或權限。" };
  }

  var sheet = ss.getSheetByName("美股區"); 
  if (!sheet) {
    return { success: false, message: "❌ 找不到『美股區』分頁，請檢查名稱。" };
  }
  
  // 🛡️ 台灣時間週日(0)或週一(1)美股休市不執行
  var dayOfWeek = new Date().getDay();
  if (dayOfWeek === 0 || dayOfWeek === 1) {
    return { success: false, message: "⏸️ 美股休市時段，未執行美股快照。" };
  }

  // ⭐️ 強制試算表先完成所有即時公式計算
  SpreadsheetApp.flush();
  
  var startRow = 15;

  // 🚀 尋找最新寫入列（歷史紀錄無上限持續向下累積）
  var targetRow = getNextAvailableRow(sheet, startRow, 1);

  // 📥 抓取即時指標數據（精準對齊表_4）
  var rawTotalValue = parseNumber(sheet.getRange("B10").getValue()); // 總市值(USD)
  var todayProfit   = parseNumber(sheet.getRange("D10").getValue()); // 今日總損益
  var changePct     = parseNumber(sheet.getRange("E10").getValue()); // 今日漲跌幅
  var beta          = parseNumber(sheet.getRange("C10").getValue()); // 整體Beta (C10)
  
  // 🛠️ 修正：納指100當日漲跌改抓取 B12，匯差損益修正抓取 D12
  var nasdaqPct     = parseNumber(sheet.getRange("B12").getValue()); // 納指100當日漲跌 (B12)
  var fxProfit      = parseNumber(sheet.getRange("D12").getValue()); // 匯差損益(NT) (D12)
  
  // 📐 動態計算美股 Alpha 值：E10 - (C10 * B12)
  var alphaValue = changePct - (beta * nasdaqPct);
  
  var timeZone = ss.getSpreadsheetTimeZone() || "Asia/Taipei";
  var date = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");

  // 🧠 市值防呆保底 (若抓到 0 則沿用上一交易日)
  if (rawTotalValue <= 0 && targetRow > startRow) {
    rawTotalValue = parseNumber(sheet.getRange("B" + (targetRow - 1)).getValue());
  }

  // 📝 批次寫入 A 到 G 欄：[日期, 總市值, 今日損益, 今日漲跌幅, 匯差損益(NT), Alpha, 納指100漲跌]
  sheet.getRange(targetRow, 1, 1, 7).setValues([[
    date, rawTotalValue, todayProfit, changePct, fxProfit, alphaValue, nasdaqPct
  ]] );

  SpreadsheetApp.flush();

  // 📊 更新面積圖
  try {
    updateSystemChart(sheet, "美股總市值趨勢 (USD)", 14, 9, targetRow, 1, 2, startRow);
  } catch(chartErr) {
    Logger.log("圖表更新略過: " + chartErr.message);
  }

  return { 
    success: true, 
    message: "✅ 美股紀錄與圖表更新完成（第 " + targetRow + " 列）",
    row: targetRow 
  };
}

// ==========================================
// --- 3. 通用圖表更新核心引擎 ---
// ==========================================
function updateSystemChart(sheet, title, anchorRow, anchorCol, lastDataRow, xCol, yCol, dataStartRow) {
  var startRow = dataStartRow || 15;
  var charts = sheet.getCharts();
  
  var dataLength = lastDataRow - startRow + 1;
  if (dataLength < 1) dataLength = 1;

  var xRange = sheet.getRange(startRow, xCol, dataLength, 1);
  var yRange = sheet.getRange(startRow, yCol, dataLength, 1);

  // 🛡️ 數值批次讀取與邊界計算
  var yValues = sheet.getRange(startRow, yCol, dataLength, 1).getValues().flatMap(function(row) {
    var val = parseNumber(row[0]);
    return val > 0 ? [val] : [];
  });
  
  var yBound = minMaxNums_(yValues);
  var minVal = yValues.length > 0 ? yBound.min : 0;
  var maxVal = yValues.length > 0 ? yBound.max : 10000;
  
  var padding = (maxVal - minVal) || (minVal * 0.02) || 1000;
  var chartMin = Math.floor(minVal - padding * 0.5);
  var chartMax = Math.ceil(maxVal + padding * 0.5);

  if (chartMin === chartMax) {
    chartMin = Math.max(0, chartMin - 1000);
    chartMax = chartMax + 1000;
  }

  // 🧹 強化清除機制
  for (var i = 0; i < charts.length; i++) {
    var c = charts[i];
    var cTitle = c.getOptions().get('title');
    var containerInfo = c.getContainerInfo();
    var isSameAnchor = containerInfo && containerInfo.getAnchorRow() === anchorRow && containerInfo.getAnchorColumn() === anchorCol;

    if (cTitle === title || isSameAnchor) {
      sheet.removeChart(c);
    }
  }

  var primaryColor = '#0f172a';
  var gridLineColor = '#e2e8f0';

  // 🚀 建立全新圖表
  var finalChart = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .setPosition(anchorRow, anchorCol, 5, 5)
    .clearRanges()
    .addRange(xRange)
    .addRange(yRange)
    .setOption('title', title)
    .setOption('width', 580)          
    .setOption('height', 280)
    .setOption('colors', [primaryColor])   
    .setOption('areaOpacity', 0.4)          
    .setOption('lineWidth', 3)              
    .setOption('pointSize', 5)              
    .setOption('curveType', 'function')    
    .setOption('legend.position', 'none')  
    .setOption('chartArea.width', '75%')   
    .setOption('chartArea.left', '18%')    
    .setOption('vAxis.viewWindowMode', 'explicit')
    .setOption('vAxis.viewWindow.min', chartMin) 
    .setOption('vAxis.viewWindow.max', chartMax) 
    .setOption('vAxis.gridlines.color', gridLineColor)
    .setOption('vAxis.textStyle.color', '#475569')
    .setOption('vAxis.textStyle.bold', true)
    .build();

  sheet.insertChart(finalChart);
}

// ==========================================
// --- 4. 基礎輔助函式 (強化支援網址/ID動態解析) ---
// ==========================================
function resolveSpreadsheet(input) {
  // 1. 若已經是 Spreadsheet 物件
  if (input && typeof input.getSheetByName === 'function') {
    return input;
  }
  
  // 2. 若傳入的是字串 (網址或 ID)
  if (typeof input === 'string' && input.trim() !== '') {
    var str = input.trim();
    var match = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
    var id = (match && match[1]) ? match[1] : str;
    try {
      return SpreadsheetApp.openById(id);
    } catch(err) {
      Logger.log("❌ openById 失敗: " + err.message);
    }
  }

  // 3. 原生環境綁定保底
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch(e) {}

  return null;
}

function parseNumber(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  var clean = String(val).replace(/[\$,]/g, '').replace(/%/g, '').trim();
  var num = parseFloat(clean);
  if (isNaN(num)) return 0;
  if (String(val).indexOf('%') !== -1) return num / 100;
  return num;
}

function getNextAvailableRow(sheet, startRow, checkCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;
  
  var numRows = lastRow - startRow + 1;
  var values = sheet.getRange(startRow, checkCol, numRows, 1).getValues();
  
  var lastFilledIndex = -1;
  for (var i = 0; i < values.length; i++) {
    var cell = values[i][0];
    if (cell !== "" && cell !== null && cell !== undefined) {
      lastFilledIndex = i;
    }
  }
  
  return (lastFilledIndex === -1) ? startRow : (startRow + lastFilledIndex + 1);
}