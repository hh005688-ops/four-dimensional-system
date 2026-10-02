// ============================================================================
// 🚀 四維系統 - 智慧資料轉移旗艦引擎 (migration.gs 旗艦修正版)
// ============================================================================

/**
 * ⚡ 執行每日紀錄與轉移前的門神把關
 */
function runDailyRecord() {
  var accessCheck = checkLocalSystemAccess_();
  if (!accessCheck.success) {
    SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  executeMigration(null, ss); // 修正為正確的函式名稱
}

function executeMigration(ui, targetSs) {
  // 🛡️ 防呆自動補正
  if (!targetSs) targetSs = SpreadsheetApp.getActiveSpreadsheet();
  if (!ui) {
    try { ui = SpreadsheetApp.getUi(); } catch(e) { return; }
  }

  // 📢 升級版提示：加入規格限制與手動擴充說明
  const promptMessage = 
    "【🏛️ 四維系統 - 一鍵資料轉移特助】\n\n" +
    "📌 系統轉移支援規格：\n" +
    "1. 台股/美股持股：預設轉移前 6 筆自選標的。\n" +
    "2. 期貨部位：國內期貨轉移前 7 筆，海外期貨轉移前 7 筆。\n" +
    "3. 負債項目：各類負債（券商/信貸/房貸/其他）預設轉移前 3 筆。\n" +
    "4. 歷史快照：自動動態完整搬移【每日紀錄】、【月結紀錄】與走勢圖表。\n\n" +
    "💡 貼心提醒：\n" +
    "本轉移僅支援寫入新表格之「預設欄位與預設列數」。\n" +
    "若您舊表的持股或負債筆數較多，超出部分請於轉移完成後「手動插入列」並參考《使用手冊》公式說明複製套用。\n\n" +
    "請在下方貼上【舊試算表網址 (URL) 或 ID】：";

  const response = ui.prompt("🚀 舊版數據一鍵轉移", promptMessage, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("轉移作業已取消。");
    return;
  }

  const rawInput = response.getResponseText().trim();
  if (!rawInput) {
    ui.alert("❌ 輸入內容為空，轉移作業已終止。");
    return;
  }

  let spreadsheetId = rawInput;
  const match = rawInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) spreadsheetId = match[1];

  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    ui.alert("❌ 無法連線至舊試算表！\n請確認具備檢視/編輯權限。\n錯誤資訊：" + e.message);
    return;
  }

  const confirm = ui.alert(
    "準備開始轉移",
    `已成功連線來源試算表：\n「${sourceSs.getName()}」\n\n即將對齊寫入新表，是否繼續？`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    ui.alert("轉移作業已終止。");
    return;
  }

  let logs = [];

  // A. 【台股區】(A~F 共 6 欄快照)
  try {
    const src = sourceSs.getSheetByName("台股區");
    const tgt = targetSs.getSheetByName("台股區");
    if (src && tgt) {
      smartMigrateStockHolding(src, tgt, 'TW');
      const count = autoLocateAndMigrateSnapshot(src, tgt, 14, 6);
      logs.push(`✅ 台股區轉移成功（持股已轉移，歷史快照 ${count} 筆）`);
    } else {
      logs.push("⚠️ 跳過台股區：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 台股區轉移失敗：" + e.message);
  }

  // B. 【美股區】(A~G 共 7 欄快照)
  try {
    const src = sourceSs.getSheetByName("美股區");
    const tgt = targetSs.getSheetByName("美股區");
    if (src && tgt) {
      smartMigrateStockHolding(src, tgt, 'US');
      const count = autoLocateAndMigrateSnapshot(src, tgt, 14, 7);
      logs.push(`✅ 美股區轉移成功（持股與匯率已轉移，歷史快照 ${count} 筆）`);
    } else {
      logs.push("⚠️ 跳過美股區：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 美股區轉移失敗：" + e.message);
  }

  // C. 【每日紀錄】
  try {
    const src = sourceSs.getSheetByName("每日紀錄");
    const tgt = targetSs.getSheetByName("每日紀錄");
    if (src && tgt) {
      const count = autoLocateAndMigrateSnapshot(src, tgt, 1, 14);
      logs.push(`✅ 每日紀錄轉移成功（歷史快照資料庫 ${count} 筆）`);
    } else {
      logs.push("⚠️ 跳過每日紀錄：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 每日紀錄轉移失敗：" + e.message);
  }

  // D. 【期貨區】
  try {
    const src = sourceSs.getSheetByName("期貨區");
    const tgt = targetSs.getSheetByName("期貨區");
    if (src && tgt) {
      const count = smartMigrateFuturesSheet(src, tgt);
      logs.push(`✅ 期貨區轉移成功（動態對齊 ${count} 筆部位與資金）`);
    } else {
      logs.push("⚠️ 跳過期貨區：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 期貨區轉移失敗：" + e.message);
  }

  // E. 【負債/資金】
  try {
    const src = sourceSs.getSheetByName("負債/資金");
    const tgt = targetSs.getSheetByName("負債/資金");
    if (src && tgt) {
      strictMigrateDebtSheet(src, tgt);
      logs.push("✅ 負債/資金轉移成功（四類負債分類明細 + 現金參數已對齊）");
    } else {
      logs.push("⚠️ 跳過負債/資金：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 負債/資金轉移失敗：" + e.message);
  }

  // F. 【後台數據】
  try {
    const src = sourceSs.getSheetByName("後台數據");
    const tgt = targetSs.getSheetByName("後台數據");
    if (src && tgt) {
      smartMigrateBackendSymbols(src, tgt);
      logs.push("✅ 後台數據轉移成功（自選監控標的已同步）");
    } else {
      logs.push("⚠️ 跳過後台數據：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 後台數據轉移失敗：" + e.message);
  }

  // G. 【月結紀錄】
  try {
    const src = sourceSs.getSheetByName("月結紀錄");
    const tgt = targetSs.getSheetByName("月結紀錄");
    if (src && tgt) {
      const count = autoLocateAndMigrateSnapshot(src, tgt, 1, 7);
      logs.push(`✅ 月結紀錄轉移成功（歷史數據 ${count} 筆）`);
    } else {
      logs.push("⚠️ 跳過月結紀錄：找不到對應分頁");
    }
  } catch (e) {
    logs.push("❌ 月結紀錄轉移失敗：" + e.message);
  }

  // H. 【圖表自動重建】(精準清除與重建)
  try {
    migrateCharts(targetSs);
    logs.push("✅ 台股與美股市值趨勢圖重建完成");
  } catch (e) {
    logs.push("⚠️ 圖表重建略過：" + e.message);
  }

  SpreadsheetApp.flush();

  const finalSummary = 
    "🎉 轉移作業執行完成！\n\n" +
    "執行報告：\n" + logs.join("\n") + "\n\n" +
    "━━━━━━━━━━━━━━━━━━━━\n" +
    "💡 後續提醒：\n" +
    "• 部位若超過預設上限（股票6筆/期貨7筆/各類負債3筆），請手動補齊剩餘項目。\n" +
    "• 更多擴充技巧請查閱《使用手冊(必讀)》分頁。";

  ui.alert(finalSummary);
}

// ============================================================================
// 🛠️ 核心智慧對齊輔助模組
// ============================================================================

/**
 * 1. 台股/美股持股參數動態遷移
 */
function smartMigrateStockHolding(srcSheet, tgtSheet, marketType) {
  tgtSheet.getRange("A2:A7").clearContent();
  tgtSheet.getRange("E2:E7").setValue(0);

  if (marketType === 'TW') {
    tgtSheet.getRange("M2:M7").setValue(0);
    tgtSheet.getRange("O2:O7").clearContent();
  } else if (marketType === 'US') {
    tgtSheet.getRange("N2:N7").clearContent();
  }

  const maxRows = Math.min(srcSheet.getLastRow(), 7);
  if (maxRows < 2) return;
  
  const srcData = srcSheet.getRange(2, 1, maxRows - 1, srcSheet.getLastColumn() || 20).getValues();
  
  const ignoreKeywords = [
    "CELLIMAGE", "加權指數", "加權", "匯率", "大盤", "今日漲跌幅", 
    "今日漲跌", "今日損益", "總市值", "ALPHA", "BETA", "USD", "TPE:", "日期"
  ];

  let targetRowOffset = 0;

  for (let i = 0; i < srcData.length; i++) {
    if (targetRowOffset >= 6) break;

    const row = srcData[i];
    const rawVal = row[0];
    if (rawVal === "" || rawVal === null || rawVal === undefined) continue;

    const rawCode = String(rawVal).trim();
    if (!rawCode) continue;

    const isInvalidKeyword = ignoreKeywords.some(kw => rawCode.toUpperCase().includes(kw));
    if (isInvalidKeyword) continue;

    const numVal = Number(rawCode);
    if (!isNaN(numVal)) {
      if (rawCode.includes(".") || numVal > 999999 || numVal <= 0) continue;
    }

    const tgtRow = 2 + targetRowOffset;

    tgtSheet.getRange(tgtRow, 1).setValue(rawCode);
    const qty = Number(row[4]);
    tgtSheet.getRange(tgtRow, 5).setValue(!isNaN(qty) && qty > 0 ? qty : 0);

    if (marketType === 'TW') {
      const pledge = Number(row[12]);
      if (!isNaN(pledge) && pledge > 0) tgtSheet.getRange(tgtRow, 13).setValue(pledge);

      if (row[14] !== undefined && row[14] !== "") {
        tgtSheet.getRange(tgtRow, 15).setValue(row[14]);
      }
    } else if (marketType === 'US') {
      if (row[13] !== undefined && row[13] !== "") {
        tgtSheet.getRange(tgtRow, 14).setValue(row[13]);
      }
    }

    targetRowOffset++;
  }
}

/**
 * 2. 期貨區動態遷移模組
 */
function smartMigrateFuturesSheet(srcSheet, tgtSheet) {
  let migratedCount = 0;

  tgtSheet.getRange("A2:A8").clearContent();
  tgtSheet.getRange("D2:F8").clearContent();
  tgtSheet.getRange("I2:J8").clearContent();
  tgtSheet.getRange("L2:L8").clearContent();
  tgtSheet.getRange("A11").clearContent();
  tgtSheet.getRange("E11").clearContent();

  tgtSheet.getRange("A16:A22").clearContent();
  tgtSheet.getRange("D16:F22").clearContent();
  tgtSheet.getRange("I16:J22").clearContent();
  tgtSheet.getRange("L16:L22").clearContent();
  tgtSheet.getRange("A25").clearContent();
  tgtSheet.getRange("F25").clearContent();

  const srcMaxRow = srcSheet.getLastRow();
  const srcMaxCol = Math.max(srcSheet.getLastColumn(), 12);
  if (srcMaxRow < 2) return 0;

  const srcData = srcSheet.getRange(1, 1, srcMaxRow, srcMaxCol).getValues();

  for (let r = 1; r <= 7; r++) {
    if (r >= srcData.length) break;
    const row = srcData[r];
    const symbol = String(row[0] || '').trim();

    if (symbol && symbol !== "期貨總入金(本金)") {
      const tgtRow = r + 1;
      tgtSheet.getRange(tgtRow, 1).setValue(symbol);
      if (row[3] !== undefined && row[3] !== "") tgtSheet.getRange(tgtRow, 4).setValue(row[3]);
      if (row[4] !== undefined && row[4] !== "") tgtSheet.getRange(tgtRow, 5).setValue(row[4]);
      if (row[5] !== undefined && row[5] !== "") tgtSheet.getRange(tgtRow, 6).setValue(row[5]);
      if (row[8] !== undefined && row[8] !== "") tgtSheet.getRange(tgtRow, 9).setValue(row[8]);
      if (row[9] !== undefined && row[9] !== "") tgtSheet.getRange(tgtRow, 10).setValue(row[9]);
      if (row[11] !== undefined && row[11] !== "") tgtSheet.getRange(tgtRow, 12).setValue(row[11]);
      migratedCount++;
    }
  }

  try {
    const twDeposit = Number(srcSheet.getRange("A11").getValue()) || 0;
    if (twDeposit > 0) tgtSheet.getRange("A11").setValue(twDeposit);
    
    const twMargin = Number(srcSheet.getRange("E11").getValue()) || 0;
    if (twMargin > 0) tgtSheet.getRange("E11").setValue(twMargin);
  } catch(e) {}

  for (let r = 15; r <= 21; r++) {
    if (r >= srcData.length) break;
    const row = srcData[r];
    const symbol = String(row[0] || '').trim();

    if (symbol && symbol !== "期貨總入金") {
      const tgtRow = r + 1;
      tgtSheet.getRange(tgtRow, 1).setValue(symbol);
      if (row[3] !== undefined && row[3] !== "") tgtSheet.getRange(tgtRow, 4).setValue(row[3]);
      if (row[4] !== undefined && row[4] !== "") tgtSheet.getRange(tgtRow, 5).setValue(row[4]);
      if (row[5] !== undefined && row[5] !== "") tgtSheet.getRange(tgtRow, 6).setValue(row[5]);
      if (row[8] !== undefined && row[8] !== "") tgtSheet.getRange(tgtRow, 9).setValue(row[8]);
      if (row[9] !== undefined && row[9] !== "") tgtSheet.getRange(tgtRow, 10).setValue(row[9]);
      if (row[11] !== undefined && row[11] !== "") tgtSheet.getRange(tgtRow, 12).setValue(row[11]);
      migratedCount++;
    }
  }

  try {
    const usDeposit = Number(srcSheet.getRange("A25").getValue()) || 0;
    if (usDeposit > 0) tgtSheet.getRange("A25").setValue(usDeposit);
    
    const usMargin = Number(srcSheet.getRange("F25").getValue()) || 0;
    if (usMargin > 0) tgtSheet.getRange("F25").setValue(usMargin);
  } catch(e) {}

  return migratedCount;
}

/**
 * 3. 負債/資金結構智慧動態轉移
 */
function strictMigrateDebtSheet(srcSheet, tgtSheet) {
  tgtSheet.getRange("A5:E7").clearContent();
  tgtSheet.getRange("A9:E11").clearContent();
  tgtSheet.getRange("A13:C15").clearContent();
  tgtSheet.getRange("E13:E15").clearContent();
  tgtSheet.getRange("H13:H15").setValue(0);
  tgtSheet.getRange("A17:E19").clearContent();

  const srcData = srcSheet.getDataRange().getValues();
  if (srcData.length < 5) return;

  const targetSections = {
    "券商": { startRow: 5, maxRows: 3 },
    "信貸": { startRow: 9, maxRows: 3 },
    "房貸": { startRow: 13, maxRows: 3 },
    "其他": { startRow: 17, maxRows: 3 }
  };

  let currentCategory = "";
  let categoryCounts = { "券商": 0, "信貸": 0, "房貸": 0, "其他": 0 };

  for (let r = 0; r < srcData.length; r++) {
    const row = srcData[r];
    const firstCell = String(row[0] || '').trim();

    if (firstCell.includes("負債總計") || firstCell.includes("負債餘額") || firstCell.includes("當月配息")) {
      continue;
    }

    if (targetSections[firstCell]) {
      currentCategory = firstCell;
      continue;
    }

    const loanAmount = Number(row[1]);
    if (currentCategory && firstCell && !isNaN(loanAmount) && loanAmount > 0) {
      const config = targetSections[currentCategory];
      const count = categoryCounts[currentCategory];

      if (count < config.maxRows) {
        const tgtRow = config.startRow + count;

        tgtSheet.getRange(tgtRow, 1).setValue(firstCell);
        tgtSheet.getRange(tgtRow, 2).setValue(loanAmount);
        if (row[2] !== undefined && row[2] !== "") tgtSheet.getRange(tgtRow, 3).setValue(row[2]);

        if (currentCategory === "房貸") {
          if (row[4] !== undefined && row[4] !== "") tgtSheet.getRange(tgtRow, 5).setValue(row[4]);
          if (row[7] !== undefined && !isNaN(Number(row[7]))) tgtSheet.getRange(tgtRow, 8).setValue(Number(row[7]));
        } else {
          if (row[3] !== undefined && row[3] !== "") tgtSheet.getRange(tgtRow, 4).setValue(row[3]);
          if (row[4] !== undefined && row[4] !== "") tgtSheet.getRange(tgtRow, 5).setValue(row[4]);
        }

        categoryCounts[currentCategory]++;
      }
    }
  }

  for (let r = 0; r < srcData.length; r++) {
    for (let c = 0; c < srcData[r].length; c++) {
      const cellText = String(srcData[r][c] || '').trim();
      
      if (cellText === "現金" && srcData[r][c + 1] !== undefined) {
        const cashVal = Number(srcData[r][c + 1]);
        if (!isNaN(cashVal) && cashVal > 0) tgtSheet.getRange("D21").setValue(cashVal);
      }
      
      if (cellText.includes("當月配息") && srcData[r][c + 1] !== undefined) {
        const divVal = Number(srcData[r][c + 1]);
        if (!isNaN(divVal) && divVal > 0) tgtSheet.getRange("B23").setValue(divVal);
      }
    }
  }
}

/**
 * 4. 歷史快照資料庫動態定位與搬移 (純值批次搬移)
 */
function autoLocateAndMigrateSnapshot(srcSheet, tgtSheet, headerRow, dataCols) {
  const lastRow = srcSheet.getLastRow();
  if (lastRow <= headerRow) return 0;

  const numRows = lastRow - headerRow;
  const srcRange = srcSheet.getRange(headerRow + 1, 1, numRows, dataCols);
  const srcValues = srcRange.getValues();

  const validRows = srcValues.filter(r => r[0] !== "" && r[0] !== null);
  if (validRows.length === 0) return 0;

  tgtSheet.getRange(headerRow + 1, 1, validRows.length, dataCols).setValues(validRows);
  return validRows.length;
}

/**
 * 5. 後台數據自選標的代碼動態同步
 */
function smartMigrateBackendSymbols(srcSheet, tgtSheet) {
  const maxCols = Math.min(srcSheet.getLastColumn(), tgtSheet.getLastColumn(), 30);
  if (maxCols < 6) return;

  const symbols = srcSheet.getRange(1, 6, 1, maxCols - 5).getValues();
  if (symbols && symbols[0]) {
    tgtSheet.getRange(1, 6, 1, symbols[0].length).setValues(symbols);
  }
}

/**
 * 6. 圖表自動重建模組
 */
function migrateCharts(targetSs) {
  const twSheet = targetSs.getSheetByName("台股區");
  const usSheet = targetSs.getSheetByName("美股區");
  
  if (twSheet) rebuildStockChart(twSheet, "台股總市值趨勢", 15, 14, 9);
  if (usSheet) rebuildStockChart(usSheet, "美股總市值趨勢 (USD)", 15, 14, 9);
}

function rebuildStockChart(sheet, title, dataStartRow, anchorRow, anchorCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  const charts = sheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    sheet.removeChart(charts[i]);
  }

  const numRows = lastRow - dataStartRow + 1;
  const xRange = sheet.getRange(dataStartRow, 1, numRows, 1);
  const yRange = sheet.getRange(dataStartRow, 2, numRows, 1);

  const primaryColor = '#0f172a';
  const gridLineColor = '#e2e8f0';

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(xRange)
    .addRange(yRange)
    .setPosition(anchorRow, anchorCol, 5, 5)
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
    .setOption('vAxis.gridlines.color', gridLineColor)
    .setOption('vAxis.textStyle.color', '#475569')
    .setOption('vAxis.textStyle.bold', true)
    .build();

  sheet.insertChart(chart);
}

/**
 * 7. 通用安全複製函式
 */
function robustCopy(srcSheet, tgtSheet, rangeA1) {
  try {
    const srcRange = srcSheet.getRange(rangeA1);
    if (srcRange.getRow() > srcSheet.getMaxRows() || srcRange.getColumn() > srcSheet.getMaxColumns()) return;
    
    const values = srcRange.getValues();
    const hasData = values.some(row => row.some(cell => cell !== "" && cell !== null));
    if (!hasData) return;

    tgtSheet.getRange(rangeA1).setValues(values);
  } catch (e) {
    Logger.log(`略過複製區間 ${rangeA1}: ${e.message}`);
  }
}