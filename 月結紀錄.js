
/**
 * 按鈕專用無參數入口：指派給按鈕「recordMonthlySummaryMain」 (已加門神)
 */
function recordMonthlySummaryMain() {
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}
  recordMonthlySummary();
}

/**
 * 排程專用入口 (已加門神)
 */
function runDailyRecord() {
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  recordMonthlySummary(ss);
}

/**
 * 四維系統：一鍵月結紀錄 (純後端全自動防呆計算版 + 門神安全防護)
 */
function recordMonthlySummary(ss) {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}

  // 1. 試算表物件安全解析
  if (!ss || typeof ss.getSheetByName !== 'function') {
    if (typeof resolveSpreadsheet === 'function') {
      ss = resolveSpreadsheet(ss);
    }
    if (!ss || typeof ss.getSheetByName !== 'function') {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
  }

  if (!ss) {
    Logger.log("❌ 無法獲取試算表物件");
    return;
  }

  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null;
  }

  // 2. 分頁物件綁定
  const debtSheet = ss.getSheetByName("負債/資金") || ss.getSheetByName("負債與資金") || ss.getSheetByName("負債");
  const recordSheet = ss.getSheetByName("月結紀錄") || ss.getSheetByName("月結記錄");
  const dailySheet = ss.getSheetByName("每日紀錄") || ss.getSheetByName("每日記錄");

  if (!debtSheet || !recordSheet || !dailySheet) {
    const msg = "❌ 缺少必要工作表，請確認分頁名稱（月結紀錄、負債/資金、每日紀錄）。";
    if (ui) ui.alert(msg);
    Logger.log(msg);
    return;
  }

  // 3. 從「負債/資金」抓取維持率與負債餘額
  const rawRatio = debtSheet.getRange("F22").getValue();
  let rawDebt = debtSheet.getRange("B22").getValue();
  if (!rawDebt || rawDebt === 0) rawDebt = debtSheet.getRange("B21").getValue();

  let curRatio = 0;
  if (typeof rawRatio === 'number') {
    curRatio = rawRatio;
  } else if (typeof rawRatio === 'string') {
    const cleanStr = rawRatio.replace(/[%,\s]/g, '');
    curRatio = rawRatio.includes('%') ? (Number(cleanStr) / 100 || 0) : (Number(cleanStr) || 0);
  }

  let curDebt = 0;
  if (typeof rawDebt === 'number') {
    curDebt = rawDebt;
  } else if (typeof rawDebt === 'string') {
    const cleanStr = rawDebt.replace(/[NT$,\s]/gi, '');
    curDebt = Number(cleanStr) || 0;
  }

  const today = new Date();
  let inputYear = today.getFullYear();
  let inputMonth = today.getMonth() + 1;
  let isConfirmed = true;

  // 4. UI 互動詢問
  if (ui) {
    const response = ui.prompt(
      "📊 一鍵月結紀錄", 
      "請輸入要紀錄的月份 (預設為 " + inputMonth + " 月，支援如 2026/8 或 8)：", 
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const text = response.getResponseText().trim();
      if (text.includes("/") || text.includes("-")) {
        const parts = text.split(/[\/\-]/);
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        if (!isNaN(y) && !isNaN(m) && m >= 1 && m <= 12) {
          inputYear = y;
          inputMonth = m;
        }
      } else {
        const parsed = parseInt(text, 10);
        if (!isNaN(parsed) && parsed >= 1 && parsed <= 12) {
          inputMonth = parsed;
        }
      }
    } else {
      isConfirmed = false;
    }
  }

  if (!isConfirmed) return;

  const targetMonthIndex = inputMonth - 1;
  const targetYearStr = String(inputYear);
  const targetMonthPadded = inputMonth < 10 ? "0" + inputMonth : String(inputMonth);
  const targetFormatStandard = targetYearStr + "/" + targetMonthPadded;

  // 5. 批次讀取「每日紀錄」進行後端精確計算
  const dailyLastRow = dailySheet.getLastRow();
  let dailyData = [];
  if (dailyLastRow > 1) {
    dailyData = dailySheet.getRange(2, 1, dailyLastRow - 1, 8).getValues();
  }

  function parseToDate(val) {
    if (!val) return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }

  let prevMonthLastRecord = null; 
  let targetMonthLastRecord = null; 
  let marketProduct = 1; 
  let hasMarketData = false;

  const startOfCurrentMonth = new Date(inputYear, targetMonthIndex, 1);
  const endOfCurrentMonth = new Date(inputYear, targetMonthIndex + 1, 0, 23, 59, 59);

  for (let i = 0; i < dailyData.length; i++) {
    const row = dailyData[i];
    const rowDate = parseToDate(row[0]);
    if (!rowDate) continue;

    const rowTotalAsset = Number(row[1]) || 0;
    const rowNetAsset = Number(row[2]) || 0;
    const rawMarketGain = row[7];
    let rowMarketReturn = 0;
    if (typeof rawMarketGain === 'number') {
      rowMarketReturn = rawMarketGain;
    } else if (typeof rawMarketGain === 'string') {
      rowMarketReturn = Number(rawMarketGain.replace(/[%,\s]/g, '')) / 100 || 0;
    }

    if (rowDate < startOfCurrentMonth) {
      prevMonthLastRecord = { totalAsset: rowTotalAsset, netAsset: rowNetAsset, date: rowDate };
    }

    if (rowDate >= startOfCurrentMonth && rowDate <= endOfCurrentMonth) {
      targetMonthLastRecord = { totalAsset: rowTotalAsset, netAsset: rowNetAsset, date: rowDate };
      marketProduct *= (1 + rowMarketReturn);
      hasMarketData = true;
    }
  }

  // 6. 後端計算各項核心指標
  const startAsset = prevMonthLastRecord ? prevMonthLastRecord.totalAsset : 0;
  const prevNetAsset = prevMonthLastRecord ? prevMonthLastRecord.netAsset : 0;
  const curNetAsset = targetMonthLastRecord ? targetMonthLastRecord.netAsset : 0;

  const netProfit = (targetMonthLastRecord && prevMonthLastRecord) ? (curNetAsset - prevNetAsset) : 0;
  const monthReturn = (startAsset > 0) ? (netProfit / startAsset) : 0;
  const marketReturn = hasMarketData ? (marketProduct - 1) : 0;
  const alpha = monthReturn - marketReturn;

  // 7. 掃描「月結紀錄」定位行號
  const maxRows = Math.max(recordSheet.getLastRow(), 1);
  const dateData = recordSheet.getRange(1, 1, maxRows, 1).getValues();
  let targetRow = 0;
  let firstEmptyRow = 0;

  for (let i = 0; i < dateData.length; i++) {
    const val = dateData[i][0];
    const rowNum = i + 1;

    if (rowNum > 1 && (val === "" || val === null || val === undefined) && firstEmptyRow === 0) {
      firstEmptyRow = rowNum;
    }

    if (rowNum === 1 || !val) continue;

    if (val instanceof Date) {
      if (val.getFullYear() === inputYear && val.getMonth() === targetMonthIndex) {
        targetRow = rowNum;
        break;
      }
    } else {
      const strVal = String(val).trim();
      if (strVal.includes(targetFormatStandard) || 
          strVal === (targetYearStr + "/" + inputMonth) ||
          strVal === (targetYearStr + "-" + targetMonthPadded)) {
        targetRow = rowNum;
        break;
      }
    }
  }

  const finalRow = targetRow > 0 ? targetRow : (firstEmptyRow > 0 ? firstEmptyRow : recordSheet.getLastRow() + 1);
  const standardDateObj = new Date(inputYear, targetMonthIndex, 1);

  // 8. 批次寫入純數值
  const outputRow = [[
    standardDateObj, 
    startAsset,      
    netProfit,       
    monthReturn,     
    alpha,           
    curRatio,        
    curDebt,         
    marketReturn     
  ]];

  recordSheet.getRange(finalRow, 1, 1, 8).setValues(outputRow);

  recordSheet.getRange(finalRow, 1).setNumberFormat("yyyy/mm");
  recordSheet.getRange(finalRow, 2, 1, 2).setNumberFormat("#,##0");
  recordSheet.getRange(finalRow, 4, 1, 3).setNumberFormat("0.00%");
  recordSheet.getRange(finalRow, 7).setNumberFormat("$#,##0");
  recordSheet.getRange(finalRow, 8).setNumberFormat("0.00%");

  SpreadsheetApp.flush();

  // 9. 結果通知反饋
  const formattedRatio = (curRatio * 100).toFixed(2) + "%";
  const formattedDebt = "$" + Math.round(curDebt).toLocaleString();
  const formattedProfit = "$" + Math.round(netProfit).toLocaleString();
  const formattedReturn = (monthReturn * 100).toFixed(2) + "%";

  Logger.log("✅ " + targetFormatStandard + " 純後端月結紀錄完成！寫入第 " + finalRow + " 列");

  if (ui) {
    ui.alert(
      "✅ " + targetFormatStandard + " 月結紀錄同步成功！\n\n" +
      "📍 紀錄位置：第 " + finalRow + " 列\n" +
      "💰 本月淨獲利：" + formattedProfit + " (" + formattedReturn + ")\n" +
      "📈 月底維持率：" + formattedRatio + "\n" +
      "💳 月底負債總額：" + formattedDebt + "\n\n" +
      "⚡ 已以後端高防呆模式寫入純數值，不受公式誤刪影響！"
    );
  }
}