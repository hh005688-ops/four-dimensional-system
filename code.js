/**
 * 🏛️ 四維系統：通用獨立 Web App 與母表本地防護引擎 (架構精簡優化版)
 * 嚴格遵守：批次讀寫、記憶體 2D 陣列運算、完整 try...catch 與防呆機制
 */

var MASTER_SHEET_ID = "1_IWJgJJXkd3BQ-rUTJOX6-UTlvMo7CsSwwrLSIPngQE";

// Web App 入口已改為 api-gateway.js 的 doGet / doPost JSON API（Vercel 前端專用）

/**
 * 🔑 彈出輸入框讓買家輸入金鑰，並聯動中央總台自動啟用
 */
function showActivationPrompt_() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "🏛️ 四維系統 - 授權啟動中心",
      "請輸入您的專屬購買授權金鑰：",
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    var inputKey = response.getResponseText();
    if (!inputKey || !inputKey.trim()) {
      ui.alert("⚠️ 提示", "輸入的金鑰不得為空！", ui.ButtonSet.OK);
      return;
    }
    
    var result = activateSystemWithKey(MASTER_SHEET_ID, inputKey);
    ui.alert(result.success ? "🎉 啟動成功" : "🔒 啟動受阻", result.message, ui.ButtonSet.OK);
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ 系統異常", "發生未預期的錯誤: " + err.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 🔑 核心守門員：檢查本地授權狀態 (自用版永久放行)
 */
function checkLocalSystemAccess_() {
  return { success: true, message: "✅ 本地授權有效 (自用版)" };
}

/**
 * 🚀 買家初次啟動與驗證綁定
 */
function activateSystemWithKey(masterSheetId, inputKey) {
  try {
    var currentSs = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheetId = currentSs.getId();
    var cleanKey = String(inputKey || "").trim();
    if (!cleanKey) return { success: false, message: "⚠️ 請輸入有效的購買金鑰！" };

    var masterSs = SpreadsheetApp.openById(masterSheetId);
    var keySheet = masterSs.getSheetByName("金鑰驗證");
    if (!keySheet) return { success: false, message: "🔒 授權驗證失敗：無法連線至中央金鑰總台。" };

    var lastRow = keySheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "🔒 中央金鑰資料庫無紀錄。" };

    var data = keySheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var targetRowIndex = -1;
    var isMatched = false;

    for (var i = 0; i < data.length; i++) {
      var rowKey = String(data[i][0] || "").trim();
      var rowSheetId = String(data[i][3] || "").trim();

      if (rowKey === cleanKey && (!rowSheetId || rowSheetId === currentSheetId)) {
        targetRowIndex = i + 2;
        isMatched = true;
        break;
      }
    }

    if (!isMatched) return { success: false, message: "🔒 啟用失敗：金鑰錯誤、或已被其他人綁定！" };

    keySheet.getRange(targetRowIndex, 2).setValue("已啟用");
    keySheet.getRange(targetRowIndex, 3).setValue(new Date());
    keySheet.getRange(targetRowIndex, 4).setValue(currentSheetId);
    SpreadsheetApp.flush();

    var docProps = PropertiesService.getDocumentProperties();
    docProps.setProperty("FOUR_DIM_ACTIVATED", "TRUE");
    docProps.setProperty("BOUND_SHEET_ID", currentSheetId);

    return { success: true, message: "🎉 恭喜！四維系統解鎖成功！" };
  } catch (err) {
    return { success: false, message: "❌ 啟動過程發生異常: " + err.message };
  }
}

/**
 * 🛠️ 強固型試算表個體取得器（自用版絕對保底升級版）
 */
function getSpreadsheetInstance(sheetIdentifier) {
  if (!sheetIdentifier) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  var val = String(sheetIdentifier).trim();
  var id = val;
  if (val.indexOf('/d/') !== -1) {
    var parts = val.split('/d/');
    if (parts.length > 1) id = parts[1].split('/')[0];
  }
  
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    Logger.log("❌ openById 失敗，改用getActiveSpreadsheet保底: " + err.message);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * 📥 首頁儀表板資料請見 數據讀取.js 的 getInitDataUniversal
 * （必須以「總覽」B6 / B12 / A15 / C15 為準，禁止自行加總以免覆蓋正確數字）
 */

/**
 * ⚡ 執行資金調度並回寫
 */
function executeDispatchUniversalLite_(sheetIdentifier, params) {
  try {
    var ss = getSpreadsheetInstance(sheetIdentifier);
    if (!ss) return { success: false, message: "❌ 試算表連線失敗" };

    var type = params.type;
    var amount = Number(params.amount) || 0;
    var debtSheet = ss.getSheetByName("負債/資金");
    if (!debtSheet) return { success: false, message: "❌ 找不到『負債/資金』分頁" };

    var cashCell = debtSheet.getRange("D21");
    var currentCash = Number(cashCell.getValue()) || 0;

    switch (type) {
      case 'DIVIDEND':
        var divCell = debtSheet.getRange("B23");
        divCell.setValue((Number(divCell.getValue()) || 0) + amount);
        cashCell.setValue(currentCash + amount);
        break;
      case 'INFLOW':
        cashCell.setValue(currentCash + amount);
        break;
      case 'PLEDGE_BORROW':
        var pBorrowCell = debtSheet.getRange("B" + params.row);
        pBorrowCell.setValue((Number(pBorrowCell.getValue()) || 0) + amount);
        cashCell.setValue(currentCash + amount);
        break;
      case 'PLEDGE_REPAY':
        var pRepayCell = debtSheet.getRange("B" + params.row);
        pRepayCell.setValue(Math.max(0, (Number(pRepayCell.getValue()) || 0) - amount));
        cashCell.setValue(Math.max(0, currentCash - amount));
        break;
      case 'FUTURES_IN':
      case 'FUTURES_OUT':
        var futSheet = ss.getSheetByName("期貨區");
        if (futSheet) {
          var futCell = futSheet.getRange(params.market === 'TW' ? "A11" : "A25");
          var curFut = Number(futCell.getValue()) || 0;
          futCell.setValue(type === 'FUTURES_IN' ? curFut + amount : Math.max(0, curFut - amount));
        }
        cashCell.setValue(type === 'FUTURES_IN' ? Math.max(0, currentCash - amount) : currentCash + amount);
        break;
      case 'LOAN_PAY':
        var periodCell = debtSheet.getRange("F" + params.row);
        periodCell.setValue((Number(periodCell.getValue()) || 0) + 1);
        cashCell.setValue(Math.max(0, currentCash - amount));
        break;
      case 'STOCK_BUY':
      case 'STOCK_SELL':
        var sSheet = ss.getSheetByName(params.market === 'TW' ? "台股區" : "美股區");
        if (sSheet) {
          var qtyCell = sSheet.getRange("E" + params.row);
          var curQty = Number(qtyCell.getValue()) || 0;
          var deltaQty = Number(params.qty || 0);
          qtyCell.setValue(type === 'STOCK_BUY' ? curQty + deltaQty : Math.max(0, curQty - deltaQty));
        }
        if (amount > 0) {
          cashCell.setValue(type === 'STOCK_BUY' ? Math.max(0, currentCash - amount) : currentCash + amount);
        }
        break;
    }

    SpreadsheetApp.flush();
    return { success: true, message: "✅ 調度成功完成！" };
  } catch (e) {
    return { success: false, message: "❌ 寫入失敗: " + e.message };
  }
}

/**
 * 📈 新增期貨部位並回寫（批次效能優化與欄位兼容版）
 */
function appendFuturesPosition(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: "❌ 系統忙碌中（併發鎖定超時），請稍後再試！" };
  }
  
  try {
    var ss = getSpreadsheetInstance(data.sheetIdentifier);
    if (!ss) throw new Error("❌ 無法開啟試算表，請檢查 ID/網址！");

    const sheet = ss.getSheetByName('期貨區');
    if (!sheet) throw new Error("❌ 找不到「期貨區」分頁！");
    
    let startRow, maxCheckRow;
    if (data.marketType === 'domestic') {
      startRow = 2;
      maxCheckRow = 15;
    } else {
      startRow = 16;
      maxCheckRow = 30;
    }
    
    // 批次讀取 B 欄範圍來尋找第一個空白列
    const rangeToCheck = sheet.getRange(startRow, 2, maxCheckRow - startRow + 1, 1);
    const values = rangeToCheck.getValues();
    
    let targetRow = -1;
    for (let i = 0; i < values.length; i++) {
      if (!String(values[i][0]).trim()) {
        targetRow = startRow + i;
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error("❌ 該期貨區段空間已滿（無可用空白列），請手動在試算表中新增列數！");
    }
    
    // 兼容前端傳遞的 contractCode 或 colB
    const contractVal = String(data.contractCode || data.colB || "").trim();
    const unit = Number(data.colE) || 0;
    const price = Number(data.colF) || 0;
    const closePrice = (data.colJ !== "" && data.colJ !== null && data.colJ !== undefined) ? Number(data.colJ) : "";
    
    if (!contractVal) {
      throw new Error("❌ 合約代碼不得為空！");
    }
    if (unit <= 0) {
      throw new Error("❌ 單位口數必須大於 0！");
    }
    if (price <= 0) {
      throw new Error("❌ 本期買入均價必須大於 0！");
    }

    // 🚀 嚴格遵守批次寫入原則：建構整列資料陣列一次性回寫 (A 到 J 欄，共 10 欄)
    var currentRowValues = sheet.getRange(targetRow, 1, 1, 10).getValues()[0];
    
    currentRowValues[1] = contractVal;                                    // B 欄: 合約代碼 / 月份名稱
    currentRowValues[3] = String(data.colD || "多").trim();                 // D 欄: 多/空
    currentRowValues[4] = unit;                                             // E 欄: 單位口數
    currentRowValues[5] = price;                                            // F 欄: 本期買入均價
    currentRowValues[8] = String(data.colI || "加碼").trim();               // I 欄: 狀態 (加碼/平倉)
    currentRowValues[9] = (closePrice !== "") ? closePrice : "";            // J 欄: 平倉價 (選填)

    // 一次性整包回寫至試算表
    sheet.getRange(targetRow, 1, 1, 10).setValues([currentRowValues]);
    
    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: "🎉 成功將期貨部位建倉寫入「期貨區」第 " + targetRow + " 列！" 
    };
    
  } catch (error) {
    return { success: false, message: "❌ 執行失敗: " + error.message };
  } finally {
    lock.releaseLock();
  }
}