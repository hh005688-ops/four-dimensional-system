// ============================================================================
// 🏛️ 四維系統：基礎核心、網頁入口與統一授權派發模組 (v4.9 旗艦整合版)
// ============================================================================

var DEFAULT_SHEET_ID = "1GapDSh1Ge_Dpm40kXun5pOZQ7IU9sLQcwHVuJZwNNwM";

/**
 * 1. 取得目標試算表物件 (支援網址、ID 或物件傳入)
 */
function getTargetSpreadsheet_(sheetIdentifier) {
  var ss = null;
  if (sheetIdentifier && typeof sheetIdentifier.getSheetByName === 'function') {
    return sheetIdentifier;
  }
  var target = (typeof sheetIdentifier === 'string' && sheetIdentifier.trim() !== '') 
               ? sheetIdentifier.trim() 
               : DEFAULT_SHEET_ID;

  var match = target.match(/\/d\/([a-zA-Z0-9-_]+)/);
  var id = (match && match[1]) ? match[1] : target;
  
  try {
    ss = SpreadsheetApp.openById(id);
  } catch(err) {
    Logger.log("openById 失敗: " + err.message);
  }

  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  return ss;
}

/**
 * 2. 數值解析與防呆工具
 */
function parseNum_(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  var cleaned = String(val).replace(/[$,%]/g, '').trim();
  var num = Number(cleaned);
  return isNaN(num) ? 0 : num;
}

/** 避免 Math.min.apply 在長陣列上觸發 Maximum call stack size exceeded */
function minMaxNums_(arr) {
  var min = Infinity;
  var max = -Infinity;
  var list = arr || [];
  for (var i = 0; i < list.length; i++) {
    var v = Number(list[i]);
    if (isNaN(v)) continue;
    if (v < min) min = v;
    if (v > max) max = v;
  }
  if (min === Infinity) return { min: 0, max: 0 };
  return { min: min, max: max };
}

/**
 * 3. Web App 網頁端點入口已移至 api-gateway.js（doGet / doPost JSON API）
 */

// ============================================================================
// 🛡️ 4. 網頁端/前端請求「統一派發與授權總司令官」(大範圍鎖表核心)
// ============================================================================
function dispatchUniversalAction(sheetIdentifier, actionType, params) {
  // A. 取得目標試算表
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) {
    return { success: false, message: "❌ 無法開啟試算表或未提供有效的識別碼" };
  }
  
  // 🚀 自用版已移除門神檢查，直接安全分派到對應模組執行
  try {
    params = params || {};
    switch (actionType) {
      case "CASH_DISPATCH":
        return executeDispatchUniversal(sheetIdentifier, params);

      case "BATCH_SAVE_TARGET_RATIOS":
        params.type = "BATCH_SAVE_TARGET_RATIOS";
        return executeDispatchUniversal(sheetIdentifier, params);
        
      case "STOCK_TRADE":
        return executeStockTradeUniversal(sheetIdentifier, params);
        
      case "PLEDGE_TRANSFER":
        return executePledgeTransferUniversal(sheetIdentifier, params);
        
      case "FUTURES_TRANSFER":
        return executeFuturesTransferUniversal(sheetIdentifier, params);
        
      case "ADD_STOCK":
        if (typeof addNewStockTargetUniversal === "function") {
          return addNewStockTargetUniversal(sheetIdentifier, params);
        }
        params.type = "ADD_STOCK";
        return executeDispatchUniversal(sheetIdentifier, params);
        
      default:
        return { success: false, message: "❌ 未知的動作類型: " + actionType };
    }
  } catch (err) {
    return { success: false, message: "❌ 系統執行例外: " + err.message };
  }
}

/**
 * 5. 輔助檢查特定試算表物件是否已解鎖 (讀取文件屬性)
 */
function checkSpreadsheetUnlockedUniversal_(ss) {
  try {
    // 優先檢查該文件專屬的 Properties
    var cache = ss.getDocProperties ? ss.getDocProperties() : PropertiesService.getDocumentProperties();
    // 註：若透過 openById 遠端檢查，亦可搭配取得文件屬性
    var isUnlocked = cache.getProperty('SYSTEM_UNLOCKED');
    return (isUnlocked === 'TRUE');
  } catch (e) {
    // 保底：若在當前作用中表格
    try {
      var localCache = PropertiesService.getDocumentProperties();
      return (localCache.getProperty('SYSTEM_UNLOCKED') === 'TRUE');
    } catch(err) {
      return false;
    }
  }
}