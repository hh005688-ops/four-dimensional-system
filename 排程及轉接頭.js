// ============================================================================
// 🏛️ 四維系統 - 本地端完整主控中心 (取代 D4MS 雲端轉接)
// ============================================================================

/**
 * 🛡️ 本地門神檢查（本地獨立版直接放行）
 */
function runWithAccessCheck_(actionCallback) {
  if (typeof actionCallback === 'function') {
    actionCallback();
  }
}

/**
 * 🔑 授權解鎖引導
 */
function showActivationPrompt() {
  SpreadsheetApp.getUi().alert("系統提示", "目前為本地端獨立執行版本，無需雲端解鎖。", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// ⏰ 全自動排程安裝中心
// ==========================================
function setupDailyTrigger() {
  runWithAccessCheck_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const targetFunctions = ['recordDailySnapshot', 'recordTaiwanStock', 'recordUSStock', 'handleDropdownEdit'];
    const triggers = ScriptApp.getUserTriggers(ss);
    
    for (var i = 0; i < triggers.length; i++) {
      if (targetFunctions.indexOf(triggers[i].getHandlerFunction()) !== -1) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // 美股快照排程 (每日清晨)
    ScriptApp.newTrigger('recordUSStock')
      .timeBased().everyDays(1).atHour(6).nearMinute(15).create();

    // 台股快照排程 (每日下午收盤後)
    ScriptApp.newTrigger('recordTaiwanStock')
      .timeBased().everyDays(1).atHour(14).nearMinute(0).create();

    // 每日總覽快照排程
    ScriptApp.newTrigger('recordDailySnapshot')
      .timeBased().everyDays(1).atHour(14).nearMinute(15).create();

    // 編輯觸發監聽
    ScriptApp.newTrigger('handleDropdownEdit')
      .forSpreadsheet(ss).onEdit().create();

    ss.toast("🎉 本地端全自動排程與監聽已成功啟用！", "系統啟用成功", 6);
  });
}

// ============================================================================
// 🚀 選單／排程請直接呼叫各模組的實作函式（禁止與實作同名再自我呼叫）
// ⚠️ 舊版曾在此定義 recordTaiwanStock(){ recordTaiwanStock(ss) }，會蓋掉
//    「台美股紀錄.js」的實作並造成 Maximum call stack size exceeded。
// ============================================================================

function sendDailyBriefingMenu() {
  runWithAccessCheck_(function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sendDailyBriefing(ss);
    ss.toast("⚡ 每日盤後速報已觸發！", "四維決策中心", 3);
  });
}

function sendWeeklyDeepReportMenu() {
  runWithAccessCheck_(function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sendWeeklyDeepReport(ss);
    ss.toast("📊 週末深度週報已觸發！", "四維決策中心", 3);
  });
}

function sendMonthlyDeepReportMenu() {
  runWithAccessCheck_(function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sendMonthlyDeepReport(ss);
    ss.toast("🏆 月度資產大報告已觸發！", "四維決策中心", 3);
  });
}

function updateAllFutures() {
  runWithAccessCheck_(function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof updateFuturesQuotes === "function") updateFuturesQuotes(ss);
    if (typeof updateOverseasFuturesQuotes === "function") updateOverseasFuturesQuotes(ss);
    ss.toast("⚡ 國內與海外期貨行情已全部同步更新！", "四維決策中心", 3);
  });
}

function migrateFromOldVersion() {
  SpreadsheetApp.getUi().alert("系統提示", "目前為本地獨立版本，無需執行雲端版本移轉。", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// 🧠 智腦側邊欄與本機轉接中心
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
    if (typeof callDeepSeekAgent === 'function') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      return callDeepSeekAgent("deepseek-chat", "你是一個高冷、說話幽默犀利、極度重視數據邏輯的四維系統首席分析師。", prompt, ss);
    }
    return "❌ 本地端尚未定義 AI 核心呼叫函式";
  } catch(e) {
    return "❌ 執行智腦 API 失敗：" + e.toString();
  }
}

// 🌐 點選單時觸發，彈出 Web App 專屬連結視窗
function openWebAppDirectly() {
  var webAppUrl = "https://script.google.com/macros/s/AKfycbzfx-TbbU-s1-DVmllcGD_xqnzoPPOcfqvQy6P6rRSLQ2Q123433SUNZgLlRTKXagtykg/exec"; // 👈 請換成您的真實 Web App 網址
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\' , Roboto, sans-serif; text-align: center; padding: 25px; background-color: #f8fafc; border-radius: 8px;">' +
      '<h3 style="color: #0f172a; margin-bottom: 10px;">🚀 四維系統遠端 Web App</h3>' +
      '<p style="color: #475569; font-size: 14px; margin-bottom: 20px;">請點擊下方按鈕前往專屬操作介面：</p>' +
      '<a href="' + webAppUrl + '" target="_blank" style="display: inline-block; background-color: #0f172a; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">立即開啟 Web App 網頁</a>' +
    '</div>'
  ).setWidth(380).setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "🔗 外部系統導航");
}