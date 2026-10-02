/**
 * =========================================================================
 * 🏆 四維高槓桿財務風控系統 (LineNotification.gs v4.7 純推播版)
 * 特點：專職 DeepSeek 引擎分析 + 主動式 LINE 定時推播 (已移除即時問答 doPost)
 * =========================================================================
 */

/**
 * 取得當前動態時間字串
 */
function getCurrentTimeString() {
  const now = new Date();
  return Utilities.formatDate(now, "GMT+8", "yyyy年MM月dd日 HH:mm");
}

/**
 * AI 首席分析師基本人設 (DeepSeek 專用)
 */
function getSystemInstruction() {
  return "你是一個高冷、說話幽默犀利、極度重視數據邏輯、不灌心靈雞湯的四維高槓桿財務系統首席分析師。現在的真實時間是 " + getCurrentTimeString() + "。請嚴格基於當前提供的事實數據進行冷酷分析，嚴禁使用過往對話的歷史記憶猜測數據或日期。";
}

function notifyPushUser_(title, msg) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg || ""), title || "LINE 推播", 8);
  } catch (e) {
    Logger.log((title || "LINE") + ": " + msg);
  }
}

function sendDailyBriefing(ss) {
  return runLineBriefing_("盤後速報", ss, false, function(data) {
    return "請根據以下最新的盤後事實數據，為使用者生成一份【專業、幽默】的「每日盤後快報」。\n\n" +
      "【四維系統當前即時財務現況】：\n" +
      "- 總資產市值：" + data.overviewSummary.總資產市值 + " 元\n" +
      "- 負債總計：" + data.overviewSummary.負債總計 + " 元\n" +
      "- 實質淨資產：" + data.overviewSummary.淨資產 + " 元\n" +
      "- 今日單日獲利：" + data.overviewSummary.今日獲利 + " 元\n" +
      "- 當前質押維持率：【" + data.overviewSummary.當前質押維持率 + "】\n" +
      "- 實質槓桿倍數：【" + data.overviewSummary.實質槓桿 + "】倍\n\n" +
      "請聚焦在今日損益與質押維持率的安全跳動狀況，用你一貫高冷、數據導向的幽默風格，在 150 字內快速點評，不要廢話。";
  }, "⚡ 【四維系統 盤後即時速報】\n\n");
}

function sendWeeklyDeepReport(ss) {
  return runLineBriefing_("週末週報", ss, true, function(data) {
    return "請根據以下【百分之百真實】的財務數據，為使用者生成一期充滿洞察力、長文深度的「四維系統個人財務與高槓桿風險週報分析」。\n\n" +
      "【四維系統頂部核心數據】：\n" +
      "- 總資產市值：" + data.overviewSummary.總資產市值 + " 元\n" +
      "- 負債總計：" + data.overviewSummary.負債總計 + " 元\n" +
      "- 實質淨資產：" + data.overviewSummary.淨資產 + " 元\n" +
      "- 今日獲利：" + data.overviewSummary.今日獲利 + " 元\n" +
      "- 關鍵風控指標：當前質押維持率【" + data.overviewSummary.當前質押維持率 + "】、實質槓桿【" + data.overviewSummary.實質槓桿 + "】、資產總beta【" + data.overviewSummary.資產總beta + "】、夏普值【" + data.overviewSummary.夏普值 + "】。\n\n" +
      "【其餘四大核心分頁完整數據 (JSON)】：\n" +
      "- 1. 台股配置與質押狀況：" + data.extraDetails.台股區資料 + "\n" +
      "- 2. 美股標的持倉明細：" + data.extraDetails.美股區資料 + "\n" +
      "- 3. 期貨微型槓桿部位：" + data.extraDetails.期貨區資料 + "\n" +
      "- 4. 負債結構與利息流動：" + data.extraDetails.負債與資金資料 + "\n\n" +
      "請發揮你的語意理解能力，綜合以上分頁的所有事實數據，進行深度的風險勘驗報告。必須包含：\n" +
      "1. 台美股（正2、航太、半導體）高Beta配置與期貨部位在過去一週的波動與風控點評。\n" +
      "2. 評估整體負債結構（質押維持率、信貸、保單借款）的利息成本與安全防線。\n" +
      "3. 給予下週市場震盪時的硬核操作或心理防禦策略建議（保持幽默冷酷口吻，不灌雞湯）。";
  }, "📊 【四維系統 週末硬核深度週報】\n\n");
}

function sendMonthlyDeepReport(ss) {
  return runLineBriefing_("月度體檢", ss, true, function(data) {
    return "請根據以下【四維系統跨月核心財務現況】，為使用者生成一期具備「宏觀戰略視角、高冷幽默」的「月份個人資產風控與負債體檢大報告」。\n\n" +
      "【四維系統頂部核心數據】：\n" +
      "- 總資產市值：" + data.overviewSummary.總資產市值 + " 元\n" +
      "- 負債總計：" + data.overviewSummary.負債總計 + " 元\n" +
      "- 實質淨資產：" + data.overviewSummary.淨資產 + " 元\n" +
      "- 當前質押維持率：【" + data.overviewSummary.當前質押維持率 + "】\n" +
      "- 實質槓桿：【" + data.overviewSummary.實質槓桿 + "】倍\n" +
      "- 夏普值：【" + data.overviewSummary.夏普值 + "】\n\n" +
      "【過去近一整個月歷史紀錄區（JSON）】：\n" +
      data.dailySummaryRaw + "\n\n" +
      "請發揮你的數據洞察力，拉長到「整個月份 (約22-25個交易日)」的時間維度進行冷酷分析。必須包含：\n" +
      "1. 檢視本月整體資產規模與實質淨資產的推移推演。\n" +
      "2. 針對當前的「實質槓桿倍數」與「負債結構」進行月份健康度評估，精算利息成本是否正在蠶食現金水量。\n" +
      "3. 給予使用者下一個月在維持率控管、自律現金流上的核心戰略建議（嚴禁灌雞湯）。";
  }, "🏆 【四維系統 月度資產大局體檢】\n\n");
}

function runLineBriefing_(label, ss, withExtra, buildPrompt, linePrefix) {
  try {
    ss = (typeof resolveSpreadsheet === "function") ? resolveSpreadsheet(ss) : ss;
    if (!ss || typeof ss.getSheetByName !== "function") {
      notifyPushUser_("推播失敗", "❌ 找不到試算表物件");
      return { success: false, message: "找不到試算表" };
    }
    notifyPushUser_("LINE 推播", "⏳ 正在產生「" + label + "」…");
    var data = fetchFourDimensionsData(ss, withExtra);
    if (!data) {
      notifyPushUser_("推播失敗", "❌ 讀不到總覽資料，請確認有「總覽」分頁");
      return { success: false, message: "讀不到總覽" };
    }
    var promptData = buildPrompt(data);
    var aiReport = callDeepSeekAgent("deepseek-chat", getSystemInstruction(), promptData, ss);
    var pushRes = sendToLineBot(linePrefix + aiReport, ss);
    notifyPushUser_(pushRes.success ? "✅ 推播完成" : "❌ 推播失敗", pushRes.message);
    return pushRes;
  } catch (err) {
    notifyPushUser_("推播例外", "❌ " + err.message);
    return { success: false, message: err.message };
  }
}

// 🛠️ 【核心數據中樞】
function fetchFourDimensionsData(ss, withExtra) {
  ss = (typeof resolveSpreadsheet === "function") ? resolveSpreadsheet(ss) : ss;
  if (!ss) {
    Logger.log("❌ 找不到可用的試算表物件。");
    return null;
  }
  
  const mainSheetName = (typeof CONFIG !== 'undefined' && CONFIG.MAIN_SHEET) ? CONFIG.MAIN_SHEET : "總覽";
  const logSheetName = (typeof CONFIG !== 'undefined' && CONFIG.LOG_SHEET) ? CONFIG.LOG_SHEET : "每日紀錄";

  const sheet1 = ss.getSheetByName(mainSheetName);
  const sheet2 = ss.getSheetByName(logSheetName) || ss.getSheetByName("每日紀錄") || ss.getSheetByName("每日記錄");
  const sheet3 = ss.getSheetByName("台股區"); 
  const sheet4 = ss.getSheetByName("美股區"); 
  const sheet5 = ss.getSheetByName("期貨區"); 
  const sheet6 = ss.getSheetByName("負債/資金"); 
  
  if (!sheet1) {
    Logger.log("找不到【總覽】分頁。");
    return null;
  }
  
  const ovValues = sheet1.getRange(1, 1, 15, 6).getValues();
  
  let dailySummaryRaw = "無每日紀錄資料";
  if (sheet2) {
    try {
      const lastRow = sheet2.getLastRow();
      if (lastRow > 1) {
        const startRow = Math.max(1, lastRow - 24); 
        const numRows = lastRow - startRow + 1;
        const lastCol = Math.min(12, Math.max(1, sheet2.getLastColumn()));
        dailySummaryRaw = JSON.stringify(sheet2.getRange(startRow, 1, numRows, lastCol).getValues());
      }
    } catch(err) {
      dailySummaryRaw = "歷史紀錄讀取失敗：" + err.toString();
    }
  }
  
  const getSheetJson = (sheet, maxRows, maxCols) => {
    if (!sheet) return "分頁尚未建立或名稱不符";
    try {
      const lastRow = Math.min(sheet.getLastRow() || 0, maxRows || 20);
      const lastCol = Math.min(sheet.getLastColumn() || 0, maxCols || 12);
      if (lastRow === 0 || lastCol === 0) return "此分頁目前內容為空";
      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      return JSON.stringify(data);
    } catch(err) {
      return "讀取此分頁時發生被動攔截";
    }
  };

  const maintRatioRaw = ovValues[14][2]; 
  var extraDetails = { "台股區資料": "略", "美股區資料": "略", "期貨區資料": "略", "負債與資金資料": "略" };
  if (withExtra) {
    extraDetails = {
      "台股區資料": getSheetJson(sheet3, 15, 11),
      "美股區資料": getSheetJson(sheet4, 15, 11),
      "期貨區資料": getSheetJson(sheet5, 30, 12),
      "負債與資金資料": getSheetJson(sheet6, 25, 8)
    };
  }

  return {
    overviewSummary: {
      "總資產市值": ovValues[5][1],
      "負債總計": ovValues[11][1],
      "真_質押借款金額": ovValues[7][1],
      "淨資產": ovValues[14][0],
      "今日獲利": ovValues[14][1],
      "當前質押維持率": typeof maintRatioRaw === 'number' ? (maintRatioRaw * 100).toFixed(2) + "%" : maintRatioRaw,
      "實質槓桿": ovValues[3][4],
      "資產總beta": ovValues[5][4],
      "資產總alpha": ovValues[7][4],
      "夏普值": ovValues[13][4]
    },
    dailySummaryRaw: dailySummaryRaw,
    extraDetails: extraDetails
  };
}

// 🔀 【DeepSeek API 專屬呼叫函式】
function callDeepSeekAgent(modelName, systemInstruction, prompt, ss) {
  var docProps = PropertiesService.getDocumentProperties();
  var userProps = PropertiesService.getUserProperties();
  var scriptProps = PropertiesService.getScriptProperties();

  const apiKey = (docProps ? docProps.getProperty("DEEPSEEK_API_KEY") : null) ||
                 (userProps ? userProps.getProperty("DEEPSEEK_API_KEY") : null) ||
                 (scriptProps ? scriptProps.getProperty("DEEPSEEK_API_KEY") : null);
  
  if (!apiKey) {
    Logger.log("❌ 未設定 DEEPSEEK_API_KEY");
    return "⚠️ 系統未設定 DeepSeek API Key，請先於屬性設定中填入金鑰。";
  }

  const url = "https://api.deepseek.com/chat/completions"; 
  
  const payload = {
    "model": modelName || "deepseek-chat",
    "messages": [
      { "role": "system", "content": systemInstruction },
      { "role": "user", "content": prompt }
    ],
    "temperature": 0.2, 
    "stream": false
  };
  
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "headers": { "Authorization": "Bearer " + apiKey.trim() },
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) {
      return json.choices[0].message.content.trim();
    }
    return "DeepSeek 思考後未返回文字：" + response.getContentText();
  } catch(e) {
    return "DeepSeek 核心連線或解析錯誤：" + e.toString();
  }
}

// 🚀 【LINE 推播發送主控端】
function sendToLineBot(message, ss) {
  var docProps = PropertiesService.getDocumentProperties();
  var userProps = PropertiesService.getUserProperties();
  var scriptProps = PropertiesService.getScriptProperties();

  const token = (docProps ? docProps.getProperty("LINE_CHANNEL_ACCESS_TOKEN") : null) || 
                (userProps ? userProps.getProperty("LINE_CHANNEL_ACCESS_TOKEN") : null) || 
                (scriptProps ? scriptProps.getProperty("LINE_CHANNEL_ACCESS_TOKEN") : "");

  const userId = (docProps ? docProps.getProperty("MY_USER_ID") : null) || 
                 (userProps ? userProps.getProperty("MY_USER_ID") : null) || 
                 (scriptProps ? scriptProps.getProperty("MY_USER_ID") : "");
  
  if (!token || !userId) {
    Logger.log("❌ 缺少 LINE_CHANNEL_ACCESS_TOKEN 或 MY_USER_ID，無法發送推播");
    return { success: false, message: "❌ 未設定 LINE Token 或 User ID（文件／使用者／指令碼屬性）" };
  }

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = { "to": userId, "messages": [{ "type": "text", "text": String(message || "").substring(0, 4900) }] };
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "headers": {"Authorization": "Bearer " + token}, 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      return { success: true, message: "✅ LINE 已送出（HTTP " + code + "）" };
    }
    return { success: false, message: "❌ LINE API HTTP " + code + "：" + String(res.getContentText()).substring(0, 180) };
  } catch (err) {
    return { success: false, message: "❌ LINE 連線失敗：" + err.message };
  }
}

function testDirectLinePush() {
  var docProps = PropertiesService.getDocumentProperties();
  var scriptProps = PropertiesService.getScriptProperties();

  const token = (docProps ? docProps.getProperty("LINE_CHANNEL_ACCESS_TOKEN") : null) || 
                (scriptProps ? scriptProps.getProperty("LINE_CHANNEL_ACCESS_TOKEN") : "");

  const userId = (docProps ? docProps.getProperty("MY_USER_ID") : null) || 
                 (scriptProps ? scriptProps.getProperty("MY_USER_ID") : "");

  Logger.log("Token 前 6 碼: " + (token ? token.substring(0, 6) : "無"));
  Logger.log("User ID: " + (userId ? userId : "無"));

  if (!token || !userId) {
    Logger.log("❌ 缺少 Token 或 User ID");
    return;
  }

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = { "to": userId, "messages": [{ "type": "text", "text": "🔔 四維系統 LINE 推播連線測試成功！" }] };
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "headers": {"Authorization": "Bearer " + token}, 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  
  const res = UrlFetchApp.fetch(url, options);
  Logger.log("LINE 回應狀態碼: " + res.getResponseCode());
  Logger.log("LINE 回應內容: " + res.getContentText());
}

/**
 * 試算表物件解析請使用 台美股紀錄.js 的 resolveSpreadsheet（避免此檔覆寫後無法 openById）
 */