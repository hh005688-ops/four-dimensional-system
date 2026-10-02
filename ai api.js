/**
 * =========================================================================
 * 🏆 四維智腦 - 側邊欄專屬 API 核心 (API核心.gs 完整版)
 * 特點：多層級金鑰讀取、Gemini / DeepSeek 自動降級切換、完整例外防呆
 * =========================================================================
 */

/**
 * 側邊欄專用核心 API 主入口 (供前端 main.gs / 側邊欄呼叫)
 * @param {string} prompt - 使用者輸入的指令或選區數據
 * @param {Object} [ss] - 前端傳入的試算表物件
 */
function CALL_CORE_API_BACKEND(prompt, ss) {
  // 🛡️ 1. 提示詞防呆
  if (!prompt || typeof prompt !== "string") {
    prompt = "請進行簡單的系統連線測試，回覆：『四維智腦側邊欄連線正常！』";
  }

  // 🛡️ 2. 解析試算表物件
  ss = resolveSpreadsheet(ss);

  // 🛡️ 3. 多層級讀取 API 金鑰 (Script > Document > User)
  var scriptProps = PropertiesService.getScriptProperties();
  var docProps    = PropertiesService.getDocumentProperties();
  var userProps   = PropertiesService.getUserProperties();

  const geminiKey = (scriptProps ? scriptProps.getProperty("GEMINI_API_KEY") : null) ||
                    (docProps ? docProps.getProperty("GEMINI_API_KEY") : null) ||
                    (userProps ? userProps.getProperty("GEMINI_API_KEY") : null);

  const deepseekKey = (scriptProps ? scriptProps.getProperty("DEEPSEEK_API_KEY") : null) ||
                      (docProps ? docProps.getProperty("DEEPSEEK_API_KEY") : null) ||
                      (userProps ? userProps.getProperty("DEEPSEEK_API_KEY") : null);

  // 🚀 主力：優先呼叫 Gemini (速度快、回應即時)
  if (geminiKey) {
    return callGeminiApi_(prompt, geminiKey);
  }

  // 🛡️ 備援：若無 Gemini 則降級切換 DeepSeek
  if (deepseekKey) {
    return callDeepSeekApi_(prompt, deepseekKey);
  }

  return "⚠️ 系統未檢測到有效金鑰：請於後端專案「專案設定 → 指令碼屬性」中填入 GEMINI_API_KEY 或 DEEPSEEK_API_KEY。";
}

/**
 * 內部底層：Gemini API 呼叫模組
 */
function callGeminiApi_(prompt, apiKey) {
  if (!apiKey) return "⚠️ 系統設定缺失：未設定 Gemini API 金鑰。";

  const MODEL_NAME = "gemini-2.5-flash"; 
  const SYSTEM_DESC = "你是一個專業投資特助，專精四維系統。風格專業、嚴謹、邏輯清晰，提供客觀的金融數據分析與建議，不使用戲謔語氣。";
  
  const url = "https://generativelanguage.googleapis.com/v1/models/" + MODEL_NAME + ":generateContent?key=" + apiKey.trim();
  const payload = { 
    "contents": [{ 
      "parts": [{ "text": SYSTEM_DESC + "\n\n用戶提問或數據：\n" + prompt }] 
    }] 
  };
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };
  
  try {
    const res = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(res.getContentText());
    if (res.getResponseCode() === 200 && json.candidates && json.candidates[0].content) {
      return json.candidates[0].content.parts[0].text.trim();
    }
    return "⚠️ Gemini API 配額異常或金鑰無效：" + res.getContentText();
  } catch (e) { 
    return "❌ Gemini 連線失敗：" + e.message; 
  }
}

/**
 * 內部底層：DeepSeek API 備援呼叫模組
 */
function callDeepSeekApi_(prompt, apiKey) {
  if (!apiKey) return "⚠️ 系統設定缺失：未設定 DeepSeek API 金鑰。";

  const SYSTEM_DESC = "你是一個專業投資特助，專精四維系統。風格專業、嚴謹、邏輯清晰，提供客觀的金融數據分析與建議。";
  const url = "https://api.deepseek.com/chat/completions";
  const payload = {
    "model": "deepseek-chat",
    "messages": [
      { "role": "system", "content": SYSTEM_DESC },
      { "role": "user", "content": prompt }
    ],
    "temperature": 0.3
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": "Bearer " + apiKey.trim() },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(res.getContentText());
    if (res.getResponseCode() === 200 && json.choices && json.choices[0].message) {
      return json.choices[0].message.content.trim();
    }
    return "⚠️ DeepSeek API 回應異常：" + res.getContentText();
  } catch (e) {
    return "❌ DeepSeek 連線失敗：" + e.message;
  }
}

/**
 * 側邊欄專用 API 金鑰檢測工具
 */
function testApiKey() {
  var scriptProps = PropertiesService.getScriptProperties();
  var docProps    = PropertiesService.getDocumentProperties();
  var userProps   = PropertiesService.getUserProperties();

  const gmKey = (scriptProps ? scriptProps.getProperty("GEMINI_API_KEY") : null) ||
                (docProps ? docProps.getProperty("GEMINI_API_KEY") : null) ||
                (userProps ? userProps.getProperty("GEMINI_API_KEY") : null);

  const dsKey = (scriptProps ? scriptProps.getProperty("DEEPSEEK_API_KEY") : null) ||
                (docProps ? docProps.getProperty("DEEPSEEK_API_KEY") : null) ||
                (userProps ? userProps.getProperty("DEEPSEEK_API_KEY") : null);

  if (gmKey) {
    Logger.log("✅ 側邊欄 Gemini 金鑰就緒，前 4 碼為：" + gmKey.trim().substring(0, 4));
  } else if (dsKey) {
    Logger.log("✅ 側邊欄備用 DeepSeek 金鑰就緒，前 4 碼為：" + dsKey.trim().substring(0, 4));
  } else {
    Logger.log("❌ 側邊欄未偵測到任何 API Key！");
  }
}

/**
 * 智腦特助側邊欄 HTML 輸出 (供前端呼叫)
 */
function getBrainSidebarHtml() {
  // 若後端有專屬的智腦 HTML (例如 Brain.html)，若無則建立基礎介面
  try {
    return HtmlService.createHtmlOutputFromFile('側邊欄ai')
      .setTitle('🧠 四維智腦特助')
      .setWidth(360);
  } catch (e) {
    return HtmlService.createHtmlOutput('<h3 style="color:#58a6ff;padding:15px;background:#0d1117;font-family:sans-serif;">🧠 四維智腦特助已連線</h3><p style="color:#8b949e;padding:0 15px;background:#0d1117;font-family:sans-serif;">請於輸入框或對話介面下達指令。</p>')
      .setTitle('🧠 四維智腦特助')
      .setWidth(360);
  }
}

/**
 * 開啟四維調度器側邊欄
 */
function showFourDimensionsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('側邊欄操作')
    .setTitle('⚡ 四維資產即時調度器')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}
