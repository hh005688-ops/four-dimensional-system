/**
 * 🏛️ 四維系統：Vercel 前端專用 GAS JSON API（doGet / doPost）
 * 讀寫邏輯仍委派既有模組；此檔只負責解析、路由、JSON 輸出與例外防呆。
 * 前端請用 text/plain POST JSON，以避免 CORS preflight。
 */

function doGet(e) {
  return routeGasApi_(e);
}

function doPost(e) {
  return routeGasApi_(e);
}

function jsonOutput_(obj) {
  var payload = (obj && typeof obj === "object") ? obj : { success: false, message: "❌ 空回應" };
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseApiPayload_(e) {
  var payload = {};
  if (e && e.postData && e.postData.contents) {
    try {
      payload = JSON.parse(e.postData.contents) || {};
    } catch (err) {
      throw new Error("POST JSON 解析失敗: " + err.message);
    }
  }
  if (e && e.parameter) {
    var keys = Object.keys(e.parameter);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (payload[k] === undefined || payload[k] === "") {
        payload[k] = e.parameter[k];
      }
    }
  }
  if (typeof payload.params === "string") {
    try { payload.params = JSON.parse(payload.params); } catch (err2) { payload.params = {}; }
  }
  if (!payload.params || typeof payload.params !== "object") payload.params = {};
  return payload;
}

function routeGasApi_(e) {
  try {
    var payload = parseApiPayload_(e);
    var action = String(payload.action || "").trim();
    if (!action) {
      return jsonOutput_({
        success: true,
        message: "✅ 四維 GAS API 就緒",
        hint: "POST { action, sheetId, params }，Content-Type 請用 text/plain"
      });
    }
    var result = dispatchApiAction_(action, payload);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, message: "❌ API 例外: " + err.message });
  }
}

function dispatchApiAction_(action, payload) {
  var sheetId = payload.sheetId || payload.sheetIdentifier || "";
  var params = payload.params || {};
  var actionType = payload.actionType || params.actionType || "";

  switch (action) {
    case "health":
      return { success: true, message: "✅ 四維 GAS API 就緒" };

    case "getInitDataUniversal":
      return getInitDataUniversal(sheetId);

    case "dispatchUniversalAction":
      if (!actionType) return { success: false, message: "❌ 缺少 actionType" };
      return dispatchUniversalAction(sheetId, actionType, params);

    case "getDailyLogHistoryUniversal":
      return getDailyLogHistoryUniversal(sheetId);

    case "getStockLogHistoryUniversal":
      return getStockLogHistoryUniversal(sheetId);

    case "getSnapshotHistoryDataUniversal":
      return getSnapshotHistoryDataUniversal(sheetId);

    case "triggerTaiwanStockSnapshotUniversal":
      return triggerTaiwanStockSnapshotUniversal(sheetId);

    case "triggerUSStockSnapshotUniversal":
      return triggerUSStockSnapshotUniversal(sheetId);

    case "triggerDailySnapshotUniversal":
      return triggerDailySnapshotUniversal(sheetId);

    case "refreshBetaAlphaUniversal":
      return refreshBetaAlphaUniversal(sheetId);

    case "refreshMarketDataUniversal":
      return refreshMarketDataUniversal(sheetId);

    case "fetchStockInfoUniversal":
      return fetchStockInfoUniversal(payload.market || params.market, payload.code || params.code);

    default:
      return { success: false, message: "❌ 未知的 API action: " + action };
  }
}
