// 轉送層：先在伺服器端檢查自用密碼，通過後才帶著暗號轉呼叫 GAS
const { send, checkPassword, readBody } = require("./_lib");

// 只放行網頁實際會用到的 GAS action，其餘一律擋下
const ALLOWED_ACTIONS = new Set([
  "getInitDataUniversal",
  "dispatchUniversalAction",
  "getDailyLogHistoryUniversal",
  "getStockLogHistoryUniversal",
  "getSnapshotHistoryDataUniversal",
  "triggerTaiwanStockSnapshotUniversal",
  "triggerUSStockSnapshotUniversal",
  "triggerDailySnapshotUniversal",
  "refreshBetaAlphaUniversal",
  "refreshMarketDataUniversal",
  "fetchStockInfoUniversal",
]);

module.exports = async function handler(req, res) {
  if (req.method !== "POST") return send(res, 405, { success: false, message: "僅接受 POST" });
  if (!checkPassword(req, res)) return;

  const { GAS_URL, GAS_SECRET, SHEET_ID } = process.env;
  if (!GAS_URL || !GAS_SECRET || !SHEET_ID) {
    return send(res, 500, { success: false, message: "❌ 伺服器尚未設定 GAS_URL / GAS_SECRET / SHEET_ID" });
  }

  const body = readBody(req);
  const action = String(body.action || "");
  if (!ALLOWED_ACTIONS.has(action)) {
    return send(res, 400, { success: false, message: "❌ 不允許的 action" });
  }

  // 試算表 ID 一律由伺服器環境變數決定，忽略網頁傳來的值
  const forward = {
    action: action,
    sheetId: SHEET_ID,
    actionType: body.actionType,
    params: body.params && typeof body.params === "object" ? body.params : {},
    market: body.market,
    code: body.code,
    secret: GAS_SECRET,
  };

  let text;
  try {
    const r = await fetch(GAS_URL, {
      method: "POST",
      redirect: "follow",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(forward),
    });
    text = await r.text();
  } catch (err) {
    return send(res, 502, { success: false, message: "❌ 無法連線到 GAS" });
  }

  let data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    return send(res, 502, { success: false, message: "❌ GAS 回傳格式錯誤（請確認已重新部署）" });
  }
  send(res, 200, data);
};
