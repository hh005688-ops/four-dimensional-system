// 共用工具（檔名開頭是底線，Vercel 不會把它當成 API 入口）
const crypto = require("crypto");

function sha256(text) {
  return crypto.createHash("sha256").update(String(text), "utf8").digest();
}

// 固定時間比對，避免被「猜時間差」破解
function safeEqual(a, b) {
  return crypto.timingSafeEqual(sha256(a), sha256(b));
}

function send(res, status, obj) {
  res.status(status).setHeader("Cache-Control", "no-store");
  res.json(obj);
}

// 檢查自用密碼。通過回傳 true；不通過會直接回應並回傳 false。
function checkPassword(req, res) {
  const expected = process.env.APP_PASSWORD;
  if (!expected) {
    send(res, 500, { success: false, message: "❌ 伺服器尚未設定 APP_PASSWORD" });
    return false;
  }
  const given = req.headers["x-app-password"];
  if (typeof given !== "string" || given === "" || !safeEqual(given, expected)) {
    send(res, 401, { success: false, message: "🔒 密碼錯誤或未提供", code: "UNAUTHORIZED" });
    return false;
  }
  return true;
}

function readBody(req) {
  let b = req.body;
  if (typeof b === "string") {
    try { b = JSON.parse(b); } catch (e) { b = {}; }
  }
  return b && typeof b === "object" ? b : {};
}

module.exports = { send, checkPassword, readBody };
