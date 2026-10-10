// 只負責驗證密碼：網頁第一次輸入密碼時用來確認對不對
const { send, checkPassword } = require("./_lib");

module.exports = function handler(req, res) {
  if (req.method !== "POST") return send(res, 405, { success: false, message: "僅接受 POST" });
  if (!checkPassword(req, res)) return;
  send(res, 200, { success: true });
};
