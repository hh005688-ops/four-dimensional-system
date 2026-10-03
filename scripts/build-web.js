const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const web = path.join(root, "web");

try {
  buildWeb();
} catch (err) {
  console.error("❌ build-web.js 失敗：", err.message);
  process.exit(1);
}

function resolveSource(name) {
  const exact = path.join(root, name);
  if (fs.existsSync(exact)) return exact;
  var entries = [];
  try {
    entries = fs.readdirSync(root);
  } catch (e) {
    throw new Error("無法讀取專案根目錄：" + e.message);
  }
  const found = entries.find(function (f) {
    return f.toLowerCase() === String(name).toLowerCase();
  });
  if (found) return path.join(root, found);
  throw new Error(
    "找不到來源檔「" + name + "」。Linux/Vercel 檔名大小寫必須完全一致。根目錄現有：" +
      entries.filter(function (f) { return /\.(html|js|css|json|png)$/i.test(f); }).join(", ")
  );
}

function read(name) {
  return fs.readFileSync(resolveSource(name), "utf8");
}

function inner(text, tag) {
  const m = text.match(new RegExp("<" + tag + "[^>]*>([\\s\\S]*)</" + tag + ">", "i"));
  return m ? m[1].trim() : text;
}

function stripStrayCloseScript(html) {
  return String(html).replace(/<\/script>\s*$/i, "").trim();
}

function buildWeb() {
  fs.mkdirSync(path.join(web, "css"), { recursive: true });
  fs.mkdirSync(path.join(web, "js"), { recursive: true });

  let css = inner(read("風格及css變數.html"), "style");
  const asset = read("資產圖表.html");
  const styleMatch = asset.match(/<style>([\s\S]*?)<\/style>/);
  if (styleMatch) css += "\n" + styleMatch[1].trim();
  fs.writeFileSync(path.join(web, "css", "app.css"), css, "utf8");

  fs.writeFileSync(path.join(web, "js", "app.js"), inner(read("互動邏輯.html"), "script"), "utf8");

  const parts = asset.split(/<script>/);
  let assetHtml = parts[0].replace(/<style>[\s\S]*?<\/style>/g, "").trim();
  const assetJs = (parts[1] || "").replace(/<\/script>\s*$/, "").trim();
  fs.writeFileSync(path.join(web, "js", "charts.js"), assetJs, "utf8");

  // 根目錄 index.html 是獨立 Tailwind 精簡頁，不是中控台骨架。
  // 完整戰情室由 app-shell.html + 首頁 / 資金調度 / 資產圖表 / 互動邏輯 組裝。
  let html = read("app-shell.html");
  html = html.replace("@@INCLUDE_HOME@@", stripStrayCloseScript(read("首頁.html")));
  html = html.replace("@@INCLUDE_DISPATCH@@", stripStrayCloseScript(read("資金調度.html")));
  html = html.replace("@@INCLUDE_CHARTS_HTML@@", assetHtml);

  if (html.indexOf("cdn.tailwindcss.com") !== -1 || html.indexOf("FOUR-DIMENSIONAL SYSTEM") !== -1) {
    throw new Error("組裝結果誤用了精簡 Tailwind 頁，已中止以免蓋掉中控台。");
  }
  if (html.indexOf("tabHome") === -1 || html.indexOf("dispNetWorth") === -1 || html.indexOf("./js/app.js") === -1) {
    throw new Error("組裝結果缺少中控台必要區塊，已中止以免寫入錯誤介面。");
  }

  const publicDir = path.join(root, "public");
  const webPublic = path.join(web, "public");
  if (fs.existsSync(publicDir)) {
    fs.mkdirSync(webPublic, { recursive: true });
    fs.readdirSync(publicDir).forEach(function (f) {
      if (/\.png$/i.test(f)) fs.copyFileSync(path.join(publicDir, f), path.join(webPublic, f));
    });
  }

  const manifestSrc = path.join(root, "manifest.json");
  if (fs.existsSync(manifestSrc)) {
    fs.copyFileSync(manifestSrc, path.join(web, "manifest.json"));
  }

  fs.writeFileSync(path.join(web, "index.html"), html, "utf8");
  console.log("built", web);
}
