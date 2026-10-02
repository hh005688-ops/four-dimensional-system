const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const web = path.join(root, "web");
fs.mkdirSync(path.join(web, "css"), { recursive: true });
fs.mkdirSync(path.join(web, "js"), { recursive: true });

const read = (name) => fs.readFileSync(path.join(root, name), "utf8");
const inner = (text, tag) => {
  const m = text.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*)</${tag}>`, "i"));
  return m ? m[1].trim() : text;
};

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

let html = read("Index.html");
html = html.replace(/<base target="_top">\s*/, "");
html = html.replace(
  "<?!= HtmlService.createHtmlOutputFromFile('風格及css變數').getContent(); ?>",
  '<link rel="stylesheet" href="./css/app.css">'
);
html = html.replace("<?!= HtmlService.createHtmlOutputFromFile('首頁').getContent(); ?>", read("首頁.html"));
html = html.replace("<?!= HtmlService.createHtmlOutputFromFile('資金調度').getContent(); ?>", read("資金調度.html"));
html = html.replace("<?!= HtmlService.createHtmlOutputFromFile('資產圖表').getContent(); ?>", assetHtml);
html = html.replace(
  "<?!= HtmlService.createHtmlOutputFromFile('互動邏輯').getContent(); ?>",
  '<script src="./js/app.js"></script>\n  <script src="./js/charts.js"></script>'
);
html = html.replace("<head>", '<head>\n  <meta charset="UTF-8">\n  <title>四維掌上中控台</title>');
const iconTags =
  '\n  <!-- iOS 專用 -->\n' +
  '  <link rel="apple-touch-icon" sizes="180x180" href="/public/icon.png">\n' +
  '  <!-- Android 與通用 PWA 專用 -->\n' +
  '  <link rel="icon" type="image/png" sizes="192x192" href="/public/icon.png">\n' +
  '  <!-- 設定瀏覽器網址列的主題色 -->\n' +
  '  <meta name="theme-color" content="#1e293b">';
if (html.indexOf('rel="apple-touch-icon"') === -1) {
  html = html.replace("</head>", iconTags + "\n</head>");
}

const publicIcon = path.join(root, "public", "icon.png");
const webPublic = path.join(web, "public");
if (fs.existsSync(publicIcon)) {
  fs.mkdirSync(webPublic, { recursive: true });
  fs.copyFileSync(publicIcon, path.join(webPublic, "icon.png"));
}

fs.writeFileSync(path.join(web, "index.html"), html, "utf8");
console.log("built", web);
