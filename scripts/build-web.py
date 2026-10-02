# -*- coding: utf-8 -*-
from pathlib import Path
import re

root = Path(__file__).resolve().parents[1]
web = root / "web"
(css_dir := web / "css").mkdir(parents=True, exist_ok=True)
(js_dir := web / "js").mkdir(parents=True, exist_ok=True)

def read(name):
    return (root / name).read_text(encoding="utf-8")

def inner(text, tag):
    m = re.search(rf"<{tag}[^>]*>(.*)</{tag}>", text, re.S | re.I)
    return m.group(1).strip() if m else text

css = inner(read("風格及css變數.html"), "style")
asset = read("資產圖表.html")
m = re.search(r"<style>(.*?)</style>", asset, re.S)
if m:
    css += "\n" + m.group(1).strip()
(css_dir / "app.css").write_text(css, encoding="utf-8")

app_js = inner(read("互動邏輯.html"), "script")
(js_dir / "app.js").write_text(app_js, encoding="utf-8")

parts = re.split(r"<script>", asset, maxsplit=1)
asset_html = parts[0]
asset_html = re.sub(r"<style>.*?</style>", "", asset_html, flags=re.S).strip()
asset_js = re.sub(r"</script>\s*$", "", parts[1]).strip()
(js_dir / "charts.js").write_text(asset_js, encoding="utf-8")

html = read("Index.html")
html = re.sub(r'<base target="_top">\s*', "", html)
html = html.replace(
    "<?!= HtmlService.createHtmlOutputFromFile('風格及css變數').getContent(); ?>",
    '<link rel="stylesheet" href="./css/app.css">',
)
html = html.replace(
    "<?!= HtmlService.createHtmlOutputFromFile('首頁').getContent(); ?>",
    read("首頁.html"),
)
html = html.replace(
    "<?!= HtmlService.createHtmlOutputFromFile('資金調度').getContent(); ?>",
    read("資金調度.html"),
)
html = html.replace(
    "<?!= HtmlService.createHtmlOutputFromFile('資產圖表').getContent(); ?>",
    asset_html,
)
html = html.replace(
    "<?!= HtmlService.createHtmlOutputFromFile('互動邏輯').getContent(); ?>",
    '<script src="./js/app.js"></script>\n  <script src="./js/charts.js"></script>',
)
html = html.replace(
    "<head>",
    "<head>\n  <meta charset=\"UTF-8\">\n  <title>四維掌上中控台</title>",
    1,
)
(web / "index.html").write_text(html, encoding="utf-8")
print("built", web)
