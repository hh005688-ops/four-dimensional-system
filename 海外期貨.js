// ============================================================================
// 🏛️ 四維系統 - 海外期貨行情解析核心 (海外期貨價格.gs - 門神安全防護版)
// ============================================================================

/**
 * 🔑 本地門神檢查：驗證此複本是否已解鎖授權
 */
function checkSystemAccess_() {
  try {
    var cache = PropertiesService.getDocumentProperties();
    var isUnlocked = cache.getProperty('SYSTEM_UNLOCKED');
    if (isUnlocked === 'TRUE') {
      return true;
    }
    SpreadsheetApp.getUi().alert("🔒 授權失敗", "此試算表尚未啟用或授權金鑰不符，請點擊上方選單進行【啟動/解鎖系統】！", SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ 系統錯誤", "驗證過程發生異常: " + err.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * ⚡ 海外期貨行情更新專用入口 (已改名避免與其他模組的 runDailyRecord 衝突)
 */
function updateOverseasFuturesMain() {
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  updateOverseasFuturesQuotes(ss);
}

/**
 * ⚡ 極速批次更新海外期貨行情 (鎖定 16~22 列部位明細)
 * 支援副本建立免設定、自動解析 CME 代碼與中文月份
 */
function updateOverseasFuturesQuotes(ss) {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}

  // 海外期貨字典表
  const OVERSEAS_FUT_DICT = {
    // 美股四大指數 (標準 / 微型)
    'NQ':  { symbol: 'NQ=F',  name: '那斯達克期貨' },
    'MNQ': { symbol: 'MNQ=F', name: '微型那指期貨' },
    'ES':  { symbol: 'ES=F',  name: '標普500期貨' },
    'MES': { symbol: 'MES=F', name: '微型標普期貨' },
    'YM':  { symbol: 'YM=F',  name: '道瓊指數期貨' },
    'MYM': { symbol: 'MYM=F', name: '微型道瓊期貨' },
    'RTY': { symbol: 'RTY=F', name: '羅素2000期貨' },
    'M2K': { symbol: 'M2K=F', name: '微型羅素期貨' },

    // 能源 / 貴金屬 / 原物料
    'CL':  { symbol: 'CL=F',  name: '輕原油期貨' },
    'MCL': { symbol: 'MCL=F', name: '微型原油期貨' },
    'GC':  { symbol: 'GC=F',  name: '黃金期貨' },
    'MGC': { symbol: 'MGC=F', name: '微型黃金期貨' },
    'SI':  { symbol: 'SI=F',  name: '白銀期貨' },
    'HG':  { symbol: 'HG=F',  name: '高級銅期貨' },
    'NG':  { symbol: 'NG=F',  name: '天然氣期貨' },

    // 美債 / 外匯
    'ZN':  { symbol: 'ZN=F',  name: '美國10年公債' },
    'ZB':  { symbol: 'ZB=F',  name: '美國30年公債' },
    'DX':  { symbol: 'DX-Y.NYB', name: '美元指數期貨' }
  };

  const CME_MONTH_MAP = {
    'F': '01月', 'G': '02月', 'H': '03月', 'J': '04月',
    'K': '05月', 'M': '06月', 'N': '07月', 'Q': '08月',
    'U': '09月', 'V': '10月', 'X': '11月', 'Z': '12月'
  };

  // 1. 試算表物件解析 (相容副本建立)
  if (!ss || typeof ss.getSheetByName !== 'function') {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }

  if (!ss) {
    Logger.log("❌ 找不到可用的試算表物件");
    return;
  }

  const sheet = ss.getSheetByName('期貨區');
  if (!sheet) {
    Logger.log("⚠️ 找不到「期貨區」工作表，跳過執行");
    return;
  }

  const startRow = 16;
  const endRow = 22;
  const numRows = endRow - startRow + 1;

  // 2. 批次讀取 A 欄海外期貨代號
  const codeList = sheet.getRange(startRow, 1, numRows, 1).getValues();
  const namePriceOutputs = [];

  for (let i = 0; i < codeList.length; i++) {
    const rawCode = String(codeList[i][0] || '').trim().toUpperCase();

    if (!rawCode) {
      namePriceOutputs.push(["", ""]);
      continue;
    }

    let commodity = rawCode;
    let monthLabel = "";

    const matchCME = rawCode.match(/^([A-Z0-9]{2,4})([FGHJKMNQUVXZ])(\d{2})$/);
    const matchNum = rawCode.match(/^([A-Z0-9]{2,4})(\d{2,6})$/);

    if (matchCME && OVERSEAS_FUT_DICT[matchCME[1]]) {
      commodity = matchCME[1];
      const monthLetter = matchCME[2];
      const year = matchCME[3];
      monthLabel = ` 20${year}${monthLetter}(${CME_MONTH_MAP[monthLetter]})`;
    } else if (matchNum && OVERSEAS_FUT_DICT[matchNum[1]]) {
      commodity = matchNum[1];
      monthLabel = ` ${matchNum[2]}`;
    }

    const meta = OVERSEAS_FUT_DICT[commodity] || OVERSEAS_FUT_DICT[rawCode];
    const lookupSymbol = meta ? meta.symbol : (rawCode.includes('=F') ? rawCode : `${rawCode}=F`);
    const baseName = meta ? meta.name : `${commodity}期貨`;
    const finalDisplayName = baseName + monthLabel;

    // 抓取 Yahoo 現價 (數值防呆)
    const price = fetchSingleYahooV8Price_(lookupSymbol);
    namePriceOutputs.push([finalDisplayName, price > 0 ? price : ""]);
  }

  // 3. 一次性回寫 B 欄 (名稱) 與 C 欄 (當前價位)
  sheet.getRange(startRow, 2, numRows, 2).setValues(namePriceOutputs);
  SpreadsheetApp.flush();
  Logger.log("✅ 海外期貨行情已更新完成！");
}

/**
 * 🛠️ 專用底層：抓取 Yahoo Finance v8 即時現價 (純數值輸出)
 */
function fetchSingleYahooV8Price_(symbol) {
  if (!symbol) return 0;

  const querySymbol = String(symbol).trim().toUpperCase();
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(querySymbol)}?interval=1d&range=5d`;
  
  const options = {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    },
    timeoutInMilliseconds: 6000
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() !== 200) {
      Logger.log(`⚠️ 期貨 [${querySymbol}] 抓取失敗，HTTP 狀態碼: ${res.getResponseCode()}`);
      return 0;
    }

    const json = JSON.parse(res.getContentText());
    const result = json.chart && json.chart.result && json.chart.result[0];

    if (!result || !result.meta) {
      return 0;
    }

    // 取得即時現價或最後收盤價，強制轉型防呆
    const price = Number(result.meta.regularMarketPrice) || Number(result.meta.chartPreviousClose) || 0;
    return price;
  } catch (e) {
    Logger.log(`⚠️ 期貨 [${querySymbol}] 請求例外: ${e.message}`);
    return 0;
  }
}