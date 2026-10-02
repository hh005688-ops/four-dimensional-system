// ============================================================================
// 🏛️ 四維系統 - 台指期貨行情解析核心 (台指期價格.gs - 門神安全防護版)
// ============================================================================

const CORE_FUTURES_DICT = {
  'TX': '臺股期貨',   'TXF': '臺股期貨',
  'MTX': '小型臺指',  'MXF': '小型臺指',
  'TMF': '微型臺指',
  'TE': '電子期貨',   'EXF': '電子期貨',
  'TF': '金融期貨',   'FXF': '金融期貨'
};

/**
 * ⚡ 國內期貨行情更新專用入口 (已加門神)
 */
function updateFuturesQuotesMain() {
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  updateFuturesQuotes(ss);
}

function updateFuturesQuotes(ss) {
  // 🚪 門神檢查：沒授權直接擋下
  var accessCheck = checkLocalSystemAccess_();
if (!accessCheck.success) {
  SpreadsheetApp.getUi().alert("🔒 授權失敗", accessCheck.message, SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}

  if (!ss || typeof ss.getSheetByName !== 'function') {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch(e) {}
  }

  if (!ss) {
    throw new Error("此為後端 Library 函式，請由前端試算表之 main.gs 傳入 ss 執行！");
  }

  const sheet = ss.getSheetByName('期貨區');
  if (!sheet) {
    throw new Error("找不到「期貨區」工作表，請確認工作表名稱是否完全相符！");
  }

  const allTaifexData = fetchAllTaifexData_();
  if (!allTaifexData || allTaifexData.length === 0) {
    throw new Error("無法連線至期交所 OpenAPI 或目前非開盤/報價時段");
  }

  const startRow = 2;
  const endRow = 8;
  const numRows = endRow - startRow + 1;

  const codeList = sheet.getRange(startRow, 1, numRows, 1).getValues();
  const namePriceOutputs = [];

  for (let i = 0; i < codeList.length; i++) {
    const rawCode = String(codeList[i][0] || '').trim().toUpperCase();

    if (!rawCode) {
      namePriceOutputs.push(["", ""]);
      continue;
    }

    const match = findTaifexFromMemory_(rawCode, allTaifexData);
    if (match) {
      namePriceOutputs.push([match.name, match.price]);
    } else {
      namePriceOutputs.push(["查無此合約", 0]);
    }
  }

  sheet.getRange(startRow, 2, numRows, 2).setValues(namePriceOutputs);
}

/**
 * 抓取期交所 OpenAPI 全市場資料 (含快取)
 */
function fetchAllTaifexData_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = "TAIFEX_ALL_FUTURES_RAW";
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  const url = "https://openapi.taifex.com.tw/v1/DailyMarketReportFut";
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true,
      headers: { "Accept": "application/json" }
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (Array.isArray(data) && data.length > 0) {
        try {
          cache.put(cacheKey, JSON.stringify(data), 60);
        } catch (e) {}
        return data;
      }
    }
  } catch (err) {
    Logger.log("TAIFEX API 請求失敗: " + err.message);
  }
  return null;
}

/**
 * 記憶體高速比對期貨商品 (指數期貨 + 個股期貨 雙模相容)
 */
function findTaifexFromMemory_(rawInput, dataList) {
  if (!rawInput || !dataList) return null;

  let sym = rawInput.trim().toUpperCase();
  let targetMonth = "";

  const monthMatch = sym.match(/^([A-Z]+)(\d{2,6})$/);
  if (monthMatch) {
    sym = monthMatch[1];
    let rawMonth = monthMatch[2];
    
    if (rawMonth.length === 3 && rawMonth.startsWith('0')) {
      rawMonth = rawMonth.slice(1);
    }
    if (rawMonth.length === 2) {
      targetMonth = `${new Date().getFullYear()}${rawMonth}`;
    } else {
      targetMonth = rawMonth;
    }
  }

  const INDEX_ALIAS_MAP = {
    'TXF': 'TX', 'TX': 'TX',
    'MXF': 'MTX', 'MTX': 'MTX',
    'TMF': 'TMF',
    'EXF': 'TE', 'TE': 'TE',
    'FXF': 'TF', 'TF': 'TF'
  };

  let searchSyms = [];
  if (INDEX_ALIAS_MAP[sym]) {
    searchSyms = [INDEX_ALIAS_MAP[sym], sym];
  } else {
    const baseStockSym = (sym.endsWith('F') && sym.length === 3) ? sym.slice(0, 2) : sym;
    searchSyms = [sym, baseStockSym, baseStockSym + 'F'];
  }

  const matchedList = dataList.filter(item => {
    const code = String(item["Contract"] || item["Symbol"] || '').trim().toUpperCase();
    const month = String(item["ContractMonth(Week)"] || item["ContractMonth"] || '').trim();
    const isCodeMatch = searchSyms.includes(code);
    const isStandardMonth = !month.includes('/') && !month.includes('W');
    return isCodeMatch && isStandardMonth;
  });

  if (matchedList.length === 0) return null;

  let targetItem = null;
  if (targetMonth) {
    targetItem = matchedList.find(item => {
      const month = String(item["ContractMonth(Week)"] || item["ContractMonth"] || '').trim();
      return month.endsWith(targetMonth) || month.endsWith(targetMonth.slice(-2));
    });

    if (!targetItem) {
      return {
        code: sym,
        month: targetMonth,
        name: `查無 ${targetMonth} 合約`,
        price: 0
      };
    }
  } else {
    targetItem = matchedList[0];
  }

  const finalCode = String(targetItem["Contract"] || targetItem["Symbol"] || sym).trim().toUpperCase();
  const finalMonth = String(targetItem["ContractMonth(Week)"] || targetItem["ContractMonth"] || '').trim();

  let rawPrice = targetItem["Last"] || targetItem["SettlementPrice"] || targetItem["ClosingPrice"] || targetItem["Open"] || "0";
  const priceVal = parseFloat(String(rawPrice).replace(/,/g, '').trim()) || 0;

  const STOCK_FUT_DICT = {
    'CD': '台積電期', 'CDF': '台積電期', 'QV': '小台積期', 'QVF': '小台積期',
    'DH': '鴻海期貨', 'DHF': '鴻海期貨',
    'CC': '聯電期貨', 'CCF': '聯電期貨',
    'FR': '台達電期', 'FRF': '台達電期', 'RV': '小台達期', 'RVF': '小台達期',
    'IJ': '聯發科期', 'IJF': '聯發科期', 'QK': '小發科期', 'QKF': '小發科期',
    'QF': '廣達期貨', 'QFF': '廣達期貨',
    'LX': '國巨期貨', 'LXF': '國巨期貨', 'QE': '小國巨期', 'QEF': '小國巨期',
    'DQ': '大立光期', 'DQF': '大立光期', 'RG': '小立光期', 'RGF': '小立光期',
    'CG': '仁寶期貨', 'CGF': '仁寶期貨',
    'FQ': '光寶科期', 'FQF': '光寶科期',
    'FS': '金寶期貨', 'FSF': '金寶期貨',
    'FT': '華通期貨', 'FTF': '華通期貨',
    'CU': '中環期貨', 'CUF': '中環期貨',
    'OD': '為升期貨', 'ODF': '為升期貨',
    'CQ': '友達期貨', 'CQF': '友達期貨',
    'DP': '群創期貨', 'DPF': '群創期貨',
    'CZ': '長榮期貨', 'CZF': '長榮期貨', 'OY': '小長榮期', 'OYF': '小長榮期',
    'DB': '陽明期貨', 'DBF': '陽明期貨',
    'DK': '萬海期貨', 'DKF': '萬海期貨',
    'BA': '台泥期貨', 'BAF': '台泥期貨',
    'CA': '台塑期貨', 'CAF': '台塑期貨',
    'CB': '南亞期貨', 'CBF': '南亞期貨'
  };

  const chineseName = CORE_FUTURES_DICT[finalCode] || STOCK_FUT_DICT[finalCode] || `${finalCode}期貨`;

  return {
    code: finalCode,
    month: finalMonth,
    name: finalMonth ? `${chineseName} ${finalMonth}` : chineseName,
    price: priceVal
  };
}

/**
 * 智慧匹配單一期交所商品行情 (自訂公式專用)
 */
function getTaifexItem_(ticker) {
  if (!ticker) return null;
  const rawInput = String(ticker).trim().toUpperCase();

  const cache = CacheService.getScriptCache();
  const cacheKey = 'FUT_QUOTE_' + rawInput;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  const allData = fetchAllTaifexData_();
  if (!allData) return null;

  const result = findTaifexFromMemory_(rawInput, allData);
  if (result) {
    try {
      cache.put(cacheKey, JSON.stringify(result), 60);
    } catch(e) {}
  }
  return result;
}

/**
 * 前端自訂公式：抓取期貨名稱
 */
function GET_FUTURES_NAME(ticker) {
  if (!ticker) return "";
  const item = getTaifexItem_(ticker);
  return item ? item.name : "查無名稱";
}

/**
 * 前端自訂公式：抓取期貨價格
 */
function GET_FUTURES_PRICE(ticker) {
  if (!ticker) return 0;
  const item = getTaifexItem_(ticker);
  return item ? item.price : 0;
}