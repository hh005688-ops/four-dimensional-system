/**
 * 🏛️ 四維系統：資金調度、現貨買賣、質押劃撥與期貨操作模組
 */
function getTargetSpreadsheet_(sheetIdentifier) {
  if (sheetIdentifier && typeof sheetIdentifier.getSheetByName === "function") {
    return sheetIdentifier;
  }
  var target = (typeof sheetIdentifier === "string" && sheetIdentifier.trim() !== "")
    ? sheetIdentifier.trim()
    : (typeof DEFAULT_SHEET_ID !== "undefined" ? DEFAULT_SHEET_ID : "");
  if (!target) {
    try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { return null; }
  }
  var match = target.match(/\/d\/([a-zA-Z0-9-_]+)/);
  var id = (match && match[1]) ? match[1] : target;
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    Logger.log("openById 失敗: " + err.message);
  }
  try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  return null;
}

function executeDispatchUniversal(sheetIdentifier, params) {
  try {
    var ss = getTargetSpreadsheet_(sheetIdentifier);
    if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

    // 🛡️ 雙保險：不管是 type 還是 actionType 都能自動抓到
    var type = params.type || params.actionType || "";
    var amount = Number(params.amount) || 0;
    var row = Number(params.row) || 0;

    var debtSheet = ss.getSheetByName("負債/資金");
    var mainSheet = ss.getSheetByName("總覽");
    if (!debtSheet) return { success: false, message: "❌ 找不到『負債/資金』分頁" };

    // 依據第二段架構：現金水庫以「負債/資金 D21」為主
    var cashCell = debtSheet.getRange("D21");
    var currentCash = Number(cashCell.getValue()) || 0;

    // 同步取得總覽 B5 現金水位（若存在）
    var mainCashCell = mainSheet ? mainSheet.getRange("B5") : null;
    var mainCurrentCash = mainCashCell ? (Number(mainCashCell.getValue()) || 0) : 0;

    var message = "";

    switch (type) {
      case 'DIVIDEND':
        // 1. 累加當月配息金額 (負債/資金 B23)
        var divCell = debtSheet.getRange("B23");
        var currentDiv = Number(divCell.getValue()) || 0;
        divCell.setValue(currentDiv + amount);

        // 2. 增加現金水庫餘額 (負債/資金 D21)
        cashCell.setValue(currentCash + amount);
        
        // 3. 同步更新總覽分頁的現金水位 (B5)
        if (mainCashCell) {
          mainCashCell.setValue(mainCurrentCash + amount);
        }
        message = "✅ 股息入帳成功！水庫 +" + amount.toLocaleString() + " 元";
        break;

      case 'INFLOW':
        cashCell.setValue(currentCash + amount);
        if (mainCashCell) {
          mainCashCell.setValue(mainCurrentCash + amount);
        }
        message = "✅ 外部注水成功！水庫 +" + amount.toLocaleString() + " 元";
        break;

      case 'LOAN_PAY':
        if (row < 9 || row > 17) return { success: false, message: "❌ 貸款項目列號無效" };
        
        // 1. 扣除水庫金額 (D21 與 總覽 B5)
        cashCell.setValue(currentCash - amount);
        if (mainCashCell) {
          mainCashCell.setValue(mainCurrentCash - amount);
        }
        
        // 2. 已繳期數 F 欄 (+1)
        var curPeriod = Number(debtSheet.getRange(row, 6).getValue()) || 0;
        debtSheet.getRange(row, 6).setValue(curPeriod + 1);
        
        // 3. 賸餘本金 G 欄（同步扣除本次還款金額）
        var balanceCell = debtSheet.getRange(row, 7);
        var currentBalance = Number(balanceCell.getValue()) || 0;
        balanceCell.setValue(Math.max(0, currentBalance - amount));

        message = "✅ 貸款扣繳成功！水庫扣除 $" + amount.toLocaleString() + "，期數推進至第 " + (curPeriod + 1) + " 期";
        break;

      case 'PLEDGE_BORROW':
        if (row < 5 || row > 7) return { success: false, message: "❌ 質押機構列號無效" };
        
        cashCell.setValue(currentCash + amount);
        if (mainCashCell) {
          mainCashCell.setValue(mainCurrentCash + amount);
        }
        
        var curBorrow = Number(debtSheet.getRange(row, 2).getValue()) || 0;
        debtSheet.getRange(row, 2).setValue(curBorrow + amount);
        message = "⚡ 質押撥款成功！水庫 +" + amount.toLocaleString() + " 元";
        break;

      case 'PLEDGE_REPAY':
        if (row < 5 || row > 7) return { success: false, message: "❌ 質押機構列號無效" };
        
        cashCell.setValue(currentCash - amount);
        if (mainCashCell) {
          mainCashCell.setValue(mainCurrentCash - amount);
        }
        
        var curRepay = Number(debtSheet.getRange(row, 2).getValue()) || 0;
        debtSheet.getRange(row, 2).setValue(Math.max(0, curRepay - amount));
        message = "🛡️ 質押還本成功！水庫扣除 $" + amount.toLocaleString() + " 元";
        break;

      // 🎯 批次儲存目標佔比 (寫入 K 欄，即第 11 欄)
      case 'BATCH_SAVE_TARGET_RATIOS':
        var updates = params.updates || [];
        var twS = ss.getSheetByName("台股區");
        var usS = ss.getSheetByName("美股區");
        
        updates.forEach(function(u) {
          var r = Number(u.row);
          var val = Number(u.targetRatio) || 0;
          var formattedVal = val > 1 ? val / 100 : val;
          
          if (twS && r >= 2 && r <= twS.getLastRow()) {
            twS.getRange(r, 11).setValue(formattedVal);
          }
          if (usS && r >= 2 && r <= usS.getLastRow()) {
            usS.getRange(r, 11).setValue(formattedVal);
          }
        });
        message = "✅ 目標佔比已成功批次儲存至 K 欄！";
        break;

      // 🎯 建立全新標的並初始化
      case 'ADD_STOCK':
        var market = params.market;
        var code = String(params.code).trim().toUpperCase();
        var qty = Number(params.qty) || 0;
        var cashTotal = Number(params.cashTotal) || 0;
        var targetRatio = Number(params.targetRatio) || 0;

        var targetSheet = ss.getSheetByName(market === "TW" ? "台股區" : "美股區");
        if (!targetSheet) return { success: false, message: "❌ 找不到對應市場分頁" };

        var lastR = Math.max(targetSheet.getLastRow(), 7);
        var aValues = targetSheet.getRange("A2:A" + lastR).getValues();
        var targetRow = -1;
        for (var idx = 0; idx < aValues.length; idx++) {
          if (!String(aValues[idx][0]).trim()) {
            targetRow = idx + 2;
            break;
          }
        }
        if (targetRow === -1) targetRow = lastR + 1;

        targetSheet.getRange(targetRow, 1).setValue(code);
        if (qty > 0) targetSheet.getRange(targetRow, 5).setValue(qty);
        if (targetRatio > 0) {
          targetSheet.getRange(targetRow, 11).setValue(targetRatio > 1 ? targetRatio / 100 : targetRatio);
        }

        if (cashTotal > 0) {
          cashCell.setValue(Math.max(0, currentCash - cashTotal));
          if (mainCashCell) {
            mainCashCell.setValue(Math.max(0, mainCurrentCash - cashTotal));
          }
        }

        message = "🎉 成功建立新標的 [" + market + "] " + code + "！水庫已連帶扣除 $" + cashTotal.toLocaleString() + " 元";
        break;

      default:
        return { success: false, message: "❌ 未知的金流類型: " + type };
    }

    SpreadsheetApp.flush();
    return { success: true, message: message };
  } catch (err) {
    return { success: false, message: "❌ 執行回寫失敗: " + err.message };
  }
}

// 2. 現貨買賣回寫
function executeStockTradeUniversal(sheetIdentifier, params) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    var market = params.market;
    var row = Number(params.row);
    var action = params.action;
    var qty = Number(params.qty) || 0;
    var cashDiff = Number(params.cashDiff) || 0;

    if (qty <= 0) return { success: false, message: "❌ 變動股數必須大於 0" };

    var targetSheet = ss.getSheetByName(market === "TW" ? "台股區" : "美股區");
    var mainSheet = ss.getSheetByName("總覽");
    if (!targetSheet) return { success: false, message: "❌ 找不到對應市場分頁" };

    var currentQty = Number(targetSheet.getRange(row, 5).getValue()) || 0;
    var newQty = (action === "BUY") ? (currentQty + qty) : (currentQty - qty);
    if (newQty < 0) return { success: false, message: "❌ 庫存不足，無法賣出！" };

    targetSheet.getRange(row, 5).setValue(newQty);

    if (cashDiff > 0 && mainSheet) {
      var currentCash = Number(mainSheet.getRange("B5").getValue()) || 0;
      var newCash = (action === "BUY") ? (currentCash - cashDiff) : (currentCash + cashDiff);
      mainSheet.getRange("B5").setValue(newCash);
    }

    var codeName = targetSheet.getRange(row, 1).getValue();
    return { 
      success: true, 
      message: "✅ 現貨回寫完成！" + codeName + " " + (action === "BUY" ? "買進" : "賣出") + " " + qty + " 股 (最新庫存: " + newQty + " 股)" 
    };
  } catch (err) {
    return { success: false, message: "❌ 現貨回寫失敗: " + err.message };
  }
}

// 3. 質押劃撥回寫
function executePledgeTransferUniversal(sheetIdentifier, params) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    var row = Number(params.row);
    var action = params.action;
    var qty = Number(params.qty) || 0;

    if (qty <= 0) return { success: false, message: "❌ 劃撥股數必須大於 0" };

    var twSheet = ss.getSheetByName("台股區");
    if (!twSheet) return { success: false, message: "❌ 找不到『台股區』分頁" };

    var totalQty = Number(twSheet.getRange(row, 5).getValue()) || 0;
    var pledgedQty = Number(twSheet.getRange(row, 9).getValue()) || 0;
    var unpledgedQty = Math.max(0, totalQty - pledgedQty);

    if (action === "LOCK") {
      if (qty > unpledgedQty) return { success: false, message: "❌ 未質押股數不足！(最多可撥: " + unpledgedQty + " 股)" };
      twSheet.getRange(row, 9).setValue(pledgedQty + qty);
      return { success: true, message: "🔒 質押劃撥成功！鎖定 " + qty + " 股，質押總數: " + (pledgedQty + qty) + " 股" };
    } else {
      if (qty > pledgedQty) return { success: false, message: "❌ 質押股數不足！(最多可解: " + pledgedQty + " 股)" };
      twSheet.getRange(row, 9).setValue(pledgedQty - qty);
      return { success: true, message: "🔓 解除質押成功！釋出 " + qty + " 股，剩餘質押: " + (pledgedQty - qty) + " 股" };
    }
  } catch (err) {
    return { success: false, message: "❌ 質押劃撥失敗: " + err.message };
  }
}

// 4. 期貨保證金出入金回寫
function executeFuturesTransferUniversal(sheetIdentifier, params) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    var market = params.market;
    var action = params.action;
    var amount = Number(params.amount) || 0;

    if (amount <= 0) return { success: false, message: "❌ 金額必須大於 0" };

    var fSheet = ss.getSheetByName("期貨區");
    var mainSheet = ss.getSheetByName("總覽");
    if (!fSheet || !mainSheet) return { success: false, message: "❌ 找不到『期貨區』或『總覽』分頁" };

    var cash = Number(mainSheet.getRange("B5").getValue()) || 0;
    var targetCell = (market === "TW") ? "B11" : "B25";
    var curMargin = Number(fSheet.getRange(targetCell).getValue()) || 0;

    if (action === "IN") {
      mainSheet.getRange("B5").setValue(cash - amount);
      fSheet.getRange(targetCell).setValue(curMargin + amount);
      return { success: true, message: "📥 期貨入金成功！水庫扣除 $" + amount.toLocaleString() + "，保證金本金為 $" + (curMargin + amount).toLocaleString() };
    } else {
      if (amount > curMargin) return { success: false, message: "❌ 出金金額大於目前保證金總入金！" };
      mainSheet.getRange("B5").setValue(cash + amount);
      fSheet.getRange(targetCell).setValue(curMargin - amount);
      return { success: true, message: "📤 期貨出金成功！水庫增加 $" + amount.toLocaleString() + "，保證金本金為 $" + (curMargin - amount).toLocaleString() };
    }
  } catch (err) {
    return { success: false, message: "❌ 期貨回寫失敗: " + err.message };
  }
}

// 🆕 【新增】獲取期貨現有合約代碼清單（供前端下拉選單動態載入）
function getExistingContractsUniversal(sheetIdentifier, marketType) {
  try {
    var ss = getTargetSpreadsheet_(sheetIdentifier);
    if (!ss) return [];

    var sheet = ss.getSheetByName("期貨操作") || ss.getSheetByName("期貨區");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var contracts = [];

    // 國內期貨從第 2 行起 (index 1)，海外期貨從第 16 行起 (index 15)
    var startRow = (marketType === 'domestic') ? 1 : 15;
    var endRow = (marketType === 'domestic') ? 14 : data.length;

    for (var i = startRow; i < endRow && i < data.length; i++) {
      var contractCode = data[i][1]; // 假設代碼在 B 欄 (index 1)
      if (contractCode && !contracts.includes(String(contractCode))) {
        contracts.push(String(contractCode));
      }
    }
    return contracts;
  } catch (error) {
    console.error("getExistingContractsUniversal 錯誤: " + error.message);
    return [];
  }
}

// 🆕 【新增】執行期貨部位建倉 / 加碼 / 平倉回寫
function executeFuturesPositionUniversal(sheetIdentifier, params) {
  try {
    var ss = getTargetSpreadsheet_(sheetIdentifier);
    if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

    var marketType = params.marketType; // 'domestic' 或 'overseas'
    var contract = String(params.contract || "").trim();
    var direction = params.colD;        // 多 / 空
    var qty = Number(params.colE) || 0; // 單位口數
    var avgPrice = Number(params.colF) || 0; // 本期買入均價
    var status = params.colI;           // 狀態 (加碼/平倉)
    var closePrice = params.colJ !== "" ? Number(params.colJ) : ""; // 平倉價

    if (!contract) return { success: false, message: "❌ 合約代碼不得為空" };
    if (qty <= 0) return { success: false, message: "❌ 單位口數必須大於 0" };
    if (avgPrice <= 0) return { success: false, message: "❌ 買入均價必須大於 0" };

    var sheet = ss.getSheetByName("期貨操作") || ss.getSheetByName("期貨區");
    if (!sheet) return { success: false, message: "❌ 找不到期貨相關分頁" };

    // 組合要寫入的一列資料
    var rowData = [
      new Date(),       // A欄: 時間戳記
      contract,         // B欄: 合約代碼
      marketType,       // C欄: 市場分區
      direction,        // D欄: 多/空
      qty,              // E欄: 單位口數
      avgPrice,         // F欄: 本期買入均價
      status,           // I欄: 狀態
      closePrice        // J欄: 平倉價
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { success: true, message: "🚀 期貨部位回寫成功！[" + contract + "] " + direction + " " + qty + " 口" };
  } catch (error) {
    return { success: false, message: "❌ 期貨部位回寫失敗: " + error.message };
  }
}

// 5. 新增全新股票標的
function addNewStockTargetUniversal(sheetIdentifier, params) {
  var ss = getTargetSpreadsheet_(sheetIdentifier);
  if (!ss) return { success: false, message: "❌ 無法開啟試算表" };

  try {
    var market = params.market;
    var code = String(params.code).trim().toUpperCase();
    var name = String(params.name).trim();
    var qty = Number(params.qty) || 0;
    var targetRatio = Number(params.targetRatio) || 0;
    var leverage = Number(params.leverage) || 1;

    if (!code || !name) return { success: false, message: "❌ 股票代碼與名稱不得為空！" };

    var targetSheet = ss.getSheetByName(market === "TW" ? "台股區" : "美股區");
    if (!targetSheet) return { success: false, message: "❌ 找不到對應市場分頁" };

    var data = targetSheet.getRange("A2:A7").getValues();
    var targetRow = -1;

    for (var i = 0; i < data.length; i++) {
      var curCode = String(data[i][0]).trim().toUpperCase();
      if (curCode === code) {
        return { success: false, message: "⚠️ 該標的 (" + code + ") 已存在於持倉中！" };
      }
      if (curCode === "" && targetRow === -1) {
        targetRow = 2 + i;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: "❌ 持倉插槽已滿（上限 6 檔），請先清理閒置標的！" };
    }

    targetSheet.getRange(targetRow, 1).setValue(code);
    targetSheet.getRange(targetRow, 2).setValue(name);
    targetSheet.getRange(targetRow, 5).setValue(qty);
    
    if (market === "TW") {
      targetSheet.getRange(targetRow, 9).setValue(0);
      targetSheet.getRange(targetRow, 11).setValue(targetRatio > 1 ? targetRatio / 100 : targetRatio);
      targetSheet.getRange(targetRow, 14).setValue(leverage);
    } else {
      targetSheet.getRange(targetRow, 10).setValue(targetRatio > 1 ? targetRatio / 100 : targetRatio);
      targetSheet.getRange(targetRow, 13).setValue(leverage);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: "🎉 成功新增標的 [" + market + "] " + code + " " + name + "！(配置於第 " + targetRow + " 列)"
    };
  } catch (err) {
    return { success: false, message: "❌ 新增標的失敗: " + err.message };
  }
}

/**
 * 🏛️ 四維系統：自動獲取股票繁體中文名稱 API (Yahoo Search 官方 JSON 解析版)
 */
function fetchStockInfoUniversal(market, code) {
  var cleanCode = String(code).trim().toUpperCase();
  if (!cleanCode) return { success: false, message: "代碼為空" };

  try {
    if (market === "TW") {
      var queryCode = cleanCode + ".TW";
      var searchUrl = "https://query2.finance.yahoo.com/v1/finance/search?q=" + encodeURIComponent(cleanCode) + "&lang=zh-Hant-TW&region=TW&quotesCount=6";
      var options = {
        "headers": {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        "muteHttpExceptions": true
      };

      try {
        var res = UrlFetchApp.fetch(searchUrl, options);
        if (res.getResponseCode() === 200) {
          var json = JSON.parse(res.getContentText());
          if (json.quotes && json.quotes.length > 0) {
            for (var i = 0; i < json.quotes.length; i++) {
              var q = json.quotes[i];
              var sym = String(q.symbol || "").toUpperCase();
              if (sym === cleanCode + ".TW" || sym === cleanCode + ".TWO" || sym === cleanCode) {
                var twName = q.longname || q.shortname || "";
                twName = twName.replace(/\.TW/i, "").replace(/走勢圖/g, "").trim();
                if (twName) {
                  return { success: true, name: twName, price: 0 };
                }
              }
            }
            var first = json.quotes[0];
            var fName = first.longname || first.shortname || "";
            fName = fName.replace(/\.TW/i, "").replace(/走勢圖/g, "").trim();
            if (fName) {
              return { success: true, name: fName, price: 0 };
            }
          }
        }
      } catch(e) {}

      try {
        var twUrl = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_" + cleanCode + ".tw|otc_" + cleanCode + ".tw&json=1&delay=0";
        var twRes = UrlFetchApp.fetch(twUrl, options);
        var twJson = JSON.parse(twRes.getContentText());
        if (twJson.msgArray && twJson.msgArray.length > 0) {
          var info = twJson.msgArray[0];
          var name = String(info.n || info.nf).trim();
          if (name && !name.match(/^[A-Za-z\s]+$/)) {
            return { success: true, name: name, price: Number(info.z) || 0 };
          }
        }
      } catch(e) {}

      var twDict = {
        "0050": "元大台灣50",
        "0056": "元大高股息",
        "00631L": "元大台灣50正2",
        "00675L": "富邦臺灣加權正2",
        "00680L": "元大美債20正2",
        "00687B": "國泰20年美債",
        "00878": "國泰永續高股息",
        "00919": "群益台灣精選高息",
        "00926": "凱基全球菁英55",
        "00929": "復華台灣科技優息",
        "00935": "野村臺灣新科技50",
        "00963": "中信全球高股息",
        "00965": "元大航太防衛科技",
        "2330": "台積電",
        "2317": "鴻海",
        "2454": "聯發科"
      };
      if (twDict[cleanCode]) {
        return { success: true, name: twDict[cleanCode], price: 0 };
      }

    } else {
      var usUrl = "https://query2.finance.yahoo.com/v1/finance/search?q=" + encodeURIComponent(cleanCode) + "&quotesCount=1";
      var usRes = UrlFetchApp.fetch(usUrl, { muteHttpExceptions: true });
      if (usRes.getResponseCode() === 200) {
        var usJson = JSON.parse(usRes.getContentText());
        if (usJson.quotes && usJson.quotes.length > 0) {
          var u = usJson.quotes[0];
          return { success: true, name: u.shortname || u.longname || u.symbol, price: 0 };
        }
      }
    }

    return { success: false, message: "查無標的繁中名稱" };
  } catch (err) {
    return { success: false, message: "查詢失敗: " + err.message };
  }
}