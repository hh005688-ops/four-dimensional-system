// @ts-nocheck
// ==========================================
// ⚡ 四維側邊欄服務與調度核心 (公式自動抓名稱/精簡版)
// ==========================================

function getInitData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var debtSheet = ss.getSheetByName("負債/資金");
  var pledges = [];
  var loans = [];
  var currentCat = "";

  if (debtSheet) {
    var lastRow = debtSheet.getLastRow();
    if (lastRow >= 4) {
      var debtData = debtSheet.getRange(4, 1, lastRow - 3, 7).getValues();

      for (var i = 0; i < debtData.length; i++) {
        var name = String(debtData[i][0] || "").trim();
        var pay = Number(debtData[i][3]) || 0;
        var row = i + 4; 
        if (!name) continue;

        if (name === "券商" || name === "信貸" || name === "房貸" || name === "其他") {
          currentCat = name;
          continue;
        }

        if (currentCat === "券商" || currentCat === "其他") {
      // 🚀 過濾掉包含「總計」、「餘額」、「金額」、「CellImage」或名稱為空的雜訊行
      if (name && name.indexOf("總計") === -1 && name.indexOf("餘額") === -1 && name.indexOf("CellImage") === -1 && name.indexOf("金額") === -1) {
        pledges.push({ name: currentCat + " - " + name, row: row });
      }
    }
        else if (currentCat === "信貸" || currentCat === "房貸") {
          loans.push({ name: "[" + currentCat + "] " + name, row: row, defaultPay: pay });
        }
      }
    }
  }

  var twSheet = ss.getSheetByName("台股區");
  var usSheet = ss.getSheetByName("美股區");
  var twStocks = [];
  var usStocks = [];

  // 🚀 台股區：動態讀取，但具備「智慧邊界探測」
  if (twSheet) {
    var twLast = Math.max(twSheet.getLastRow(), 2);
    var twData = twSheet.getRange(2, 1, twLast - 1, 5).getValues();
    for (var j = 0; j < twData.length; j++) {
      var rawCode = twData[j][0];
      
      // 過濾空值或 Date 物件
      if (!rawCode || rawCode instanceof Date || typeof rawCode === 'object') continue;
      
      var code = String(rawCode).trim();
      
      // 🛡️ 邊界探測：一旦遇到下方統計區的標題，直接 break 結束台股讀取迴圈！
      if (code.indexOf("加權") !== -1 || code.indexOf("今日漲跌") !== -1 || code.indexOf("日期") !== -1) {
        break; 
      }
      
      // 🛡️ 格式防呆：排除包含中文、CellImage、空白、或是純小數的公式結果
      if (code.indexOf("CellImage") !== -1 || /[\u4e00-\u9fa5]/.test(code) || code.indexOf(" ") !== -1) continue;
      if (!isNaN(code) && code.indexOf('.') !== -1) continue; 

      twStocks.push({ 
        code: code, 
        qty: Number(twData[j][4]) || 0, 
        row: j + 2, 
        type: 'TW' 
      });
    }
  }

  // 🚀 美股區：動態讀取與智慧邊界探測
  if (usSheet) {
    var usLast = Math.max(usSheet.getLastRow(), 2);
    var usData = usSheet.getRange(2, 1, usLast - 1, 5).getValues();
    for (var k = 0; k < usData.length; k++) {
      var rawCodeUs = usData[k][0];
      
      if (!rawCodeUs || rawCodeUs instanceof Date || typeof rawCodeUs === 'object') continue;
      
      var codeUs = String(rawCodeUs).trim();
      
      // 🛡️ 邊界探測：遇到匯率或日期，直接 break 結束美股讀取迴圈！
      if (codeUs.indexOf("匯率") !== -1 || codeUs.indexOf("日期") !== -1) {
        break;
      }
      
      // 🛡️ 格式防呆
      if (codeUs.indexOf("CellImage") !== -1 || /[\u4e00-\u9fa5]/.test(codeUs) || codeUs.indexOf(" ") !== -1) continue;
      if (!isNaN(codeUs) && codeUs.indexOf('.') !== -1) continue;

      usStocks.push({ 
        code: codeUs, 
        qty: Number(usData[k][4]) || 0, 
        row: k + 2, 
        type: 'US' 
      });
    }
  }

  return {
    pledges: pledges,
    loans: loans,
    stocks: twStocks.concat(usStocks)
  };
}

function executeDispatch(params) {
  if (!params) {
    return { success: false, message: "❌ 錯誤：未收到前端傳遞的參數" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var type = String(params.type || "").trim();
  var amount = Number(params.amount) || 0;
  
  var debtSheet = ss.getSheetByName("負債/資金");
  if (!debtSheet) {
    return { success: false, message: "找不到「負債/資金」分頁！" };
  }

  var cashCell = debtSheet.getRange("D21");
  var currentCash = Number(cashCell.getValue()) || 0;

  // 🚀 1. 現貨買賣處理 (完全獨立的 if 判斷，絕不崩潰)
  if (type === 'STOCK_BUY' || type === 'STOCK_SELL') {
    var sheetName = (params.market === 'TW') ? "台股區" : "美股區";
    var targetSheet = ss.getSheetByName(sheetName);
    
    if (!targetSheet) {
      return { success: false, message: "❌ 找不到市場分頁：「" + sheetName + "」" };
    }

    var targetRow = Number(params.row) || 0;
    
    if (params.isNewStock) {
      if (!params.code) {
        return { success: false, message: "❌ 新增標的代碼不得為空！" };
      }
      
      targetRow = 2;
      while (targetRow <= 8 && targetSheet.getRange(targetRow, 1).getValue() !== "") {
        targetRow++;
      }
      
      if (targetRow > 8) {
        return { success: false, message: "❌ 該市場清單已滿（上限 7 檔），無法再新增！" };
      }
      
      targetSheet.getRange(targetRow, 1).setValue(params.code); // A 欄寫入代碼
      targetSheet.getRange(targetRow, 5).setValue(0);          // E 欄初始化股數為 0
    }

    if (targetRow < 2) {
      return { success: false, message: "❌ 寫入行數異常 (Row: " + targetRow + ")" };
    }

    var qtyCell = targetSheet.getRange("E" + targetRow);
    var currentQty = Number(qtyCell.getValue()) || 0;
    var deltaQty = Number(params.qty) || 0;

    ss.toast("🎯 成功寫入 [" + sheetName + "] A" + targetRow + " / E" + targetRow, "自動化通知", 4);

    if (type === 'STOCK_BUY') {
      qtyCell.setValue(currentQty + deltaQty);
      if (amount > 0) cashCell.setValue(Math.max(0, currentCash - amount));
    } else {
      qtyCell.setValue(Math.max(0, currentQty - deltaQty));
      if (amount > 0) cashCell.setValue(currentCash + amount);
    }
    
    return { success: true, message: "調度成功完成！" };
  }

  // 🚀 2. 其他財務動作處理
  if (type === 'DIVIDEND') {
    var divCell = debtSheet.getRange("B23");
    divCell.setValue((Number(divCell.getValue()) || 0) + amount);
    cashCell.setValue(currentCash + amount);
  } 
  else if (type === 'INFLOW') {
    cashCell.setValue(currentCash + amount);
  } 
  else if (type === 'PLEDGE_BORROW') {
    var rowBorrow = Number(params.row);
    var pBorrowCell = debtSheet.getRange("B" + rowBorrow);
    pBorrowCell.setValue((Number(pBorrowCell.getValue()) || 0) + amount);
    cashCell.setValue(currentCash + amount);
  } 
  else if (type === 'PLEDGE_REPAY') {
    var rowRepay = Number(params.row);
    var pRepayCell = debtSheet.getRange("B" + rowRepay);
    pRepayCell.setValue(Math.max(0, (Number(pRepayCell.getValue()) || 0) - amount));
    cashCell.setValue(Math.max(0, currentCash - amount));
  } 
  else if (type === 'FUTURES_IN') {
    var futSheetIn = ss.getSheetByName("期貨區");
    if (futSheetIn) {
      var futCellIn = futSheetIn.getRange(params.market === 'TW' ? "A11" : "A25");
      futCellIn.setValue((Number(futCellIn.getValue()) || 0) + amount);
    }
    cashCell.setValue(Math.max(0, currentCash - amount));
  } 
  else if (type === 'FUTURES_OUT') {
    var futSheetOut = ss.getSheetByName("期貨區");
    if (futSheetOut) {
      var futCellOut = futSheetOut.getRange(params.market === 'TW' ? "A11" : "A25");
      futCellOut.setValue(Math.max(0, (Number(futCellOut.getValue()) || 0) - amount));
    }
    cashCell.setValue(currentCash + amount);
  } 
  else if (type === 'LOAN_PAY') {
    var rowLoan = Number(params.row);
    var periodCell = debtSheet.getRange("F" + rowLoan);
    periodCell.setValue((Number(periodCell.getValue()) || 0) + 1);
    cashCell.setValue(Math.max(0, currentCash - amount));
  }

  return { success: true, message: "調度成功完成！" };
}

// 🛡️ 嚴格過濾台美股：只允許標準代碼格式，徹底封殺日期、時間、物件字串
function isValidStockCode(code) {
  if (!code) return false;
  var text = String(code).trim();
  if (!text) return false;
  var upper = text.toUpperCase();
  if (upper.indexOf("MON") !== -1 || upper.indexOf("TUE") !== -1 || upper.indexOf("WED") !== -1 ||
      upper.indexOf("THU") !== -1 || upper.indexOf("FRI") !== -1 || upper.indexOf("SAT") !== -1 ||
      upper.indexOf("SUN") !== -1 || upper.indexOf("GMT") !== -1 || upper.indexOf("CELLIMAGE") !== -1 ||
      text.indexOf("加權") !== -1 || text.indexOf("日期") !== -1 || text.indexOf("今日漲跌") !== -1 ||
      text.indexOf("匯率") !== -1 || text.indexOf("總市值") !== -1 || /[\u4e00-\u9fa5]/.test(text)) {
    return false;
  }
  if (/^\d{7,}$/.test(text) || /^\d+\.\d+$/.test(text)) {
    return false;
  }
  if (text.length > 12) {
    return false;
  }
  return /^[A-Za-z0-9.-]+$/.test(text);
}

function isInventoryHoldingRow_(rawCode, qty) {
  if (rawCode instanceof Date) return false;
  if (typeof rawCode === "object" && rawCode) return false;
  if (!(Number(qty) > 0)) return false;
  return isValidStockCode(rawCode);
}

function isStockAreaBoundary_(rawCode) {
  if (rawCode instanceof Date) return true;
  var text = String(rawCode || "").trim();
  return text.indexOf("加權") !== -1 || text.indexOf("今日漲跌") !== -1 ||
         text.indexOf("日期") !== -1 || text.indexOf("匯率") !== -1;
}