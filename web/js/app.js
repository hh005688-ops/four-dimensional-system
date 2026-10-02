var nestPieChartInstance = null;
  var globalData = null;
  var DEFAULT_SHEET_ID = "";
  var isNewStockMode = false;
  var ALLOC_TW_COLORS = ['#10b981', '#d4af37', '#22d3ee', '#84cc16', '#f59e0b', '#34d399'];
  var ALLOC_US_COLORS = ['#6366f1', '#f97316', '#ec4899', '#38bdf8', '#a855f7', '#fb7185'];
  var ALLOC_CASH_COLOR = '#fbbf24';
  var ALLOC_TW_INNER = '#059669';
  var ALLOC_US_INNER = '#4f46e5';

  function initTheme() {
    var savedTheme = localStorage.getItem('4D_THEME') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    document.getElementById('themeToggleBtn').innerText = (savedTheme === 'light') ? '🌙' : '☀️';
  }

  function toggleTheme() {
    var current = document.documentElement.getAttribute('data-theme');
    var next = (current === 'light') ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('4D_THEME', next);
    document.getElementById('themeToggleBtn').innerText = (next === 'light') ? '🌙' : '☀️';
    if (globalData) renderNestedAllocationChart(globalData);
  }

  function getCurrentSheetId() {
    var inputVal = document.getElementById('sheetIdInput') ? document.getElementById('sheetIdInput').value.trim() : '';
    var localVal = localStorage.getItem('4D_SHEET_ID');
    return inputVal || localVal || DEFAULT_SHEET_ID;
  }

  function getGasApiUrl() {
    var inputVal = document.getElementById('gasApiUrlInput') ? document.getElementById('gasApiUrlInput').value.trim() : '';
    return inputVal || localStorage.getItem('4D_GAS_API') || '';
  }

  function callGasApi(action, extra) {
    extra = extra || {};
    var url = getGasApiUrl();
    if (!url) {
      return Promise.reject(new Error('尚未設定 GAS Web App 網址（設定頁貼上 /exec 連結）'));
    }
    var body = {
      action: action,
      sheetId: extra.sheetId !== undefined ? extra.sheetId : getCurrentSheetId(),
      actionType: extra.actionType,
      params: extra.params || {},
      market: extra.market,
      code: extra.code
    };
    return fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function(res) {
      return res.text();
    }).then(function(text) {
      try {
        return JSON.parse(text);
      } catch (err) {
        throw new Error('API 回傳不是 JSON，請確認已部署 Web App 且權限為「任何人」');
      }
    });
  }

  function runDispatchAction(actionType, params, loader, onSuccess, failPrefix) {
    callGasApi('dispatchUniversalAction', { actionType: actionType, params: params })
      .then(function(res) {
        if (loader) loader.style.display = 'none';
        handleResponseGuard(res, onSuccess);
      })
      .catch(function(err) {
        if (loader) loader.style.display = 'none';
        alert((failPrefix || '❌ 執行失敗: ') + err.message);
      });
  }

  window.onload = function() {
    initTheme();
    if (document.getElementById('gasApiUrlInput')) {
      document.getElementById('gasApiUrlInput').value = getGasApiUrl();
    }
    var savedId = getCurrentSheetId();
    if (document.getElementById('sheetIdInput')) {
      document.getElementById('sheetIdInput').value = savedId;
    }
    if (getGasApiUrl() && savedId) {
      connectSheet();
    }
  };

  /**
   * 🛡️ 核心防護攔截器：檢查後端回傳結果是否為授權失敗
   */
  function handleResponseGuard(res, successCallback) {
    if (!res) {
      alert("❌ 伺服器未回應或發生未知錯誤");
      return false;
    }
    // 🚨 關鍵攔截：若後端回傳 success 為 false，直接彈出警告提示！
    if (res.success === false) {
      alert(res.message || "🔒 授權失敗：此試算表尚未啟用或金鑰不符，請先回到 Google 試算表進行解鎖！");
      return false;
    }
    if (typeof successCallback === 'function') {
      successCallback(res);
    }
    return true;
  }

  function connectSheet() {
    var apiUrl = getGasApiUrl();
    if (apiUrl && document.getElementById('gasApiUrlInput')) {
      localStorage.setItem('4D_GAS_API', apiUrl);
    }
    var val = getCurrentSheetId();
    if (!getGasApiUrl()) return alert("請先在設定頁貼上 GAS Web App 網址（/exec）！");
    if (!val) return alert("請輸入試算表網址或 ID！");
    
    document.getElementById('connStatus').innerText = "⏳ 正在讀取帳本架構...";
    document.getElementById('topStatus').innerText = "連線中...";

    callGasApi('getInitDataUniversal', { sheetId: val })
      .then(function(res) {
        if (!handleResponseGuard(res, function(data) {
          localStorage.setItem('4D_SHEET_ID', val);
          globalData = data;
          globalData.stocks = (globalData.stocks || []).filter(isInventoryStock_);
          
          document.getElementById('connStatus').innerText = "✅ 已連線：" + data.sheetName;
          document.getElementById('topStatus').innerText = "🟢 " + data.sheetName + " 已連線";
          
          document.getElementById('dispNetWorth').innerText = "$" + Math.round(Number(data.netWorth) || 0).toLocaleString();
          document.getElementById('dispTotalAssets').innerText = "$" + Math.round(Number(data.totalAssets) || 0).toLocaleString();
          document.getElementById('dispTotalLiab').innerText = "$" + Math.round(Number(data.totalLiabilities) || 0).toLocaleString();
          document.getElementById('dispCash').innerText = "$" + Math.round(Number(data.currentCash) || 0).toLocaleString();
          document.getElementById('dispMaintenance').innerText = Number(data.pledgeMaintenance).toFixed(2) + "%";
          document.getElementById('dispLeverage').innerText = Number(data.realLeverage || data.coverageRatio).toFixed(2) + " x";
          document.getElementById('dispAssetBeta').innerText = Number(data.assetBeta).toFixed(2);

          renderAssetAllocation(globalData);
          renderNestedAllocationChart(globalData);
          renderStockList(globalData.stocks);
          renderFutures(data.futures);
          renderDebtList(data.pledges, data.loans);
          initDispatchSelects(globalData);
        })) {
          document.getElementById('connStatus').innerText = "🔒 授權失敗，未解鎖";
          document.getElementById('topStatus').innerText = "未授權";
        }
      })
      .catch(function(err) {
        document.getElementById('connStatus').innerText = "❌ 連線異常: " + err.message;
        document.getElementById('topStatus').innerText = "連線失敗";
      });
  }

  function renderAssetAllocation(res) {
    var box = document.getElementById('assetAllocationContainer');
    var badge = document.getElementById('assetRatioCount');
    var stocks = ((res && res.stocks) ? res.stocks : []).filter(isInventoryStock_);
    var twStocks = stocks.filter(function(s) { return s.market === 'TW'; });
    var usStocks = stocks.filter(function(s) { return s.market === 'US'; });

    if (badge) {
      badge.innerText = '台股 ' + twStocks.length + ' 檔｜美股 ' + usStocks.length + ' 檔';
    }

    if (!res || stocks.length === 0) {
      box.innerHTML = '<div style="text-align:center; color:var(--subtext); padding: 10px;">無持倉數據</div>';
      return;
    }

    var twColors = ALLOC_TW_COLORS;
    var usColors = ALLOC_US_COLORS;

    var html = '';
    html += buildMarketAllocationSection_('🇹🇼 台股配置', twStocks, twColors, '台股');
    html += buildMarketAllocationSection_('🇺🇸 美股配置', usStocks, usColors, '美股');

    var cash = Number(res.currentCash) || 0;
    var totalAssets = Number(res.totalAssets) || 0;
    var cashRatio = totalAssets > 0 ? ((cash / totalAssets) * 100).toFixed(1) : '0.0';
    html += '<div style="margin-top: 6px; padding-top: 10px; border-top: 1px solid var(--border);">' +
              '<div style="display:flex; justify-content:space-between; font-size:12px; color:var(--subtext); margin-bottom:4px;">' +
                '<span>💰 現金水庫（佔總資產）</span>' +
                '<span style="color:' + ALLOC_CASH_COLOR + '; font-weight:bold;">$' + Math.round(cash).toLocaleString() + ' (' + cashRatio + '%)</span>' +
              '</div>' +
              '<div class="alloc-track">' +
                '<div class="alloc-fill" style="background:' + ALLOC_CASH_COLOR + '; width:' + Math.min(100, Math.max(0, Number(cashRatio))) + '%;"></div>' +
              '</div>' +
            '</div>';

    box.innerHTML = html;
  }

  function buildNestedAllocationData_(res) {
    var stocks = ((res && res.stocks) ? res.stocks : []).filter(isInventoryStock_);
    var twStocks = stocks.filter(function(s) { return s.market === 'TW'; });
    var usStocks = stocks.filter(function(s) { return s.market === 'US'; });
    var cash = Math.max(0, Number(res && res.currentCash) || 0);
    var twTotal = sumMarketValue_(twStocks);
    var usTotal = sumMarketValue_(usStocks);
    var pieTotal = twTotal + usTotal + cash;

    var inner = [];
    if (twTotal > 0) inner.push({ name: '🇹🇼 台股', value: twTotal, itemStyle: { color: ALLOC_TW_INNER } });
    if (usTotal > 0) inner.push({ name: '🇺🇸 美股', value: usTotal, itemStyle: { color: ALLOC_US_INNER } });
    if (cash > 0) inner.push({ name: '💰 現金', value: cash, itemStyle: { color: ALLOC_CASH_COLOR } });

    var outer = [];
    twStocks.forEach(function(s, idx) {
      var val = Number(s.marketVal) || 0;
      if (val <= 0) return;
      outer.push({
        name: s.code,
        value: val,
        itemStyle: { color: ALLOC_TW_COLORS[idx % ALLOC_TW_COLORS.length] }
      });
    });
    usStocks.forEach(function(s, idx) {
      var val = Number(s.marketVal) || 0;
      if (val <= 0) return;
      outer.push({
        name: s.code,
        value: val,
        itemStyle: { color: ALLOC_US_COLORS[idx % ALLOC_US_COLORS.length] }
      });
    });
    if (cash > 0) {
      outer.push({ name: '現金', value: cash, itemStyle: { color: ALLOC_CASH_COLOR } });
    }

    return { inner: inner, outer: outer, pieTotal: pieTotal };
  }

  function renderNestedAllocationChart(res) {
    var chartDom = document.getElementById('nestedAllocChart');
    var badge = document.getElementById('nestedAllocBadge');
    if (!chartDom) return;

    if (typeof echarts === 'undefined') {
      chartDom.innerText = '⚠️ 圖表引擎尚未載入';
      return;
    }

    var pack = buildNestedAllocationData_(res);
    if (badge) {
      badge.innerText = pack.pieTotal > 0
        ? ('配置總額 $' + Math.round(pack.pieTotal).toLocaleString() + '＝100%')
        : '台股＋美股＋現金 = 100%';
    }

    if (!pack.pieTotal || (pack.inner.length === 0 && pack.outer.length === 0)) {
      if (nestPieChartInstance) {
        nestPieChartInstance.dispose();
        nestPieChartInstance = null;
      }
      chartDom.innerHTML = '<div style="text-align:center; color:var(--subtext); padding: 40px 10px;">尚無台股／美股／現金可繪製</div>';
      return;
    }

    var isLight = document.documentElement.getAttribute('data-theme') === 'light';
    var textColor = isLight ? '#0f172a' : '#f3f4f6';
    var subColor = isLight ? '#64748b' : '#9ca3af';
    var borderColor = isLight ? '#ffffff' : '#121824';

    if (nestPieChartInstance) nestPieChartInstance.dispose();
    nestPieChartInstance = echarts.init(chartDom, isLight ? undefined : 'dark');

    nestPieChartInstance.setOption({
      backgroundColor: 'transparent',
      tooltip: {
        trigger: 'item',
        formatter: function(p) {
          var pct = pack.pieTotal > 0 ? ((p.value / pack.pieTotal) * 100).toFixed(1) : '0.0';
          return p.seriesName + '<br/>' + p.name + '　$' + Math.round(p.value).toLocaleString() + '（' + pct + '%）';
        }
      },
      series: [
        {
          name: '大類佔比',
          type: 'pie',
          radius: ['18%', '42%'],
          center: ['50%', '50%'],
          startAngle: 90,
          avoidLabelOverlap: true,
          itemStyle: { borderColor: borderColor, borderWidth: 2 },
          label: {
            color: textColor,
            fontSize: 11,
            fontWeight: 'bold',
            formatter: function(p) {
              var pct = pack.pieTotal > 0 ? ((p.value / pack.pieTotal) * 100).toFixed(1) : '0.0';
              return p.name + '\n' + pct + '%';
            }
          },
          labelLine: { length: 8, length2: 6 },
          data: pack.inner
        },
        {
          name: '標的佔比',
          type: 'pie',
          radius: ['52%', '72%'],
          center: ['50%', '50%'],
          startAngle: 90,
          avoidLabelOverlap: true,
          itemStyle: { borderColor: borderColor, borderWidth: 2 },
          label: {
            color: subColor,
            fontSize: 10,
            formatter: function(p) {
              var pct = pack.pieTotal > 0 ? ((p.value / pack.pieTotal) * 100).toFixed(1) : '0.0';
              if (Number(pct) < 3) return '';
              return p.name + ' ' + pct + '%';
            }
          },
          labelLine: { length: 10, length2: 8 },
          data: pack.outer
        }
      ]
    });
  }

  function resizeNestedAllocationChart() {
    if (nestPieChartInstance) nestPieChartInstance.resize();
  }

  if (typeof window !== 'undefined') {
    window.addEventListener('resize', function() {
      resizeNestedAllocationChart();
    });
  }

  function isInventoryStock_(s) {
    if (!s) return false;
    var qty = Number(s.totalQty) || 0;
    if (qty <= 0) return false;
    var code = String(s.code || '').trim();
    if (!code) return false;
    var upper = code.toUpperCase();
    if (upper.indexOf('CELLIMAGE') !== -1 || upper.indexOf('GMT') !== -1) return false;
    if (code.indexOf('加權') !== -1 || code.indexOf('日期') !== -1 || code.indexOf('今日漲跌') !== -1 || code.indexOf('匯率') !== -1) return false;
    if (/[\u4e00-\u9fa5]/.test(code)) return false;
    if (/^\d+\.\d+$/.test(code) || /^\d{7,}$/.test(code)) return false;
    if (code.length > 12) return false;
    return /^[A-Za-z0-9.-]+$/.test(code);
  }

  function sumMarketValue_(list) {
    return (list || []).reduce(function(sum, s) {
      return sum + (Number(s.marketVal) || 0);
    }, 0);
  }

  function buildMarketAllocationSection_(title, list, colors, emptyLabel) {
    var total = sumMarketValue_(list);
    var header = '<div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:8px;">' +
                   '<span style="font-size:13px; font-weight:700; color:var(--text);">' + title + '</span>' +
                   '<span style="font-size:11px; color:var(--subtext);">市值 $' + Math.round(total).toLocaleString() + '｜權重獨立計 100%</span>' +
                 '</div>';

    if (!list || list.length === 0) {
      return '<div style="margin-bottom: 16px; padding: 10px 12px; background:var(--card-sub); border:1px solid var(--border); border-radius:10px;">' +
               header +
               '<div style="text-align:center; color:var(--subtext); padding: 6px; font-size:12px;">無' + emptyLabel + '持倉</div>' +
             '</div>';
    }

    var rows = list.map(function(s, idx) {
      var mVal = Number(s.marketVal) || 0;
      var ratioNum = total > 0 ? (mVal / total) * 100 : 0;
      var ratio = ratioNum.toFixed(1);
      var color = colors[idx % colors.length];
      var target = Number(s.targetRatio);
      var hasTarget = !isNaN(target) && target > 0;
      var targetHtml = hasTarget
        ? '<span style="font-size:11px; color:var(--subtext); margin-left:6px;">目標 ' + target.toFixed(1) + '%</span>'
        : '';

      return '<div style="margin-bottom: 10px;">' +
               '<div style="display:flex; justify-content:space-between; font-size:13px; margin-bottom:4px; gap:8px;">' +
                 '<span><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + color + ';margin-right:6px;vertical-align:middle;"></span>' + s.code + ' ' + s.name + targetHtml + '</span>' +
                 '<span style="color:' + color + '; font-weight:bold; white-space:nowrap;">$' + Math.round(mVal).toLocaleString() + ' (' + ratio + '%)</span>' +
               '</div>' +
               '<div class="alloc-track">' +
                 '<div class="alloc-fill" style="background:' + color + '; width:' + Math.min(100, Math.max(2, ratioNum)) + '%;"></div>' +
               '</div>' +
             '</div>';
    }).join('');

    return '<div style="margin-bottom: 16px; padding: 10px 12px; background:var(--card-sub); border:1px solid var(--border); border-radius:10px;">' +
             header + rows +
           '</div>';
  }

  function renderStockList(stocks) {
    stocks = (stocks || []).filter(isInventoryStock_);
    var box = document.getElementById('stockListContainer');
    document.getElementById('stockCount').innerText = (stocks ? stocks.length : 0) + " 檔標的";
    if (!stocks || stocks.length === 0) {
      box.innerHTML = '<div style="text-align:center; color:var(--subtext); padding: 10px;">無持倉</div>';
      return;
    }
    var html = '';
    stocks.forEach(function(s) {
      var subText = (s.market === 'TW') 
        ? ('質押 ' + s.pledgedQty.toLocaleString() + ' 股 | 未質押 ' + s.unpledgedQty.toLocaleString() + ' 股')
        : ('美股折台幣 $' + Math.round(s.marketVal).toLocaleString());
      
      html += '<div class="list-item">' +
              '<div>' +
                '<div class="name">[' + s.market + '] ' + s.code + ' ' + s.name + '</div>' +
                '<div class="sub">' + subText + '</div>' +
              '</div>' +
              '<div class="data" style="color: var(--primary);">' + s.totalQty.toLocaleString() + ' 股</div>' +
            '</div>';
    });
    box.innerHTML = html;
  }

  function renderFutures(fut) {
    if (!fut) return;
    document.getElementById('twFutPos').innerText = "未平倉：" + fut.twPosition;
    document.getElementById('twFutMargin').innerText = "$" + Number(fut.twEquity).toLocaleString();
    document.getElementById('usFutPos').innerText = "未平倉：" + fut.usPosition + " (匯率 " + fut.usRate + ")";
    document.getElementById('usFutMargin').innerText = "$" + Number(fut.usEquityTwd).toLocaleString();
  }

  function renderDebtList(pledges, loans) {
    var box = document.getElementById('debtListContainer');
    var html = '';
    if (pledges) {
      pledges.forEach(function(p) {
        html += '<div class="list-item">' +
                '<div>' +
                  '<div class="name">' + p.name + '</div>' +
                  '<div class="sub">年利率 ' + p.rate + '%</div>' +
                '</div>' +
                '<div class="data" style="color: var(--danger);">$' + Number(p.balance).toLocaleString() + '</div>' +
              '</div>';
      });
    }
    if (loans) {
      loans.forEach(function(l) {
        var subText = (l.defaultPay > 0) ? ('已繳 ' + l.periods + ' 期 |月還 $' + l.defaultPay.toLocaleString()) : '循環借款/固定本金';
        html += '<div class="list-item">' +
                '<div>' +
                  '<div class="name">' + l.name + '</div>' +
                  '<div class="sub">' + subText + '</div>' +
                '</div>' +
                '<div class="data" style="color: var(--danger);">$' + Number(l.balance).toLocaleString() + '</div>' +
              '</div>';
      });
    }
    box.innerHTML = html || '<div style="text-align:center; color:var(--subtext); padding: 10px;">無負債</div>';
  }

  function initDispatchSelects(res) {
    if (res.stocks && res.stocks.length > 0) {
      
      // 1. 💡 補回：填滿「既有標的買賣」的下拉選單
      var stockOpt = res.stocks.map(function(s){ 
        return '<option value="'+s.row+'" data-market="'+s.market+'">['+s.market+'] '+s.code+' '+s.name+' (現有: '+s.totalQty+')</option>'; 
      }).join('');
      var tradeSelect = document.getElementById('stockTradeSelect');
      if (tradeSelect) {
        tradeSelect.innerHTML = stockOpt;
      }

      // 2. 渲染「K 欄目標佔比調整」的輸入清單 (第二頁上方)
      renderExistingStocksRatioInputs(res.stocks);

      // 3. 填滿質押標的選單
      var twStocks = res.stocks.filter(function(s){ return s.market === 'TW'; });
      var pledgeOpt = twStocks.map(function(s){ 
        return '<option value="'+s.row+'">'+s.code+' '+s.name+' (未質押: '+s.unpledgedQty+' / 質押: '+s.pledgedQty+')</option>'; 
      }).join('');
      var pledgeSelect = document.getElementById('pledgeStockSelect');
      if (pledgeSelect) {
        pledgeSelect.innerHTML = pledgeOpt;
      }
    }
    renderCashSubInputs();
  }

  // 🔄 動態渲染金流子選項（支援股息入帳、貸款、質押）
function renderCashSubInputs() {
  if (!globalData) return;
  var type = document.getElementById('cashActionType').value;
  var box = document.getElementById('targetSelectBox');
  var label = document.getElementById('targetSelectLabel');
  var sel = document.getElementById('targetSelectDropdown');
  var divBox = document.getElementById('dividendInfoBox');
  var amtInput = document.getElementById('cashAmountInput');

  // 預設全部隱藏
  box.style.display = 'none';
  divBox.style.display = 'none';
  amtInput.readOnly = false;

  if (type === 'DIVIDEND') {
    box.style.display = 'block';
    divBox.style.display = 'block';
    label.innerText = '選擇股息入帳標的：';
    
    // 過濾出所有台股或有庫存的標的
    var twStocks = (globalData.stocks || []).filter(function(s){ return s.market === 'TW'; });
    sel.innerHTML = twStocks.map(function(s){ 
      return '<option value="'+s.code+'" data-name="'+s.name+'" data-shares="'+s.totalQty+'">['+s.code+'] '+s.name+' (庫存: '+s.totalQty+'股)</option>'; 
    }).join('');
    
    amtInput.readOnly = true; // 總金額由系統自動計算
    onTargetDropdownChange(); // 初始化連動

  } else if (type === 'LOAN_PAY') {
    box.style.display = 'block';
    label.innerText = '選擇貸款項目：';
    sel.innerHTML = (globalData.loans || []).map(function(l){ 
      return '<option value="'+l.row+'" data-pay="'+l.defaultPay+'">'+l.name+' (月還: $'+l.defaultPay+')</option>'; 
    }).join('');
    if (sel.options.length > 0) {
      amtInput.value = sel.options[0].getAttribute('data-pay') || 0;
    }
  } else if (type === 'PLEDGE_BORROW' || type === 'PLEDGE_REPAY') {
    box.style.display = 'block';
    label.innerText = '選擇質押機構：';
    sel.innerHTML = (globalData.pledges || []).map(function(p){ 
      return '<option value="'+p.row+'">'+p.name+' (餘額: $'+p.balance.toLocaleString()+')</option>'; 
    }).join('');
    amtInput.value = '';
  } else {
    amtInput.value = '';
  }
}

// 🎯 當股息下拉選單切換時，自動更新股數與超連結
function onTargetDropdownChange() {
  var type = document.getElementById('cashActionType').value;
  if (type !== 'DIVIDEND') return;

  var sel = document.getElementById('targetSelectDropdown');
  var opt = sel.options[sel.selectedIndex];
  if (!opt) return;

  var code = opt.value;
  var shares = Number(opt.getAttribute('data-shares')) || 0;

  // 更新介面股數顯示
  document.getElementById('displayShares').innerText = shares.toLocaleString();

  // 更新 Goodinfo 與 MoneyDJ 查詢超連結
  document.getElementById('goodinfoBtn').href = "https://goodinfo.tw/tw/StockDetail.asp?STOCK_ID=" + code;
  document.getElementById('moneydjBtn').href = "https://www.moneydj.com/ETF/X/Basic/Basic0005.xdjhtm?etfid=" + code + ".TW";

  // 重新計算總配息
  calculateTotalDividend();
}

// 🧮 依據總庫存股數 × 每股配息計算總金額
function calculateTotalDividend() {
  var sel = document.getElementById('targetSelectDropdown');
  var opt = sel.options[sel.selectedIndex];
  if (!opt) return;

  var shares = Number(opt.getAttribute('data-shares')) || 0;
  var dps = Number(document.getElementById('dividendPerShareInput').value) || 0;
  var total = Math.round(shares * dps);

  document.getElementById('cashAmountInput').value = total;
}

  function submitCashDispatch() {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 找不到試算表 ID！");
    var type = document.getElementById('cashActionType').value;
    var amt = Number(document.getElementById('cashAmountInput').value || 0);
    var sel = document.getElementById('targetSelectDropdown');
    var row = (sel && sel.value) ? sel.value : 0;
    if (amt <= 0) return alert("⚠️ 請輸入有效金額！");

    var loader = document.getElementById('cashLoader');
    loader.style.display = 'block';
    runDispatchAction('CASH_DISPATCH', { type: type, amount: amt, row: row }, loader, function(data) {
      alert(data.message || "執行成功");
      document.getElementById('cashAmountInput').value = '';
      connectSheet();
    });
  }

  function submitStockTrade() {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 找不到試算表 ID！");
    var sel = document.getElementById('stockTradeSelect');
    if (!sel || sel.options.length === 0) return alert("⚠️ 尚未載入股票標的！");

    var row = sel.value;
    var market = sel.options[sel.selectedIndex].getAttribute('data-market');
    var action = document.getElementById('stockTradeAction').value;
    var qty = Number(document.getElementById('stockTradeQty').value || 0);
    var cashDiff = Number(document.getElementById('stockTradeCash').value || 0);
    if (qty <= 0) return alert("⚠️ 請輸入有效變動股數！");

    var loader = document.getElementById('stockLoader');
    loader.style.display = 'block';
    runDispatchAction('STOCK_TRADE', { market: market, row: row, action: action, qty: qty, cashDiff: cashDiff }, loader, function(data) {
      alert(data.message || "執行成功");
      document.getElementById('stockTradeQty').value = '';
      connectSheet();
    });
  }

  // 📊 渲染現有標的目標佔比調整清單 (分別獨立渲染台股與美股)
  function renderExistingStocksRatioInputs(stocks) {
    var twContainer = document.getElementById('twStocksRatioContainer');
    var usContainer = document.getElementById('usStocksRatioContainer');

    var twStocks = (stocks || []).filter(function(s){ return s.market === 'TW'; });
    var usStocks = (stocks || []).filter(function(s){ return s.market === 'US'; });

    // 1. 渲染台股清單
    if (twStocks.length === 0) {
      twContainer.innerHTML = '<div style="text-align:center; color:var(--subtext); padding: 6px; font-size:12px;">無台股持倉</div>';
    } else {
      twContainer.innerHTML = twStocks.map(function(s) {
        return buildRatioRowHtml(s);
      }).join('');
    }

    // 2. 渲染美股清單
    if (usStocks.length === 0) {
      usContainer.innerHTML = '<div style="text-align:center; color:var(--subtext); padding: 6px; font-size:12px;">無美股持倉</div>';
    } else {
      usContainer.innerHTML = usStocks.map(function(s) {
        return buildRatioRowHtml(s);
      }).join('');
    }

    updateTotalStockRatios();
  }

  // 🛠️ 產生單行佔比輸入框的 HTML 輔助函式
  function buildRatioRowHtml(s) {
    var currentRatio = Number(s.targetRatio) || 0;
    return '<div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:6px; padding:4px 6px; background:var(--card-sub); border-radius:4px;">' +
             '<div style="font-size:12px; flex:1;"><strong>' + s.code + '</strong> ' + s.name + '</div>' +
             '<div style="display:flex; align-items:center; gap:6px;">' +
               '<input type="number" class="existing-ratio-input" data-market="' + s.market + '" data-row="' + s.row + '" value="' + currentRatio + '" step="0.1" min="0" max="100" style="width:70px; margin-bottom:0; text-align:right; padding:4px;" oninput="updateTotalStockRatios()">' +
               '<span style="font-size:12px;">%</span>' +
             '</div>' +
           '</div>';
  }

  // 📈 即時計算並更新台股與美股各自的總佔比
  function updateTotalStockRatios() {
    var inputs = document.querySelectorAll('.existing-ratio-input');
    var twTotal = 0, usTotal = 0;

    inputs.forEach(function(inp) {
      var val = Number(inp.value) || 0;
      if (inp.getAttribute('data-market') === 'TW') {
        twTotal += val;
      } else {
        usTotal += val;
      }
    });

    var twIndicator = document.getElementById('twTotalRatioIndicator');
    var usIndicator = document.getElementById('usTotalRatioIndicator');

    if (twIndicator) {
      twIndicator.innerText = '📊 台股總佔比：' + twTotal.toFixed(1) + '%';
      twIndicator.style.color = (Math.abs(twTotal - 100) < 0.1) ? 'var(--accent)' : 'var(--warning, #f59e0b)';
    }
    if (usIndicator) {
      usIndicator.innerText = '📊 美股總佔比：' + usTotal.toFixed(1) + '%';
      usIndicator.style.color = (Math.abs(usTotal - 100) < 0.1) ? 'var(--accent)' : 'var(--warning, #f59e0b)';
    }
  }

  // 💾 一次性儲存目前市場目標佔比 (寫入 K 欄)
  function submitBatchSaveTargetRatios() {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 找不到試算表 ID！");

    var inputs = document.querySelectorAll('.existing-ratio-input');
    var updates = [];
    inputs.forEach(function(inp) {
      updates.push({
        row: inp.getAttribute('data-row'),
        targetRatio: Number(inp.value) || 0
      });
    });

    var loader = document.getElementById('stockLoader');
    loader.style.display = 'block';
    runDispatchAction('BATCH_SAVE_TARGET_RATIOS', { updates: updates }, loader, function(data) {
      alert(data.message || "目標佔比儲存成功！");
      connectSheet();
    }, '❌ 儲存失敗: ');
  }

  // 🔄 雙頁面切換控制 (既有標的買賣 vs 新增標的與佔比調整)
  function toggleStockMode() {
    isNewStockMode = !isNewStockMode;
    document.getElementById('stockTradeExistingMode').style.display = isNewStockMode ? 'none' : 'block';
    document.getElementById('stockTradeNewMode').style.display = isNewStockMode ? 'block' : 'none';
    document.getElementById('toggleStockModeBtn').innerText = isNewStockMode ? '🔙 返回既有標的買賣' : '➕ 新增標的與佔比調整';
  }

  // 🚀 提交新增全新標的 (名稱交由試算表公式自動帶入，無需前端傳送名稱)
  function submitAddNewStock() {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 找不到試算表 ID！");
    
    var market = document.getElementById('newStockMarket').value;
    var code = document.getElementById('newStockCode').value.trim();
    var qty = Number(document.getElementById('newStockQty').value || 0);
    var cashTotal = Number(document.getElementById('newStockCashTotal').value || 0);
    var ratio = Number(document.getElementById('newStockRatio').value || 0);
    
    if (!code) return alert("⚠️ 請輸入股票代碼！");

    var loader = document.getElementById('stockLoader');
    loader.style.display = 'block';
    runDispatchAction('ADD_STOCK', {
      market: market,
      code: code,
      qty: qty,
      cashTotal: cashTotal,
      targetRatio: ratio
    }, loader, function(data) {
      alert(data.message || "建立並初始化完成");
      document.getElementById('newStockCode').value = '';
      document.getElementById('newStockQty').value = '';
      document.getElementById('newStockCashTotal').value = '0';
      document.getElementById('newStockRatio').value = '';
      toggleStockMode();
      connectSheet();
    });
  }

  function submitPledgeTransfer() {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 找不到試算表 ID！");
    var sel = document.getElementById('pledgeStockSelect');
    if (!sel || sel.options.length === 0) return alert("⚠️ 尚未載入台股質押標的！");
    var row = sel.value;
    var action = document.getElementById('pledgeAction').value;
    var qty = Number(document.getElementById('pledgeQtyInput').value || 0);
    if (qty <= 0) return alert("⚠️ 請輸入有效劃撥股數！");

    var loader = document.getElementById('pledgeLoader');
    loader.style.display = 'block';
    runDispatchAction('PLEDGE_TRANSFER', { row: row, action: action, qty: qty }, loader, function(data) {
      alert(data.message || "執行成功");
      document.getElementById('pledgeQtyInput').value = '';
      connectSheet();
    });
  }

  function submitFuturesTransfer() {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 找不到試算表 ID！");
    var market = document.getElementById('futMarketSelect').value;
    var action = document.getElementById('futActionSelect').value;
    var amt = Number(document.getElementById('futAmountInput').value || 0);
    if (amt <= 0) return alert("⚠️ 請輸入有效金額！");

    var loader = document.getElementById('futLoader');
    loader.style.display = 'block';
    runDispatchAction('FUTURES_TRANSFER', { market: market, action: action, amount: amt }, loader, function(data) {
      alert(data.message || "執行成功");
      document.getElementById('futAmountInput').value = '';
      connectSheet();
    });
  }

  function triggerSnapshot(market) {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 請先連線試算表！");
    var loader = document.getElementById('snapshotLoader');
    loader.style.display = 'block';
    var runner = (market === 'TW') ? 'triggerTaiwanStockSnapshotUniversal' : (market === 'US' ? 'triggerUSStockSnapshotUniversal' : 'triggerDailySnapshotUniversal');

    callGasApi(runner, { sheetId: sheetId })
      .then(function(res) {
        loader.style.display = 'none';
        handleResponseGuard(res, function(data) {
          alert(data.message || "快照完成！");
          connectSheet();
        });
      })
      .catch(function(err) {
        loader.style.display = 'none';
        alert("❌ 失敗: " + err.message);
      });
  }

  function triggerRefresh(type) {
    var sheetId = getCurrentSheetId();
    if (!sheetId) return alert("⚠️ 請先連線試算表！");
    var loader = document.getElementById('snapshotLoader');
    loader.style.display = 'block';
    var runner = (type === 'BETA') ? 'refreshBetaAlphaUniversal' : 'refreshMarketDataUniversal';

    callGasApi(runner, { sheetId: sheetId })
      .then(function(res) {
        loader.style.display = 'none';
        handleResponseGuard(res, function(data) {
          alert(data.message || "刷新完成！");
          connectSheet();
        });
      })
      .catch(function(err) {
        loader.style.display = 'none';
        alert("❌ 失敗: " + err.message);
      });
  }

  function autoFetchStockName() {
    var market = document.getElementById('newStockMarket').value;
    var code = document.getElementById('newStockCode').value.trim();
    var nameInput = document.getElementById('newStockName');
    if (!code) return;
    nameInput.placeholder = "⏳ 查詢中...";

    callGasApi('fetchStockInfoUniversal', { market: market, code: code })
      .then(function(res) {
        if (res && res.success && res.name) nameInput.value = res.name;
        else nameInput.placeholder = "請手動輸入";
      })
      .catch(function() { nameInput.placeholder = "請手動輸入"; });
  }

  function switchDispatchSub(subId, el) {
    document.querySelectorAll('.dispatch-sub').forEach(function(sub) { sub.style.display = 'none'; });
    document.querySelectorAll('.pill-item').forEach(function(item) { item.classList.remove('active'); });
    document.getElementById(subId).style.display = 'block';
    el.classList.add('active');
  }

// 🔄 切換「既有合約選單」與「手動新增合約」模式
var isCustomContractMode = false;

function toggleContractInputMode() {
  isCustomContractMode = !isCustomContractMode;
  var selectEl = document.getElementById("existingContractSelect");
  var inputEl = document.getElementById("customContractInput");
  var btnEl = document.getElementById("toggleContractModeBtn");

  if (isCustomContractMode) {
    selectEl.style.display = "none";
    selectEl.required = false;
    inputEl.style.display = "block";
    inputEl.required = true;
    inputEl.value = "";
    
    // 🎨 切換為手動輸入模式時的樣式（改為醒目的橘色系提示）
    btnEl.innerText = "📋 切換回現有合約選單";
    btnEl.style.background = "#fef3c7";
    btnEl.style.borderColor = "#d97706";
    btnEl.style.color = "#b45309";
  } else {
    inputEl.style.display = "none";
    inputEl.required = false;
    selectEl.style.display = "block";
    selectEl.required = true;
    
    // 🎨 切回選單模式時的樣式（恢復原本的清爽藍色系）
    btnEl.innerText = "➕ 切換手動新增合約";
    btnEl.style.background = "#e0f2fe";
    btnEl.style.borderColor = "#0284c7";
    btnEl.style.color = "#0369a1";
  }
}