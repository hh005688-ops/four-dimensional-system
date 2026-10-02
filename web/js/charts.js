var cachedOverviewData = null;
  var cachedStockData = null;
  var globalChartInstance = null;
  
  var rawSnapshotTables = {
    general: [],
    taiwan: [],
    us: []
  };

  function switchMasterView(type) {
    document.querySelectorAll('.snapshot-table-pane').forEach(function(pane) { 
      pane.style.display = 'none'; 
    });

    if (type === 'OVERVIEW') {
      document.getElementById('tableGeneralContainer').style.display = 'block';
      if (cachedOverviewData) renderOverviewChart(cachedOverviewData);
    } else if (type === 'TW') {
      document.getElementById('tableTaiwanContainer').style.display = 'block';
      if (cachedStockData && cachedStockData.tw) {
        renderSingleStockChart('台股市值', cachedStockData.tw.dates, cachedStockData.tw.values, '#10b981');
      }
    } else if (type === 'US') {
      document.getElementById('tableUsContainer').style.display = 'block';
      if (cachedStockData && cachedStockData.us) {
        renderSingleStockChart('美股市值 (折合台幣)', cachedStockData.us.dates, cachedStockData.us.values, '#f59e0b');
      }
    }

    filterCustomRange();
  }

  function loadSnapshotCharts() {
    var sheetId = "";
    try {
      sheetId = typeof getCurrentSheetId === 'function' ? getCurrentSheetId() : "";
    } catch(e) { console.warn(e); }

    var chartDom = document.getElementById('dailyChart');
    if (chartDom) chartDom.innerHTML = '⏳ 正在向後端取得歷史紀錄...';

    if (typeof callGasApi !== 'function') return;

    callGasApi('getDailyLogHistoryUniversal', { sheetId: sheetId })
      .then(function(res) {
        if (res && res.success) {
          cachedOverviewData = res;
          var view = document.getElementById('masterViewSelector').value;
          if (view === 'OVERVIEW') renderOverviewChart(res);
        }
      })
      .catch(function(err) { console.warn(err); });

    callGasApi('getStockLogHistoryUniversal', { sheetId: sheetId })
      .then(function(res) {
        if (res && res.success) {
          cachedStockData = res;
          var view = document.getElementById('masterViewSelector').value;
          if (view !== 'OVERVIEW') switchMasterView(view);
        }
      })
      .catch(function(err) { console.warn(err); });
  }

  function loadSnapshotHistoryTable() {
    var sheetId = "";
    try {
      sheetId = typeof getCurrentSheetId === 'function' ? getCurrentSheetId() : "";
    } catch(e) { console.warn(e); }
    if (!sheetId) return;
    if (typeof callGasApi !== 'function') return;

    callGasApi('getSnapshotHistoryDataUniversal', { sheetId: sheetId })
      .then(function(res) {
        if (!res || !res.success) return;
        rawSnapshotTables.general = res.general || [];
        rawSnapshotTables.taiwan = res.taiwan || [];
        rawSnapshotTables.us = res.us || [];
        filterCustomRange();
      })
      .catch(function(err) { console.warn(err); });
  }

  function renderFilteredTables(startDate, endDate) {
    function filterData(dataList) {
      if (!startDate && !endDate) return dataList.slice().reverse();
      return dataList.filter(function(row) {
        var d = new Date(row.date);
        if (startDate && d < startDate) return false;
        if (endDate && d > endDate) return false;
        return true;
      }).reverse();
    }

    var filteredGen = filterData(rawSnapshotTables.general);
    var filteredTw = filterData(rawSnapshotTables.taiwan);
    var filteredUs = filterData(rawSnapshotTables.us);

    var genBody = document.getElementById('generalTableBody');
    if (genBody) {
      genBody.innerHTML = filteredGen.length > 0 ? filteredGen.map(function(row) {
        var pCol = (Number(row.todayProfit) >= 0) ? 'var(--accent, #10b981)' : 'var(--danger, #ef4444)';
        var rCol = (Number(row.todayReturn) >= 0) ? 'var(--accent, #10b981)' : 'var(--danger, #ef4444)';
        return '<tr style="border-bottom: 1px solid var(--border);">' +
          '<td style="padding: 6px 8px;"><strong>' + row.date + '</strong></td>' +
          '<td style="padding: 6px 8px; text-align: right;">$' + Number(row.totalAsset).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right;">$' + Number(row.netAsset).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right; color:' + pCol + ';">$' + Math.round(row.todayProfit).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right; color:' + rCol + ';">' + (row.todayReturn * 100).toFixed(2) + '%</td>' +
          '<td style="padding: 6px 8px; text-align: right;">$' + Math.round(row.cumProfit).toLocaleString() + '</td>' +
          '</tr>';
      }).join('') : '<tr><td colspan="6" style="text-align:center; padding:20px; color:var(--subtext);">此區間無總覽紀錄</td></tr>';
    }

    var twBody = document.getElementById('taiwanTableBody');
    if (twBody) {
      twBody.innerHTML = filteredTw.length > 0 ? filteredTw.map(function(row) {
        var pCol = (Number(row.profit) >= 0) ? 'var(--accent, #10b981)' : 'var(--danger, #ef4444)';
        var rCol = (Number(row.returnRate) >= 0) ? 'var(--accent, #10b981)' : 'var(--danger, #ef4444)';
        return '<tr style="border-bottom: 1px solid var(--border);">' +
          '<td style="padding: 6px 8px;"><strong>' + row.date + '</strong></td>' +
          '<td style="padding: 6px 8px; text-align: right;">$' + Number(row.marketVal).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right; color:' + pCol + ';">$' + Math.round(row.profit).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right; color:' + rCol + ';">' + (row.returnRate * 100).toFixed(2) + '%</td>' +
          '<td style="padding: 6px 8px; text-align: right;">' + (row.benchReturn * 100).toFixed(2) + '%</td>' +
          '<td style="padding: 6px 8px; text-align: right;">' + (row.alpha * 100).toFixed(2) + '%</td>' +
          '</tr>';
      }).join('') : '<tr><td colspan="6" style="text-align:center; padding:20px; color:var(--subtext);">此區間無台股明細</td></tr>';
    }

    var usBody = document.getElementById('usTableBody');
    if (usBody) {
      usBody.innerHTML = filteredUs.length > 0 ? filteredUs.map(function(row) {
        var pCol = (Number(row.profit) >= 0) ? 'var(--accent, #10b981)' : 'var(--danger, #ef4444)';
        var rCol = (Number(row.returnRate) >= 0) ? 'var(--accent, #10b981)' : 'var(--danger, #ef4444)';
        return '<tr style="border-bottom: 1px solid var(--border);">' +
          '<td style="padding: 6px 8px;"><strong>' + row.date + '</strong></td>' +
          '<td style="padding: 6px 8px; text-align: right;">$' + Number(row.marketVal).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right; color:' + pCol + ';">$' + Math.round(row.profit).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right; color:' + rCol + ';">' + (row.returnRate * 100).toFixed(2) + '%</td>' +
          '<td style="padding: 6px 8px; text-align: right;">$' + Number(row.fxProfit).toLocaleString() + '</td>' +
          '<td style="padding: 6px 8px; text-align: right;">' + (row.alpha * 100).toFixed(2) + '%</td>' +
          '</tr>';
      }).join('') : '<tr><td colspan="6" style="text-align:center; padding:20px; color:var(--subtext);">此區間無美股明細</td></tr>';
    }
  }

  function filterCustomRange() {
    var startVal = document.getElementById('customStartDate').value;
    var endVal = document.getElementById('customEndDate').value;
    var startDate = startVal ? new Date(startVal) : null;
    var endDate = endVal ? new Date(endVal) : null;
    renderFilteredTables(startDate, endDate);
  }

  function setChartRange(type, btnEl) {
    var btns = btnEl.parentNode.querySelectorAll('.range-btn');
    btns.forEach(function(b) { b.classList.remove('active'); });
    btnEl.classList.add('active');

    var view = document.getElementById('masterViewSelector').value;
    var dates = [];
    if (view === 'OVERVIEW' && cachedOverviewData) {
      dates = cachedOverviewData.dates;
    } else if (cachedStockData && cachedStockData.tw) {
      dates = cachedStockData.tw.dates;
    }

    if (!dates || dates.length === 0) return;

    var lastDateStr = dates[dates.length - 1];
    var lastDate = new Date(lastDateStr);
    var startDate = new Date(lastDate);

    if (type === 'ALL') {
      if (globalChartInstance) globalChartInstance.dispatchAction({ type: 'dataZoom', start: 0, end: 100 });
      document.getElementById('customStartDate').value = '';
      document.getElementById('customEndDate').value = '';
      renderFilteredTables(null, null);
      return;
    } else if (type === '1W') {
      startDate.setDate(lastDate.getDate() - 7);
    } else if (type === '1M') {
      startDate.setMonth(lastDate.getMonth() - 1);
    } else if (type === '1Q') {
      startDate.setMonth(lastDate.getMonth() - 3);
    } else if (type === '6M') {
      startDate.setMonth(lastDate.getMonth() - 6);
    } else if (type === '1Y') {
      startDate.setFullYear(lastDate.getFullYear() - 1);
    }

    document.getElementById('customStartDate').value = startDate.toISOString().split('T')[0];
    document.getElementById('customEndDate').value = lastDate.toISOString().split('T')[0];
    renderFilteredTables(startDate, lastDate);

    if (globalChartInstance) {
      var startIndex = 0;
      for (var i = 0; i < dates.length; i++) {
        if (new Date(dates[i]) >= startDate) { startIndex = i; break; }
      }
      var startPercent = (startIndex / (dates.length - 1)) * 100;
      globalChartInstance.dispatchAction({ type: 'dataZoom', start: Math.max(0, startPercent), end: 100 });
    }
  }

  function formatMoneyAxis(value) {
    if (value >= 1000000) return (value / 1000000).toFixed(1) + 'M';
    if (value >= 1000) return (value / 1000).toFixed(0) + 'k';
    return value;
  }

  function renderOverviewChart(dataObj) {
    var chartDom = document.getElementById('dailyChart');
    if (!chartDom) return;
    if (globalChartInstance) globalChartInstance.dispose();

    globalChartInstance = echarts.init(chartDom, 'dark');
    var option = {
      backgroundColor: 'transparent',
      tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
      legend: { data: ['總資產', '淨資產', '動態回撤率'], textStyle: { color: '#ccc', fontSize: 11 }, top: 0 },
      dataZoom: [{ type: 'slider', xAxisIndex: [0, 1], start: 0, end: 100, bottom: '2%', height: 16 }, { type: 'inside', xAxisIndex: [0, 1] }],
      grid: [{ left: '12%', right: '3%', top: '15%', height: '48%' }, { left: '12%', right: '3%', top: '70%', height: '18%' }],
      xAxis: [
        { type: 'category', data: dataObj.dates, gridIndex: 0, boundaryGap: false, axisLabel: { show: false } },
        { type: 'category', data: dataObj.dates, gridIndex: 1, boundaryGap: false, axisLabel: { color: '#aaa', fontSize: 10 } }
      ],
      yAxis: [
        { type: 'value', gridIndex: 0, scale: false, splitLine: { lineStyle: { color: '#333' } }, axisLabel: { color: '#aaa', fontSize: 11, formatter: formatMoneyAxis } },
        { type: 'value', gridIndex: 1, scale: false, inverse: true, max: 0, splitLine: { lineStyle: { color: '#333' } }, axisLabel: { color: '#aaa', fontSize: 11, formatter: '{value}%' } }
      ],
      series: [
        { name: '總資產', type: 'line', xAxisIndex: 0, yAxisIndex: 0, data: dataObj.assets, itemStyle: { color: '#4169E1' }, showSymbol: false, lineStyle: { width: 2 } },
        { name: '淨資產', type: 'line', xAxisIndex: 0, yAxisIndex: 0, data: dataObj.netWorths, itemStyle: { color: '#00FA9A' }, showSymbol: false, lineStyle: { width: 2 } },
        { name: '動態回撤率', type: 'line', xAxisIndex: 1, yAxisIndex: 1, data: dataObj.marketDDs, itemStyle: { color: '#FF4500' }, showSymbol: false, areaStyle: { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: 'rgba(255, 69, 0, 0.4)' }, { offset: 1, color: 'rgba(255, 69, 0, 0.02)' }]) }, lineStyle: { width: 1.5 } }
      ]
    };
    globalChartInstance.setOption(option);
  }

  function renderSingleStockChart(name, dates, values, colorHex) {
    var chartDom = document.getElementById('dailyChart');
    if (!chartDom) return;
    if (globalChartInstance) globalChartInstance.dispose();

    globalChartInstance = echarts.init(chartDom, 'dark');
    var option = {
      backgroundColor: 'transparent',
      tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
      legend: { data: [name], textStyle: { color: '#ccc', fontSize: 11 }, top: 0 },
      dataZoom: [{ type: 'slider', start: 0, end: 100, bottom: '2%', height: 16 }, { type: 'inside' }],
      grid: { left: '12%', right: '3%', top: '15%', bottom: '20%' },
      xAxis: { type: 'category', data: dates, boundaryGap: false, axisLabel: { color: '#aaa', fontSize: 10 } },
      yAxis: { type: 'value', scale: false, splitLine: { lineStyle: { color: '#333' } }, axisLabel: { color: '#aaa', fontSize: 11, formatter: formatMoneyAxis } },
      series: [{
        name: name,
        type: 'line',
        data: values,
        itemStyle: { color: colorHex },
        showSymbol: false,
        areaStyle: { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: colorHex + '66' }, { offset: 1, color: colorHex + '05' }]) },
        lineStyle: { width: 2 }
      }]
    };
    globalChartInstance.setOption(option);
  }

  // 🏛️ 分頁切換總指揮函式
  function switchMainTab(tabId, el) {
    document.querySelectorAll('.tab-content').forEach(function(tab) {
      tab.classList.remove('active');
    });
    document.querySelectorAll('.nav-item').forEach(function(item) {
      item.classList.remove('active');
    });
    
    var targetTab = document.getElementById(tabId);
    if (targetTab) {
      targetTab.classList.add('active');
    }
    if (el) {
      el.classList.add('active');
    }

    if (tabId === 'tabHome' && typeof resizeNestedAllocationChart === 'function') {
      setTimeout(resizeNestedAllocationChart, 60);
    }
    // 當切換到快照頁籤時，自動觸發歷史走勢圖與歷史明細表格載入
    if (tabId === 'tabSnapshot') {
      if (typeof loadSnapshotCharts === 'function') {
        loadSnapshotCharts();
      }
      if (typeof loadSnapshotHistoryTable === 'function') {
        loadSnapshotHistoryTable();
      }
    }
  }