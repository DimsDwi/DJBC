/**
 * OceanWave BI Dashboard Application Logic
 * Data Source: Rekap_Penerimaan_2026_CLEAN (1).xlsx (Sheet: DENPASAR_CLEAN)
 */

document.addEventListener('DOMContentLoaded', async () => {
  // Register Chart.js Datalabels plugin if available
  if (typeof ChartDataLabels !== 'undefined') {
    Chart.register(ChartDataLabels);
  }

  // Default dataset verified against Rekap_Penerimaan_2026_CLEAN (1).xlsx
  const defaultDataset = [
    { "No Bulan": 1, "Bulan": "JANUARI", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 152617000, "Cukai Tembakau": 424586000, "Cukai MMEA": 102981160000, "Cukai Lainnya": 1730729000, "Cukai Total": 105136475000, "Total Akhir": 105289092000, "Total Kumulatif TLKV": 105289092000 },
    { "No Bulan": 2, "Bulan": "FEBRUARI", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 137137000, "Cukai Tembakau": 2343856000, "Cukai MMEA": 88352014000, "Cukai Lainnya": 926023000, "Cukai Total": 91621893000, "Total Akhir": 91759030000, "Total Kumulatif TLKV": 197048122000 },
    { "No Bulan": 3, "Bulan": "MARET", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 84529000, "Cukai Tembakau": 1700193000, "Cukai MMEA": 68235181000, "Cukai Lainnya": 938983000, "Cukai Total": 70874357000, "Total Akhir": 70958886000, "Total Kumulatif TLKV": 268007008000 },
    { "No Bulan": 4, "Bulan": "APRIL", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 145923000, "Cukai Tembakau": 1335715000, "Cukai MMEA": 90957534000, "Cukai Lainnya": 716535000, "Cukai Total": 93009784000, "Total Akhir": 93155707000, "Total Kumulatif TLKV": 361162715000 },
    { "No Bulan": 5, "Bulan": "MEI", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 227349000, "Cukai Tembakau": 1537945000, "Cukai MMEA": 98184026000, "Cukai Lainnya": 159218000, "Cukai Total": 99881189000, "Total Akhir": 100108538000, "Total Kumulatif TLKV": 461271253000 },
    { "No Bulan": 6, "Bulan": "JUNI", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 100809000, "Cukai Tembakau": 3420730000, "Cukai MMEA": 133570845000, "Cukai Lainnya": 836030000, "Cukai Total": 137827605000, "Total Akhir": 137928414000, "Total Kumulatif TLKV": 599199667000 },
    { "No Bulan": 7, "Bulan": "JULI", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 142196000, "Cukai Tembakau": 1561957000, "Cukai MMEA": 123474168000, "Cukai Lainnya": 1480902000, "Cukai Total": 126517027000, "Total Akhir": 126659223000, "Total Kumulatif TLKV": 725858890000 },
    { "No Bulan": 8, "Bulan": "AGUSTUS", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 122792000, "Cukai Tembakau": 2024219000, "Cukai MMEA": 128321733000, "Cukai Lainnya": 121329000, "Cukai Total": 130467281000, "Total Akhir": 130590073000, "Total Kumulatif TLKV": 856448963000 },
    { "No Bulan": 9, "Bulan": "SEPTEMBER", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 25504000, "Cukai Tembakau": 1113956000, "Cukai MMEA": 21658605000, "Cukai Lainnya": 0, "Cukai Total": 22772561000, "Total Akhir": 22798065000, "Total Kumulatif TLKV": 879247028000 },
    { "No Bulan": 10, "Bulan": "OKTOBER", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 0, "Cukai Tembakau": 0, "Cukai MMEA": 0, "Cukai Lainnya": 0, "Cukai Total": 0, "Total Akhir": 0, "Total Kumulatif TLKV": 879247028000 },
    { "No Bulan": 11, "Bulan": "NOVEMBER", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 0, "Cukai Tembakau": 0, "Cukai MMEA": 0, "Cukai Lainnya": 0, "Cukai Total": 0, "Total Akhir": 0, "Total Kumulatif TLKV": 879247028000 },
    { "No Bulan": 12, "Bulan": "DESEMBER", "Tahun": 2026, "Kantor": "DENPASAR", "Bea Masuk": 0, "Cukai Tembakau": 0, "Cukai MMEA": 0, "Cukai Lainnya": 0, "Cukai Total": 0, "Total Akhir": 0, "Total Kumulatif TLKV": 879247028000 }
  ];

  // Dynamic raw dataset (load custom uploaded data if available in localStorage)
  let rawDataset = [...defaultDataset];
  const savedDataStr = localStorage.getItem('djbc_custom_dataset');
  if (savedDataStr) {
    try {
      const parsed = JSON.parse(savedDataStr);
      if (Array.isArray(parsed) && parsed.length > 0) {
        rawDataset = parsed;
      }
    } catch (e) {
      console.warn('Gagal memuat dataset tersimpan:', e);
    }
  }

  // Global Filter & Sort States
  let state = {
    tahun: '2026',
    selectedMonths: ['ALL'],
    kantor: 'ALL',
    jenisPenerimaan: 'Total Akhir',
    activeCrossMonth: null
  };

  // Table Sorting State (for Page 1 and Page 2)
  let sortStateP1 = { col: null, dir: 'asc' };
  let sortStateP2 = { col: null, dir: 'asc' };

  // Chart References
  const charts = {
    p1Trend: null,
    p1Bar: null,
    p1Kumulatif: null,
    p2Trend: null,
    p2Donut: null,
    p2Mom: null
  };

  // --- Formatting Helpers ---
  function formatBN(val) {
    if (val === null || val === undefined || isNaN(val)) return '0.0 BN';
    const bn = val / 1000000000;
    return bn.toFixed(1) + ' BN';
  }

  function formatIDR(val) {
    if (val === null || val === undefined || isNaN(val)) return 'Rp 0';
    return 'Rp ' + Math.round(val).toLocaleString('id-ID');
  }

  function formatPercent(val) {
    if (val === null || val === undefined || isNaN(val)) return '-';
    const sign = val > 0 ? '+' : '';
    return sign + val.toFixed(1) + '%';
  }

  // --- Month Helper ---
  function isMonthSelected(monthName) {
    if (state.selectedMonths.includes('ALL')) return true;
    return state.selectedMonths.includes(monthName);
  }

  // --- Filter Logic ---
  function getFilteredData() {
    return rawDataset.filter(d => {
      if (state.tahun !== 'ALL' && String(d['Tahun']) !== state.tahun) return false;
      if (!isMonthSelected(d['Bulan'])) return false;
      if (state.kantor !== 'ALL' && d['Kantor'] !== state.kantor) return false;
      if (state.activeCrossMonth && d['Bulan'] !== state.activeCrossMonth) return false;
      return true;
    });
  }

  // --- UI Navigation ---
  const navButtons = document.querySelectorAll('.nav-menu button');
  const pages = document.querySelectorAll('.dashboard-page');
  const pageTitle = document.getElementById('current-page-title');
  const pageSubtitle = document.getElementById('current-page-subtitle');

  const pageMeta = {
    'page-overview': { title: 'Executive Overview', subtitle: 'Monitoring Realisasi Penerimaan Bea dan Cukai 2026' },
    'page-analysis': { title: 'Analysis & Details', subtitle: 'Analisis Komposisi Cukai & Perubahan Bulanan (% MoM)' },
    'page-explore': { title: 'Explore Data', subtitle: 'Eksplorasi Data Mentah dan Ekspor Laporan' }
  };

  navButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetPage = btn.getAttribute('data-page');
      
      navButtons.forEach(b => b.classList.remove('active'));
      pages.forEach(p => p.classList.remove('active'));
      
      btn.classList.add('active');
      document.getElementById(targetPage).classList.add('active');

      if (pageMeta[targetPage]) {
        pageTitle.innerText = pageMeta[targetPage].title;
        pageSubtitle.innerText = pageMeta[targetPage].subtitle;
      }

      renderDashboard();
    });
  });

  // --- Multi-Select Month Dropdown Logic ---
  const btnDropdownBulan = document.getElementById('btn-dropdown-bulan');
  const menuDropdownBulan = document.getElementById('menu-dropdown-bulan');
  const labelDropdownBulan = document.getElementById('label-dropdown-bulan');
  const chkBulanAll = document.getElementById('chk-bulan-all');
  const chkBulanItems = document.querySelectorAll('.chk-bulan');

  btnDropdownBulan.addEventListener('click', (e) => {
    e.stopPropagation();
    menuDropdownBulan.classList.toggle('open');
  });

  document.addEventListener('click', () => {
    menuDropdownBulan.classList.remove('open');
  });

  menuDropdownBulan.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  function updateSelectedMonthsFromUI() {
    if (chkBulanAll.checked) {
      state.selectedMonths = ['ALL'];
      labelDropdownBulan.innerText = 'Semua Bulan';
    } else {
      const checked = Array.from(chkBulanItems).filter(c => c.checked).map(c => c.value);
      if (checked.length === 0) {
        state.selectedMonths = [];
        labelDropdownBulan.innerText = 'Pilih Bulan';
      } else if (checked.length === 12) {
        chkBulanAll.checked = true;
        state.selectedMonths = ['ALL'];
        labelDropdownBulan.innerText = 'Semua Bulan';
      } else if (checked.length === 1) {
        state.selectedMonths = checked;
        const name = checked[0];
        labelDropdownBulan.innerText = name.charAt(0) + name.slice(1).toLowerCase();
      } else {
        state.selectedMonths = checked;
        labelDropdownBulan.innerText = `${checked.length} Bulan Dipilih`;
      }
    }
    state.activeCrossMonth = null;
    renderDashboard();
  }

  chkBulanAll.addEventListener('change', () => {
    chkBulanItems.forEach(c => c.checked = chkBulanAll.checked);
    updateSelectedMonthsFromUI();
  });

  chkBulanItems.forEach(c => {
    c.addEventListener('change', () => {
      const allChecked = Array.from(chkBulanItems).every(c => c.checked);
      chkBulanAll.checked = allChecked;
      updateSelectedMonthsFromUI();
    });
  });

  // --- Other Filter Event Listeners ---
  const selectTahun = document.getElementById('filter-tahun');
  const selectKantor = document.getElementById('filter-kantor');
  const selectJenis = document.getElementById('selector-jenis-penerimaan');
  const btnReset = document.getElementById('btn-reset-filters');

  selectTahun.addEventListener('change', (e) => { state.tahun = e.target.value; renderDashboard(); });
  selectKantor.addEventListener('change', (e) => { state.kantor = e.target.value; renderDashboard(); });
  selectJenis.addEventListener('change', (e) => { state.jenisPenerimaan = e.target.value; renderDashboard(); });

  btnReset.addEventListener('click', () => {
    state.tahun = '2026';
    state.selectedMonths = ['ALL'];
    state.kantor = 'ALL';
    state.jenisPenerimaan = 'Total Akhir';
    state.activeCrossMonth = null;

    sortStateP1 = { col: null, dir: 'asc' };
    sortStateP2 = { col: null, dir: 'asc' };

    selectTahun.value = '2026';
    selectKantor.value = 'ALL';
    selectJenis.value = 'Total Akhir';

    chkBulanAll.checked = true;
    chkBulanItems.forEach(c => c.checked = true);
    labelDropdownBulan.innerText = 'Semua Bulan';

    renderDashboard();
  });

  // --- Table Sorting Event Listeners (Fitur 4: Sortir Kolom Interaktif) ---
  document.querySelectorAll('#table-p1-detail th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (sortStateP1.col === col) {
        sortStateP1.dir = (sortStateP1.dir === 'asc') ? 'desc' : 'asc';
      } else {
        sortStateP1.col = col;
        sortStateP1.dir = (col === 'Bulan') ? 'asc' : 'desc';
      }
      renderPage1(getFilteredData());
    });
  });

  document.querySelectorAll('#table-p2-detail th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (sortStateP2.col === col) {
        sortStateP2.dir = (sortStateP2.dir === 'asc') ? 'desc' : 'asc';
      } else {
        sortStateP2.col = col;
        sortStateP2.dir = (col === 'Bulan') ? 'asc' : 'desc';
      }
      renderPage2(getFilteredData());
    });
  });

  // --- Core Render Orchestrator ---
  function renderDashboard() {
    const filtered = getFilteredData();

    // 1. Render Page 1 (Overview)
    renderPage1(filtered);

    // 2. Render Page 2 (Analysis)
    renderPage2(filtered);

    // 3. Render Page 3 (Explore)
    renderPage3(rawDataset);
  }

  // ==========================================
  // PAGE 1 LOGIC
  // ==========================================
  function renderPage1(data) {
    // KPI Cards
    const totalPenerimaan = data.reduce((s, r) => s + (r['Total Akhir'] || 0), 0);
    const mmea = data.reduce((s, r) => s + (r['Cukai MMEA'] || 0), 0);
    const tembakau = data.reduce((s, r) => s + (r['Cukai Tembakau'] || 0), 0);
    const lainnya = data.reduce((s, r) => s + (r['Cukai Lainnya'] || 0), 0);

    document.getElementById('kpi-p1-total').innerText = formatBN(totalPenerimaan);
    document.getElementById('kpi-p1-mmea').innerText = formatBN(mmea);
    document.getElementById('kpi-p1-tembakau').innerText = formatBN(tembakau);
    document.getElementById('kpi-p1-lainnya').innerText = formatBN(lainnya);

    // Dynamic Subtexts based on actual composition
    const pctMMEA = totalPenerimaan > 0 ? ((mmea / totalPenerimaan) * 100).toFixed(1) : '0.0';
    const pctHT = totalPenerimaan > 0 ? ((tembakau / totalPenerimaan) * 100).toFixed(1) : '0.0';
    const pctLain = totalPenerimaan > 0 ? ((lainnya / totalPenerimaan) * 100).toFixed(1) : '0.0';

    const subTotal = document.getElementById('kpi-p1-total-sub');
    if (subTotal) subTotal.innerText = `Total Realisasi 2026 (100.0%)`;
    const subMmea = document.getElementById('kpi-p1-mmea-sub');
    if (subMmea) subMmea.innerText = `${pctMMEA}% dari Total Penerimaan`;
    const subHt = document.getElementById('kpi-p1-tembakau-sub');
    if (subHt) subHt.innerText = `${pctHT}% dari Total Penerimaan`;
    const subLain = document.getElementById('kpi-p1-lainnya-sub');
    if (subLain) subLain.innerText = `${pctLain}% dari Total Penerimaan`;

    // Fitur 1: Executive Highlights (100% fakta data aktual Excel)
    const activeRows = rawDataset.filter(r => (r['Total Akhir'] || 0) > 0 && isMonthSelected(r['Bulan']));
    if (activeRows.length > 0) {
      const bestRow = activeRows.reduce((prev, curr) => (curr['Total Akhir'] > prev['Total Akhir']) ? curr : prev, activeRows[0]);
      const worstRow = activeRows.reduce((prev, curr) => (curr['Total Akhir'] < prev['Total Akhir']) ? curr : prev, activeRows[0]);
      const sumActive = activeRows.reduce((s, r) => s + (r['Total Akhir'] || 0), 0);
      const avgMonthVal = sumActive / activeRows.length;
      const sumActiveMMEA = activeRows.reduce((s, r) => s + (r['Cukai MMEA'] || 0), 0);
      const mmeaPct = sumActive > 0 ? ((sumActiveMMEA / sumActive) * 100).toFixed(1) : '0.0';

      const bestPct = sumActive > 0 ? (((bestRow['Total Akhir'] || 0) / sumActive) * 100).toFixed(1) : '0.0';
      const worstPct = sumActive > 0 ? (((worstRow['Total Akhir'] || 0) / sumActive) * 100).toFixed(1) : '0.0';

      const elBestMonth = document.getElementById('hl-best-month');
      const elBestVal = document.getElementById('hl-best-value');
      const elWorstMonth = document.getElementById('hl-worst-month');
      const elWorstVal = document.getElementById('hl-worst-value');
      const elAvgMonth = document.getElementById('hl-avg-month');
      const elMmeaShare = document.getElementById('hl-mmea-share');

      if (elBestMonth) elBestMonth.innerText = bestRow['Bulan'];
      if (elBestVal) elBestVal.innerText = `${formatBN(bestRow['Total Akhir'])} (${bestPct}%)`;
      if (elWorstMonth) elWorstMonth.innerText = worstRow['Bulan'];
      if (elWorstVal) elWorstVal.innerText = `${formatBN(worstRow['Total Akhir'])} (${worstPct}%)`;
      if (elAvgMonth) elAvgMonth.innerText = formatBN(avgMonthVal);
      if (elMmeaShare) elMmeaShare.innerText = `${formatBN(sumActiveMMEA)} (${mmeaPct}%)`;
    }

    const metric = state.jenisPenerimaan;
    document.getElementById('trend-subtitle').innerText = `Tren realisasi ${metric} per bulan (2026)`;

    // Monthly data arrays for charts (Jan - Des)
    const monthLabels = rawDataset.map(d => d['Bulan']);
    const metricValues = rawDataset.map(d => {
      if (!isMonthSelected(d['Bulan'])) return 0;
      if (state.activeCrossMonth && d['Bulan'] !== state.activeCrossMonth) return 0;
      return d[metric] || 0;
    });

    // Chart 1: Trend Line/Area
    if (charts.p1Trend) charts.p1Trend.destroy();
    const ctxTrend = document.getElementById('chart-p1-trend').getContext('2d');
    
    const gradient = ctxTrend.createLinearGradient(0, 0, 0, 300);
    gradient.addColorStop(0, 'rgba(56, 189, 248, 0.4)');
    gradient.addColorStop(1, 'rgba(56, 189, 248, 0.0)');

    charts.p1Trend = new Chart(ctxTrend, {
      type: 'line',
      data: {
        labels: monthLabels,
        datasets: [{
          label: metric,
          data: metricValues.map(v => v / 1000000000), // in Billion
          borderColor: '#38bdf8',
          borderWidth: 3,
          backgroundColor: gradient,
          fill: true,
          tension: 0.35,
          pointBackgroundColor: '#0284c7',
          pointBorderColor: '#ffffff',
          pointHoverRadius: 7,
          pointRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const rawVal = rawDataset[ctx.dataIndex][metric];
                const pct = totalPenerimaan > 0 ? ((rawVal / totalPenerimaan) * 100).toFixed(1) : 0;
                return `${metric}: ${ctx.raw.toFixed(1)} BN (${pct}% kontribusi | ${formatIDR(rawVal)})`;
              }
            }
          },
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: '#94a3b8', font: { size: 11 } } },
          y: { 
            grid: { color: 'rgba(255, 255, 255, 0.06)' }, 
            ticks: { color: '#94a3b8', callback: (v) => v + ' BN' } 
          }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const idx = elements[0].index;
            const clickedMonth = monthLabels[idx];
            state.activeCrossMonth = (state.activeCrossMonth === clickedMonth) ? null : clickedMonth;
            renderDashboard();
          }
        }
      }
    });

    // Chart 2: Monthly Bar Chart
    if (charts.p1Bar) charts.p1Bar.destroy();
    const ctxBar = document.getElementById('chart-p1-bar').getContext('2d');
    
    // MoM for bar chart datalabels
    const momValuesP1 = [];
    for (let i = 0; i < rawDataset.length; i++) {
      const valCurrent = rawDataset[i][metric] || 0;
      if (i === 0) {
        momValuesP1.push(null);
      } else {
        const valPrev = rawDataset[i - 1][metric] || 0;
        if (valPrev === 0) {
          momValuesP1.push(null);
        } else {
          const pctChange = ((valCurrent - valPrev) / valPrev) * 100;
          momValuesP1.push(pctChange);
        }
      }
    }

    charts.p1Bar = new Chart(ctxBar, {
      type: 'bar',
      data: {
        labels: monthLabels,
        datasets: [{
          label: metric,
          data: metricValues.map(v => v / 1000000000),
          backgroundColor: metricValues.map((v, i) => 
            state.activeCrossMonth === monthLabels[i] ? '#06b6d4' : (isMonthSelected(monthLabels[i]) ? 'rgba(56, 189, 248, 0.85)' : 'rgba(255, 255, 255, 0.1)')
          ),
          borderRadius: 6,
          hoverBackgroundColor: '#38bdf8'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const rawVal = rawDataset[ctx.dataIndex][metric];
                const pct = totalPenerimaan > 0 ? ((rawVal / totalPenerimaan) * 100).toFixed(1) : 0;
                const mom = momValuesP1[ctx.dataIndex];
                const momStr = mom !== null ? (mom > 0 ? `+${mom.toFixed(1)}% MoM` : `${mom.toFixed(1)}% MoM`) : 'MoM: -';
                return [`${ctx.raw.toFixed(1)} BN (${pct}% kontribusi)`, momStr];
              }
            }
          },
          datalabels: {
            display: true,
            anchor: 'end',
            align: 'top',
            color: '#38bdf8',
            font: { weight: 'bold', size: 9 },
            formatter: (val, ctx) => {
              const mom = momValuesP1[ctx.dataIndex];
              if (mom === null || mom === undefined || isNaN(mom)) return '';
              const sign = mom > 0 ? '+' : '';
              return `${sign}${mom.toFixed(1)}%`;
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: '#94a3b8', font: { size: 10 } } },
          y: { grid: { color: 'rgba(255, 255, 255, 0.06)' }, ticks: { color: '#94a3b8' } }
        },
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const idx = elements[0].index;
            const clickedMonth = monthLabels[idx];
            state.activeCrossMonth = (state.activeCrossMonth === clickedMonth) ? null : clickedMonth;
            renderDashboard();
          }
        }
      }
    });

    // Chart Tambahan: Akumulasi Penerimaan Berjalan (YTD) - Realisasi Aktual dari Excel
    const ctxKumulatif = document.getElementById('chart-p1-kumulatif');
    if (ctxKumulatif) {
      if (charts.p1Kumulatif) charts.p1Kumulatif.destroy();
      const ctxK = ctxKumulatif.getContext('2d');

      // Realisasi kumulatif YTD dari kolom asli 'Total Kumulatif TLKV'
      const ytdValues = rawDataset.map(d => (d['Total Kumulatif TLKV'] || 0) / 1000000000);

      charts.p1Kumulatif = new Chart(ctxK, {
        type: 'line',
        data: {
          labels: monthLabels,
          datasets: [
            {
              label: 'Realisasi Kumulatif (YTD)',
              data: ytdValues,
              borderColor: '#10b981',
              borderWidth: 3,
              backgroundColor: 'rgba(16, 185, 129, 0.12)',
              fill: true,
              tension: 0.3,
              pointBackgroundColor: '#10b981',
              pointBorderColor: '#ffffff',
              pointRadius: 5
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              display: true,
              position: 'top',
              labels: { color: '#f8fafc', font: { size: 11 } }
            },
            tooltip: {
              callbacks: {
                label: (ctx) => `${ctx.dataset.label}: ${ctx.raw.toFixed(1)} BN (${formatIDR(rawDataset[ctx.dataIndex]['Total Kumulatif TLKV'])})`
              }
            },
            datalabels: {
              display: (ctx) => ctx.dataIndex <= 8,
              anchor: 'end',
              align: 'top',
              color: '#10b981',
              font: { weight: 'bold', size: 9 },
              formatter: (v) => `${v.toFixed(1)} BN`
            }
          },
          scales: {
            x: { grid: { display: false }, ticks: { color: '#94a3b8' } },
            y: {
              grid: { color: 'rgba(255, 255, 255, 0.06)' },
              ticks: { color: '#94a3b8', callback: (v) => v + ' BN' }
            }
          }
        }
      });
    }

    // Table Overview with MoM, Data-Bars, and Interactive Sorting
    const tbody = document.querySelector('#table-p1-detail tbody');
    tbody.innerHTML = '';

    let sumBM = 0, sumMMEA = 0, sumHT = 0, sumLain = 0, sumCukaiTotal = 0, sumTotal = 0;

    rawDataset.forEach(row => {
      const isFiltered = isMonthSelected(row['Bulan']) &&
                         (!state.activeCrossMonth || state.activeCrossMonth === row['Bulan']);

      const bm = isFiltered ? (row['Bea Masuk'] || 0) : 0;
      const mmeaVal = isFiltered ? (row['Cukai MMEA'] || 0) : 0;
      const htVal = isFiltered ? (row['Cukai Tembakau'] || 0) : 0;
      const lainVal = isFiltered ? (row['Cukai Lainnya'] || 0) : 0;
      const cTotal = isFiltered ? (row['Cukai Total'] || 0) : 0;
      const tAkhir = isFiltered ? (row['Total Akhir'] || 0) : 0;

      sumBM += bm;
      sumMMEA += mmeaVal;
      sumHT += htVal;
      sumLain += lainVal;
      sumCukaiTotal += cTotal;
      sumTotal += tAkhir;
    });

    // Prepare table items
    let tableP1Rows = rawDataset.map((row, i) => {
      const isFiltered = isMonthSelected(row['Bulan']) &&
                         (!state.activeCrossMonth || state.activeCrossMonth === row['Bulan']);

      const bm = isFiltered ? (row['Bea Masuk'] || 0) : 0;
      const mmeaVal = isFiltered ? (row['Cukai MMEA'] || 0) : 0;
      const htVal = isFiltered ? (row['Cukai Tembakau'] || 0) : 0;
      const lainVal = isFiltered ? (row['Cukai Lainnya'] || 0) : 0;
      const cTotal = isFiltered ? (row['Cukai Total'] || 0) : 0;
      const tAkhir = isFiltered ? (row['Total Akhir'] || 0) : 0;
      const pctShare = sumTotal > 0 ? (tAkhir / sumTotal) * 100 : 0;

      let momNum = null;
      let momBadge = '<span class="badge-mom neutral">-</span>';
      if (i > 0) {
        const prevAkhir = rawDataset[i - 1]['Total Akhir'] || 0;
        const currAkhir = row['Total Akhir'] || 0;
        if (prevAkhir > 0) {
          momNum = ((currAkhir - prevAkhir) / prevAkhir) * 100;
          if (momNum > 0) {
            momBadge = `<span class="badge-mom pos">+${momNum.toFixed(1)}%</span>`;
          } else if (momNum < 0) {
            momBadge = `<span class="badge-mom neg">${momNum.toFixed(1)}%</span>`;
          } else {
            momBadge = `<span class="badge-mom neutral">0.0%</span>`;
          }
        }
      }

      return {
        row,
        order: row['No Bulan'],
        Bulan: row['Bulan'],
        'Bea Masuk': bm,
        'Cukai MMEA': mmeaVal,
        'Cukai Tembakau': htVal,
        'Cukai Lainnya': lainVal,
        'Cukai Total': cTotal,
        'Total Akhir': tAkhir,
        pctShare: pctShare,
        mom: momNum,
        momBadge: momBadge
      };
    });

    // Apply sort if specified
    if (sortStateP1.col) {
      tableP1Rows.sort((a, b) => {
        let valA = a[sortStateP1.col];
        let valB = b[sortStateP1.col];
        if (valA === null || valA === undefined) valA = -999999999999;
        if (valB === null || valB === undefined) valB = -999999999999;
        if (typeof valA === 'string') {
          return sortStateP1.dir === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
        return sortStateP1.dir === 'asc' ? (valA - valB) : (valB - valA);
      });
    }

    // Update Header Sort Indicators
    document.querySelectorAll('#table-p1-detail th.sortable').forEach(th => {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.getAttribute('data-sort') === sortStateP1.col) {
        th.classList.add(sortStateP1.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });

    // Render table rows with data-bars
    tableP1Rows.forEach(item => {
      const tr = document.createElement('tr');
      if (!isMonthSelected(item.Bulan)) {
        tr.style.opacity = '0.4';
      }
      if (state.activeCrossMonth === item.Bulan) {
        tr.style.backgroundColor = 'rgba(6, 182, 212, 0.15)';
      }

      const databarHtml = item.pctShare > 0 ? `
        <div class="databars-cell">
          <div class="databar-bg">
            <div class="databar-fill" style="width: ${Math.min(item.pctShare, 100)}%;"></div>
          </div>
          <span class="databar-label">${item.pctShare.toFixed(1)}%</span>
        </div>
      ` : `<span style="color: var(--text-dim);">-</span>`;

      tr.innerHTML = `
        <td style="font-weight: 600;">${item.Bulan}</td>
        <td class="num">${formatIDR(item['Bea Masuk'])}</td>
        <td class="num">${formatIDR(item['Cukai MMEA'])}</td>
        <td class="num">${formatIDR(item['Cukai Tembakau'])}</td>
        <td class="num">${formatIDR(item['Cukai Lainnya'])}</td>
        <td class="num" style="color: #38bdf8;">${formatIDR(item['Cukai Total'])}</td>
        <td class="num" style="font-weight:700; color: #f8fafc;">${formatIDR(item['Total Akhir'])}</td>
        <td class="num">${databarHtml}</td>
        <td class="num">${item.momBadge}</td>
      `;
      tbody.appendChild(tr);
    });

    const trTotal = document.createElement('tr');
    trTotal.className = 'total-row';
    trTotal.innerHTML = `
      <td>TOTAL</td>
      <td class="num">${formatIDR(sumBM)}</td>
      <td class="num">${formatIDR(sumMMEA)}</td>
      <td class="num">${formatIDR(sumHT)}</td>
      <td class="num">${formatIDR(sumLain)}</td>
      <td class="num" style="color: #38bdf8;">${formatIDR(sumCukaiTotal)}</td>
      <td class="num" style="font-weight:800; color: #38bdf8;">${formatIDR(sumTotal)}</td>
      <td class="num" style="font-weight:800; color: #38bdf8;">100.0%</td>
      <td class="num" style="color: var(--text-dim);">—</td>
    `;
    tbody.appendChild(trTotal);
  }


  // ==========================================
  // PAGE 2 LOGIC
  // ==========================================
  function renderPage2(data) {
    // KPI Cards Page 2
    const totalPenerimaan = data.reduce((s, r) => s + (r['Total Akhir'] || 0), 0);
    const mmea = data.reduce((s, r) => s + (r['Cukai MMEA'] || 0), 0);
    const tembakau = data.reduce((s, r) => s + (r['Cukai Tembakau'] || 0), 0);

    document.getElementById('kpi-p2-total').innerText = formatBN(totalPenerimaan);
    document.getElementById('kpi-p2-mmea').innerText = formatBN(mmea);
    document.getElementById('kpi-p2-tembakau').innerText = formatBN(tembakau);

    // Dynamic Subtexts Page 2 based on actual filtered data
    const pctMmeaP2 = totalPenerimaan > 0 ? ((mmea / totalPenerimaan) * 100).toFixed(1) : '0.0';
    const pctHtP2 = totalPenerimaan > 0 ? ((tembakau / totalPenerimaan) * 100).toFixed(1) : '0.0';

    const subTotalP2 = document.getElementById('kpi-p2-total-sub');
    if (subTotalP2) subTotalP2.innerText = `Total Realisasi Filtered`;
    const subMmeaP2 = document.getElementById('kpi-p2-mmea-sub');
    if (subMmeaP2) subMmeaP2.innerText = `${pctMmeaP2}% dari Total Realisasi`;
    const subHtP2 = document.getElementById('kpi-p2-tembakau-sub');
    if (subHtP2) subHtP2.innerText = `${pctHtP2}% dari Total Realisasi`;

    const metric = state.jenisPenerimaan;
    document.getElementById('p2-trend-title').innerText = `Trend ${metric} 2026`;

    const monthLabels = rawDataset.map(d => d['Bulan']);
    const metricValues = rawDataset.map(d => {
      if (!isMonthSelected(d['Bulan'])) return 0;
      if (state.activeCrossMonth && d['Bulan'] !== state.activeCrossMonth) return 0;
      return d[metric] || 0;
    });

    // Visual 1: Trend Cukai Dynamic Line Chart
    if (charts.p2Trend) charts.p2Trend.destroy();
    const ctxTrend2 = document.getElementById('chart-p2-trend').getContext('2d');

    charts.p2Trend = new Chart(ctxTrend2, {
      type: 'line',
      data: {
        labels: monthLabels,
        datasets: [{
          label: metric,
          data: metricValues.map(v => v / 1000000000),
          borderColor: '#818cf8',
          borderWidth: 3,
          backgroundColor: 'rgba(129, 140, 248, 0.15)',
          fill: true,
          tension: 0.3,
          pointBackgroundColor: '#6366f1',
          pointBorderColor: '#ffffff',
          pointRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const rawVal = rawDataset[ctx.dataIndex][metric];
                const pct = totalPenerimaan > 0 ? ((rawVal / totalPenerimaan) * 100).toFixed(1) : 0;
                return `${metric}: ${ctx.raw.toFixed(1)} BN (${pct}% kontribusi | ${formatIDR(rawVal)})`;
              }
            }
          },
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: '#94a3b8' } },
          y: { grid: { color: 'rgba(255, 255, 255, 0.06)' }, ticks: { color: '#94a3b8', callback: (v) => v + ' BN' } }
        }
      }
    });

    // Visual 2: Composition Donut Chart
    // Priority: cross-filter > month filter > all data (NOT driven by Jenis Penerimaan slicer)
    const hasMonthFilter = !state.selectedMonths.includes('ALL') && state.selectedMonths.length < 12;
    const hasCrossFilter = !!state.activeCrossMonth;

    let donutSource;
    let donutLabel;
    if (hasCrossFilter) {
      // User clicked a specific month bar/point → show that month only
      donutSource = rawDataset.filter(r => r['Bulan'] === state.activeCrossMonth);
      donutLabel = state.activeCrossMonth;
    } else if (hasMonthFilter) {
      // User selected specific months via multi-select → show those months
      donutSource = rawDataset.filter(r => isMonthSelected(r['Bulan']));
      donutLabel = `${state.selectedMonths.length} Bulan Dipilih`;
    } else {
      // No filter at all → show full year (all data)
      donutSource = rawDataset;
      donutLabel = 'Keseluruhan 2026';
    }
    
    const sumBeaMasuk = donutSource.reduce((s, r) => s + (r['Bea Masuk'] || 0), 0);
    const sumMMEAComp = donutSource.reduce((s, r) => s + (r['Cukai MMEA'] || 0), 0);
    const sumHTComp = donutSource.reduce((s, r) => s + (r['Cukai Tembakau'] || 0), 0);
    const sumLainComp = donutSource.reduce((s, r) => s + (r['Cukai Lainnya'] || 0), 0);
    const compTotal = sumBeaMasuk + sumMMEAComp + sumHTComp + sumLainComp;

    // Update donut chart subtitle dynamically
    const donutSubEl = document.getElementById('p2-donut-subtitle');
    if (donutSubEl) {
      if (hasCrossFilter) {
        donutSubEl.innerText = `Komposisi bulan ${donutLabel}`;
      } else if (hasMonthFilter) {
        donutSubEl.innerText = `Komposisi ${donutLabel}`;
      } else {
        donutSubEl.innerText = 'Komposisi Penerimaan 2026 — Keseluruhan';
      }
    }

    if (charts.p2Donut) charts.p2Donut.destroy();
    const ctxDonut = document.getElementById('chart-p2-donut').getContext('2d');

    charts.p2Donut = new Chart(ctxDonut, {
      type: 'doughnut',
      data: {
        labels: ['Cukai MMEA', 'Cukai Tembakau', 'Cukai Lainnya', 'Bea Masuk'],
        datasets: [{
          data: [sumMMEAComp, sumHTComp, sumLainComp, sumBeaMasuk].map(v => v / 1000000000),
          backgroundColor: ['#38bdf8', '#f59e0b', '#10b981', '#6366f1'],
          borderWidth: 2,
          borderColor: '#1e293b'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { color: '#f8fafc', font: { size: 11 }, padding: 14,
              generateLabels: (chart) => {
                const data = chart.data;
                const rawVals = [sumMMEAComp, sumHTComp, sumLainComp, sumBeaMasuk];
                return data.labels.map((label, i) => {
                  const pct = compTotal > 0 ? ((rawVals[i] / compTotal) * 100).toFixed(1) : 0;
                  return {
                    text: `${label}  ${pct}%`,
                    fillStyle: data.datasets[0].backgroundColor[i],
                    strokeStyle: '#1e293b',
                    lineWidth: 2,
                    index: i
                  };
                });
              }
            }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const valBN = ctx.raw.toFixed(1);
                const rawVal = [sumMMEAComp, sumHTComp, sumLainComp, sumBeaMasuk][ctx.dataIndex];
                const pct = compTotal > 0 ? ((rawVal / compTotal) * 100).toFixed(1) : 0;
                return [`${ctx.label}`, `${valBN} BN  (${pct}% dari total)`];
              }
            }
          },
          datalabels: {
            color: '#ffffff',
            font: { weight: 'bold', size: 11 },
            formatter: (value, ctx) => {
              const rawVal = [sumMMEAComp, sumHTComp, sumLainComp, sumBeaMasuk][ctx.dataIndex];
              const pct = compTotal > 0 ? ((rawVal / compTotal) * 100).toFixed(1) : 0;
              return parseFloat(pct) > 1 ? `${pct}%` : '';
            }
          }
        },
        cutout: '65%'
      }
    });

    // Visual 3: % Change MoM Column Chart
    const momValues = [];
    const momColors = [];

    for (let i = 0; i < rawDataset.length; i++) {
      const valCurrent = rawDataset[i][metric] || 0;
      if (i === 0) {
        momValues.push(null);
        momColors.push('rgba(255,255,255,0.1)');
      } else {
        const valPrev = rawDataset[i - 1][metric] || 0;
        if (valPrev === 0) {
          momValues.push(null);
          momColors.push('rgba(255,255,255,0.1)');
        } else {
          const pctChange = ((valCurrent - valPrev) / valPrev) * 100;
          momValues.push(pctChange);
          momColors.push(pctChange >= 0 ? '#10b981' : '#ef4444');
        }
      }
    }

    if (charts.p2Mom) charts.p2Mom.destroy();
    const ctxMom = document.getElementById('chart-p2-mom').getContext('2d');

    charts.p2Mom = new Chart(ctxMom, {
      type: 'bar',
      data: {
        labels: monthLabels,
        datasets: [{
          label: `% Perubahan ${metric} MoM`,
          data: momValues,
          backgroundColor: momColors,
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                if (ctx.raw === null) return 'Januari: Tidak ada pembanding bulan sebelumnya';
                const sign = ctx.raw > 0 ? '+' : '';
                return `% Perubahan MoM: ${sign}${ctx.raw.toFixed(1)}%`;
              }
            }
          },
          datalabels: {
            anchor: (ctx) => (ctx.dataset.data[ctx.dataIndex] || 0) >= 0 ? 'end' : 'start',
            align: (ctx) => (ctx.dataset.data[ctx.dataIndex] || 0) >= 0 ? 'top' : 'bottom',
            color: '#f8fafc',
            font: { weight: 'bold', size: 10 },
            formatter: (val) => {
              if (val === null) return '-';
              const sign = val > 0 ? '+' : '';
              return `${sign}${val.toFixed(1)}%`;
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: '#94a3b8' } },
          y: {
            grid: { color: 'rgba(255, 255, 255, 0.06)' },
            ticks: {
              color: '#94a3b8',
              callback: (v) => v + '%'
            }
          }
        }
      }
    });

    // Table Detail Page 2 with Data-Bars and Interactive Sorting
    const tbody2 = document.querySelector('#table-p2-detail tbody');
    tbody2.innerHTML = '';

    let sumTotalPage2 = 0;
    let sumMmeaP2 = 0, sumHtP2 = 0, sumLainP2 = 0, sumCTotalP2 = 0;
    rawDataset.forEach(row => {
      const isFiltered = isMonthSelected(row['Bulan']) &&
                         (!state.activeCrossMonth || state.activeCrossMonth === row['Bulan']);
      if (isFiltered) {
        sumTotalPage2 += (row['Total Akhir'] || 0);
        sumMmeaP2 += (row['Cukai MMEA'] || 0);
        sumHtP2 += (row['Cukai Tembakau'] || 0);
        sumLainP2 += (row['Cukai Lainnya'] || 0);
        sumCTotalP2 += (row['Cukai Total'] || 0);
      }
    });

    let tableP2Rows = rawDataset.map((row, i) => {
      const isFiltered = isMonthSelected(row['Bulan']) &&
                         (!state.activeCrossMonth || state.activeCrossMonth === row['Bulan']);

      const mmeaVal = isFiltered ? (row['Cukai MMEA'] || 0) : 0;
      const htVal = isFiltered ? (row['Cukai Tembakau'] || 0) : 0;
      const lainVal = isFiltered ? (row['Cukai Lainnya'] || 0) : 0;
      const cTotal = isFiltered ? (row['Cukai Total'] || 0) : 0;
      const tAkhir = isFiltered ? (row['Total Akhir'] || 0) : 0;

      const pctShare = sumTotalPage2 > 0 ? (tAkhir / sumTotalPage2) * 100 : 0;

      const momPct = momValues[i];
      let badgeHtml = '<span class="badge-mom neutral">-</span>';
      if (momPct !== null && !isNaN(momPct)) {
        if (momPct > 0) {
          badgeHtml = `<span class="badge-mom pos">+${momPct.toFixed(1)}%</span>`;
        } else if (momPct < 0) {
          badgeHtml = `<span class="badge-mom neg">${momPct.toFixed(1)}%</span>`;
        } else {
          badgeHtml = `<span class="badge-mom neutral">0.0%</span>`;
        }
      }

      return {
        row,
        order: row['No Bulan'],
        Bulan: row['Bulan'],
        'Cukai MMEA': mmeaVal,
        'Cukai Tembakau': htVal,
        'Cukai Lainnya': lainVal,
        'Cukai Total': cTotal,
        'Total Akhir': tAkhir,
        pctShare: pctShare,
        mom: momPct,
        momBadge: badgeHtml
      };
    });

    // Apply sort if specified
    if (sortStateP2.col) {
      tableP2Rows.sort((a, b) => {
        let valA = a[sortStateP2.col];
        let valB = b[sortStateP2.col];
        if (valA === null || valA === undefined) valA = -999999999999;
        if (valB === null || valB === undefined) valB = -999999999999;
        if (typeof valA === 'string') {
          return sortStateP2.dir === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
        return sortStateP2.dir === 'asc' ? (valA - valB) : (valB - valA);
      });
    }

    // Update Header Sort Indicators Page 2
    document.querySelectorAll('#table-p2-detail th.sortable').forEach(th => {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.getAttribute('data-sort') === sortStateP2.col) {
        th.classList.add(sortStateP2.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });

    tableP2Rows.forEach(item => {
      const tr = document.createElement('tr');
      if (!isMonthSelected(item.Bulan)) {
        tr.style.opacity = '0.4';
      }

      const databarHtml = item.pctShare > 0 ? `
        <div class="databars-cell">
          <div class="databar-bg">
            <div class="databar-fill" style="width: ${Math.min(item.pctShare, 100)}%;"></div>
          </div>
          <span class="databar-label">${item.pctShare.toFixed(1)}%</span>
        </div>
      ` : `<span style="color: var(--text-dim);">-</span>`;

      tr.innerHTML = `
        <td style="font-weight: 600;">${item.Bulan}</td>
        <td class="num">${formatIDR(item['Cukai MMEA'])}</td>
        <td class="num">${formatIDR(item['Cukai Tembakau'])}</td>
        <td class="num">${formatIDR(item['Cukai Lainnya'])}</td>
        <td class="num" style="color: #38bdf8;">${formatIDR(item['Cukai Total'])}</td>
        <td class="num" style="font-weight:700;">${formatIDR(item['Total Akhir'])}</td>
        <td class="num">${databarHtml}</td>
        <td class="num">${item.momBadge}</td>
      `;
      tbody2.appendChild(tr);
    });

    // Total Row Page 2
    const trTotal2 = document.createElement('tr');
    trTotal2.className = 'total-row';
    trTotal2.innerHTML = `
      <td>TOTAL</td>
      <td class="num">${formatIDR(sumMmeaP2)}</td>
      <td class="num">${formatIDR(sumHtP2)}</td>
      <td class="num">${formatIDR(sumLainP2)}</td>
      <td class="num" style="color: #38bdf8;">${formatIDR(sumCTotalP2)}</td>
      <td class="num" style="font-weight:800; color: #38bdf8;">${formatIDR(sumTotalPage2)}</td>
      <td class="num" style="font-weight:800; color: #38bdf8;">100.0%</td>
      <td class="num" style="color: var(--text-dim);">—</td>
    `;
    tbody2.appendChild(trTotal2);
  }


  // ==========================================
  // PAGE 3 LOGIC (EXPLORE & EXPORT)
  // ==========================================
  function renderPage3(data) {
    const tbody3 = document.querySelector('#table-p3-raw tbody');
    if (!tbody3) return;

    const searchTerm = (document.getElementById('search-input')?.value || '').toLowerCase();

    tbody3.innerHTML = '';

    data.forEach(r => {
      const rowStr = `${r['No Bulan']} ${r['Bulan']} ${r['Tahun']} ${r['Kantor']} ${r['Total Akhir']}`.toLowerCase();
      if (searchTerm && !rowStr.includes(searchTerm)) return;

      const tr = document.createElement('tr');
      if (!isMonthSelected(r['Bulan'])) {
        tr.style.opacity = '0.4';
      }

      tr.innerHTML = `
        <td>${r['No Bulan']}</td>
        <td style="font-weight: 600;">${r['Bulan']}</td>
        <td>${r['Tahun']}</td>
        <td>${r['Kantor']}</td>
        <td class="num">${formatIDR(r['Bea Masuk'])}</td>
        <td class="num">${formatIDR(r['Cukai Tembakau'])}</td>
        <td class="num">${formatIDR(r['Cukai MMEA'])}</td>
        <td class="num">${formatIDR(r['Cukai Lainnya'])}</td>
        <td class="num" style="color: #38bdf8;">${formatIDR(r['Cukai Total'])}</td>
        <td class="num" style="font-weight: 700; color: #f8fafc;">${formatIDR(r['Total Akhir'])}</td>
        <td class="num" style="color: #94a3b8;">${formatIDR(r['Total Kumulatif TLKV'])}</td>
      `;
      tbody3.appendChild(tr);
    });
  }

  // Search input event
  document.getElementById('search-input')?.addEventListener('input', () => {
    renderPage3(rawDataset);
  });

  // CSV Export logic
  document.getElementById('btn-export-csv')?.addEventListener('click', () => {
    let csv = 'No Bulan,Bulan,Tahun,Kantor,Bea Masuk,Cukai Tembakau,Cukai MMEA,Cukai Lainnya,Cukai Total,Total Akhir,Total Kumulatif TLKV\n';
    rawDataset.forEach(r => {
      if (isMonthSelected(r['Bulan'])) {
        csv += `${r['No Bulan']},${r['Bulan']},${r['Tahun']},${r['Kantor']},${r['Bea Masuk']},${r['Cukai Tembakau']},${r['Cukai MMEA']},${r['Cukai Lainnya']},${r['Cukai Total']},${r['Total Akhir']},${r['Total Kumulatif TLKV']}\n`;
      }
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'Rekap_Penerimaan_DJBC_BaliNusra_2026.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });

  // ==========================================
  // FITUR: UPLOAD FILE EXCEL / CSV BARU
  // ==========================================
  const fileInputExcel = document.getElementById('file-input-excel');
  const btnUploadFile = document.getElementById('btn-upload-file');
  const btnResetData = document.getElementById('btn-reset-data');

  function updateResetButtonVisibility() {
    if (!btnResetData) return;
    const hasCustomData = !!localStorage.getItem('djbc_custom_dataset');
    btnResetData.style.display = hasCustomData ? 'flex' : 'none';
  }
  updateResetButtonVisibility();

  if (btnUploadFile && fileInputExcel) {
    btnUploadFile.addEventListener('click', () => {
      fileInputExcel.value = ''; // Reset input agar bisa re-upload file sama
      fileInputExcel.click();
    });

    fileInputExcel.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      if (typeof XLSX === 'undefined') {
        alert('Library pembaca Excel (SheetJS) belum termuat. Periksa koneksi internet Anda.');
        return;
      }

      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Cari sheet 'DENPASAR_CLEAN' atau gunakan sheet pertama
        let targetSheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('DENPASAR')) || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[targetSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: 0 });

        if (!jsonData || jsonData.length === 0) {
          alert('File Excel kosong atau format tidak dikenali.');
          return;
        }

        // Helper normalisasi nama kolom
        function findVal(row, possibleKeys) {
          const rowKeys = Object.keys(row);
          for (let p of possibleKeys) {
            const foundKey = rowKeys.find(k => k.trim().toLowerCase() === p.trim().toLowerCase());
            if (foundKey !== undefined) {
              const val = row[foundKey];
              return typeof val === 'number' ? val : (parseFloat(String(val).replace(/[^0-9.-]+/g, '')) || 0);
            }
          }
          return 0;
        }

        function findStr(row, possibleKeys, fallback) {
          const rowKeys = Object.keys(row);
          for (let p of possibleKeys) {
            const foundKey = rowKeys.find(k => k.trim().toLowerCase() === p.trim().toLowerCase());
            if (foundKey !== undefined) {
              return String(row[foundKey]).trim();
            }
          }
          return fallback;
        }

        // Map data ke format standar
        const standardMonths = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'];
        
        const mappedData = jsonData.map((row, idx) => {
          const noBulan = findVal(row, ['No Bulan', 'No', 'Bulan Ke', 'No.']) || (idx + 1);
          let bulanName = findStr(row, ['Bulan', 'Month', 'Nama Bulan'], standardMonths[idx] || `BULAN ${idx + 1}`).toUpperCase();
          const tahun = findVal(row, ['Tahun', 'Year']) || 2026;
          const kantor = findStr(row, ['Kantor', 'KPPBC', 'KPBC'], 'DENPASAR').toUpperCase();

          const bm = findVal(row, ['Bea Masuk', 'BM', 'Bea_Masuk']);
          const ht = findVal(row, ['Cukai Tembakau', 'HT', 'Cukai_Tembakau', 'Hasil Tembakau']);
          const mmea = findVal(row, ['Cukai MMEA', 'MMEA', 'Cukai_MMEA']);
          const lain = findVal(row, ['Cukai Lainnya', 'Lainnya', 'Cukai_Lainnya']);
          const cTotal = findVal(row, ['Cukai Total', 'Total Cukai', 'Cukai_Total']) || (ht + mmea + lain);
          const tAkhir = findVal(row, ['Total Akhir', 'Total', 'Total_Akhir', 'Total Penerimaan']) || (bm + cTotal);
          const kumulatif = findVal(row, ['Total Kumulatif TLKV', 'Kumulatif', 'Kumulatif TLKV', 'Total Kumulatif']) || tAkhir;

          return {
            "No Bulan": Number(noBulan),
            "Bulan": bulanName,
            "Tahun": Number(tahun),
            "Kantor": kantor,
            "Bea Masuk": bm,
            "Cukai Tembakau": ht,
            "Cukai MMEA": mmea,
            "Cukai Lainnya": lain,
            "Cukai Total": cTotal,
            "Total Akhir": tAkhir,
            "Total Kumulatif TLKV": kumulatif
          };
        });

        // Simpan ke memory dan localStorage
        rawDataset = mappedData;
        localStorage.setItem('djbc_custom_dataset', JSON.stringify(mappedData));

        // Update footer info jika ada
        const footerSource = document.querySelector('.source-badge');
        if (footerSource) {
          footerSource.innerHTML = `
            <svg width="14" height="14" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
            Data: ${file.name}
          `;
        }

        updateResetButtonVisibility();
        state.activeCrossMonth = null;
        renderDashboard();

        alert(`✅ File "${file.name}" berhasil diunggah!\nTotal baris data: ${mappedData.length}\nDashboard telah diperbarui secara otomatis.`);

      } catch (err) {
        console.error('Error membaca file Excel:', err);
        alert('Terjadi kesalahan saat memproses file Excel: ' + err.message);
      }
    });
  }

  // Tombol Reset ke data bawaan awal
  if (btnResetData) {
    btnResetData.addEventListener('click', () => {
      if (confirm('Kembalikan dashboard ke data awal Rekap 2026? Data file upload akan dihapus dari browser.')) {
        localStorage.removeItem('djbc_custom_dataset');
        rawDataset = [...defaultDataset];
        updateResetButtonVisibility();

        const footerSource = document.querySelector('.source-badge');
        if (footerSource) {
          footerSource.innerHTML = `
            <svg width="14" height="14" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
            Data: Rekap_Penerimaan_2026.xlsx
          `;
        }

        state.activeCrossMonth = null;
        renderDashboard();
        alert('Dashboard telah dikembalikan ke data default.');
      }
    });
  }

  // Initial Render
  renderDashboard();
});
