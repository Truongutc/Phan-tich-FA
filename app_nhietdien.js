/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app_nhietdien.js (Nhiệt điện: POW, NT2, PPC, QTP)
   Dữ liệu thuần từ data/{TICKER}.json (không live-fetch Vietcap trong
   trình duyệt) — xem template_nhietdien.py cho pipeline tính toán.
   ════════════════════════════════════════════════════════ */

'use strict';

// ── Chart instances ──────────────────────────────────────
let chartRevenueGm = null;
let chartEbitNpat = null;
let chartFcf = null;
let chartValuation = null;
let chartPeerComps = null;

const CFG = { icon: '🔥', label: 'Nhiệt Điện', color: '#dc2626' };

const CHART_DEFAULTS = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { labels: { color: '#8892a4', font: { family: 'Inter', size: 11 }, boxWidth: 12 } },
        tooltip: { backgroundColor: 'rgba(10,16,32,0.95)', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1, titleColor: '#f0f2f8', bodyColor: '#8892a4', padding: 12, cornerRadius: 8 },
    },
    scales: {
        x: { ticks: { color: '#545f74', font: { size: 10 }, maxRotation: 45 }, grid: { color: 'rgba(255,255,255,0.04)' } },
        y: { ticks: { color: '#545f74', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,0.04)' } }
    }
};

// ═══════════════════════════════════════════════════════════
// ROUTING / NAVIGATION
// ═══════════════════════════════════════════════════════════
function goHome() { window.location.href = 'index.html'; }

function showAnalysisView(ticker) {
    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';
    document.getElementById('nav-breadcrumb').innerHTML = `
        <span class="breadcrumb-home" style="cursor:pointer" onclick="goHome()">Tổng quan</span>
        <span class="breadcrumb-sep">›</span>
        <span class="breadcrumb-sector" style="color:${CFG.color}">${CFG.icon} ${CFG.label}</span>
        <span class="breadcrumb-sep">›</span>
        <span class="breadcrumb-ticker">${ticker}</span>
    `;
    const badge = document.getElementById('sector-badge-nav');
    badge.style.display = 'flex';
    badge.style.borderColor = CFG.color + '50';
    badge.style.color = CFG.color;
    document.getElementById('back-btn').style.display = 'inline-flex';
    const ribbon = document.getElementById('sector-ribbon');
    ribbon.textContent = CFG.label.toUpperCase();
    ribbon.style.background = CFG.color;
    document.documentElement.style.setProperty('--sector-color', CFG.color);
}

document.addEventListener('DOMContentLoaded', async () => {
    const params = new URLSearchParams(location.search);
    const autoTicker = params.get('ticker');
    if (autoTicker) {
        await loadStockDashboard(autoTicker);
    } else {
        await loadOverview();
    }
});

// ═══════════════════════════════════════════════════════════
// OVERVIEW (stock list)
// ═══════════════════════════════════════════════════════════
async function loadOverview() {
    try {
        const resp = await fetch('data/index.json');
        const stocks = await resp.json();
        renderOverview(stocks);
    } catch {
        document.getElementById('sector-groups-container').innerHTML =
            '<div class="loading-state card">Chưa có dữ liệu. Hãy chạy phân tích ít nhất một mã Nhiệt điện.</div>';
    }
}

function isNhietDienSector(sectorStr) {
    const s = (sectorStr || '').toLowerCase();
    return s.includes('nhiệt điện') || s.includes('nhiet dien') || s.includes('thermal power');
}

function renderOverview(stocks) {
    const ndStocks = (stocks || []).filter(s => isNhietDienSector(s.sector));
    window._allStocks = ndStocks;
    _groupAndRender(ndStocks);

    const inp = document.getElementById('stock-search-input');
    if (inp) {
        inp.addEventListener('input', () => {
            const q = inp.value.trim().toLowerCase();
            document.getElementById('search-clear').style.opacity = q ? '1' : '0';
            if (!q) { _groupAndRender(window._allStocks); return; }
            const filtered = window._allStocks.filter(s =>
                s.ticker.toLowerCase().includes(q) || (s.companyName || '').toLowerCase().includes(q));
            _renderSearchResults(filtered, q);
        });
    }
}

function clearSearch() {
    const inp = document.getElementById('stock-search-input');
    if (inp) inp.value = '';
    document.getElementById('search-clear').style.opacity = '0';
    if (window._allStocks) _groupAndRender(window._allStocks);
}

function _renderSearchResults(filtered, query) {
    const container = document.getElementById('sector-groups-container');
    if (filtered.length === 0) {
        container.innerHTML = `<div class="loading-state card">Không tìm thấy mã nào khớp với "${query}"</div>`;
        return;
    }
    container.innerHTML = '';
    const card = document.createElement('div');
    card.className = 'card sector-group';
    card.innerHTML = `<div class="sector-group-header"><h3 style="color:var(--text-muted)">🔍 Kết quả tìm kiếm</h3>
        <span class="sector-count-badge">${filtered.length} mã</span></div><div class="stocks-grid" id="grid-search"></div>`;
    container.appendChild(card);
    const grid = card.querySelector('#grid-search');
    filtered.forEach(s => grid.appendChild(_makeStockCard(s)));
}

function _groupAndRender(stocks) {
    const container = document.getElementById('sector-groups-container');
    container.innerHTML = '';
    if (!stocks || stocks.length === 0) {
        container.innerHTML = '<div class="loading-state card">Chưa có mã Nhiệt điện nào được phân tích.</div>';
        return;
    }
    const groupDiv = document.createElement('div');
    groupDiv.className = 'card sector-group';
    groupDiv.innerHTML = `
        <div class="sector-group-header">
            <div class="sector-group-dot" style="background:${CFG.color}"></div>
            <h3 style="color:${CFG.color}">${CFG.icon} ${CFG.label}</h3>
            <span class="sector-count-badge">${stocks.length} mã</span>
        </div>
        <div class="stocks-grid" id="grid-nd"></div>
    `;
    container.appendChild(groupDiv);
    const grid = groupDiv.querySelector('#grid-nd');
    stocks.forEach(s => grid.appendChild(_makeStockCard(s)));
}

function _makeStockCard(s) {
    const card = document.createElement('div');
    card.className = 'stock-card';
    card.style.setProperty('--sector-color', CFG.color);
    card.innerHTML = `<div class="stock-card-ticker" style="color:${CFG.color}">${s.ticker}</div>`;
    card.onclick = () => loadStockDashboard(s.ticker);
    return card;
}

// ═══════════════════════════════════════════════════════════
// LOAD STOCK DASHBOARD
// ═══════════════════════════════════════════════════════════
async function loadStockDashboard(ticker) {
    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';
    document.getElementById('ticker-badge').textContent = ticker;
    document.getElementById('company-name').textContent = 'Đang tải dữ liệu...';
    destroyCharts();

    const localJson = await fetch(`data/${ticker}.json`).then(r => r.ok ? r.json() : null).catch(() => null);
    if (!localJson) {
        document.getElementById('company-name').textContent = `Không tìm thấy dữ liệu cho ${ticker}`;
        showAnalysisView(ticker);
        return;
    }

    showAnalysisView(ticker);

    const companyName = localJson.companyName || ticker;
    const currentPrice = localJson.currentPrice || 0;

    document.getElementById('company-name').textContent = companyName;
    document.getElementById('company-sector').textContent = `${CFG.icon} ${CFG.label}`;

    const btnPdf = document.getElementById('download-pdf');
    const btnExcel = document.getElementById('download-excel');
    if (localJson.gdrivePdfUrl) { btnPdf.href = localJson.gdrivePdfUrl; btnPdf.classList.remove('hidden'); } else btnPdf.classList.add('hidden');
    if (localJson.gdriveExcelUrl) { btnExcel.href = localJson.gdriveExcelUrl; btnExcel.classList.remove('hidden'); } else btnExcel.classList.add('hidden');

    renderValuationSnapshot(localJson, currentPrice);
    renderFinancialSnapshotTable(localJson);
    renderThesisAndRisks(localJson.thesis, localJson.risks);
    renderMoatScorecard(localJson.moats);
    renderPESTLE(localJson.pestle);
    renderFuelPrices(localJson.fuelPrices);
    renderPeerCompsTable(localJson.peerComps, ticker);

    // Charts
    renderRevenueGmChart(localJson);
    renderEbitNpatChart(localJson);
    renderFcfChart(localJson);
    renderValuationChart(localJson, currentPrice);
    renderPeerCompsChart(localJson.peerComps, ticker);
}

function destroyCharts() {
    [chartRevenueGm, chartEbitNpat, chartFcf, chartValuation, chartPeerComps].forEach(c => { if (c) c.destroy(); });
    chartRevenueGm = chartEbitNpat = chartFcf = chartValuation = chartPeerComps = null;
}

function formatNumber(num, decimals = 0) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Number(num).toLocaleString('vi-VN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// ═══════════════════════════════════════════════════════════
// VALUATION SNAPSHOT
// ═══════════════════════════════════════════════════════════
function renderValuationSnapshot(localJson, currPrice) {
    const val = localJson.valuation || {};
    const fmt = (n) => n ? Math.round(n).toLocaleString('vi-VN') + ' đ' : '-';
    const upside = (val.upside || 0) * 100;
    const recText = upside >= 15 ? 'MUA' : upside <= -15 ? 'BÁN' : 'THEO DÕI';
    const recColor = recText === 'MUA' ? '#10b981' : recText === 'BÁN' ? '#ef4444' : '#f59e0b';
    const el = (id) => document.getElementById(id);

    if (el('snap-curr-price')) el('snap-curr-price').textContent = fmt(currPrice);
    if (el('snap-target-price')) el('snap-target-price').textContent = fmt(val.fair_blend);
    if (el('snap-upside')) {
        el('snap-upside').textContent = (upside >= 0 ? '+' : '') + upside.toFixed(1) + '%';
        el('snap-upside').style.color = upside >= 10 ? '#10b981' : upside >= 0 ? '#f59e0b' : '#ef4444';
    }
    if (el('snap-recommend')) { el('snap-recommend').textContent = recText; el('snap-recommend').style.background = recColor; }
    if (el('snap-dcf')) el('snap-dcf').textContent = fmt(val.fair_dcf);
    if (el('snap-evebitda')) el('snap-evebitda').textContent = fmt(val.fair_ev_ebitda) + (val.target_ev_ebitda ? ` (${val.target_ev_ebitda}x)` : '');
    if (el('snap-pb')) el('snap-pb').textContent = fmt(val.fair_pb) + (val.target_pb ? ` (${val.target_pb}x)` : '');
    if (el('snap-asset')) el('snap-asset').textContent = fmt(val.fair_asset);
    if (el('snap-pe')) el('snap-pe').textContent = fmt(val.fair_pe) + (val.target_pe ? ` (${val.target_pe}x)` : '');
    if (el('snap-wacc')) el('snap-wacc').textContent = (val.wacc != null ? (val.wacc * 100).toFixed(2) : '-') + '%';
    if (el('snap-coe')) el('snap-coe').textContent = (val.coe != null ? (val.coe * 100).toFixed(2) : '-') + '%';
    if (el('snap-beta')) el('snap-beta').textContent = (val.beta != null ? val.beta.toFixed(2) : '-');
}

// ═══════════════════════════════════════════════════════════
// FINANCIAL SNAPSHOT TABLE
// ═══════════════════════════════════════════════════════════
function renderFinancialSnapshotTable(localJson) {
    const data = localJson.data;
    if (!data || !data.years) return;
    const headerRow = document.getElementById('fs-header-row');
    const tbody = document.getElementById('web-financial-snapshot-tbody');
    headerRow.innerHTML = '<th>Chỉ tiêu (tỷ VND)</th>' + data.years.map(y => `<th>${y}</th>`).join('');
    const gm = (data.revenue || []).map((r, i) => (r ? (data.grossProfit[i] / r * 100) : null));
    const npm = (data.revenue || []).map((r, i) => (r ? (data.npat[i] / r * 100) : null));
    const rows = [
        { label: 'Doanh thu thuần', arr: data.revenue },
        { label: 'Lợi nhuận gộp', arr: data.grossProfit },
        { label: 'Biên LNG (%)', arr: gm, isPct: true },
        { label: 'EBIT', arr: data.ebit },
        { label: 'LNST cổ đông mẹ', arr: data.npat, highlight: true },
        { label: 'Biên LNST (%)', arr: npm, isPct: true },
    ];
    tbody.innerHTML = rows.map(r => `<tr><td><strong>${r.label}</strong></td>${
        (r.arr || []).map(v => `<td class="${r.highlight ? 'highlight' : ''}">${v == null ? '-' : (r.isPct ? v.toFixed(1) + '%' : formatNumber(v))}</td>`).join('')
    }</tr>`).join('');
}

// ═══════════════════════════════════════════════════════════
// FUEL PRICES
// ═══════════════════════════════════════════════════════════
function renderFuelPrices(fuelPrices) {
    const grid = document.getElementById('fuel-prices-grid');
    if (!fuelPrices || !grid) return;
    const items = [
        { label: 'Than nhiệt (Newcastle)', unit: 'USD/tấn', data: fuelPrices.coal },
        { label: 'Khí tự nhiên (Henry Hub)', unit: 'USD/MMBtu', data: fuelPrices.gas },
        { label: 'Dầu thô (Brent)', unit: 'USD/thùng', data: fuelPrices.oil },
        { label: 'Tỷ giá USD/VND', unit: 'VND/USD', data: fuelPrices.usdvnd },
    ];
    grid.innerHTML = items.map(it => `
        <div class="bank-kpi-item">
            <span class="bank-kpi-label">${it.label}</span>
            <span class="bank-kpi-value">${it.data ? formatNumber(it.data.value, it.data.value < 20 ? 2 : 0) : '-'}</span>
            <span style="font-size:0.7rem;color:var(--text-muted)">${it.unit}${it.data && it.data.source ? ' · ' + it.data.source : ''}</span>
        </div>
    `).join('');
}

// ═══════════════════════════════════════════════════════════
// PEER COMPS TABLE
// ═══════════════════════════════════════════════════════════
function renderPeerCompsTable(peerComps, currentTicker) {
    const tbody = document.getElementById('peer-comps-tbody');
    if (!tbody) return;
    if (!peerComps || !peerComps.perTicker) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center">Chưa có dữ liệu peer comps.</td></tr>';
        return;
    }
    const rows = Object.entries(peerComps.perTicker).map(([t, v]) => {
        const mark = t === currentTicker ? `<strong style="color:${CFG.color}">${t} ★</strong>` : t;
        return `<tr>
            <td>${mark}</td>
            <td>${v.ev_ebitda ? v.ev_ebitda.toFixed(2) + 'x' : '-'}</td>
            <td>${v.pb ? v.pb.toFixed(2) + 'x' : '-'}</td>
            <td>${v.pe ? v.pe.toFixed(2) + 'x' : '-'}</td>
            <td>${formatNumber(v.current_price)}</td>
            <td>${formatNumber(v.market_cap)}</td>
        </tr>`;
    });
    tbody.innerHTML = rows.join('') || '<tr><td colspan="6" style="text-align:center">Không có dữ liệu.</td></tr>';
}

// ═══════════════════════════════════════════════════════════
// QUALITATIVE RENDERERS
// ═══════════════════════════════════════════════════════════
function renderThesisAndRisks(thesisList, risksList) {
    document.getElementById('thesis-list').innerHTML = (thesisList || []).map(t => `<li>${t}</li>`).join('') || '<li>Chưa có dữ liệu.</li>';
    document.getElementById('risks-list').innerHTML = (risksList || []).map(r => `<li>${r}</li>`).join('') || '<li>Chưa có dữ liệu.</li>';
}

function renderMoatScorecard(moats) {
    const container = document.getElementById('moat-scorecard-list');
    if (!moats) { container.innerHTML = ''; return; }
    container.innerHTML = '';
    Object.entries(moats).forEach(([key, item]) => {
        const pct = (item.score / 5) * 100;
        const div = document.createElement('div');
        div.className = 'moat-item';
        div.innerHTML = `
            <div class="moat-header"><span>${key}</span><span class="moat-score">${item.score}/5</span></div>
            <div class="moat-bar-bg"><div class="moat-bar-fill" style="width:0%;background:linear-gradient(90deg,${CFG.color}80,${CFG.color})"></div></div>
            <div class="moat-desc">${item.desc}</div>
        `;
        container.appendChild(div);
        setTimeout(() => div.querySelector('.moat-bar-fill').style.width = pct + '%', 50);
    });
}

function renderPESTLE(pestle) {
    const container = document.getElementById('pestle-grid-container');
    if (!pestle || !container) { if (container) container.innerHTML = ''; return; }
    const pestleColors = { Political: '#3b82f6', Economic: '#10b981', Social: '#f59e0b', Technological: '#a855f7', Legal: '#ef4444', Environmental: '#22c55e' };
    container.innerHTML = Object.entries(pestle).map(([key, text]) => `
        <div class="pestle-box">
            <div class="pestle-box-header">
                <span>${key}</span>
                <span class="pestle-tag" style="background:${(pestleColors[key] || '#64748b') + '25'};color:${pestleColors[key] || '#64748b'}">${key[0]}</span>
            </div>
            <div class="pestle-box-text">${text}</div>
        </div>
    `).join('');
}

// ═══════════════════════════════════════════════════════════
// CHARTS
// ═══════════════════════════════════════════════════════════
function renderRevenueGmChart(localJson) {
    const data = localJson.data;
    if (!data || !data.years) return;
    const ctx = document.getElementById('revenueGmChart');
    const gm = (data.revenue || []).map((r, i) => (r ? (data.grossProfit[i] / r * 100) : null));
    chartRevenueGm = new Chart(ctx, {
        data: {
            labels: data.years,
            datasets: [
                { type: 'bar', label: 'Doanh thu (tỷ VND)', data: data.revenue, backgroundColor: CFG.color + 'aa', yAxisID: 'y' },
                { type: 'line', label: 'Biên LNG (%)', data: gm, borderColor: '#60a5fa', backgroundColor: '#60a5fa', tension: 0.3, yAxisID: 'y1' },
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: CHART_DEFAULTS.scales.x,
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'tỷ VND', color: '#8892a4' } },
                y1: { position: 'right', ticks: { color: '#60a5fa' }, grid: { display: false }, title: { display: true, text: '%', color: '#60a5fa' } },
            }
        }
    });
}

function renderEbitNpatChart(localJson) {
    const data = localJson.data;
    if (!data || !data.years) return;
    const ctx = document.getElementById('ebitNpatChart');
    chartEbitNpat = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.years,
            datasets: [
                { label: 'EBIT', data: data.ebit, backgroundColor: '#f59e0b' },
                { label: 'LNST cổ đông mẹ', data: data.npat, backgroundColor: '#10b981' },
            ]
        },
        options: CHART_DEFAULTS
    });
}

function renderFcfChart(localJson) {
    const rows = localJson.fcRows;
    if (!rows || !rows.length) return;
    const ctx = document.getElementById('fcfChart');
    chartFcf = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: rows.map(r => `Năm +${r.year_offset}`),
            datasets: [
                { label: 'FCFF (tỷ VND)', data: rows.map(r => r.fcff), backgroundColor: '#3b82f6' },
                { label: 'PV(FCFF)', data: rows.map(r => r.fcff_pv), backgroundColor: '#8b5cf6' },
            ]
        },
        options: CHART_DEFAULTS
    });
}

function renderValuationChart(localJson, currPrice) {
    const val = localJson.valuation || {};
    const ctx = document.getElementById('valuationChart');
    const labels = ['DCF', 'EV/EBITDA', 'P/B', 'Asset', 'Blend', 'Giá TT'];
    const values = [val.fair_dcf, val.fair_ev_ebitda, val.fair_pb, val.fair_asset, val.fair_blend, currPrice];
    chartValuation = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'VND/cp',
                data: values,
                backgroundColor: labels.map(l => l === 'Blend' ? '#10b981' : l === 'Giá TT' ? '#64748b' : CFG.color + '99'),
            }]
        },
        options: { ...CHART_DEFAULTS, plugins: { ...CHART_DEFAULTS.plugins, legend: { display: false } } }
    });
}

function renderPeerCompsChart(peerComps, currentTicker) {
    if (!peerComps || !peerComps.perTicker) return;
    const ctx = document.getElementById('peerCompsChart');
    const tickers = Object.keys(peerComps.perTicker);
    const evEbitda = tickers.map(t => peerComps.perTicker[t].ev_ebitda);
    const pb = tickers.map(t => peerComps.perTicker[t].pb);
    const pe = tickers.map(t => peerComps.perTicker[t].pe);
    chartPeerComps = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: tickers,
            datasets: [
                { label: 'EV/EBITDA (x)', data: evEbitda, backgroundColor: '#f59e0b' },
                { label: 'P/B (x)', data: pb, backgroundColor: '#3b82f6' },
                { label: 'P/E (x)', data: pe, backgroundColor: '#10b981' },
            ]
        },
        options: CHART_DEFAULTS
    });
}
