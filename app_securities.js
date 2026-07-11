/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app_securities.js (Chứng khoán / CTCK)
   Dữ liệu thuần từ data/{TICKER}.json + data/peer_benchmark_securities.json
   (không live-fetch Vietcap trong trình duyệt — đơn giản & ổn định hơn)
   ════════════════════════════════════════════════════════ */

'use strict';

// ── Chart instances ──────────────────────────────────────
let chartSegmentRevenue = null;
let chartSegmentMix = null;
let chartRevNpat = null;
let chartQuarterlyNpatRoe = null;
let chartRoePbCorrelation = null;
let chartMarginLeverage = null;
let chartMarketShare = null;
let chartPE = null;
let chartPB = null;
let chartFvtplComp = null;
let chartFvtplGainLoss = null;
let chartFvtplHoldings = null;
let chartFvtplHoldingsQuarterly = null;
let chartSegmentRevenueQuarterly = null;
let chartSegmentGrossProfit = null;
let chartSegmentGrossProfitQuarterly = null;

if (typeof ChartDataLabels !== 'undefined') { Chart.register(ChartDataLabels); }

const SEGMENT_COLORS = {
    MoiGioi: '#3b82f6', Margin: '#f59e0b', TuDoanh: '#10b981', IB_LuuKy: '#8b5cf6', QLQ: '#ec4899',
};
const SEGMENT_LABELS_FALLBACK = {
    MoiGioi: 'Môi giới', Margin: 'Cho vay Margin', TuDoanh: 'Tự doanh (FVTPL+AFS)', IB_LuuKy: 'IB + Lưu ký', QLQ: 'Quản lý quỹ',
};
const CFG = { icon: '📈', label: 'Chứng Khoán', color: '#3b82f6' };

const CHART_DEFAULTS = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { labels: { color: '#8892a4', font: { family: 'Inter', size: 11 }, boxWidth: 12 } },
        tooltip: { backgroundColor: 'rgba(10,16,32,0.95)', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1, titleColor: '#f0f2f8', bodyColor: '#8892a4', padding: 12, cornerRadius: 8 },
        datalabels: { display: false } // Mặc định tắt để tránh rối các line chart
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
            '<div class="loading-state card">Chưa có dữ liệu. Hãy chạy phân tích ít nhất một mã CTCK.</div>';
    }
}

function isSecuritiesSector(sectorStr) {
    const s = (sectorStr || '').toLowerCase();
    return s.includes('chứng khoán') || s.includes('chung khoan') || s.includes('securities');
}

function renderOverview(stocks) {
    const ctckStocks = (stocks || []).filter(s => isSecuritiesSector(s.sector));
    window._allStocks = ctckStocks;
    _groupAndRender(ctckStocks);

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
        container.innerHTML = '<div class="loading-state card">Chưa có mã CTCK nào được phân tích.</div>';
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
        <div class="stocks-grid" id="grid-ctck"></div>
    `;
    container.appendChild(groupDiv);
    const grid = groupDiv.querySelector('#grid-ctck');
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
    renderSegmentMixGrid(localJson);
    renderQuarterlyEvaluation(localJson);
    renderFinancialSnapshotTable(localJson);
    renderThesisAndRisks(localJson.thesis, localJson.risks);
    renderMoatScorecard(localJson.moats);
    renderPESTLE(localJson.pestle);
    renderMacroLiquidity(localJson);
    renderCommentary(localJson.comments);
    renderSegmentPctGrid(localJson);
    loadPeerBenchmark(ticker);
    loadFvtplHoldings(ticker);

    // Charts
    renderSegmentRevenueChart(localJson);
    renderSegmentMixChart(localJson);
    renderSegmentRevenueQuarterlyChart(localJson);
    renderSegmentGrossProfitChart(localJson);
    renderSegmentGrossProfitQuarterlyChart(localJson);
    renderRevNpatChart(localJson);
    renderQuarterlyNpatRoeChart(localJson);
    renderRoePbCorrelationChart(localJson);
    renderMarginLeverageChart(localJson);
    renderMarketShareChart(ticker);
    renderPEChart(localJson);
    renderPBChart(localJson);
}

function destroyCharts() {
    [chartSegmentRevenue, chartSegmentMix, chartRevNpat, chartQuarterlyNpatRoe,
     chartRoePbCorrelation, chartMarginLeverage, chartMarketShare, chartPE, chartPB,
     chartFvtplComp, chartFvtplGainLoss, chartFvtplHoldings, chartFvtplHoldingsQuarterly,
     chartSegmentRevenueQuarterly, chartSegmentGrossProfit, chartSegmentGrossProfitQuarterly].forEach(c => { if (c) c.destroy(); });
    chartSegmentRevenue = chartSegmentMix = chartRevNpat = chartQuarterlyNpatRoe =
    chartRoePbCorrelation = chartMarginLeverage = chartMarketShare = chartPE = chartPB =
    chartFvtplComp = chartFvtplGainLoss = chartFvtplHoldings = chartFvtplHoldingsQuarterly =
    chartSegmentRevenueQuarterly = chartSegmentGrossProfit = chartSegmentGrossProfitQuarterly = null;
}

function calcMedian(arr) {
    const valid = (arr || []).filter(v => v !== null && v !== undefined && !isNaN(v)).sort((a, b) => a - b);
    if (!valid.length) return null;
    const m = Math.floor(valid.length / 2);
    return valid.length % 2 !== 0 ? valid[m] : (valid[m - 1] + valid[m]) / 2;
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
    const target = val.base;
    const upside = val.upsidePct ?? 0;
    const recText = val.recommendation || 'THEO DÕI';
    const recColor = recText === 'MUA' ? '#10b981' : recText === 'BÁN' ? '#ef4444' : '#f59e0b';
    const el = (id) => document.getElementById(id);
    if (el('snap-curr-price')) el('snap-curr-price').textContent = fmt(currPrice);
    if (el('snap-target-price')) el('snap-target-price').textContent = fmt(target);
    if (el('snap-upside')) {
        el('snap-upside').textContent = (upside >= 0 ? '+' : '') + upside.toFixed(1) + '%';
        el('snap-upside').style.color = upside >= 10 ? '#10b981' : upside >= 0 ? '#f59e0b' : '#ef4444';
    }
    if (el('snap-recommend')) { el('snap-recommend').textContent = recText; el('snap-recommend').style.background = recColor; }
    if (el('snap-coe')) el('snap-coe').textContent = (val.coe != null ? val.coe.toFixed(2) : '-') + '%';
    if (el('snap-pb-target')) el('snap-pb-target').textContent = val.methods?.pb ? fmt(val.methods.pb) : '-';
    if (el('snap-pe-target')) el('snap-pe-target').textContent = val.methods?.pe ? fmt(val.methods.pe) : '-';
    if (el('snap-bear')) el('snap-bear').textContent = val.bear ? fmt(val.bear) : '-';
    if (el('snap-bull')) el('snap-bull').textContent = val.bull ? fmt(val.bull) : '-';
    if (el('snap-pb-lower')) el('snap-pb-lower').textContent = val.lowerTarget ? fmt(val.lowerTarget) : '-';
    if (el('snap-pb-upper')) el('snap-pb-upper').textContent = val.upperTarget ? fmt(val.upperTarget) : '-';
    const lev = localJson.marginLeverage || {};
    if (el('snap-margin-lev')) {
        el('snap-margin-lev').textContent = (lev.latest != null ? lev.latest.toFixed(2) : '-') + 'x';
        el('snap-margin-lev').style.color = lev.latest >= 2.0 ? '#ef4444' : lev.latest >= 1.8 ? '#f59e0b' : '#10b981';
    }

    document.getElementById('val-bear').textContent = val.bear ? formatNumber(val.bear) + ' VND' : '-';
    document.getElementById('val-base').textContent = val.base ? formatNumber(val.base) + ' VND' : '-';
    document.getElementById('val-bull').textContent = val.bull ? formatNumber(val.bull) + ' VND' : '-';
}

// ═══════════════════════════════════════════════════════════
// SEGMENT MIX GRID (metrics-grid at top)
// ═══════════════════════════════════════════════════════════
function renderSegmentMixGrid(localJson) {
    const seg = localJson.segments;
    const grid = document.getElementById('segment-mix-grid');
    if (!seg || !grid) return;
    const names = seg.names || [];
    grid.innerHTML = names.map(n => {
        const pct = (seg.pctNow?.[n] || 0) * 100;
        const label = (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n;
        return `<div class="metric-card">
            <div class="metric-label">${label}</div>
            <div class="metric-value" style="color:${SEGMENT_COLORS[n] || '#64748b'}">${pct.toFixed(1)}%</div>
            <div class="metric-sub">% Tổng doanh thu dự phóng</div>
        </div>`;
    }).join('');
}

function renderSegmentPctGrid(localJson) {
    const seg = localJson.segments;
    const grid = document.getElementById('segment-pct-grid');
    if (!seg || !grid) return;
    const names = seg.names || [];
    grid.innerHTML = names.map(n => {
        const pct = (seg.pctNow?.[n] || 0) * 100;
        const label = (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n;
        return `<div class="bank-kpi-item"><span class="bank-kpi-label">${label}</span><span class="bank-kpi-value">${pct.toFixed(1)}%</span></div>`;
    }).join('');
}

// ═══════════════════════════════════════════════════════════
// KQKD QUARTERLY EVALUATION CARD
// ═══════════════════════════════════════════════════════════
function renderQuarterlyEvaluation(localJson) {
    const container = document.getElementById('quarterly-ytd-evaluation-container');
    const qu = localJson.quarterlyUpdate;
    if (!container) return;
    if (!qu || !qu.quarter_label) {
        container.innerHTML = '<div class="loading-state">Chưa có dữ liệu quý gần nhất.</div>';
        return;
    }
    const yoyRevColor = qu.yoy_rev_pct >= 0 ? '#10b981' : '#ef4444';
    const yoyNpatColor = qu.yoy_npat_pct >= 0 ? '#10b981' : '#ef4444';
    const yoyRevTxt = qu.yoy_rev_pct != null ? `${qu.yoy_rev_pct >= 0 ? '+' : ''}${qu.yoy_rev_pct}% YoY` : 'không có dữ liệu cùng kỳ';
    const yoyNpatTxt = qu.yoy_npat_pct != null ? `${qu.yoy_npat_pct >= 0 ? '+' : ''}${qu.yoy_npat_pct}% YoY` : 'không có dữ liệu cùng kỳ';

    let cumHtml = '';
    if (qu.n_known_q > 0) {
        cumHtml = `
            <div class="metric-card">
                <div class="metric-label">Lũy kế ${qu.n_known_q}/4 quý — Doanh thu</div>
                <div class="metric-value">${formatNumber(qu.cum_rev)}</div>
                <div class="metric-sub">tỷ VND · ${qu.pct_of_annual_est_rev != null ? qu.pct_of_annual_est_rev + '% ước tính cả năm' : ''}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Lũy kế ${qu.n_known_q}/4 quý — LNST</div>
                <div class="metric-value">${formatNumber(qu.cum_npat)}</div>
                <div class="metric-sub">tỷ VND · ${qu.pct_of_annual_est_npat != null ? qu.pct_of_annual_est_npat + '% ước tính cả năm' : ''}</div>
            </div>`;
    }

    container.innerHTML = `
        <div class="metrics-grid" style="margin-top:0">
            <div class="metric-card highlight">
                <div class="metric-label">Doanh thu ${qu.quarter_label}</div>
                <div class="metric-value">${formatNumber(qu.rev)}</div>
                <div class="metric-sub" style="color:${yoyRevColor}">${yoyRevTxt}</div>
            </div>
            <div class="metric-card highlight">
                <div class="metric-label">LNST ${qu.quarter_label}</div>
                <div class="metric-value">${formatNumber(qu.npat)}</div>
                <div class="metric-sub" style="color:${yoyNpatColor}">${yoyNpatTxt}</div>
            </div>
            ${cumHtml}
        </div>
        <p style="margin-top:12px;font-size:0.85em;color:var(--text-muted)">
            ${qu.n_known_q > 0
                ? `Phần còn lại của năm ${qu.cur_fc_year} (${4 - qu.n_known_q} quý) được ước tính theo giả định bottom-up — xem chi tiết công thức blend trong Excel Model (sheet 04b_Dien_Bien_Quy).`
                : `Chưa có quý nào của năm ${qu.cur_fc_year} được công bố — ước tính cả năm hiện dựa hoàn toàn trên giả định bottom-up.`}
        </p>
    `;
}

// ═══════════════════════════════════════════════════════════
// FINANCIAL SNAPSHOT TABLE
// ═══════════════════════════════════════════════════════════
function renderFinancialSnapshotTable(localJson) {
    const data = localJson.data;
    if (!data || !data.years) return;
    const headerRow = document.getElementById('fs-header-row');
    const tbody = document.getElementById('web-financial-snapshot-tbody');
    const nHist = data.years.length - (localJson.quarterlyUpdate ? 3 : 0);
    headerRow.innerHTML = '<th>Chỉ tiêu (tỷ VND)</th>' + data.years.map((y, i) =>
        `<th${i >= data.years.length - 3 ? ' style="color:#60a5fa"' : ''}>${y}${i >= data.years.length - 3 ? 'E' : 'A'}</th>`).join('');
    const rows = [
        { label: 'Doanh thu hoạt động', arr: data.revenue },
        { label: 'LNST cổ đông mẹ', arr: data.npat, highlight: true },
        { label: 'EPS (VND)', arr: data.eps },
        { label: 'VCSH (tỷ VND)', arr: data.equity },
    ];
    tbody.innerHTML = rows.map(r => `<tr><td><strong>${r.label}</strong></td>${
        (r.arr || []).map(v => `<td class="${r.highlight ? 'highlight' : ''}">${formatNumber(v)}</td>`).join('')
    }</tr>`).join('');
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

function renderPESTLE(pestleData) {
    let data = {};
    (pestleData || []).forEach(item => { if (item.factor) data[item.factor] = item.content; });
    const pestleColors = { Political: '#3b82f6', Economic: '#10b981', Social: '#f59e0b', Technological: '#a855f7', Legal: '#ef4444', Environmental: '#22c55e' };
    const container = document.getElementById('pestle-grid-container');
    container.innerHTML = Object.entries(data).map(([key, text]) => `
        <div class="pestle-box">
            <div class="pestle-box-header">
                <span>${key}</span>
                <span class="pestle-tag" style="background:${(pestleColors[key] || '#64748b') + '25'};color:${pestleColors[key] || '#64748b'}">${key[0]}</span>
            </div>
            <div class="pestle-box-text">${text}</div>
        </div>
    `).join('');
}

function renderCommentary(comments) {
    const data = comments || {};
    document.getElementById('comment-business').textContent = data.businessModel || 'Chưa có đánh giá mô hình kinh doanh.';
    document.getElementById('comment-financial').textContent = data.financialPerformance || 'Chưa có đánh giá sức khỏe tài chính.';
    document.getElementById('comment-valuation').textContent = data.valuationText || 'Chưa có đánh giá định giá.';
}

// ═══════════════════════════════════════════════════════════
// MACRO & THANH KHOẢN (từ macro_liquidity trong JSON)
// ═══════════════════════════════════════════════════════════
function renderMacroLiquidity(localJson) {
    const container = document.getElementById('macro-liquidity-container');
    if (!container) return;
    const ml = localJson.macro_liquidity;
    if (!ml || Object.keys(ml).length === 0) {
        container.innerHTML = '<div class="loading-state">Chưa có dữ liệu vĩ mô & thanh khoản.</div>';
        return;
    }

    // ADTV lịch sử
    const adtvHist = ml.adtv_hist || {};
    const adtvFc = ml.adtv_fc || {};
    const allAdtvYears = Object.keys(adtvHist).sort();
    const allFcYears = Object.keys(adtvFc).sort();
    const fmtAdtv = v => v != null ? Math.round(v).toLocaleString('vi-VN') + ' tỷ' : '-';

    // FVTPL mix
    const mix = ml.fvtpl_mix || {};
    const rates = ml.fvtpl_rates || {};
    const expectedYield = ml.fvtpl_expected_yield;
    const marketShare = ml.market_share;

    // taper
    const taper = (ml.adtv_growth_taper || []).map(g => (g * 100).toFixed(0) + '%').join(' → ');

    container.innerHTML = `
        <div class="macro-liq-section">
            <h4 class="macro-liq-title">📊 Giá trị giao dịch bình quân/phiên HOSE (tỷ VND)</h4>
            <div class="macro-adtv-grid">
                ${allAdtvYears.map(y => `
                    <div class="macro-adtv-item">
                        <div class="macro-adtv-year">${y}A</div>
                        <div class="macro-adtv-val">${fmtAdtv(adtvHist[y])}</div>
                    </div>`).join('')}
                ${allFcYears.map((y, i) => `
                    <div class="macro-adtv-item forecast">
                        <div class="macro-adtv-year">${y}E</div>
                        <div class="macro-adtv-val">${fmtAdtv(adtvFc[y])}</div>
                    </div>`).join('')}
            </div>
            <div class="macro-liq-note">Tăng trưởng dự phóng (taper): <strong>${taper}</strong> — giảm dần theo năm xa</div>
        </div>

        <div class="macro-liq-section">
            <h4 class="macro-liq-title">🏦 Cơ cấu danh mục Tự doanh (FVTPL)</h4>
            <div class="macro-fvtpl-grid">
                <div class="macro-fvtpl-item">
                    <div class="macro-fvtpl-label">CDs (Chứng chỉ tiền gửi)</div>
                    <div class="macro-fvtpl-val">${mix.CDs != null ? (mix.CDs * 100).toFixed(0) + '%' : '-'}</div>
                    <div class="macro-fvtpl-rate">R = ${rates.R_CDs != null ? (rates.R_CDs * 100).toFixed(1) + '%' : '-'}/năm</div>
                </div>
                <div class="macro-fvtpl-item">
                    <div class="macro-fvtpl-label">TP (Trái phiếu)</div>
                    <div class="macro-fvtpl-val">${mix.TP != null ? (mix.TP * 100).toFixed(0) + '%' : '-'}</div>
                    <div class="macro-fvtpl-rate">R = ${rates.R_TP != null ? (rates.R_TP * 100).toFixed(1) + '%' : '-'}/năm</div>
                </div>
                <div class="macro-fvtpl-item">
                    <div class="macro-fvtpl-label">CP (Cổ phiếu/VN-Index)</div>
                    <div class="macro-fvtpl-val">${mix.CP != null ? (mix.CP * 100).toFixed(0) + '%' : '-'}</div>
                    <div class="macro-fvtpl-rate">R = ${rates.R_VNI != null ? (rates.R_VNI * 100).toFixed(1) + '%' : '-'}/năm</div>
                </div>
                <div class="macro-fvtpl-item highlight">
                    <div class="macro-fvtpl-label">Lợi suất kỳ vọng tổng hợp</div>
                    <div class="macro-fvtpl-val" style="color:#10b981">${expectedYield != null ? (expectedYield * 100).toFixed(2) + '%' : '-'}</div>
                    <div class="macro-fvtpl-rate">Weighted average yield</div>
                </div>
            </div>
            <div class="macro-liq-note">⚠ Tỷ trọng CDs/TP/CP là PER-TICKER — cần đối chiếu với thuyết minh BCTC thực tế mỗi quý.</div>
        </div>

        <div class="macro-liq-section">
            <h4 class="macro-liq-title">🎯 Thị phần Môi giới (ngụ ý)</h4>
            <div class="macro-adtv-item">
                <div class="macro-adtv-year">Thị phần</div>
                <div class="macro-adtv-val" style="color:#f59e0b">${marketShare != null ? (marketShare * 100).toFixed(2) + '%' : '-'}</div>
            </div>
            <div class="macro-liq-note">Tính ngược từ: DT môi giới lịch sử ÷ (ADTV × 250 phiên × phí 0.15 bps)</div>
        </div>
    `;
}

// ═══════════════════════════════════════════════════════════
// PEER BENCHMARK TABLE
// ═══════════════════════════════════════════════════════════
async function loadPeerBenchmark(ticker) {
    const tbody = document.getElementById('peer-benchmark-tbody');
    const dateEl = document.getElementById('peer-update-date');
    try {
        const resp = await fetch('data/peer_benchmark_securities.json');
        if (!resp.ok) throw new Error('no data');
        const data = await resp.json();
        const peers = data.peers || [];
        if (dateEl) dateEl.innerHTML = `Số liệu thống kê thuần túy diễn giải/so sánh — KHÔNG dùng làm input định giá. Cập nhật: <strong>${data._meta?.updated || '-'}</strong>`;

        const med = (key) => calcMedian(peers.map(p => p[key]).filter(v => v != null));
        const medPB = med('pb'), medPE = med('pe'), medROE = med('roe'), medLev = med('margin_to_equity');
        const fmtX = (v, d = 2) => v != null ? v.toFixed(d) + 'x' : '-';
        const fmtPct = (v) => v != null ? v.toFixed(1) + '%' : '-';

        const medRow = `<tr style="border-top:1.5px solid rgba(255,255,255,0.12);color:#f59e0b;font-weight:600">
            <td>⚡ Trung vị ngành</td><td>-</td><td>-</td>
            <td>${fmtX(medPB)}</td><td>${fmtX(medPE, 1)}</td><td>${fmtPct(medROE)}</td><td>${fmtX(medLev)}</td>
            <td>-</td><td>-</td><td>-</td>
        </tr>`;

        const rows = peers.map(p => {
            const isCurrent = p.ticker === ticker;
            const rowStyle = isCurrent ? 'background:rgba(59,130,246,0.10);font-weight:700;' : '';
            return `<tr style="${rowStyle}">
                <td>${isCurrent ? '👉 ' : ''}${p.ticker} — ${p.name || ''}</td>
                <td>${formatNumber(p.charter_capital)}</td>
                <td>${formatNumber(p.mcap)}</td>
                <td>${fmtX(p.pb)}</td>
                <td>${fmtX(p.pe, 1)}</td>
                <td>${fmtPct(p.roe)}</td>
                <td>${fmtX(p.margin_to_equity)}</td>
                <td>${fmtPct(p.pct_margin_rev)}</td>
                <td>${fmtPct(p.pct_brokerage_rev)}</td>
                <td>${fmtPct(p.pct_tudoanh_rev)}</td>
            </tr>`;
        }).join('');

        tbody.innerHTML = medRow + rows;
    } catch (e) {
        console.warn('Peer benchmark load error:', e);
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:#ef4444">Không thể tải dữ liệu so sánh CTCK cùng ngành.</td></tr>';
    }
}

// ═══════════════════════════════════════════════════════════
// FVTPL/AFS — Cơ cấu danh mục Tự doanh theo nhóm tài sản (data/fvtpl_holdings/<TICKER>.json)
// Chỉ hiện khi có file dữ liệu cho ticker đang xem — hoàn toàn tách biệt khỏi data/<TICKER>.json
// (giống pattern peer_benchmark_securities.json / market_share_history.json).
// ═══════════════════════════════════════════════════════════
const FVTPL_HOLDING_COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#06b6d4', '#eab308', '#f97316', '#64748b', '#84cc16'];

function _fvtplPeriodSort(a, b) {
    const ay = parseInt(a.slice(0, 4), 10), by = parseInt(b.slice(0, 4), 10);
    const aq = a.includes('Q') ? parseInt(a.split('Q')[1], 10) : 0;
    const bq = b.includes('Q') ? parseInt(b.split('Q')[1], 10) : 0;
    return ay - by || aq - bq;
}

function _fvtplMergedPeriods(pool) {
    // pool: {yearly:{...}, quarterly:{...}} -> danh sách nhãn kỳ theo thời gian, nhãn năm có hậu tố (N).
    // Năm nào đã có đủ Q4 (= cùng thời điểm 31/12) trong quarterly thì bỏ điểm yearly trùng đó để
    // tránh hiển thị 2 lần cùng 1 mốc thời gian (yearly chỉ dùng cho năm KHÔNG có dữ liệu quý).
    const qYears = new Set(Object.keys(pool.quarterly || {}).map(q => q.slice(0, 4)));
    const years = Object.keys(pool.yearly || {})
        .filter(y => !qYears.has(y) || !(pool.quarterly || {})[`${y}Q4`])
        .map(y => ({ key: y, label: `${y}(N)` }));
    const quarters = Object.keys(pool.quarterly || {}).map(q => ({ key: q, label: q }));
    return [...years, ...quarters].sort((a, b) => _fvtplPeriodSort(a.key, b.key));
}

function _fvtplValueAt(pool, periodKey, isYearly, seg) {
    const bucket = isYearly ? (pool.yearly || {})[periodKey] : (pool.quarterly || {})[periodKey];
    return bucket ? (bucket[seg] || null) : null;
}

function _combineFvtplAfs(data) {
    // Gộp 2 khối fvtpl + afs (cùng schema 5 nhóm tài sản x cost/fairValue/gain/loss) thành 1 pool
    // "Tự doanh" tổng — CTCK có cả FVTPL và AFS (vd VCI) thì cơ cấu/lãi-lỗ đánh giá lại phải tính tổng
    // cả 2 vì cả hai đều là tài sản tự doanh. CTCK không có AFS (afs rỗng, vd HCM) thì gộp vào không
    // đổi kết quả so với chỉ dùng fvtpl.
    const fvtpl = data.fvtpl || {};
    const afs = data.afs || {};
    const sumField = (a, b, key) => {
        const va = a[key], vb = b[key];
        if (va == null && vb == null) return null;
        return (va || 0) + (vb || 0);
    };
    const out = { yearly: {}, quarterly: {} };
    for (const bucketName of ['yearly', 'quarterly']) {
        const fBucket = fvtpl[bucketName] || {};
        const aBucket = afs[bucketName] || {};
        const periods = new Set([...Object.keys(fBucket), ...Object.keys(aBucket)]);
        for (const pkey of periods) {
            const fSeg = fBucket[pkey] || {};
            const aSeg = aBucket[pkey] || {};
            const cats = new Set([...Object.keys(fSeg), ...Object.keys(aSeg)]);
            const merged = {};
            for (const cat of cats) {
                const fv = fSeg[cat] || {}, av = aSeg[cat] || {};
                merged[cat] = { cost: sumField(fv, av, 'cost'), fairValue: sumField(fv, av, 'fairValue'),
                                gain: sumField(fv, av, 'gain'), loss: sumField(fv, av, 'loss') };
            }
            out[bucketName][pkey] = merged;
        }
    }
    return out;
}

async function loadFvtplHoldings(ticker) {
    const section = document.getElementById('fvtpl-holdings-section');
    if (!section) return;
    try {
        const resp = await fetch(`data/fvtpl_holdings/${ticker}.json`);
        if (!resp.ok) { section.style.display = 'none'; return; }
        const data = await resp.json();
        section.style.display = '';
        renderFvtplCompositionChart(data);
        renderFvtplGainLossChart(data);
        renderFvtplHoldingsChart(data);
        renderFvtplHoldingsQuarterlyChart(data);
    } catch (e) {
        console.warn('FVTPL holdings load error:', e);
        section.style.display = 'none';
    }
}

function renderFvtplCompositionChart(data) {
    const cats = Object.entries(data.categories || {}).sort((a, b) => (a[1].order || 99) - (b[1].order || 99));
    const combined = _combineFvtplAfs(data);
    const periods = _fvtplMergedPeriods(combined);
    if (!cats.length || !periods.length) return;

    const datasets = cats.map(([key, meta]) => ({
        label: meta.label,
        backgroundColor: meta.color,
        data: periods.map(p => {
            const v = _fvtplValueAt(combined, p.key, p.label.includes('(N)'), key);
            return v && v.fairValue != null ? Math.round(v.fairValue / 1e9 * 100) / 100 : 0; // tỷ VND
        }),
        stack: 'fvtpl',
    }));

    chartFvtplComp = new Chart(document.getElementById('fvtplCompositionChart').getContext('2d'), {
        type: 'bar',
        data: { labels: periods.map(p => p.label), datasets },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, stacked: true },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Giá trị hợp lý (tỷ VND)', color: '#8892a4' } },
            },
        },
    });

    const latest = periods[periods.length - 1];
    const latestVals = cats.map(([key, meta]) => {
        const v = _fvtplValueAt(combined, latest.key, latest.label.includes('(N)'), key);
        return { label: meta.label, fv: v ? (v.fairValue || 0) : 0 };
    }).filter(x => x.fv > 0).sort((a, b) => b.fv - a.fv);
    const total = latestVals.reduce((s, x) => s + x.fv, 0);
    const top = latestVals[0];
    document.getElementById('analysis-text-fvtpl-comp').textContent = top
        ? `Tại ${latest.label}: danh mục tự doanh (FVTPL+AFS) đạt ${formatNumber(total / 1e9, 1)} tỷ VND, tập trung nhiều nhất ở "${top.label}" (${(top.fv / total * 100).toFixed(1)}% danh mục).`
        : 'Chưa đủ dữ liệu cơ cấu tự doanh.';
}

function renderFvtplGainLossChart(data) {
    const cats = Object.entries(data.categories || {}).sort((a, b) => (a[1].order || 99) - (b[1].order || 99));
    const combined = _combineFvtplAfs(data);
    const periods = _fvtplMergedPeriods(combined);
    if (!cats.length || !periods.length) return;

    const gainDatasets = cats.map(([key, meta]) => ({
        label: meta.label + ' — Lãi',
        backgroundColor: meta.color,
        data: periods.map(p => {
            const v = _fvtplValueAt(combined, p.key, p.label.includes('(N)'), key);
            return v ? Math.round((v.gain || 0) / 1e9 * 100) / 100 : 0;
        }),
        stack: 'gainloss',
    }));
    const lossDatasets = cats.map(([key, meta]) => ({
        label: meta.label + ' — Lỗ',
        backgroundColor: meta.color + '80',
        data: periods.map(p => {
            const v = _fvtplValueAt(combined, p.key, p.label.includes('(N)'), key);
            return v ? Math.round((v.loss || 0) / 1e9 * 100) / 100 : 0;
        }),
        stack: 'gainloss',
    }));

    chartFvtplGainLoss = new Chart(document.getElementById('fvtplGainLossChart').getContext('2d'), {
        type: 'bar',
        data: { labels: periods.map(p => p.label), datasets: [...gainDatasets, ...lossDatasets] },
        options: {
            ...CHART_DEFAULTS,
            plugins: { ...CHART_DEFAULTS.plugins, legend: { display: false } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, stacked: true },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Chênh lệch đánh giá lại (tỷ VND)', color: '#8892a4' } },
            },
        },
    });

    const latest = periods[periods.length - 1];
    const isY = latest.label.includes('(N)');
    const rows = cats.map(([key, meta]) => {
        const v = _fvtplValueAt(combined, latest.key, isY, key);
        return { label: meta.label, gain: v ? (v.gain || 0) : 0, loss: v ? (v.loss || 0) : 0 };
    });
    const topGain = rows.slice().sort((a, b) => b.gain - a.gain)[0];
    const topLoss = rows.slice().sort((a, b) => a.loss - b.loss)[0];
    let txt = `Tại ${latest.label}: `;
    txt += topGain && topGain.gain > 0 ? `lãi đánh giá lại chủ yếu từ "${topGain.label}" (+${formatNumber(topGain.gain / 1e9, 1)} tỷ VND).` : 'không ghi nhận lãi đánh giá lại đáng kể.';
    if (topLoss && topLoss.loss < 0) txt += ` Lỗ đánh giá lại chủ yếu từ "${topLoss.label}" (${formatNumber(topLoss.loss / 1e9, 1)} tỷ VND).`;
    document.getElementById('analysis-text-fvtpl-gainloss').textContent = txt;
}

function _renderFvtplHoldingsChartGeneric(holdings, periods, canvasId, analysisElId, chartVarSetter) {
    if (!periods.length) return;
    const tickerSet = new Set();
    periods.forEach(p => {
        const bucket = (p.label.includes('(N)') ? holdings.yearly : holdings.quarterly)[p.key] || {};
        Object.keys(bucket).forEach(t => { if (t !== 'Khac') tickerSet.add(t); });
    });
    const tickers = [...tickerSet];
    tickers.push('Khac');

    const datasets = tickers.map((t, i) => ({
        label: t,
        backgroundColor: FVTPL_HOLDING_COLORS[i % FVTPL_HOLDING_COLORS.length],
        data: periods.map(p => {
            const bucket = (p.label.includes('(N)') ? holdings.yearly : holdings.quarterly)[p.key] || {};
            const v = bucket[t];
            return v ? Math.round(v / 1e9 * 100) / 100 : 0;
        }),
        stack: 'holdings',
    }));

    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const chart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: { labels: periods.map(p => p.label), datasets },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, stacked: true },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Giá trị hợp lý (tỷ VND)', color: '#8892a4' } },
            },
        },
    });
    chartVarSetter(chart);

    const latest = periods[periods.length - 1];
    const bucket = (latest.label.includes('(N)') ? holdings.yearly : holdings.quarterly)[latest.key] || {};
    const sorted = Object.entries(bucket).filter(([t]) => t !== 'Khac').sort((a, b) => b[1] - a[1]);
    const el = document.getElementById(analysisElId);
    if (el) {
        el.textContent = sorted.length
            ? `Tại ${latest.label}, cổ phiếu niêm yết được nắm giữ lớn nhất trong danh mục tự doanh: ${sorted.slice(0, 3).map(([t, v]) => `${t} (${formatNumber(v / 1e9, 1)} tỷ)`).join(', ')}.`
            : 'Chưa có dữ liệu danh mục cổ phiếu cụ thể.';
    }
}

function renderFvtplHoldingsChart(data) {
    const holdings = data.holdings || {};
    const periods = _fvtplMergedPeriods(holdings).filter(p => p.label.includes('(N)'));
    _renderFvtplHoldingsChartGeneric(holdings, periods, 'fvtplHoldingsChart', 'analysis-text-fvtpl-holdings', c => { chartFvtplHoldings = c; });
}

function renderFvtplHoldingsQuarterlyChart(data) {
    const holdings = data.holdings || {};
    const periods = _fvtplMergedPeriods(holdings).filter(p => !p.label.includes('(N)'));
    if (!periods.length) {
        const card = document.getElementById('fvtplHoldingsQuarterlyChart')?.closest('.chart-card-split');
        if (card) card.style.display = 'none';
        return;
    }
    _renderFvtplHoldingsChartGeneric(holdings, periods, 'fvtplHoldingsQuarterlyChart', 'analysis-text-fvtpl-holdings-quarterly', c => { chartFvtplHoldingsQuarterly = c; });
}

// ═══════════════════════════════════════════════════════════
// CHARTS
// ═══════════════════════════════════════════════════════════
function renderSegmentRevenueChart(localJson) {
    const seg = localJson.segments;
    if (!seg || !seg.names) return;
    const names = seg.names;
    const nHistYears = Math.max(...names.map(n => (seg.revenueHist[n] || []).length));
    const nFcYears = Math.max(...names.map(n => (seg.revenueForecast[n] || []).length));
    const years = localJson.data?.years || [];
    const labels = years.map((y, i) => i < nHistYears ? `${y}A` : `${y}E`);

    const ctx = document.getElementById('segmentRevenueChart').getContext('2d');
    chartSegmentRevenue = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: names.map(n => ({
                label: (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n,
                data: [...(seg.revenueHist[n] || []), ...(seg.revenueForecast[n] || [])].map(v => Math.round(v)),
                backgroundColor: (SEGMENT_COLORS[n] || '#64748b') + 'cc',
                borderRadius: 3,
                stack: 'seg',
            })),
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                ...CHART_DEFAULTS.plugins,
                datalabels: {
                    display: (context) => {
                        return context.dataset.data[context.dataIndex] > 80;
                    },
                    color: '#ffffff',
                    font: { family: 'Inter', weight: '600', size: 9.5 },
                    formatter: (val) => val.toLocaleString('vi-VN')
                }
            },
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, stacked: true }, y: { ...CHART_DEFAULTS.scales.y, stacked: true } }
        },
    });

    const dominant = names.reduce((a, b) => (seg.pctNow[a] || 0) > (seg.pctNow[b] || 0) ? a : b);
    const pct = (seg.pctNow[dominant] * 100).toFixed(1);
    const label = (seg.labels && seg.labels[dominant]) || SEGMENT_LABELS_FALLBACK[dominant] || dominant;
    document.getElementById('analysis-text-segment').textContent =
        `Mảng đóng góp doanh thu lớn nhất hiện tại: ${label} (${pct}% tổng DT dự phóng). ` +
        (pct > 45 ? 'Mức độ tập trung khá cao — theo dõi rủi ro phụ thuộc 1 mảng.' : 'Cơ cấu doanh thu tương đối cân bằng giữa các mảng.');
}

function renderSegmentMixChart(localJson) {
    const seg = localJson.segments;
    if (!seg || !seg.names) return;
    const names = seg.names;
    const years = localJson.data?.years || [];
    const nHistYears = Math.max(...names.map(n => (seg.revenueHist[n] || []).length));
    const labels = years.map((y, i) => i < nHistYears ? `${y}A` : `${y}E`);
    const totals = labels.map((_, i) => names.reduce((s, n) => {
        const arr = [...(seg.revenueHist[n] || []), ...(seg.revenueForecast[n] || [])];
        return s + (arr[i] || 0);
    }, 0));

    const ctx = document.getElementById('segmentMixChart').getContext('2d');
    chartSegmentMix = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: names.map(n => {
                const arr = [...(seg.revenueHist[n] || []), ...(seg.revenueForecast[n] || [])];
                return {
                    label: (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n,
                    data: arr.map((v, i) => totals[i] ? +(v / totals[i] * 100).toFixed(1) : 0),
                    backgroundColor: (SEGMENT_COLORS[n] || '#64748b') + 'cc',
                    borderRadius: 3,
                    stack: 'pct',
                };
            }),
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                ...CHART_DEFAULTS.plugins,
                datalabels: {
                    display: (context) => {
                        return context.dataset.data[context.dataIndex] > 6;
                    },
                    color: '#ffffff',
                    font: { family: 'Inter', weight: '600', size: 9.5 },
                    formatter: (val) => val.toFixed(0) + '%'
                }
            },
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, stacked: true }, y: { ...CHART_DEFAULTS.scales.y, stacked: true, max: 100 } },
        },
    });
    document.getElementById('analysis-text-mix').textContent =
        'Theo dõi tỷ trọng Tự doanh (FVTPL+AFS) qua các năm — mảng này biến động mạnh nhất theo VN-Index, tỷ trọng tăng nhanh đồng nghĩa LNST kém ổn định hơn.';
}

function renderSegmentRevenueQuarterlyChart(localJson) {
    const seg = localJson.segments;
    const q = localJson.quarterly;
    if (!seg || !seg.names || !q || !q.segmentRevenue || !q.labels || !q.labels.length) return;
    const names = seg.names;

    const ctx = document.getElementById('segmentRevenueQuarterlyChart');
    if (!ctx) return;
    chartSegmentRevenueQuarterly = new Chart(ctx.getContext('2d'), {
        type: 'bar',
        data: {
            labels: q.labels,
            datasets: names.map(n => ({
                label: (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n,
                data: (q.segmentRevenue[n] || []).map(v => Math.round(v)),
                backgroundColor: (SEGMENT_COLORS[n] || '#64748b') + 'cc',
                borderRadius: 3,
                stack: 'segq',
            })),
        },
        options: {
            ...CHART_DEFAULTS,
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, stacked: true }, y: { ...CHART_DEFAULTS.scales.y, stacked: true } },
        },
    });

    const lastIdx = q.labels.length - 1;
    const lastVals = names.map(n => ({ n, v: (q.segmentRevenue[n] || [])[lastIdx] || 0 })).sort((a, b) => b.v - a.v);
    const top = lastVals[0];
    const label = top ? ((seg.labels && seg.labels[top.n]) || SEGMENT_LABELS_FALLBACK[top.n] || top.n) : null;
    document.getElementById('analysis-text-segment-quarterly').textContent = top
        ? `Quý ${q.labels[lastIdx]}: mảng đóng góp doanh thu lớn nhất là ${label} (${formatNumber(top.v, 0)} tỷ VND).`
        : 'Chưa đủ dữ liệu doanh thu theo quý.';
}

function renderSegmentGrossProfitChart(localJson) {
    const seg = localJson.segments;
    if (!seg || !seg.grossProfitHist) return;
    const gpSegNames = Object.keys(seg.grossProfitHist);
    if (!gpSegNames.length) return;
    const gpFc = seg.grossProfitForecast || {};
    const nHistYears = Math.max(...gpSegNames.map(n => (seg.grossProfitHist[n] || []).length));
    const nFcYears = Math.max(0, ...gpSegNames.map(n => (gpFc[n] || []).length));
    const years = (localJson.data?.years || []).map((y, i) => i < nHistYears ? `${y}A` : `${y}E`);

    const ctx = document.getElementById('segmentGrossProfitChart');
    if (!ctx) return;
    chartSegmentGrossProfit = new Chart(ctx.getContext('2d'), {
        type: 'bar',
        data: {
            labels: years,
            datasets: gpSegNames.map(n => ({
                label: (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n,
                data: [...(seg.grossProfitHist[n] || []), ...(gpFc[n] || [])].map(v => Math.round(v)),
                backgroundColor: (SEGMENT_COLORS[n] || '#64748b') + 'cc',
                borderRadius: 3,
                stack: 'gp',
            })),
        },
        options: {
            ...CHART_DEFAULTS,
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, stacked: true }, y: { ...CHART_DEFAULTS.scales.y, stacked: true } },
        },
    });

    const lastIdx = nHistYears - 1;
    const lastVals = gpSegNames.map(n => ({ n, v: (seg.grossProfitHist[n] || [])[lastIdx] || 0 })).sort((a, b) => b.v - a.v);
    const top = lastVals[0];
    const label = top ? ((seg.labels && seg.labels[top.n]) || SEGMENT_LABELS_FALLBACK[top.n] || top.n) : null;
    document.getElementById('analysis-text-segment-gp').textContent = top
        ? `Năm ${years[lastIdx] || ''}: mảng đóng góp lợi nhuận gộp lớn nhất là ${label} (${formatNumber(top.v, 0)} tỷ VND). Không gồm Quản lý quỹ (Vietcap không có dữ liệu chi phí riêng mảng này). Các năm ${nFcYears ? 'dự phóng dùng công thức Excel căn cứ tỷ lệ chi phí lịch sử từng mảng' : ''}.`
        : 'Chưa đủ dữ liệu lợi nhuận gộp theo mảng.';
}

function renderSegmentGrossProfitQuarterlyChart(localJson) {
    const q = localJson.quarterly;
    const seg = localJson.segments;
    if (!seg || !q || !q.segmentGrossProfit || !q.labels || !q.labels.length) return;
    const gpSegNames = Object.keys(q.segmentGrossProfit);
    if (!gpSegNames.length) return;

    const ctx = document.getElementById('segmentGrossProfitQuarterlyChart');
    if (!ctx) return;
    chartSegmentGrossProfitQuarterly = new Chart(ctx.getContext('2d'), {
        type: 'bar',
        data: {
            labels: q.labels,
            datasets: gpSegNames.map(n => ({
                label: (seg.labels && seg.labels[n]) || SEGMENT_LABELS_FALLBACK[n] || n,
                data: (q.segmentGrossProfit[n] || []).map(v => Math.round(v)),
                backgroundColor: (SEGMENT_COLORS[n] || '#64748b') + 'cc',
                borderRadius: 3,
                stack: 'gpq',
            })),
        },
        options: {
            ...CHART_DEFAULTS,
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, stacked: true }, y: { ...CHART_DEFAULTS.scales.y, stacked: true } },
        },
    });

    const lastIdx = q.labels.length - 1;
    const lastVals = gpSegNames.map(n => ({ n, v: (q.segmentGrossProfit[n] || [])[lastIdx] || 0 })).sort((a, b) => b.v - a.v);
    const top = lastVals[0];
    const label = top ? ((seg.labels && seg.labels[top.n]) || SEGMENT_LABELS_FALLBACK[top.n] || top.n) : null;
    document.getElementById('analysis-text-segment-gp-quarterly').textContent = top
        ? `Quý ${q.labels[lastIdx]}: mảng đóng góp lợi nhuận gộp lớn nhất là ${label} (${formatNumber(top.v, 0)} tỷ VND).`
        : 'Chưa đủ dữ liệu lợi nhuận gộp theo quý.';
}

function renderRevNpatChart(localJson) {
    const data = localJson.data;
    if (!data) return;
    const nFc = 3;
    const labels = data.years.map((y, i) => i < data.years.length - nFc ? `${y}A` : `${y}E`);
    const ctx = document.getElementById('revNpatChart').getContext('2d');
    chartRevNpat = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'Doanh thu hoạt động (tỷ VND)', data: data.revenue, backgroundColor: '#3b82f640', borderColor: '#3b82f6', borderWidth: 2, borderRadius: 6, yAxisID: 'y' },
                { label: 'LNST cổ đông mẹ (tỷ VND)', data: data.npat, type: 'line', borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)', borderWidth: 2.5, pointRadius: 5, pointBackgroundColor: '#10b981', tension: 0.4, yAxisID: 'y1' },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left' },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false } },
            },
        },
    });
    const lastRevGrowth = data.revenue.length > 1 ? ((data.revenue.at(-1) / data.revenue.at(-2) - 1) * 100).toFixed(1) : null;
    document.getElementById('analysis-text-rev-npat').textContent = lastRevGrowth
        ? `Doanh thu dự phóng năm cuối tăng ${lastRevGrowth}% so với năm trước — xem chi tiết diễn biến quý gần nhất ở phần đánh giá KQKD phía trên.`
        : 'Đang tổng hợp xu hướng doanh thu & lợi nhuận.';
}

function renderQuarterlyNpatRoeChart(localJson) {
    const q = localJson.quarterly;
    if (!q || !q.labels || !q.labels.length) return;
    const ctx = document.getElementById('quarterlyNpatRoeChart').getContext('2d');
    chartQuarterlyNpatRoe = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: q.labels,
            datasets: [
                { label: 'LNST cổ đông mẹ (tỷ VND)', data: q.npat, backgroundColor: '#16a34a99', borderColor: '#16a34a', borderWidth: 1.5, borderRadius: 4, yAxisID: 'y' },
                { label: 'ROE (%/năm, annualized)', data: q.roe, type: 'line', borderColor: '#dc2626', backgroundColor: 'rgba(220,38,38,0.1)', borderWidth: 2.2, pointRadius: 4, tension: 0.3, yAxisID: 'y1' },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left' },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false } },
            },
        },
    });
    const validRoe = q.roe.filter(v => v != null);
    const avgRoe = validRoe.length ? (validRoe.reduce((a, b) => a + b, 0) / validRoe.length).toFixed(1) : null;
    document.getElementById('analysis-text-quarterly').textContent = avgRoe
        ? `ROE bình quân annualized theo quý (12 quý gần nhất): ${avgRoe}%. Theo dõi xu hướng biến động — CTCK có ROE biến động mạnh thường do tỷ trọng Tự doanh cao.`
        : 'Đang tính toán ROE theo quý.';
}

function renderRoePbCorrelationChart(localJson) {
    const pairs = localJson.roePbCorrelation || [];
    if (!pairs.length) return;
    const ctx = document.getElementById('roePbCorrelationChart').getContext('2d');
    const points = pairs.map(p => ({ x: p.roe, y: p.pb }));
    chartRoePbCorrelation = new Chart(ctx, {
        type: 'scatter',
        data: {
            datasets: [{
                label: 'ROE vs P/B theo quý',
                data: points,
                backgroundColor: '#3b82f6aa',
                borderColor: '#1F4E78',
                pointRadius: 5,
            }],
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, title: { display: true, text: 'ROE (%/năm)', color: '#8892a4' } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'P/B (x)', color: '#8892a4' } },
            },
        },
    });
    // Hệ số tương quan Pearson đơn giản
    const n = points.length;
    let corr = null;
    if (n >= 3) {
        const mx = points.reduce((s, p) => s + p.x, 0) / n;
        const my = points.reduce((s, p) => s + p.y, 0) / n;
        const num = points.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0);
        const denX = Math.sqrt(points.reduce((s, p) => s + (p.x - mx) ** 2, 0));
        const denY = Math.sqrt(points.reduce((s, p) => s + (p.y - my) ** 2, 0));
        corr = denX && denY ? num / (denX * denY) : null;
    }
    document.getElementById('analysis-text-correlation').textContent = corr != null
        ? `Hệ số tương quan r = ${corr.toFixed(2)} — ${corr > 0.5 ? 'tương quan dương rõ rệt: thị trường trả P/B cao hơn cho quý có ROE cao hơn.' : corr > 0.2 ? 'tương quan dương nhẹ.' : 'tương quan yếu/không rõ ràng — P/B có thể chịu ảnh hưởng của các yếu tố khác ngoài ROE.'}`
        : 'Chưa đủ dữ liệu để tính tương quan.';
}

function renderMarginLeverageChart(localJson) {
    const q = localJson.quarterly;
    const lev = localJson.marginLeverage || {};
    if (!q || !q.labels || !q.labels.length) return;
    const ctx = document.getElementById('marginLeverageChart').getContext('2d');
    chartMarginLeverage = new Chart(ctx, {
        type: 'line',
        data: {
            labels: q.labels,
            datasets: [
                { label: 'Dư nợ Margin / VCSH (x)', data: q.marginLeverage, borderColor: '#1F4E78', backgroundColor: 'rgba(31,78,120,0.1)', borderWidth: 2.2, pointRadius: 4, tension: 0.2 },
                { label: 'Trần pháp lý 2,0x', data: q.labels.map(() => lev.legalCap ?? 2.0), borderColor: '#dc2626', borderDash: [6, 4], borderWidth: 1.5, pointRadius: 0 },
                { label: 'Ngưỡng cảnh báo 1,8x', data: q.labels.map(() => lev.warningThreshold ?? 1.8), borderColor: '#f59e0b', borderDash: [3, 3], borderWidth: 1.5, pointRadius: 0 },
            ],
        },
        options: { ...CHART_DEFAULTS },
    });
    document.getElementById('analysis-text-leverage').textContent =
        `Tỷ lệ hiện tại (${lev.latestLabel || '-'}) : ${lev.latest != null ? lev.latest.toFixed(2) : '-'}x. ` +
        (lev.latest >= 2.0 ? 'Đã VƯỢT trần pháp lý — dư nợ margin không thể tăng thêm nếu không tăng vốn.'
            : lev.latest >= 1.8 ? 'Đang GẦN chạm trần — dư địa tăng trưởng margin hạn chế, phụ thuộc kế hoạch tăng vốn.'
            : 'Còn dư địa tăng trưởng dư nợ margin dưới trần pháp lý.');
}

function renderPEChart(localJson) {
    const labels = localJson.quarter_labels || [];
    const data = localJson.pe_quarters || [];
    if (!labels.length) return;
    const med = calcMedian(data.filter(v => v && v > 0 && v < 60));
    const ctx = document.getElementById('peChart').getContext('2d');
    chartPE = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'P/E TTM', data, borderColor: '#2980B9', backgroundColor: '#2980B918', borderWidth: 2, pointRadius: 3, tension: 0.3, fill: true, spanGaps: true },
                { label: `Trung vị (${med ? med.toFixed(1) + 'x' : '-'})`, data: labels.map(() => med), borderColor: '#f59e0b', borderWidth: 1.5, borderDash: [5, 5], pointRadius: 0, fill: false },
            ],
        },
        options: { ...CHART_DEFAULTS },
    });
    document.getElementById('analysis-text-pe').textContent = med
        ? `P/E trung vị lịch sử: ${med.toFixed(1)}x. Trọng số P/E trong định giá chỉ 10% — chủ yếu làm bộ lọc chống nhiễu.`
        : 'Đang tính P/E trung vị lịch sử.';
}

function renderPBChart(localJson) {
    const labels = localJson.quarter_labels || [];
    const data = localJson.pb_quarters || [];
    const roeData = localJson.roe_quarters_ttm || [];
    if (!labels.length) return;
    const med = calcMedian(data.filter(v => v && v > 0));
    const ctx = document.getElementById('pbChart').getContext('2d');
    chartPB = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'P/B TTM', data, borderColor: '#27AE60', backgroundColor: '#27AE6018', borderWidth: 2, pointRadius: 3, tension: 0.3, fill: true, spanGaps: true, yAxisID: 'y' },
                { label: `Trung vị (${med ? med.toFixed(2) + 'x' : '-'})`, data: labels.map(() => med), borderColor: '#f59e0b', borderWidth: 1.5, borderDash: [5, 5], pointRadius: 0, fill: false, yAxisID: 'y' },
                { label: 'ROE (%, TTM 4 quý)', data: roeData, borderColor: '#dc2626', backgroundColor: 'rgba(220,38,38,0.08)', borderWidth: 2, pointRadius: 2.5, tension: 0.3, spanGaps: true, yAxisID: 'y1' },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'P/B (x)', color: '#8892a4' } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'ROE (%)', color: '#8892a4' } },
            },
        },
    });
    const validRoe = roeData.filter(v => v != null);
    const lastRoe = validRoe.length ? validRoe[validRoe.length - 1] : null;
    document.getElementById('analysis-text-pb').textContent = med
        ? `P/B trung vị lịch sử: ${med.toFixed(2)}x — neo chính cho định giá (trọng số 90%).` +
          (lastRoe != null ? ` ROE TTM gần nhất: ${lastRoe.toFixed(1)}%.` : '')
        : 'Đang tính P/B trung vị lịch sử.';
}

async function renderMarketShareChart(currentTicker) {
    const canvas = document.getElementById('marketShareChart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    
    try {
        const resp = await fetch('data/market_share_history.json');
        if (!resp.ok) throw new Error();
        const res = await resp.json();
        
        const quarters = res.quarters || [];
        const top5 = res.top_5 || [];
        const shareData = res.data || {};
        
        if (!quarters.length || !top5.length) return;
        
        // Bảng màu cho Top 8
        const colorPalette = {
            'SSI': '#3b82f6',
            'HCM': '#f59e0b',
            'VCI': '#ec4899',
            'VND': '#10b981',
            'MBS': '#8b5cf6',
            'TCBS': '#e02424', // Đỏ Techcombank
            'VPS': '#ea580c', // Cam đậm VPS
            'BVS': '#06b6d4', // Cyan
            'BSI': '#f43f5e', // Rose
            'SHS': '#14b8a6', // Teal
            'default': '#a855f7'
        };
        
        const datasets = top5.map(ticker => {
            const isCurrent = ticker === currentTicker;
            const baseColor = colorPalette[ticker] || colorPalette['default'];
            
            return {
                label: ticker + (isCurrent ? ' (Mã đang xem)' : ''),
                data: (shareData[ticker] || []).map(v => v * 100),
                borderColor: baseColor,
                backgroundColor: isCurrent ? baseColor + '15' : 'transparent',
                borderWidth: isCurrent ? 3.5 : 1.8,
                pointRadius: isCurrent ? 5 : 2,
                pointHoverRadius: 7,
                tension: 0.25,
                fill: isCurrent,
                spanGaps: true
            };
        });
        
        chartMarketShare = new Chart(ctx, {
            type: 'line',
            data: {
                labels: quarters,
                datasets: datasets
            },
            options: {
                ...CHART_DEFAULTS,
                plugins: {
                    ...CHART_DEFAULTS.plugins,
                    tooltip: {
                        ...CHART_DEFAULTS.plugins.tooltip,
                        callbacks: {
                            label: function(context) {
                                return ` ${context.dataset.label}: ${context.raw.toFixed(2)}%`;
                            }
                        }
                    }
                },
                scales: {
                    ...CHART_DEFAULTS.scales,
                    y: {
                        ...CHART_DEFAULTS.scales.y,
                        ticks: {
                            ...CHART_DEFAULTS.scales.y.ticks,
                            callback: function(value) {
                                return value.toFixed(1) + '%';
                            }
                        }
                    }
                }
            }
        });
        
        const shareCurrent = shareData[currentTicker] ? shareData[currentTicker][shareData[currentTicker].length - 1] * 100 : null;
        let analysisText = `Tính đến quý gần nhất (${quarters[quarters.length - 1]}), `;
        if (shareCurrent) {
            analysisText += `Thị phần ngụ ý của ${currentTicker} đạt khoảng <strong>${shareCurrent.toFixed(2)}%</strong>. `;
            const firstShare = shareData[currentTicker][0] * 100;
            const diff = shareCurrent - firstShare;
            if (diff > 0.5) {
                analysisText += `Đang có xu hướng <strong>tăng trưởng tích cực</strong> (+${diff.toFixed(1)}% so với giai đoạn ${quarters[0]}).`;
            } else if (diff < -0.5) {
                analysisText += `Đang có xu hướng <strong>sụt giảm thị phần</strong> (${diff.toFixed(1)}% so với giai đoạn ${quarters[0]}).`;
            } else {
                analysisText += `Đang đi ngang ổn định so với giai đoạn trước.`;
            }
        } else {
            analysisText += `Mã ${currentTicker} nằm ngoài Top-5 CTCK có thị phần môi giới lớn nhất nên không xuất hiện trong danh sách so sánh.`;
        }
        document.getElementById('analysis-text-market-share').innerHTML = analysisText;
        
    } catch (e) {
        console.warn('Không thể tải hoặc vẽ biểu đồ thị phần:', e);
        document.getElementById('analysis-text-market-share').textContent = 'Chưa có dữ liệu thị phần so sánh.';
    }
}
