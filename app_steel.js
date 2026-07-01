/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app_steel.js (Steel & Materials)
   Live data from Vietcap API + JSON fallback
   ════════════════════════════════════════════════════════ */

'use strict';

// ── Chart instances ──────────────────────────────────────
let chartRevNpat = null;
let chartEpsEquity = null;
let chartPE = null;
let chartPB = null;

// ── Vietcap API base ──────────────────────────────────────
const VC_BASE = 'https://trading.vietcap.com.vn/api/iq-insight-service/v1';

// ═══════════════════════════════════════════════════════════
// SECTOR CONFIGURATION (steel / materials)
// ═══════════════════════════════════════════════════════════
const SECTOR_CONFIG = {
    materials: {
        label: 'Thép & Vật Liệu',
        labelEn: 'Basic Resources',
        icon: '\u{1F529}',
        color: '#dc2626',
        incomeField: 'isb27',
        incomeField2: 'isb47',
        npat: 'isa20',
        incomeLabel: 'Doanh thu thu\u1ea7n',
        income2Label: 'LN g\u1ed9p',
        npatLabel: 'LNST',
        showPE: true,
        primaryValuation: 'EV/EBITDA + P/B + P/E',
        extraMetrics: (ratios) => {
            const latest = ratios?.[ratios.length - 1] || {};
            return [
                { label: 'ROE', value: latest.roe ? (latest.roe * 100).toFixed(1) + '%' : '-', desc: 'T\u1ef7 su\u1ea5t v\u1ed1n ch\u1ee7' },
                { label: 'Bi\u00ean g\u1ed9p', value: latest.grossMargin ? (latest.grossMargin * 100).toFixed(1) + '%' : '-', desc: 'Gross Margin' },
                { label: 'N\u1ee3/V\u1ed1n ch\u1ee7', value: latest.debtToEquity ? latest.debtToEquity.toFixed(2) + 'x' : '-', desc: '\u0110\u00f2n b\u1ea9y' },
            ];
        }
    },
    other: {
        label: 'Kh\u00e1c',
        labelEn: 'Other',
        icon: '\u{1F4CA}',
        color: '#64748b',
        incomeField: 'isa3',
        incomeField2: 'isa5',
        npat: 'isa20',
        incomeLabel: 'Doanh thu thu\u1ea7n',
        income2Label: 'L\u1ee3i nhu\u1eadn g\u1ed9p',
        npatLabel: 'L\u1ee3i nhu\u1eadn sau thu\u1ebf',
        showPE: true,
        primaryValuation: 'P/E',
        extraMetrics: (ratios) => {
            const latest = ratios?.[ratios.length - 1] || {};
            return [
                { label: 'ROE', value: latest.roe ? (latest.roe * 100).toFixed(1) + '%' : '-', desc: 'T\u1ef7 su\u1ea5t v\u1ed1n ch\u1ee7' },
                { label: 'Bi\u00ean g\u1ed9p', value: latest.grossMargin ? (latest.grossMargin * 100).toFixed(1) + '%' : '-', desc: 'Gross Margin' },
                { label: 'N\u1ee3/V\u1ed1n ch\u1ee7', value: latest.debtToEquity ? latest.debtToEquity.toFixed(2) + 'x' : '-', desc: '\u0110\u00f2n b\u1ea9y' },
            ];
        }
    }
};

// ── Map sector string → config key ────────────────────────
function getSectorKey(sectorStr) {
    const s = (sectorStr || '').toLowerCase();
    if (s.includes('basic resource') || s.includes('material') ||
        s.includes('steel') || s.includes('metal') || s.includes('mining') ||
        s.includes('thép') || s.includes('vật liệu') || s.includes('khai khoáng') ||
        s.includes('vat lieu') || s.includes('khai khoang')) return 'materials';
    return 'other';
}

// ═══════════════════════════════════════════════════════════
// NAVIGATION / ROUTING
// ═══════════════════════════════════════════════════════════
function goHome() {
    window.location.href = 'index.html';
}

function showAnalysisView(sectorKey, ticker) {
    const cfg = SECTOR_CONFIG[sectorKey] || SECTOR_CONFIG.other;

    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';

    document.getElementById('nav-breadcrumb').innerHTML = `
        <span class="breadcrumb-home" style="cursor:pointer" onclick="goHome()">T\u1ed5ng quan</span>
        <span class="breadcrumb-sep">\u203a</span>
        <span class="breadcrumb-sector" style="color:${cfg.color}">${cfg.icon} ${cfg.label}</span>
        <span class="breadcrumb-sep">\u203a</span>
        <span class="breadcrumb-ticker">${ticker}</span>
    `;

    const badge = document.getElementById('sector-badge-nav');
    badge.style.display = 'flex';
    badge.style.borderColor = cfg.color + '50';
    badge.style.color = cfg.color;
    document.getElementById('sector-icon-nav').textContent = cfg.icon;
    document.getElementById('sector-label-nav').textContent = cfg.label;

    document.getElementById('back-btn').style.display = 'inline-flex';

    const ribbon = document.getElementById('sector-ribbon');
    ribbon.textContent = cfg.label.toUpperCase();
    ribbon.style.background = cfg.color;

    document.documentElement.style.setProperty('--sector-color', cfg.color);
}

// ═══════════════════════════════════════════════════════════
// VIETCAP API CALLS (Live)
// ═══════════════════════════════════════════════════════════
async function fetchVietcap(path) {
    try {
        const r = await fetch(VC_BASE + path, {
            headers: { 'Referer': 'https://trading.vietcap.com.vn/' }
        });
        if (!r.ok) return null;
        const json = await r.json();
        return json.data || null;
    } catch {
        return null;
    }
}

async function fetchStockLive(ticker) {
    const [details, ratios, incomeYears, incomeQuarters] = await Promise.all([
        fetchVietcap(`/company/details?ticker=${ticker}`),
        fetchVietcap(`/company/${ticker}/statistics-financial`),
        fetchVietcap(`/company/${ticker}/financial-statement?section=INCOME_STATEMENT`),
        fetchVietcap(`/company/${ticker}/financial-statement?section=INCOME_STATEMENT&quarterly=true`),
    ]);
    if (incomeYears && incomeQuarters) {
        const rawQ = incomeQuarters.quarters || incomeQuarters.years || [];
        rawQ.forEach(q => {
            if (q.quarter === null || q.quarter === undefined) {
                q.quarter = q.lengthReport;
            }
        });
        incomeYears.quarters = rawQ;
    }
    return { details, ratios, incomeYears };
}

// ═══════════════════════════════════════════════════════════
// LOAD OVERVIEW
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {
    const params = new URLSearchParams(location.search);
    const autoTicker = params.get('ticker');
    if (autoTicker) {
        document.getElementById('view-overview').style.display = 'none';
        document.getElementById('view-analysis').style.display = 'flex';
        await loadStockDashboard(autoTicker);
    } else {
        await loadOverview();
    }
});

async function loadOverview() {
    try {
        const resp = await fetch('data/index.json');
        const stocks = await resp.json();
        renderOverview(stocks);
    } catch {
        document.getElementById('sector-groups-container').innerHTML =
            '<div class="loading-state card">Ch\u01b0a c\u00f3 d\u1eef li\u1ec7u. H\u00e3y ch\u1ea1y ph\u00e2n t\u00edch \u00edt nh\u1ea5t m\u1ed9t m\u00e3 c\u1ed5 phi\u1ebfu.</div>';
    }
}

function renderOverview(stocks) {
    if (!stocks || stocks.length === 0) {
        document.getElementById('sector-groups-container').innerHTML =
            '<div class="loading-state card">Ch\u01b0a c\u00f3 d\u1eef li\u1ec7u c\u1ed5 phi\u1ebfu n\u00e0o \u0111\u01b0\u1ee3c ph\u00e2n t\u00edch.</div>';
        return;
    }

    window._allStocks = stocks;

    _groupAndRender(stocks);

    const inp = document.getElementById('stock-search-input');
    if (inp) {
        inp.addEventListener('input', () => {
            const q = inp.value.trim().toLowerCase();
            document.getElementById('search-clear').style.opacity = q ? '1' : '0';
            if (!q) {
                _groupAndRender(window._allStocks);
            } else {
                const filtered = window._allStocks.filter(s =>
                    s.ticker.toLowerCase().includes(q) ||
                    (s.companyName || '').toLowerCase().includes(q)
                );
                _renderSearchResults(filtered, q);
            }
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
        container.innerHTML = `<div class="loading-state card">Kh\u00f4ng t\u00ecm th\u1ea5y m\u00e3 n\u00e0o kh\u1edbp v\u1edbi "${query}"</div>`;
        return;
    }
    container.innerHTML = '';
    const card = document.createElement('div');
    card.className = 'card sector-group';
    card.innerHTML = `
        <div class="sector-group-header">
            <h3 style="color:var(--text-muted)">\u{1F50D} K\u1ebft qu\u1ea3 t\u00ecm ki\u1ebfm</h3>
            <span class="sector-count-badge">${filtered.length} m\u00e3</span>
        </div>
        <div class="stocks-grid" id="grid-search"></div>
    `;
    container.appendChild(card);
    const grid = card.querySelector('#grid-search');
    filtered.forEach(s => {
        const key = getSectorKey(s.sector || '');
        const cfg = SECTOR_CONFIG[key];
        grid.appendChild(_makeStockCard(s, cfg));
    });
}

function _groupAndRender(stocks) {
    const groups = {};
    stocks.forEach(s => {
        const key = getSectorKey(s.sector || '');
        if (!groups[key]) groups[key] = [];
        groups[key].push(s);
    });

    const container = document.getElementById('sector-groups-container');
    container.innerHTML = '';

    const sectorOrder = ['materials'];

    sectorOrder.forEach(key => {
        if (!groups[key]) return;
        const cfg = SECTOR_CONFIG[key];
        const groupDiv = document.createElement('div');
        groupDiv.className = 'card sector-group';

        groupDiv.innerHTML = `
            <div class="sector-group-header">
                <div class="sector-group-dot" style="background:${cfg.color}"></div>
                <h3 style="color:${cfg.color}">${cfg.icon} ${cfg.label}</h3>
                <span class="sector-count-badge">${groups[key].length} m\u00e3</span>
            </div>
            <div class="stocks-grid" id="grid-${key}"></div>
        `;
        container.appendChild(groupDiv);

        const grid = groupDiv.querySelector(`#grid-${key}`);
        groups[key].forEach(s => {
            grid.appendChild(_makeStockCard(s, cfg));
        });
    });
}

function _makeStockCard(s, cfg) {
    const card = document.createElement('div');
    card.className = 'stock-card';
    card.style.setProperty('--sector-color', cfg.color);
    card.innerHTML = `
        <div class="stock-card-ticker" style="color:${cfg.color}">${s.ticker}</div>
    `;
    card.onclick = () => loadStockDashboard(s.ticker);
    return card;
}

// ═══════════════════════════════════════════════════════════
// LOAD STOCK DASHBOARD (steel-adaptive)
// ═══════════════════════════════════════════════════════════
async function loadStockDashboard(ticker) {
    document.documentElement.style.setProperty('--sector-color', '#dc2626');
    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';
    document.getElementById('view-analysis').classList.add('active-view');
    document.getElementById('ticker-badge').textContent = ticker;
    document.getElementById('company-name').textContent = '\u0110ang t\u1ea3i d\u1eef li\u1ec7u...';
    document.getElementById('company-sector').textContent = '';
    document.getElementById('nav-breadcrumb').innerHTML = `<span class="breadcrumb-home" style="cursor:pointer" onclick="goHome()">T\u1ed5ng quan</span> \u203a <span>${ticker}</span>`;
    document.getElementById('back-btn').classList.remove('hidden');

    destroyCharts();

    const [liveData, localJson] = await Promise.all([
        fetchStockLive(ticker),
        fetch(`data/${ticker}.json`).then(r => r.ok ? r.json() : null).catch(() => null)
    ]);

    const details = liveData.details || {};
    const ratiosAll = liveData.ratios || [];
    const incomeYears = liveData.incomeYears?.years || [];

    const sector = details.sector || localJson?.sector || 'General';
    const sectorKey = getSectorKey(sector);
    const cfg = SECTOR_CONFIG[sectorKey];

    const companyName = details.viOrganName || details.enOrganName || localJson?.companyName || ticker;
    const currentPrice = details.currentPrice || localJson?.currentPrice || 0;
    const marketCap = details.marketCap || localJson?.marketCap || 0;
    const shares = details.numberOfSharesMktCap || localJson?.shares || 0;

    showAnalysisView(sectorKey, ticker);

    document.getElementById('ticker-badge').textContent = ticker;
    document.getElementById('company-name').textContent = companyName;
    document.getElementById('company-sector').textContent = `${cfg.icon} ${cfg.label} \u00b7 ${sector}`;

    const reportDateEl = document.getElementById('report-date');
    if (reportDateEl) {
        if (localJson?.lastUpdated) {
            const parts = localJson.lastUpdated.split(' ');
            if (parts.length === 2) {
                const dateParts = parts[0].split('-');
                if (dateParts.length === 3) {
                    reportDateEl.innerHTML = `\u{1F550} Th\u1eddi gian l\u1eadp b\u00e1o c\u00e1o: <b>${dateParts[2]}/${dateParts[1]}/${dateParts[0]} ${parts[1]}</b>`;
                } else {
                    reportDateEl.innerHTML = `\u{1F550} Th\u1eddi gian l\u1eadp b\u00e1o c\u00e1o: <b>${localJson.lastUpdated}</b>`;
                }
            } else {
                reportDateEl.innerHTML = `\u{1F550} Th\u1eddi gian l\u1eadp b\u00e1o c\u00e1o: <b>${localJson.lastUpdated}</b>`;
            }
            reportDateEl.style.display = 'block';
        } else {
            reportDateEl.style.display = 'none';
        }
    }

    const liveBadge = `<span class="live-badge"><span class="live-dot"></span>Live \u00b7 Vietcap IQ</span>`;

    const btnPdf = document.getElementById('download-pdf');
    const btnExcel = document.getElementById('download-excel');
    if (localJson?.gdrivePdfUrl) { btnPdf.href = localJson.gdrivePdfUrl; btnPdf.classList.remove('hidden'); }
    else btnPdf.classList.add('hidden');
    if (localJson?.gdriveExcelUrl) { btnExcel.href = localJson.gdriveExcelUrl; btnExcel.classList.remove('hidden'); }
    else btnExcel.classList.add('hidden');

    const latestRatio = ratiosAll.filter(r => r.quarter <= 4).slice(-1)[0] || {};
    const ratioTTM = ratiosAll.filter(r => r.quarter <= 4).slice(-1)[0] || {};

    const currentPE = ratioTTM.pe ? ratioTTM.pe.toFixed(1) + 'x' : '-';
    const currentPB = ratioTTM.pb ? ratioTTM.pb.toFixed(2) + 'x' : '-';

    const metricsGrid = document.getElementById('dynamic-metrics-grid');
    metricsGrid.innerHTML = `
        <div class="metric-card highlight">
            <div class="metric-label">Gi\u00e1 hi\u1ec7n t\u1ea1i</div>
            <div class="metric-value">${formatNumber(currentPrice)}</div>
            <div class="metric-sub">VND ${liveBadge}</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">V\u1ed1n h\u00f3a th\u1ecb tr\u01b0\u1eddng</div>
            <div class="metric-value">${formatNumber(Math.round(marketCap / 1e12), 1)}</div>
            <div class="metric-sub">Ngh\u00ecn t\u1ef7 VND</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">P/E hi\u1ec7n t\u1ea1i</div>
            <div class="metric-value">${currentPE}</div>
            <div class="metric-sub">Trailing twelve months</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">P/B hi\u1ec7n t\u1ea1i</div>
            <div class="metric-value">${currentPB}</div>
            <div class="metric-sub">Trailing twelve months</div>
        </div>
        ${details.targetPrice ? `
        <div class="metric-card">
            <div class="metric-label">Gi\u00e1 m\u1ee5c ti\u00eau (${details.analyst || 'Vietcap'})</div>
            <div class="metric-value" style="color:#10b981">${formatNumber(details.targetPrice)}</div>
            <div class="metric-sub">${details.rating || ''} \u00b7 Upside ${details.upsideToTargetPercent ? (details.upsideToTargetPercent * 100).toFixed(1) + '%' : ''}</div>
        </div>` : ''}
    `;

    const bankKpiCard = document.getElementById('bank-kpi-card');
    const bankKpiTitle = bankKpiCard.querySelector('h3');
    if (bankKpiTitle) {
        bankKpiTitle.textContent = 'Ch\u1ec9 s\u1ed1 T\u00e0i ch\u00ednh then ch\u1ed1t';
    }

    if (ratiosAll.length > 0) {
        const extraMetrics = cfg.extraMetrics(ratiosAll.filter(r => r.quarter <= 4), localJson);
        if (extraMetrics.some(m => m.value !== '-')) {
            bankKpiCard.classList.remove('hidden');
            const grid = document.getElementById('bank-kpi-grid');
            grid.innerHTML = extraMetrics.map(m => `
                <div class="bank-kpi-item">
                    <div class="bank-kpi-label">${m.label}</div>
                    <div class="bank-kpi-value" style="color:${cfg.color}">${m.value}</div>
                    <div style="font-size:0.7rem;color:var(--text-dim)">${m.desc}</div>
                </div>
            `).join('');
        } else {
            bankKpiCard.classList.add('hidden');
        }
    } else {
        bankKpiCard.classList.add('hidden');
    }

    const ttmQuarters = ratiosAll
        .filter(r => r.quarter >= 1 && r.quarter <= 4)
        .sort((a, b) => a.yearReport !== b.yearReport ? a.yearReport - b.yearReport : a.quarter - b.quarter);

    const peData = ttmQuarters.map(r => ({ x: `${r.yearReport}-Q${r.quarter}`, y: r.pe && r.pe > 0 && r.pe < 500 ? parseFloat(r.pe.toFixed(1)) : null }));
    const pbData = ttmQuarters.map(r => ({ x: `${r.yearReport}-Q${r.quarter}`, y: r.pb && r.pb > 0 ? parseFloat(r.pb.toFixed(2)) : null }));
    const quarterLabels = ttmQuarters.map(r => `${r.yearReport}-Q${r.quarter}`);

    const annualYears = incomeYears.sort((a, b) => a.yearReport - b.yearReport).slice(-7);
    const incomeLabels = annualYears.map(r => r.yearReport.toString());
    const incomeData = annualYears.map(r => r[cfg.incomeField] ? parseFloat((r[cfg.incomeField] / 1e12).toFixed(1)) : 0);
    const income2Data = cfg.incomeField2 ? annualYears.map(r => r[cfg.incomeField2] ? parseFloat((r[cfg.incomeField2] / 1e12).toFixed(1)) : 0) : [];
    const npatData = annualYears.map(r => r[cfg.npat] ? parseFloat((r[cfg.npat] / 1e12).toFixed(1)) : 0);

    document.getElementById('chart1-title').textContent = `${cfg.incomeLabel} & ${cfg.npatLabel}`;
    document.getElementById('chart2-title').textContent = income2Data.length > 0 ? `${cfg.income2Label} & ${cfg.npatLabel}` : 'EPS & V\u1ed1n ch\u1ee7 s\u1edf h\u1eefu';

    document.getElementById('chart-pe-card').style.opacity = '1';

    renderIncomeChart(incomeLabels, incomeData, npatData, cfg.incomeLabel, cfg.npatLabel, cfg.color);
    renderChart2(incomeLabels, income2Data, npatData, cfg.income2Label, cfg.npatLabel, cfg.color);
    renderPEChart(quarterLabels, peData.map(p => p.y), cfg.color);
    renderPBChart(quarterLabels, pbData.map(p => p.y), cfg.color);

    renderQuarterlyAndYTDEvaluation(ticker, liveData, localJson, cfg);

    generateChartCommentaries(ticker, annualYears, ttmQuarters, cfg, localJson);

    renderValuationScenarios(localJson?.valuation, currentPrice, details, latestRatio, cfg);
    renderValuationSnapshot(localJson);
    renderFinancialSnapshotTable(localJson, sectorKey);

    const q = localJson || {};
    renderThesisAndRisks(q.thesis, q.risks);
    renderMoatScorecard(q.moats, cfg.color);
    renderPESTLE(q.pestle);
    renderCommentary(q.comments);
    renderFinancialTable(annualYears, cfg, incomeLabels, localJson);
}

// ═══════════════════════════════════════════════════════════
// CHART RENDERERS
// ═══════════════════════════════════════════════════════════
const CHART_DEFAULTS = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { labels: { color: '#8892a4', font: { family: 'Inter', size: 11 }, boxWidth: 12 } },
        tooltip: { backgroundColor: 'rgba(10,16,32,0.95)', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1, titleColor: '#f0f2f8', bodyColor: '#8892a4', padding: 12, cornerRadius: 8 }
    },
    scales: {
        x: { ticks: { color: '#545f74', font: { size: 10 }, maxRotation: 45 }, grid: { color: 'rgba(255,255,255,0.04)' } },
        y: { ticks: { color: '#545f74', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,0.04)' } }
    }
};

function destroyCharts() {
    [chartRevNpat, chartEpsEquity, chartPE, chartPB].forEach(c => { if (c) c.destroy(); });
    chartRevNpat = chartEpsEquity = chartPE = chartPB = null;
}

function calcMedian(arr) {
    const valid = arr.filter(v => v !== null && v !== undefined && !isNaN(v)).sort((a, b) => a - b);
    if (!valid.length) return null;
    const m = Math.floor(valid.length / 2);
    return valid.length % 2 !== 0 ? valid[m] : (valid[m - 1] + valid[m]) / 2;
}

function renderIncomeChart(labels, incomeData, npatData, incomeLabel, npatLabel, color) {
    const ctx = document.getElementById('revNpatChart').getContext('2d');
    if (chartRevNpat) chartRevNpat.destroy();
    chartRevNpat = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: incomeLabel + ' (ngh\u00ecn t\u1ef7)',
                    data: incomeData,
                    backgroundColor: color + '40',
                    borderColor: color,
                    borderWidth: 2,
                    borderRadius: 6,
                    yAxisID: 'y'
                },
                {
                    label: npatLabel + ' (ngh\u00ecn t\u1ef7)',
                    data: npatData,
                    type: 'line',
                    borderColor: '#10b981',
                    backgroundColor: 'rgba(16,185,129,0.1)',
                    borderWidth: 2.5,
                    pointRadius: 5,
                    pointBackgroundColor: '#10b981',
                    tension: 0.4,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'Ngh\u00ecn t\u1ef7 VND', color: '#545f74', font: { size: 10 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'LNST (ngh\u00ecn t\u1ef7)', color: '#545f74', font: { size: 10 } } }
            }
        }
    });
}

function renderChart2(labels, data1, data2, label1, label2, color) {
    const ctx = document.getElementById('epsEquityChart').getContext('2d');
    if (chartEpsEquity) chartEpsEquity.destroy();
    chartEpsEquity = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: label1 + ' (ngh\u00ecn t\u1ef7)',
                    data: data1,
                    backgroundColor: color + '30',
                    borderColor: color + 'aa',
                    borderWidth: 2,
                    borderRadius: 5,
                    yAxisID: 'y'
                },
                {
                    label: label2 + ' (ngh\u00ecn t\u1ef7)',
                    data: data2,
                    type: 'line',
                    borderColor: '#818cf8',
                    backgroundColor: 'rgba(129,140,248,0.1)',
                    borderWidth: 2,
                    pointRadius: 4,
                    tension: 0.4,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left' },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false } }
            }
        }
    });
}

function renderPEChart(labels, data, color) {
    const ctx = document.getElementById('peChart').getContext('2d');
    if (chartPE) chartPE.destroy();
    const med = calcMedian(data);
    const medLine = labels.map(() => med);
    chartPE = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                {
                    label: 'P/E TTM',
                    data,
                    borderColor: color,
                    backgroundColor: color + '18',
                    borderWidth: 2,
                    pointRadius: 3,
                    tension: 0.3,
                    fill: true,
                    spanGaps: true
                },
                {
                    label: `Trung v\u1ecb (${med ? med.toFixed(1) + 'x' : '-'})`,
                    data: medLine,
                    borderColor: '#f59e0b',
                    borderWidth: 1.5,
                    borderDash: [5, 5],
                    pointRadius: 0,
                    fill: false
                }
            ]
        },
        options: { ...CHART_DEFAULTS }
    });
}

function renderPBChart(labels, data, color) {
    const ctx = document.getElementById('pbChart').getContext('2d');
    if (chartPB) chartPB.destroy();
    const med = calcMedian(data);
    const medLine = labels.map(() => med);
    chartPB = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                {
                    label: 'P/B TTM',
                    data,
                    borderColor: '#818cf8',
                    backgroundColor: 'rgba(129,140,248,0.12)',
                    borderWidth: 2,
                    pointRadius: 3,
                    tension: 0.3,
                    fill: true,
                    spanGaps: true
                },
                {
                    label: `Trung v\u1ecb (${med ? med.toFixed(2) + 'x' : '-'})`,
                    data: medLine,
                    borderColor: '#f59e0b',
                    borderWidth: 1.5,
                    borderDash: [5, 5],
                    pointRadius: 0,
                    fill: false
                }
            ]
        },
        options: { ...CHART_DEFAULTS }
    });
}

// ═══════════════════════════════════════════════════════════
// DETAILED EVALUATION COMPONENTS
// ═══════════════════════════════════════════════════════════

function renderQuarterlyAndYTDEvaluation(ticker, liveData, localJson, cfg) {
    const container = document.getElementById('quarterly-ytd-evaluation-container');
    let isrecs = liveData.incomeYears?.quarters || [];

    if (isrecs.length < 5 && localJson?.income_quarterly?.length >= 5) {
        isrecs = localJson.income_quarterly.map(r => ({
            yearReport: r.yearReport,
            quarter: r.quarter,
            [cfg.incomeField]: (r.revenue || 0) * 1e9,
            [cfg.npat]: (r.npat || 0) * 1e9,
        }));
    }

    const quarters = isrecs
        .filter(q => q.quarter >= 1 && q.quarter <= 4)
        .sort((a, b) => a.yearReport !== b.yearReport ? a.yearReport - b.yearReport : a.quarter - b.quarter);

    if (quarters.length < 5) {
        container.innerHTML = `<div class="loading-state">Kh\u00f4ng \u0111\u1ee7 d\u1eef li\u1ec7u qu\u00fd \u0111\u1ec3 t\u00ednh to\u00e1n t\u0103ng tr\u01b0\u1edfng.</div>`;
        return;
    }

    const latestQ = quarters[quarters.length - 1];
    const prevQ = quarters[quarters.length - 2];
    const sameQLastYear = quarters.find(q => q.yearReport === latestQ.yearReport - 1 && q.quarter === latestQ.quarter);

    const revFieldName = cfg.incomeField;
    const npatFieldName = cfg.npat;

    const latestRev = latestQ[revFieldName] || 0;
    const prevRev = prevQ[revFieldName] || 0;
    const sameLastYearRev = sameQLastYear ? sameQLastYear[revFieldName] : 0;

    const latestNpat = latestQ[npatFieldName] || 0;
    const prevNpat = prevQ[npatFieldName] || 0;
    const sameLastYearNpat = sameQLastYear ? sameQLastYear[npatFieldName] : 0;

    const qoqRev = prevRev > 0 ? ((latestRev - prevRev) / prevRev * 100) : null;
    const yoyRev = sameLastYearRev > 0 ? ((latestRev - sameLastYearRev) / sameLastYearRev * 100) : null;

    const qoqNpat = prevNpat > 0 ? ((latestNpat - prevNpat) / prevNpat * 100) : null;
    const yoyNpat = sameLastYearNpat > 0 ? ((latestNpat - sameLastYearNpat) / sameLastYearNpat * 100) : null;

    const currentYear = latestQ.yearReport;
    const currentYearQuarters = quarters.filter(q => q.yearReport === currentYear);
    const lastYearQuarters = quarters.filter(q => q.yearReport === currentYear - 1 && q.quarter <= latestQ.quarter);

    const ytdRev = currentYearQuarters.reduce((sum, q) => sum + (q[revFieldName] || 0), 0);
    const lastYtdRev = lastYearQuarters.reduce((sum, q) => sum + (q[revFieldName] || 0), 0);

    const ytdNpat = currentYearQuarters.reduce((sum, q) => sum + (q[npatFieldName] || 0), 0);
    const lastYtdNpat = lastYearQuarters.reduce((sum, q) => sum + (q[npatFieldName] || 0), 0);

    const yoyYtdRev = lastYtdRev > 0 ? ((ytdRev - lastYtdRev) / lastYtdRev * 100) : null;
    const yoyYtdNpat = lastYtdNpat > 0 ? ((ytdNpat - lastYtdNpat) / lastYtdNpat * 100) : null;

    const formatGrowth = (val) => {
        if (val === null || val === undefined) return '<span class="badge-growth flat">0.0%</span>';
        if (val > 0) return `<span class="badge-growth up">\u25b2 +${val.toFixed(1)}%</span>`;
        if (val < 0) return `<span class="badge-growth down">\u25bc ${val.toFixed(1)}%</span>`;
        return '<span class="badge-growth flat">0.0%</span>';
    };

    let commentary = '';
    const nameRev = cfg.incomeLabel;

    if (yoyNpat && yoyNpat > 5) {
        commentary = `KQKD qu\u00fd ${latestQ.quarter}/${latestQ.yearReport} c\u1ee7a ${ticker} ghi nh\u1eadn s\u1ef1 t\u0103ng tr\u01b0\u1edfng t\u00edch c\u1ef1c, v\u1edbi l\u1ee3i nhu\u1eadn sau thu\u1ebf \u0111\u1ea1t ${formatNumber(latestNpat/1e9, 1)} t\u1ef7 \u0111\u1ed3ng (${formatGrowth(yoyNpat)} so v\u1edbi c\u00f9ng k\u1ef3). `;
    } else if (yoyNpat && yoyNpat < -5) {
        commentary = `K\u1ebft qu\u1ea3 kinh doanh qu\u00fd g\u1ea7n nh\u1ea5t cho th\u1ea5y t\u00edn hi\u1ec7u ch\u1eadm l\u1ea1i, LNST gi\u1ea3m ${formatGrowth(yoyNpat)} so v\u1edbi c\u00f9ng k\u1ef3 do \u00e1p l\u1ef1c bi\u00ean l\u1ee3i nhu\u1eadn ho\u1eb7c chi ph\u00ed t\u0103ng cao. `;
    } else {
        commentary = `K\u1ebft qu\u1ea3 kinh doanh qu\u00fd g\u1ea7n nh\u1ea5t c\u1ee7a ${ticker} duy tr\u00ec \u1edf m\u1ee9c \u1ed5n \u0111\u1ecbnh. `;
    }

    if (yoyYtdNpat) {
        commentary += `L\u0169y k\u1ebf t\u1eeb \u0111\u1ea7u n\u0103m (YTD), LNST \u0111\u1ea1t ${formatNumber(ytdNpat/1e9, 1)} t\u1ef7 \u0111\u1ed3ng, thay \u0111\u1ed5i ${yoyYtdNpat.toFixed(1)}% so v\u1edbi c\u00f9ng k\u1ef3 n\u0103m tr\u01b0\u1edbc.`;
    }

    container.innerHTML = `
        <div class="quarterly-ytd-grid" style="grid-template-columns: repeat(auto-fit, minmax(320px, 1fr))">
            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">Doanh thu Qu\u00fd g\u1ea7n nh\u1ea5t (${latestQ.quarter}Q/${latestQ.yearReport})</span>
                    <span style="font-size:0.75rem;color:var(--text-dim)">\u0110\u01a1n v\u1ecb: T\u1ef7 VND</span>
                </div>
                <div class="q-ytd-metrics">
                    <div class="q-metric-row">
                        <span class="q-metric-label">${nameRev}</span>
                        <span class="q-metric-value">${formatNumber(latestRev/1e9, 1)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">T\u0103ng tr\u01b0\u1edfng QoQ (so qu\u00fd tr\u01b0\u1edbc)</span>
                        <span>${formatGrowth(qoqRev)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">T\u0103ng tr\u01b0\u1edfng YoY (c\u00f9ng k\u1ef3)</span>
                        <span>${formatGrowth(yoyRev)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">L\u0169y k\u1ebf YTD c\u1ea3 n\u0103m</span>
                        <span class="q-metric-value">${formatNumber(ytdRev/1e9, 1)}</span>
                    </div>
                </div>
            </div>

            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">L\u1ee3i nhu\u1eadn Qu\u00fd g\u1ea7n nh\u1ea5t</span>
                    <span style="font-size:0.75rem;color:var(--text-dim)">\u0110\u01a1n v\u1ecb: T\u1ef7 VND</span>
                </div>
                <div class="q-ytd-metrics">
                    <div class="q-metric-row">
                        <span class="q-metric-label">L\u1ee3i nhu\u1eadn sau thu\u1ebf</span>
                        <span class="q-metric-value">${formatNumber(latestNpat/1e9, 1)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">T\u0103ng tr\u01b0\u1edfng QoQ (so qu\u00fd tr\u01b0\u1edbc)</span>
                        <span>${formatGrowth(qoqNpat)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">T\u0103ng tr\u01b0\u1edfng YoY (c\u00f9ng k\u1ef3)</span>
                        <span>${formatGrowth(yoyNpat)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">L\u0169y k\u1ebf LNST YTD c\u1ea3 n\u0103m</span>
                        <span class="q-metric-value">${formatNumber(ytdNpat/1e9, 1)} (${formatGrowth(yoyYtdNpat)})</span>
                    </div>
                </div>
            </div>
        </div>
        <div class="q-commentary-box">
            <strong>\u{1F4CB} Nh\u1eadn \u0111\u1ecbnh nhanh Earning Release:</strong> ${commentary}
        </div>
    `;
}

function generateChartCommentaries(ticker, annualYears, ttmQuarters, cfg, localJson) {
    const revEl = document.getElementById('analysis-text-rev-npat');
    if (revEl && annualYears.length >= 2) {
        const last = annualYears[annualYears.length - 1];
        const prev = annualYears[annualYears.length - 2];
        const revField = cfg.incomeField;
        const npatField = cfg.npat;
        const growthRev = prev[revField] > 0 ? ((last[revField] - prev[revField]) / prev[revField] * 100) : 0;
        const growthNpat = prev[npatField] > 0 ? ((last[npatField] - prev[npatField]) / prev[npatField] * 100) : 0;

        revEl.innerHTML = `
            Doanh thu n\u0103m g\u1ea7n nh\u1ea5t \u0111\u1ea1t <strong>${formatNumber(last[revField]/1e12, 1)} ngh\u00ecn t\u1ef7</strong>, t\u0103ng tr\u01b0\u1edfng <strong>${growthRev.toFixed(1)}% YoY</strong>. <br>
            LNST t\u01b0\u01a1ng \u1ee9ng \u0111\u1ea1t <strong>${formatNumber(last[npatField]/1e12, 1)} ngh\u00ecn t\u1ef7</strong> (${growthNpat >= 0 ? '+' : ''}${growthNpat.toFixed(1)}% YoY). <br>
            Nh\u00ecn chung, doanh nghi\u1ec7p \u0111ang duy tr\u00ec xu h\u01b0\u1edbng ${growthNpat > 0 ? 't\u0103ng tr\u01b0\u1edfng t\u00edch c\u1ef1c' : '\u0111i ngang/s\u1ee5t gi\u1ea3m'} v\u1ec1 m\u1eb7t hi\u1ec7u qu\u1ea3 kinh doanh c\u1ed1t l\u00f5i.
        `;
    }

    const epsEl = document.getElementById('analysis-text-eps-equity');
    if (epsEl && localJson?.ratios) {
        const roeArr = localJson.ratios.roe || [];
        const epsArr = localJson.data?.eps || [];

        if (roeArr.length > 0) {
            const lastRoe = (roeArr[roeArr.length - 1] * 100).toFixed(1);
            const lastEps = epsArr[epsArr.length - 1] ? formatNumber(epsArr[epsArr.length - 1]) : '-';

            epsEl.innerHTML = `
                T\u1ef7 su\u1ea5t sinh l\u1ee3i tr\u00ean v\u1ed1n ch\u1ee7 s\u1edf h\u1eefu (ROE) \u0111\u1ea1t <strong>${lastRoe}%</strong> \u1edf n\u0103m g\u1ea7n nh\u1ea5t. <br>
                Thu nh\u1eadp tr\u00ean m\u1ed7i c\u1ed5 ph\u1ea7n (EPS) t\u01b0\u01a1ng \u1ee9ng l\u00e0 <strong>${lastEps} VND/CP</strong>. <br>
                M\u1ee9c hi\u1ec7u qu\u1ea3 ROE ${parseFloat(lastRoe) > 15 ? '\u1edf m\u1ee9c cao (>15%), ch\u1ee9ng t\u1ecf kh\u1ea3 n\u0103ng sinh l\u1eddi hi\u1ec7u qu\u1ea3 c\u1ee7a ngu\u1ed3n v\u1ed1n.' : '\u1edf m\u1ee9c trung b\u00ecnh th\u1ea5p, c\u1ea7n theo d\u00f5i xu h\u01b0\u1edbng t\u00e1i c\u01a1 c\u1ea5u t\u00e0i s\u1ea3n.'}
            `;
        }
    } else {
        epsEl.innerHTML = `Hi\u1ec7u qu\u1ea3 s\u1eed d\u1ee5ng v\u1ed1n ch\u1ee7 s\u1edf h\u1eefu (ROE) v\u00e0 EPS \u0111\u01b0\u1ee3c t\u1ed1i \u01b0u h\u00f3a d\u1ef1a tr\u00ean c\u01a1 c\u1ea5u n\u1ee3 v\u00e0 hi\u1ec7u n\u0103ng ph\u00e2n b\u1ed5 v\u1ed1n t\u00e0i s\u1ea3n c\u1ee7a doanh nghi\u1ec7p.`;
    }

    const peEl = document.getElementById('analysis-text-pe');
    const pbEl = document.getElementById('analysis-text-pb');

    if (ttmQuarters.length > 0) {
        const lastRatio = ttmQuarters[ttmQuarters.length - 1];
        const peVals = ttmQuarters.map(q => q.pe).filter(v => v > 0);
        const pbVals = ttmQuarters.map(q => q.pb).filter(v => v > 0);

        const peMedian = calcMedian(peVals);
        const pbMedian = calcMedian(pbVals);

        if (peEl && lastRatio.pe) {
            const diffPe = ((lastRatio.pe - peMedian) / peMedian * 100).toFixed(1);
            peEl.innerHTML = `
                H\u1ec7 s\u1ed1 P/E trailing hi\u1ec7n t\u1ea1i l\u00e0 <strong>${lastRatio.pe.toFixed(1)}x</strong>. <br>
                Trung v\u1ecb l\u1ecbch s\u1eed c\u1ee7a doanh nghi\u1ec7p l\u00e0 <strong>${peMedian ? peMedian.toFixed(1) + 'x' : '-'}</strong>. <br>
                \u0110\u1ecbnh gi\u00e1 P/E \u0111ang ${lastRatio.pe > peMedian ? `cao h\u01a1n trung v\u1ecb (${diffPe}%), ph\u1ea3n \u00e1nh k\u1ef3 v\u1ecdng t\u0103ng tr\u01b0\u1edfng l\u1edbn ho\u1eb7c gi\u00e1 \u0111ang \u1edf v\u00f9ng \u0111\u1eaft.` : `th\u1ea5p h\u01a1n trung v\u1ecb (${Math.abs(diffPe)}%), cho th\u1ea5y bi\u00ean an to\u00e0n \u0111\u1ecbnh gi\u00e1 t\u01b0\u01a1ng \u0111\u1ed1i r\u1ebb.`}
            `;
        }

        if (pbEl && lastRatio.pb) {
            const diffPb = ((lastRatio.pb - pbMedian) / pbMedian * 100).toFixed(1);
            pbEl.innerHTML = `
                H\u1ec7 s\u1ed1 P/B trailing hi\u1ec7n t\u1ea1i l\u00e0 <strong>${lastRatio.pb.toFixed(2)}x</strong>. <br>
                Trung v\u1ecb l\u1ecbch s\u1eed c\u1ee7a doanh nghi\u1ec7p l\u00e0 <strong>${pbMedian ? pbMedian.toFixed(2) + 'x' : '-'}</strong>. <br>
                V\u00f9ng \u0111\u1ecbnh gi\u00e1 P/B hi\u1ec7n t\u1ea1i ${lastRatio.pb > pbMedian ? `n\u1eb1m tr\u00ean trung v\u1ecb l\u1ecbch s\u1eed (${diffPb}%)` : `n\u1eb1m d\u01b0\u1edbi trung v\u1ecb l\u1ecbch s\u1eed (${Math.abs(diffPb)}%)`} l\u00e0 c\u01a1 s\u1edf \u0111\u1ec3 c\u00e2n nh\u1eafc t\u00edch l\u0169y.
            `;
        }
    }
}

// ═══════════════════════════════════════════════════════════
// FINANCIAL SNAPSHOT TABLE
// ═══════════════════════════════════════════════════════════
function renderFinancialSnapshotTable(localJson, sectorKey) {
    const card  = document.getElementById('web-financial-snapshot-card');
    const tbody = document.getElementById('web-financial-snapshot-tbody');
    if (!card || !tbody) return;

    const data = localJson?.data;
    if (!data || !data.years || !data.years.length) { card.style.display = 'none'; return; }

    card.style.display = '';
    const years = data.years.slice(-6);

    for (let i = 0; i < 6; i++) {
        const th = document.getElementById(`fs-th-${i}`);
        if (th) {
            if (i < years.length) {
                th.textContent = years[i];
                th.style.color = i >= years.length - 3 ? '#60a5fa' : '';
            } else {
                th.textContent = '-';
            }
        }
    }

    const revenue = data.revenue || [];
    const npat = data.npat || [];
    const ratios = localJson?.ratios || {};
    const grossMargin = ratios.gross_margin || [];
    const roe = ratios.roe || [];
    const debtEq = ratios.debt_to_equity || [];

    const fmtN = (v) => v != null ? Math.round(v / 1e9).toLocaleString('vi-VN') : '-';
    const fmtPct = (v) => v != null ? (v * 100).toFixed(1) + '%' : '-';
    const fmtX = (v) => v != null ? v.toFixed(2) + 'x' : '-';

    const rows = [
        { label: 'Doanh thu thu\u1ea7n (t\u1ef7 VND)', arr: revenue, fmt: (v) => fmtN(v) },
        { label: 'LNST (t\u1ef7 VND)', arr: npat, fmt: (v) => fmtN(v) },
        { label: 'Bi\u00ean l\u1ee3i nhu\u1eadn g\u1ed9p (%)', arr: grossMargin, fmt: fmtPct },
        { label: 'ROE (%)', arr: roe, fmt: fmtPct },
        { label: 'N\u1ee3/V\u1ed1n ch\u1ee7 (x)', arr: debtEq, fmt: fmtX },
    ];

    tbody.innerHTML = rows.map(r => {
        const cells = years.map((y, i) => {
            const yrIdx = data.years.indexOf(y);
            const val = yrIdx !== -1 ? r.arr[yrIdx] : undefined;
            return `<td>${val != null ? r.fmt(val) : '-'}</td>`;
        }).join('');
        return `<tr><td>${r.label}</td>${cells}</tr>`;
    }).join('');
}

// ═══════════════════════════════════════════════════════════
// VALUATION SNAPSHOT
// ═══════════════════════════════════════════════════════════
function renderValuationSnapshot(localJson) {
    const val = localJson?.valuation;
    const currPrice = localJson?.currentPrice;
    if (!val || !currPrice) return;
    const fmt = (n) => n ? Math.round(n).toLocaleString('vi-VN') + ' \u0111' : '-';
    const upside = val.upside ?? 0;
    const recText  = val.recommend ?? (upside >= 15 ? 'MUA' : upside < -5 ? 'BAN' : 'THEO D\u00d5I');
    const recColor = recText === 'MUA' ? '#10b981' : recText === 'BAN' ? '#ef4444' : '#f59e0b';
    const el = (id) => document.getElementById(id);
    if (el('snap-curr-price'))   el('snap-curr-price').textContent   = fmt(currPrice);
    if (el('snap-target-price')) el('snap-target-price').textContent = fmt(val.weightedTarget);
    if (el('snap-upside')) {
        el('snap-upside').textContent = (upside >= 0 ? '+' : '') + upside.toFixed(1) + '%';
        el('snap-upside').style.color = upside >= 10 ? '#10b981' : upside >= 0 ? '#f59e0b' : '#ef4444';
    }
    if (el('snap-recommend')) {
        el('snap-recommend').textContent = recText;
        el('snap-recommend').style.background = recColor;
    }
    if (el('snap-coe')) el('snap-coe').textContent = (val.COE != null ? val.COE.toFixed(2) : '-') + '%';
    if (el('snap-bvps')) el('snap-bvps').textContent = val.bvpsBase ? fmt(val.bvpsBase) : '-';
    if (el('snap-pb-median')) el('snap-pb-median').textContent = val.pbMedian ? val.pbMedian.toFixed(2) + 'x' : '-';
    if (el('snap-pb-attractive')) el('snap-pb-attractive').textContent = val.pbAttractive ? val.pbAttractive.toFixed(2) + 'x' : '-';
    if (el('snap-pb-target')) el('snap-pb-target').textContent = val.pbOver ? val.pbOver.toFixed(2) + 'x' : '-';
    if (el('snap-pe-median')) el('snap-pe-median').textContent = val.peMedian ? val.peMedian.toFixed(1) + 'x' : (val.peTarget ? '-' : '-');
}

// ═══════════════════════════════════════════════════════════
// QUALITATIVE RENDERERS
// ═══════════════════════════════════════════════════════════
function renderThesisAndRisks(thesisList, risksList) {
    const defaultThesis = [
        'Doanh nghi\u1ec7p c\u00f3 v\u1ecb th\u1ebf c\u1ea1nh tranh t\u1ed1t trong ng\u00e0nh v\u1edbi n\u1ec1n t\u1ea3ng t\u00e0i ch\u00ednh l\u00e0nh m\u1ea1nh.',
        'D\u00f2ng ti\u1ec1n ho\u1ea1t \u0111\u1ed9ng \u1ed5n \u0111\u1ecbnh v\u00e0 c\u01a1 c\u1ea5u t\u00e0i ch\u00ednh h\u1ed7 tr\u1ee3 chi\u1ebfn l\u01b0\u1ee3c m\u1edf r\u1ed9ng d\u00e0i h\u1ea1n.',
        '\u0110\u1ed9i ng\u0169 qu\u1ea3n l\u00fd c\u00f3 kinh nghi\u1ec7m v\u1edbi track record th\u1ef1c thi k\u1ebf ho\u1ea1ch kinh doanh hi\u1ec7u qu\u1ea3.',
    ];
    const defaultRisks = [
        'R\u1ee7i ro kinh t\u1ebf v\u0129 m\u00f4 v\u00e0 bi\u1ebfn \u0111\u1ed9ng gi\u00e1 c\u1ea3 nguy\u00ean v\u1eadt li\u1ec7u \u0111\u1ea7u v\u00e0o.',
        'C\u1ea1nh tranh trong ng\u00e0nh ng\u00e0y c\u00e0ng gay g\u1eaft \u0111\u00f2i h\u1ecfi \u0111\u1ea7u t\u01b0 li\u00ean t\u1ee5c v\u00e0o n\u0103ng l\u1ef1c s\u1ea3n xu\u1ea5t.',
        'R\u1ee7i ro th\u1ef1c thi chi\u1ebfn l\u01b0\u1ee3c m\u1edf r\u1ed9ng v\u00e0 ph\u1ee5 thu\u1ed9c v\u00e0o chu k\u1ef3 ng\u00e0nh th\u00e9p.',
    ];

    const thesis = thesisList?.length > 0 ? thesisList : defaultThesis;
    const risks  = risksList?.length  > 0 ? risksList  : defaultRisks;

    document.getElementById('thesis-list').innerHTML = thesis.map(t => `<li>${t}</li>`).join('');
    document.getElementById('risks-list').innerHTML  = risks.map(r => `<li>${r}</li>`).join('');
}

function renderMoatScorecard(moats, color) {
    const defaultMoats = {
        'Intangible Assets': { score: 3, desc: 'Th\u01b0\u01a1ng hi\u1ec7u v\u00e0 uy t\u00edn trong ng\u00e0nh \u0111\u01b0\u1ee3c x\u00e2y d\u1ef1ng qua nhi\u1ec1u n\u0103m.' },
        'Cost Advantage':    { score: 3, desc: 'T\u1ed1i \u01b0u h\u00f3a quy tr\u00ecnh v\u1eadn h\u00e0nh gi\u00fap duy tr\u00ec bi\u00ean l\u1ee3i nhu\u1eadn c\u1ea1nh tranh.' },
        'Switching Cost':    { score: 2, desc: 'M\u1ee9c \u0111\u1ed9 g\u1eafn k\u1ebft v\u1edbi kh\u00e1ch h\u00e0ng \u1edf m\u1ee9c trung b\u00ecnh.' },
        'Efficient Scale':   { score: 3, desc: 'Quy m\u00f4 ph\u00f9 h\u1ee3p \u0111\u1ec3 v\u1eadn h\u00e0nh hi\u1ec7u qu\u1ea3 trong th\u1ecb tr\u01b0\u1eddng n\u1ed9i \u0111\u1ecba.' },
        'Network Effect':    { score: 2, desc: 'Hi\u1ec7u \u1ee9ng m\u1ea1ng l\u01b0\u1edbi c\u00f2n h\u1ea1n ch\u1ebf trong m\u00f4 h\u00ecnh kinh doanh hi\u1ec7n t\u1ea1i.' },
    };
    const data = moats || defaultMoats;
    const container = document.getElementById('moat-scorecard-list');
    container.innerHTML = '';
    Object.entries(data).forEach(([key, item]) => {
        const pct = (item.score / 5) * 100;
        const div = document.createElement('div');
        div.className = 'moat-item';
        div.innerHTML = `
            <div class="moat-header">
                <span>${key}</span>
                <span class="moat-score">${item.score}/5</span>
            </div>
            <div class="moat-bar-bg">
                <div class="moat-bar-fill" style="width:0%;background:linear-gradient(90deg,${color}80,${color})"></div>
            </div>
            <div class="moat-desc">${item.desc}</div>
        `;
        container.appendChild(div);
        setTimeout(() => div.querySelector('.moat-bar-fill').style.width = pct + '%', 50);
    });
}

function renderValuationScenarios(valData, currentPrice, details, latestRatio, cfg) {
    let bear = '-', base = '-', bull = '-';

    if (valData) {
        bear = formatNumber(valData.bear || 0) + ' VND';
        base = formatNumber(valData.base || 0) + ' VND';
        bull = formatNumber(valData.bull || 0) + ' VND';
    } else if (details.targetPrice) {
        const target = details.targetPrice;
        bear = formatNumber(Math.round(target * 0.75)) + ' VND';
        base = formatNumber(Math.round(target)) + ' VND';
        bull = formatNumber(Math.round(target * 1.2)) + ' VND';
    }

    document.getElementById('val-bear').textContent = bear;
    document.getElementById('val-base').textContent = base;
    document.getElementById('val-bull').textContent = bull;
}

function renderPESTLE(pestleData) {
    const defaultPestle = {
        Political:     'Ch\u00ednh s\u00e1ch v\u0129 m\u00f4 v\u00e0 \u0111\u1ecbnh h\u01b0\u1edbng ph\u00e1t tri\u1ec3n ng\u00e0nh c\u1ee7a Ch\u00ednh ph\u1ee7 t\u00e1c \u0111\u1ed9ng \u0111\u1ebfn m\u00f4i tr\u01b0\u1eddng kinh doanh.',
        Economic:      'T\u0103ng tr\u01b0\u1edfng GDP, l\u00e3i su\u1ea5t v\u00e0 l\u1ea1m ph\u00e1t \u1ea3nh h\u01b0\u1edfng \u0111\u1ebfn nhu c\u1ea7u v\u00e0 chi ph\u00ed ho\u1ea1t \u0111\u1ed9ng.',
        Social:        'Thay \u0111\u1ed5i xu h\u01b0\u1edbng ti\u00eau d\u00f9ng v\u00e0 nh\u00e2n kh\u1ea9u h\u1ecdc t\u1ea1o c\u1ea3 c\u01a1 h\u1ed9i v\u00e0 th\u00e1ch th\u1ee9c.',
        Technological: 'Chuy\u1ec3n \u0111\u1ed5i s\u1ed1 v\u00e0 t\u1ef1 \u0111\u1ed9ng h\u00f3a l\u00e0 xu h\u01b0\u1edbng quan tr\u1ecdng c\u1ea7n \u0111\u1ea7u t\u01b0 b\u1eaft k\u1ecbp.',
        Legal:         'Tu\u00e2n th\u1ee7 c\u00e1c quy \u0111\u1ecbnh ph\u00e1p lu\u1eadt v\u00e0 ti\u00eau chu\u1ea9n ng\u00e0nh l\u00e0 y\u00eau c\u1ea7u c\u01a1 b\u1ea3n.',
        Environmental: '\u00c1p l\u1ef1c ESG v\u00e0 b\u00e1o c\u00e1o ph\u00e1t tri\u1ec3n b\u1ec1n v\u1eefng ng\u00e0y c\u00e0ng tr\u1edf th\u00e0nh y\u00eau c\u1ea7u c\u1ee7a nh\u00e0 \u0111\u1ea7u t\u01b0.',
    };

    let data = defaultPestle;
    if (pestleData) {
        if (Array.isArray(pestleData)) {
            data = {};
            pestleData.forEach(item => {
                if (item.factor) data[item.factor] = item.content;
            });
        } else {
            data = pestleData;
        }
    }

    const pestleColors = { Political: '#3b82f6', Economic: '#10b981', Social: '#f59e0b', Technological: '#a855f7', Legal: '#ef4444', Environmental: '#22c55e' };
    const pestleSentiment = { Political: 'neutral', Economic: 'neutral', Social: 'positive', Technological: 'positive', Legal: 'neutral', Environmental: 'neutral' };

    const container = document.getElementById('pestle-grid-container');
    container.innerHTML = Object.entries(data).map(([key, text]) => `
        <div class="pestle-box">
            <div class="pestle-box-header">
                <span>${key}</span>
                <span class="pestle-tag ${pestleSentiment[key] || 'neutral'}" style="background:${(pestleColors[key] || '#64748b') + '25'};color:${pestleColors[key] || '#64748b'}">${key[0]}</span>
            </div>
            <div class="pestle-box-text">${text}</div>
        </div>
    `).join('');
}

function renderCommentary(comments) {
    const data = comments || {};
    document.getElementById('comment-business').textContent  = data.businessModel || data.overall || 'Ch\u01b0a c\u00f3 \u0111\u00e1nh gi\u00e1 m\u00f4 h\u00ecnh kinh doanh.';
    document.getElementById('comment-financial').textContent = data.financialPerformance || data.financial || 'Ch\u01b0a c\u00f3 \u0111\u00e1nh gi\u00e1 s\u1ee9c kh\u1ecfe t\u00e0i ch\u00ednh.';
    document.getElementById('comment-valuation').textContent = data.valuationText || data.valuation || 'Ch\u01b0a c\u00f3 \u0111\u00e1nh gi\u00e1 \u0111\u1ecbnh gi\u00e1.';
}

function renderFinancialTable(annualYears, cfg, labels, localJson) {
    const table = document.getElementById('financial-table');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');
    if (!table || !thead || !tbody) return;

    const mergedYears = [...labels];
    const forecastYears = [2026, 2027];
    forecastYears.forEach(fy => {
        if (!mergedYears.includes(fy.toString()) && !mergedYears.includes(fy.toString() + 'F')) {
            mergedYears.push(fy.toString() + 'F');
        }
    });

    thead.innerHTML = '<th>Ch\u1ec9 ti\u00eau</th>' + mergedYears.map(y => `<th>${y}</th>`).join('');
    tbody.innerHTML = '';

    const rows = [
        { label: cfg.incomeLabel + ' (t\u1ef7)', field: cfg.incomeField, div: 1e9 },
        { label: cfg.income2Label + ' (t\u1ef7)', field: cfg.incomeField2, div: 1e9 },
        { label: 'LNST (t\u1ef7)', field: cfg.npat, div: 1e9, highlight: true },
    ].filter(r => r.field);

    rows.forEach(row => {
        const tr = document.createElement('tr');
        let tdHtml = `<td><strong>${row.label}</strong></td>`;

        mergedYears.forEach(yearLabel => {
            const isForecast = yearLabel.endsWith('F');
            const numericYear = parseInt(yearLabel);
            let rawVal = null;

            if (!isForecast) {
                const histYearData = annualYears.find(yr => yr.yearReport === numericYear);
                if (histYearData) {
                    rawVal = histYearData[row.field];
                }
            } else {
                const fcData = localJson?.data;
                if (fcData && fcData.years) {
                    const fcIdx = fcData.years.indexOf(numericYear);
                    if (fcIdx !== -1) {
                        const fcField = row.field === cfg.incomeField ? 'revenue'
                            : row.field === cfg.npat ? 'npat'
                            : row.field === cfg.incomeField2 ? 'grossProfit'
                            : null;
                        if (fcField && fcData[fcField]) {
                            rawVal = fcData[fcField][fcIdx];
                        }
                        if (rawVal != null && rawVal < 1e9) {
                            rawVal = rawVal * 1e9;
                        }
                    }
                }
            }

            const fmt = rawVal ? formatNumber(Math.round(rawVal / (row.div || 1e9))) : '-';
            tdHtml += `<td class="${row.highlight ? 'highlight' : ''}">${fmt}</td>`;
        });

        tr.innerHTML = tdHtml;
        tbody.appendChild(tr);
    });

    document.getElementById('table-title').textContent = 'B\u1ea3ng k\u1ebft qu\u1ea3 kinh doanh (theo n\u0103m, t\u1ef7 VND)';
}

// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════
function formatNumber(num, decimals = 0) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Number(num).toLocaleString('vi-VN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
