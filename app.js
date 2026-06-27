/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app.js
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
// SECTOR CONFIGURATION
// ═══════════════════════════════════════════════════════════
const SECTOR_CONFIG = {
    banks: {
        label: 'Ngân Hàng',
        labelEn: 'Banks',
        icon: '🏦',
        color: '#3b82f6',
        // income statement field for P&L chart
        incomeField: 'isb27',        // Net Interest Income
        incomeField2: 'isb38',       // Total Operating Income (second line)
        npat: 'isa20',
        incomeLabel: 'Thu nhập lãi thuần (NII)',
        income2Label: 'Tổng thu nhập HĐ (TOI)',
        npatLabel: 'Lợi nhuận sau thuế',
        showPE: false,               // P/E less relevant for banks
        primaryValuation: 'P/B',
        extraMetrics: (ratios, localJson) => {
            const latest = ratios?.[ratios.length - 1] || {};
            // Ưu tiên đọc từ ratios tự tính toán trong localJson
            const localRatios = localJson?.ratios || {};
            
            const getVal = (key, rawLive, factor = 100, decimals = 2) => {
                if (localRatios[key] !== undefined && localRatios[key] !== null) {
                    const arr = localRatios[key];
                    const lastVal = arr[arr.length - 1];
                    return lastVal !== null && lastVal !== undefined ? (lastVal * (factor === 1 ? 1 : 100)).toFixed(decimals) + '%' : '-';
                }
                return rawLive ? (rawLive * factor).toFixed(decimals) + '%' : '-';
            };

            return [
                { label: 'NIM (Net Interest Margin)', value: getVal('nim', latest.netInterestMargin, 100, 2), desc: 'Biên lãi ròng' },
                { label: 'ROE', value: getVal('roe', latest.roe, 100, 1), desc: 'Tỷ suất vốn chủ' },
                { label: 'ROA', value: getVal('roa', latest.roa, 100, 2), desc: 'Tỷ suất tổng tài sản' },
                { label: 'NPL Ratio', value: getVal('npl', latest.npl, 100, 2), desc: 'Tỷ lệ nợ xấu' },
                { label: 'LDR', value: getVal('ldr', latest.ldrLoanDepositRatio, 100, 1), desc: 'Tỷ lệ tín dụng/tiền gửi' },
                { label: 'CASA Ratio', value: getVal('casa', latest.casaRatio, 100, 1), desc: 'Tiền gửi không kỳ hạn' },
            ];
        }
    },
    materials: {
        label: 'Thép & Vật liệu',
        labelEn: 'Basic Resources',
        icon: '🔩',
        color: '#f97316',
        incomeField: 'isa3',         // Net Revenue
        incomeField2: 'isa5',        // Gross Profit
        npat: 'isa20',
        incomeLabel: 'Doanh thu thuần',
        income2Label: 'Lợi nhuận gộp',
        npatLabel: 'Lợi nhuận sau thuế',
        showPE: true,
        primaryValuation: 'P/E',
        extraMetrics: (ratios, localJson) => {
            const latest = ratios?.[ratios.length - 1] || {};
            const localRatios = localJson?.ratios || {};
            
            const getVal = (key, rawLive, factor = 100, decimals = 1, suffix = '%') => {
                if (localRatios[key] !== undefined && localRatios[key] !== null) {
                    const arr = localRatios[key];
                    const lastVal = arr[arr.length - 1];
                    return lastVal !== null && lastVal !== undefined ? (lastVal * (factor === 1 ? 1 : 100)).toFixed(decimals) + suffix : '-';
                }
                return rawLive ? (rawLive * factor).toFixed(decimals) + suffix : '-';
            };

            return [
                { label: 'Biên gộp (Gross Margin)', value: getVal('gross_margin', latest.grossMargin, 100, 1), desc: 'Lợi nhuận gộp / Doanh thu' },
                { label: 'ROE', value: getVal('roe', latest.roe, 100, 1), desc: 'Tỷ suất vốn chủ sở hữu' },
                { label: 'Nợ/Vốn chủ', value: getVal('debt_to_equity', latest.debtToEquity, 1, 2, 'x'), desc: 'Đòn bẩy tài chính' },
                { label: 'EBITDA Margin', value: getVal('ebitMargin', latest.ebitMargin, 100, 1), desc: 'Biên EBITDA' },
            ];
        }
    },
    realestate: {
        label: 'Bất Động Sản',
        labelEn: 'Real Estate',
        icon: '🏘️',
        color: '#a855f7',
        incomeField: 'isa3',
        incomeField2: 'isa5',
        npat: 'isa20',
        incomeLabel: 'Doanh thu thuần',
        income2Label: 'Lợi nhuận gộp',
        npatLabel: 'Lợi nhuận sau thuế',
        showPE: true,
        primaryValuation: 'P/B',
        extraMetrics: (ratios, localJson) => {
            const latest = ratios?.[ratios.length - 1] || {};
            const localRatios = localJson?.ratios || {};
            
            const getVal = (key, rawLive, factor = 100, decimals = 1, suffix = '%') => {
                if (localRatios[key] !== undefined && localRatios[key] !== null) {
                    const arr = localRatios[key];
                    const lastVal = arr[arr.length - 1];
                    return lastVal !== null && lastVal !== undefined ? (lastVal * (factor === 1 ? 1 : 100)).toFixed(decimals) + suffix : '-';
                }
                return rawLive ? (rawLive * factor).toFixed(decimals) + suffix : '-';
            };

            return [
                { label: 'ROE', value: getVal('roe', latest.roe, 100, 1), desc: 'Tỷ suất vốn chủ sở hữu' },
                { label: 'Biên gộp', value: getVal('gross_margin', latest.grossMargin, 100, 1), desc: 'Gross Margin' },
                { label: 'Nợ/Vốn chủ', value: getVal('debt_to_equity', latest.debtToEquity, 1, 2, 'x'), desc: 'Đòn bẩy tài chính' },
                { label: 'ROA', value: getVal('roa', latest.roa, 100, 2), desc: 'Tỷ suất tổng tài sản' },
            ];
        }
    },
    technology: {
        label: 'Công Nghệ',
        labelEn: 'Technology',
        icon: '💻',
        color: '#22c55e',
        incomeField: 'isa3',
        incomeField2: 'isa5',
        npat: 'isa20',
        incomeLabel: 'Doanh thu thuần',
        income2Label: 'Lợi nhuận gộp',
        npatLabel: 'Lợi nhuận sau thuế',
        showPE: true,
        primaryValuation: 'P/E',
        extraMetrics: (ratios, localJson) => {
            const latest = ratios?.[ratios.length - 1] || {};
            const localRatios = localJson?.ratios || {};
            
            const getVal = (key, rawLive, factor = 100, decimals = 1, suffix = '%') => {
                if (localRatios[key] !== undefined && localRatios[key] !== null) {
                    const arr = localRatios[key];
                    const lastVal = arr[arr.length - 1];
                    return lastVal !== null && lastVal !== undefined ? (lastVal * (factor === 1 ? 1 : 100)).toFixed(decimals) + suffix : '-';
                }
                return rawLive ? (rawLive * factor).toFixed(decimals) + suffix : '-';
            };

            return [
                { label: 'Biên gộp', value: getVal('gross_margin', latest.grossMargin, 100, 1), desc: 'Gross Margin' },
                { label: 'ROE', value: getVal('roe', latest.roe, 100, 1), desc: 'Tỷ suất vốn chủ sở hữu' },
                { label: 'ROA', value: getVal('roa', latest.roa, 100, 2), desc: 'Tỷ suất tổng tài sản' },
                { label: 'Nợ/Vốn chủ', value: getVal('debt_to_equity', latest.debtToEquity, 1, 2, 'x'), desc: 'Đòn bẩy tài chính' },
            ];
        }
    },
    consumer: {
        label: 'Tiêu Dùng & Bán lẻ',
        labelEn: 'Consumer',
        icon: '🛒',
        color: '#eab308',
        incomeField: 'isa3',
        incomeField2: 'isa5',
        npat: 'isa20',
        incomeLabel: 'Doanh thu thuần',
        income2Label: 'Lợi nhuận gộp',
        npatLabel: 'Lợi nhuận sau thuế',
        showPE: true,
        primaryValuation: 'P/E',
        extraMetrics: (ratios, localJson) => {
            const latest = ratios?.[ratios.length - 1] || {};
            const localRatios = localJson?.ratios || {};
            
            const getVal = (key, rawLive, factor = 100, decimals = 1, suffix = '%') => {
                if (localRatios[key] !== undefined && localRatios[key] !== null) {
                    const arr = localRatios[key];
                    const lastVal = arr[arr.length - 1];
                    return lastVal !== null && lastVal !== undefined ? (lastVal * (factor === 1 ? 1 : 100)).toFixed(decimals) + suffix : '-';
                }
                return rawLive ? (rawLive * factor).toFixed(decimals) + suffix : '-';
            };

            return [
                { label: 'Biên gộp', value: getVal('gross_margin', latest.grossMargin, 100, 1), desc: 'Gross Margin' },
                { label: 'ROE', value: getVal('roe', latest.roe, 100, 1), desc: 'Tỷ suất vốn chủ sở hữu' },
                { label: 'Biên EBIT', value: getVal('ebitMargin', latest.ebitMargin, 100, 1), desc: 'EBIT Margin' },
                { label: 'Nợ/Vốn chủ', value: getVal('debt_to_equity', latest.debtToEquity, 1, 2, 'x'), desc: 'Đòn bẩy' },
            ];
        }
    },
    generic: {
        label: 'Công Nghiệp',
        labelEn: 'General',
        icon: '🏭',
        color: '#38bdf8',
        incomeField: 'isa3',
        incomeField2: 'isa5',

        npat: 'isa20',
        incomeLabel: 'Doanh thu thuần',
        income2Label: 'Lợi nhuận gộp',
        npatLabel: 'Lợi nhuận sau thuế',
        showPE: true,
        primaryValuation: 'P/E',
        extraMetrics: (ratios) => {
            const latest = ratios?.[ratios.length - 1] || {};
            return [
                { label: 'ROE', value: latest.roe ? (latest.roe * 100).toFixed(1) + '%' : '-', desc: 'Tỷ suất vốn chủ' },
                { label: 'Biên gộp', value: latest.grossMargin ? (latest.grossMargin * 100).toFixed(1) + '%' : '-', desc: 'Gross Margin' },
                { label: 'Nợ/Vốn chủ', value: latest.debtToEquity ? latest.debtToEquity.toFixed(2) + 'x' : '-', desc: 'Đòn bẩy' },
                { label: 'ROA', value: latest.roa ? (latest.roa * 100).toFixed(2) + '%' : '-', desc: 'Tỷ suất tổng tài sản' },
            ];
        }
    }
};

// ── Map sector string → config key ────────────────────────
function getSectorKey(sectorStr) {
    const s = (sectorStr || '').toLowerCase();
    if (s.includes('bank') || s.includes('financial service')) return 'banks';
    if (s.includes('basic resource') || s.includes('material') || s.includes('steel') || s.includes('metal') || s.includes('mining')) return 'materials';
    if (s.includes('real estate') || s.includes('property')) return 'realestate';
    if (s.includes('tech') || s.includes('software') || s.includes('it ')) return 'technology';
    if (s.includes('consumer') || s.includes('retail') || s.includes('food') || s.includes('beverage')) return 'consumer';
    return 'generic';
}

// ═══════════════════════════════════════════════════════════
// NAVIGATION / ROUTING
// ═══════════════════════════════════════════════════════════
function goHome() {
    document.getElementById('view-overview').style.display = 'flex';
    document.getElementById('view-analysis').style.display = 'none';

    // Reset breadcrumb
    document.getElementById('nav-breadcrumb').innerHTML = '<span class="breadcrumb-home">Tổng quan</span>';
    document.getElementById('sector-badge-nav').style.display = 'none';
    document.getElementById('back-btn').style.display = 'none';

    // Destroy charts to free memory
    destroyCharts();
}

function showAnalysisView(sectorKey, ticker) {
    const cfg = SECTOR_CONFIG[sectorKey] || SECTOR_CONFIG.generic;

    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';

    // Update breadcrumb
    document.getElementById('nav-breadcrumb').innerHTML = `
        <span class="breadcrumb-home" style="cursor:pointer" onclick="goHome()">Tổng quan</span>
        <span class="breadcrumb-sep">›</span>
        <span class="breadcrumb-sector" style="color:${cfg.color}">${cfg.icon} ${cfg.label}</span>
        <span class="breadcrumb-sep">›</span>
        <span class="breadcrumb-ticker">${ticker}</span>
    `;

    // Sector badge
    const badge = document.getElementById('sector-badge-nav');
    badge.style.display = 'flex';
    badge.style.borderColor = cfg.color + '50';
    badge.style.color = cfg.color;
    document.getElementById('sector-icon-nav').textContent = cfg.icon;
    document.getElementById('sector-label-nav').textContent = cfg.label;

    document.getElementById('back-btn').style.display = 'inline-flex';

    // Sector ribbon color
    const ribbon = document.getElementById('sector-ribbon');
    ribbon.textContent = cfg.label.toUpperCase();
    ribbon.style.background = cfg.color;

    // Apply sector CSS variable
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
    const [details, ratios, incomeYears] = await Promise.all([
        fetchVietcap(`/company/details?ticker=${ticker}`),
        fetchVietcap(`/company/${ticker}/statistics-financial`),
        fetchVietcap(`/company/${ticker}/financial-statement?section=INCOME_STATEMENT`),
    ]);
    return { details, ratios, incomeYears };
}

// ═══════════════════════════════════════════════════════════
// LOAD OVERVIEW
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {
    await loadOverview();
});

async function loadOverview() {
    try {
        const resp = await fetch('data/index.json');
        const stocks = await resp.json();
        renderOverview(stocks);
    } catch {
        document.getElementById('sector-groups-container').innerHTML =
            '<div class="loading-state card">Chưa có dữ liệu. Hãy chạy phân tích ít nhất một mã cổ phiếu.</div>';
    }
}

function renderOverview(stocks) {
    if (!stocks || stocks.length === 0) {
        document.getElementById('sector-groups-container').innerHTML =
            '<div class="loading-state card">Chưa có dữ liệu cổ phiếu nào được phân tích.</div>';
        return;
    }

    // Save stocks globally for search
    window._allStocks = stocks;

    // Group by sector key
    _groupAndRender(stocks);

    // Wire up search
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
        container.innerHTML = `<div class="loading-state card">Không tìm thấy mã nào khớp với "${query}"</div>`;
        return;
    }
    container.innerHTML = '';
    const card = document.createElement('div');
    card.className = 'card sector-group';
    card.innerHTML = `
        <div class="sector-group-header">
            <h3 style="color:var(--text-muted)">🔍 Kết quả tìm kiếm</h3>
            <span class="sector-count-badge">${filtered.length} mã</span>
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

    const sectorOrder = ['banks', 'materials', 'realestate', 'technology', 'consumer', 'generic'];

    sectorOrder.forEach(key => {
        if (!groups[key]) return;
        const cfg = SECTOR_CONFIG[key];
        const groupDiv = document.createElement('div');
        groupDiv.className = 'card sector-group';

        groupDiv.innerHTML = `
            <div class="sector-group-header">
                <div class="sector-group-dot" style="background:${cfg.color}"></div>
                <h3 style="color:${cfg.color}">${cfg.icon} ${cfg.label}</h3>
                <span class="sector-count-badge">${groups[key].length} mã</span>
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
        <div class="stock-card-name">${s.companyName || ''}</div>
        <div class="stock-card-meta">
            <span class="stock-card-price">${s.currentPrice ? formatNumber(s.currentPrice) + ' VND' : ''}</span>
            <span class="stock-card-arrow">→</span>
        </div>
    `;
    card.onclick = () => loadStockDashboard(s.ticker);
    return card;
}

// ═══════════════════════════════════════════════════════════
// LOAD STOCK DASHBOARD (sector-adaptive)
// ═══════════════════════════════════════════════════════════
async function loadStockDashboard(ticker) {
    // Show analysis view immediately with loading state
    document.documentElement.style.setProperty('--sector-color', '#3b82f6');
    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';
    document.getElementById('view-analysis').classList.add('active-view');
    document.getElementById('ticker-badge').textContent = ticker;
    document.getElementById('company-name').textContent = 'Đang tải dữ liệu...';
    document.getElementById('company-sector').textContent = '';
    document.getElementById('nav-breadcrumb').innerHTML = `<span class="breadcrumb-home" style="cursor:pointer" onclick="goHome()">Tổng quan</span> › <span>${ticker}</span>`;
    document.getElementById('back-btn').classList.remove('hidden');

    destroyCharts();

    // Fetch live data from Vietcap + local JSON
    const [liveData, localJson] = await Promise.all([
        fetchStockLive(ticker),
        fetch(`data/${ticker}.json`).then(r => r.ok ? r.json() : null).catch(() => null)
    ]);

    // Merge: live takes priority over local
    const details = liveData.details || {};
    const ratiosAll = liveData.ratios || [];
    const incomeYears = liveData.incomeYears?.years || [];

    // Use details for basic info
    const sector = details.sector || localJson?.sector || 'General';
    const sectorKey = getSectorKey(sector);
    const cfg = SECTOR_CONFIG[sectorKey];

    const companyName = details.viOrganName || details.enOrganName || localJson?.companyName || ticker;
    const currentPrice = details.currentPrice || localJson?.currentPrice || 0;
    const marketCap = details.marketCap || localJson?.marketCap || 0;
    const shares = details.numberOfSharesMktCap || localJson?.shares || 0;

    // Show sector view
    showAnalysisView(sectorKey, ticker);

    // ── Header ──────────────────────────────────────────
    document.getElementById('ticker-badge').textContent = ticker;
    document.getElementById('company-name').textContent = companyName;
    document.getElementById('company-sector').textContent = `${cfg.icon} ${cfg.label} · ${sector}`;

    // ── Live badge ─────────────────────────────────────
    const liveBadge = `<span class="live-badge"><span class="live-dot"></span>Live · Vietcap IQ</span>`;

    // ── Download buttons ────────────────────────────────
    const btnPdf = document.getElementById('download-pdf');
    const btnExcel = document.getElementById('download-excel');
    if (localJson?.gdrivePdfUrl) { btnPdf.href = localJson.gdrivePdfUrl; btnPdf.classList.remove('hidden'); }
    else btnPdf.classList.add('hidden');
    if (localJson?.gdriveExcelUrl) { btnExcel.href = localJson.gdriveExcelUrl; btnExcel.classList.remove('hidden'); }
    else btnExcel.classList.add('hidden');

    // ── Metrics (top row) ────────────────────────────────
    const latestRatio = ratiosAll.filter(r => r.quarter <= 4).slice(-1)[0] || {};
    const ratioTTM = ratiosAll.filter(r => r.quarter <= 4).slice(-1)[0] || {};

    const currentPE = ratioTTM.pe ? ratioTTM.pe.toFixed(1) + 'x' : '-';
    const currentPB = ratioTTM.pb ? ratioTTM.pb.toFixed(2) + 'x' : '-';

    const metricsGrid = document.getElementById('dynamic-metrics-grid');
    metricsGrid.innerHTML = `
        <div class="metric-card highlight">
            <div class="metric-label">Giá hiện tại</div>
            <div class="metric-value">${formatNumber(currentPrice)}</div>
            <div class="metric-sub">VND ${liveBadge}</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">Vốn hóa thị trường</div>
            <div class="metric-value">${formatNumber(Math.round(marketCap / 1e12), 1)}</div>
            <div class="metric-sub">Nghìn tỷ VND</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">${cfg.primaryValuation} hiện tại</div>
            <div class="metric-value">${cfg.primaryValuation === 'P/B' ? currentPB : currentPE}</div>
            <div class="metric-sub">Trailing twelve months</div>
        </div>
        ${details.targetPrice ? `
        <div class="metric-card">
            <div class="metric-label">Giá mục tiêu (${details.analyst || 'Vietcap'})</div>
            <div class="metric-value" style="color:#10b981">${formatNumber(details.targetPrice)}</div>
            <div class="metric-sub">${details.rating || ''} · Upside ${details.upsideToTargetPercent ? (details.upsideToTargetPercent * 100).toFixed(1) + '%' : ''}</div>
        </div>` : ''}
    `;

    // ── Bank KPI extra card (or sector metrics) ──────────
    const bankKpiCard = document.getElementById('bank-kpi-card');
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

    // ── Charts ───────────────────────────────────────────
    // Filter TTM quarters (quarter 1–4) and sort
    const ttmQuarters = ratiosAll
        .filter(r => r.quarter >= 1 && r.quarter <= 4)
        .sort((a, b) => a.yearReport !== b.yearReport ? a.yearReport - b.yearReport : a.quarter - b.quarter);

    // Build P/E and P/B series
    const peData = ttmQuarters.map(r => ({ x: `${r.yearReport}-Q${r.quarter}`, y: r.pe && r.pe > 0 && r.pe < 500 ? parseFloat(r.pe.toFixed(1)) : null }));
    const pbData = ttmQuarters.map(r => ({ x: `${r.yearReport}-Q${r.quarter}`, y: r.pb && r.pb > 0 ? parseFloat(r.pb.toFixed(2)) : null }));
    const quarterLabels = ttmQuarters.map(r => `${r.yearReport}-Q${r.quarter}`);

    // Build income chart from historical years
    const annualYears = incomeYears.sort((a, b) => a.yearReport - b.yearReport).slice(-7);
    const incomeLabels = annualYears.map(r => r.yearReport.toString());
    const incomeData = annualYears.map(r => r[cfg.incomeField] ? parseFloat((r[cfg.incomeField] / 1e12).toFixed(1)) : 0);
    const income2Data = cfg.incomeField2 ? annualYears.map(r => r[cfg.incomeField2] ? parseFloat((r[cfg.incomeField2] / 1e12).toFixed(1)) : 0) : [];
    const npatData = annualYears.map(r => r[cfg.npat] ? parseFloat((r[cfg.npat] / 1e12).toFixed(1)) : 0);

    // Update chart titles
    document.getElementById('chart1-title').textContent = `${cfg.incomeLabel} & ${cfg.npatLabel}`;
    document.getElementById('chart2-title').textContent = income2Data.length > 0 ? `${cfg.income2Label} & ${cfg.npatLabel}` : 'EPS & Vốn chủ sở hữu';
    document.getElementById('chart-pe-title').textContent = cfg.showPE ? 'Định giá Lịch sử P/E - TTM' : 'Định giá Lịch sử P/E - TTM (phụ)';

    // P/E chart visibility
    if (!cfg.showPE) {
        document.getElementById('chart-pe-card').style.opacity = '0.6';
    } else {
        document.getElementById('chart-pe-card').style.opacity = '1';
    }

    renderIncomeChart(incomeLabels, incomeData, npatData, cfg.incomeLabel, cfg.npatLabel, cfg.color);
    renderChart2(incomeLabels, income2Data, npatData, cfg.income2Label, cfg.npatLabel, cfg.color);
    renderPEChart(quarterLabels, peData.map(p => p.y), cfg.color);
    renderPBChart(quarterLabels, pbData.map(p => p.y), cfg.color);

    // ── Earning Release & YTD Evaluation (New Component) ──
    renderQuarterlyAndYTDEvaluation(ticker, liveData, localJson, cfg);

    // ── Generate side-by-side chart commentaries ─────────
    generateChartCommentaries(ticker, annualYears, ttmQuarters, cfg, localJson);

    // ── Valuation scenarios ──────────────────────────────
    renderValuationScenarios(localJson?.valuation, currentPrice, details, latestRatio, cfg);

    // ── Qualitative sections ─────────────────────────────
    const q = localJson || {};
    renderThesisAndRisks(q.thesis, q.risks);
    renderMoatScorecard(q.moats, cfg.color);
    renderPESTLE(q.pestle);
    renderCommentary(q.comments);
    renderFinancialTable(annualYears, cfg, incomeLabels);
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
                    label: incomeLabel + ' (nghìn tỷ)',
                    data: incomeData,
                    backgroundColor: color + '40',
                    borderColor: color,
                    borderWidth: 2,
                    borderRadius: 6,
                    yAxisID: 'y'
                },
                {
                    label: npatLabel + ' (nghìn tỷ)',
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
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'Nghìn tỷ VND', color: '#545f74', font: { size: 10 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'LNST (nghìn tỷ)', color: '#545f74', font: { size: 10 } } }
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
                    label: label1 + ' (nghìn tỷ)',
                    data: data1,
                    backgroundColor: color + '30',
                    borderColor: color + 'aa',
                    borderWidth: 2,
                    borderRadius: 5,
                    yAxisID: 'y'
                },
                {
                    label: label2 + ' (nghìn tỷ)',
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
                    label: `Trung vị (${med ? med.toFixed(1) + 'x' : '-'})`,
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
                    label: `Trung vị (${med ? med.toFixed(2) + 'x' : '-'})`,
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
// QUALITATIVE RENDERERS
// ═══════════════════════════════════════════════════════════
function renderThesisAndRisks(thesisList, risksList) {
    const defaultThesis = [
        'Doanh nghiệp có vị thế cạnh tranh tốt trong ngành với nền tảng tài chính lành mạnh.',
        'Dòng tiền hoạt động ổn định và cơ cấu tài chính hỗ trợ chiến lược mở rộng dài hạn.',
        'Đội ngũ quản lý có kinh nghiệm với track record thực thi kế hoạch kinh doanh hiệu quả.',
    ];
    const defaultRisks = [
        'Rủi ro kinh tế vĩ mô và biến động lãi suất có thể ảnh hưởng đến chi phí vốn.',
        'Cạnh tranh trong ngành ngày càng gay gắt đòi hỏi đầu tư liên tục vào năng lực cạnh tranh.',
        'Rủi ro thực thi chiến lược và phụ thuộc vào một số khách hàng/sản phẩm chủ lực.',
    ];

    const thesis = thesisList?.length > 0 ? thesisList : defaultThesis;
    const risks  = risksList?.length  > 0 ? risksList  : defaultRisks;

    document.getElementById('thesis-list').innerHTML = thesis.map(t => `<li>${t}</li>`).join('');
    document.getElementById('risks-list').innerHTML  = risks.map(r => `<li>${r}</li>`).join('');
}

function renderMoatScorecard(moats, color) {
    const defaultMoats = {
        'Intangible Assets': { score: 3, desc: 'Thương hiệu và uy tín trong ngành được xây dựng qua nhiều năm.' },
        'Cost Advantage':    { score: 3, desc: 'Tối ưu hóa quy trình vận hành giúp duy trì biên lợi nhuận cạnh tranh.' },
        'Switching Cost':    { score: 2, desc: 'Mức độ gắn kết với khách hàng ở mức trung bình.' },
        'Efficient Scale':   { score: 3, desc: 'Quy mô phù hợp để vận hành hiệu quả trong thị trường nội địa.' },
        'Network Effect':    { score: 2, desc: 'Hiệu ứng mạng lưới còn hạn chế trong mô hình kinh doanh hiện tại.' },
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
        Political:     'Chính sách vĩ mô và định hướng phát triển ngành của Chính phủ tác động đến môi trường kinh doanh.',
        Economic:      'Tăng trưởng GDP, lãi suất và lạm phát ảnh hưởng đến nhu cầu và chi phí hoạt động.',
        Social:        'Thay đổi xu hướng tiêu dùng và nhân khẩu học tạo cả cơ hội và thách thức.',
        Technological: 'Chuyển đổi số và tự động hóa là xu hướng quan trọng cần đầu tư bắt kịp.',
        Legal:         'Tuân thủ các quy định pháp luật và tiêu chuẩn ngành là yêu cầu cơ bản.',
        Environmental: 'Áp lực ESG và báo cáo phát triển bền vững ngày càng trở thành yêu cầu của nhà đầu tư.',
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
    document.getElementById('comment-business').textContent  = data.overall   || 'Chưa có đánh giá mô hình kinh doanh.';
    document.getElementById('comment-financial').textContent = data.financial  || 'Chưa có đánh giá sức khỏe tài chính.';
    document.getElementById('comment-valuation').textContent = data.valuation  || 'Chưa có đánh giá định giá.';
}

function renderFinancialTable(annualYears, cfg, labels) {
    const table = document.getElementById('financial-table');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = '<th>Chỉ tiêu</th>' + labels.map(y => `<th>${y}</th>`).join('');
    tbody.innerHTML = '';

    const rows = [
        { label: cfg.incomeLabel + ' (tỷ)', field: cfg.incomeField, div: 1e9 },
        { label: cfg.income2Label + ' (tỷ)', field: cfg.incomeField2, div: 1e9 },
        { label: 'LNST (tỷ)', field: cfg.npat, div: 1e9, highlight: true },
    ].filter(r => r.field);

    rows.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td><strong>${row.label}</strong></td>` +
            annualYears.map(yr => {
                const v = yr[row.field];
                const fmt = v ? formatNumber(Math.round(v / (row.div || 1e9))) : '-';
                return `<td class="${row.highlight ? 'highlight' : ''}">${fmt}</td>`;
            }).join('');
        tbody.appendChild(tr);
    });

    document.getElementById('table-title').textContent = 'Bảng kết quả kinh doanh (theo năm, tỷ VND)';
}

// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════
function formatNumber(num, decimals = 0) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Number(num).toLocaleString('vi-VN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// ═══════════════════════════════════════════════════════════
// NEW COMPONENTS FOR DETAILED EVALUATION
// ═══════════════════════════════════════════════════════════

function renderQuarterlyAndYTDEvaluation(ticker, liveData, localJson, cfg) {
    const container = document.getElementById('quarterly-ytd-evaluation-container');
    const ratios = liveData.ratios || [];
    const isrecs = liveData.incomeYears?.quarters || [];
    
    // Sort quarters chronologically
    const quarters = isrecs
        .filter(q => q.quarter >= 1 && q.quarter <= 4)
        .sort((a, b) => a.yearReport !== b.yearReport ? a.yearReport - b.yearReport : a.quarter - b.quarter);
    
    if (quarters.length < 5) {
        container.innerHTML = `<div class="loading-state">Không đủ dữ liệu quý để tính toán tăng trưởng.</div>`;
        return;
    }

    const latestQ = quarters[quarters.length - 1];
    const prevQ = quarters[quarters.length - 2];
    
    // Find same quarter last year
    const sameQLastYear = quarters.find(q => q.yearReport === latestQ.yearReport - 1 && q.quarter === latestQ.quarter);
    
    // 1. Calculate growth metrics
    const revFieldName = cfg.incomeField; // Net interest income for banks or net revenue for others
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

    // 2. YTD Calculation
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
        if (val > 0) return `<span class="badge-growth up">▲ +${val.toFixed(1)}%</span>`;
        if (val < 0) return `<span class="badge-growth down">▼ ${val.toFixed(1)}%</span>`;
        return '<span class="badge-growth flat">0.0%</span>';
    };

    // 3. Generate summary texts
    let commentary = '';
    const nameRev = cfg.incomeLabel;
    
    if (yoyNpat && yoyNpat > 5) {
        commentary = `KQKD quý ${latestQ.quarter}/${latestQ.yearReport} của ${ticker} ghi nhận sự tăng trưởng tích cực, với lợi nhuận sau thuế đạt ${formatNumber(latestNpat/1e9, 1)} tỷ đồng (${formatGrowth(yoyNpat)} so với cùng kỳ). `;
    } else if (yoyNpat && yoyNpat < -5) {
        commentary = `Kết quả kinh doanh quý gần nhất cho thấy tín hiệu chậm lại, LNST giảm ${formatGrowth(yoyNpat)} so với cùng kỳ do áp lực thu hẹp NIM/biên lợi nhuận hoặc chi phí tăng cao. `;
    } else {
        commentary = `Kết quả kinh doanh quý gần nhất của ${ticker} duy trì ở mức ổn định. `;
    }

    if (yoyYtdNpat) {
        commentary += `Lũy kế từ đầu năm (YTD), LNST đạt ${formatNumber(ytdNpat/1e9, 1)} tỷ đồng, thay đổi ${yoyYtdNpat.toFixed(1)}% so với cùng kỳ năm trước.`;
    }

    container.innerHTML = `
        <div class="quarterly-ytd-grid">
            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">Quý gần nhất (${latestQ.quarter}Q/${latestQ.yearReport})</span>
                    <span style="font-size:0.75rem;color:var(--text-dim)">Đơn vị: Tỷ VND</span>
                </div>
                <div class="q-ytd-metrics">
                    <div class="q-metric-row">
                        <span class="q-metric-label">${nameRev}</span>
                        <span class="q-metric-value">${formatNumber(latestRev/1e9, 1)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Tăng trưởng QoQ (so quý trước)</span>
                        <span>${formatGrowth(qoqRev)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Tăng trưởng YoY (cùng kỳ)</span>
                        <span>${formatGrowth(yoyRev)}</span>
                    </div>
                </div>
            </div>

            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">Lợi nhuận quý gần nhất</span>
                    <span style="font-size:0.75rem;color:var(--text-dim)">Đơn vị: Tỷ VND</span>
                </div>
                <div class="q-ytd-metrics">
                    <div class="q-metric-row">
                        <span class="q-metric-label">Lợi nhuận sau thuế</span>
                        <span class="q-metric-value">${formatNumber(latestNpat/1e9, 1)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Tăng trưởng QoQ (so quý trước)</span>
                        <span>${formatGrowth(qoqNpat)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Tăng trưởng YoY (cùng kỳ)</span>
                        <span>${formatGrowth(yoyNpat)}</span>
                    </div>
                </div>
            </div>

            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">Lũy kế từ đầu năm (YTD)</span>
                    <span style="font-size:0.75rem;color:var(--text-dim)">So với cùng kỳ</span>
                </div>
                <div class="q-ytd-metrics">
                    <div class="q-metric-row">
                        <span class="q-metric-label">Lũy kế Doanh thu/NII YTD</span>
                        <span class="q-metric-value">${formatNumber(ytdRev/1e9, 1)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Tăng trưởng YTD Doanh thu/NII</span>
                        <span>${formatGrowth(yoyYtdRev)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Lũy kế LNST YTD</span>
                        <span class="q-metric-value">${formatNumber(ytdNpat/1e9, 1)}</span>
                    </div>
                    <div class="q-metric-row">
                        <span class="q-metric-label">Tăng trưởng YTD LNST</span>
                        <span>${formatGrowth(yoyYtdNpat)}</span>
                    </div>
                </div>
            </div>
        </div>
        <div class="q-commentary-box">
            <strong>📋 Nhận định nhanh Earning Release:</strong> ${commentary}
        </div>
    `;
}

function generateChartCommentaries(ticker, annualYears, ttmQuarters, cfg, localJson) {
    // 1. Commentary for Revenue & NPAT chart
    const revEl = document.getElementById('analysis-text-rev-npat');
    if (revEl && annualYears.length >= 2) {
        const last = annualYears[annualYears.length - 1];
        const prev = annualYears[annualYears.length - 2];
        const revField = cfg.incomeField;
        const npatField = cfg.npat;
        
        const growthRev = prev[revField] > 0 ? ((last[revField] - prev[revField]) / prev[revField] * 100) : 0;
        const growthNpat = prev[npatField] > 0 ? ((last[npatField] - prev[npatField]) / prev[npatField] * 100) : 0;
        
        revEl.innerHTML = `
            Doanh thu/NII năm gần nhất đạt <strong>${formatNumber(last[revField]/1e12, 1)} nghìn tỷ</strong>, tăng trưởng <strong>${growthRev.toFixed(1)}% YoY</strong>. <br>
            LNST tương ứng đạt <strong>${formatNumber(last[npatField]/1e12, 1)} nghìn tỷ</strong> (${growthNpat >= 0 ? '+' : ''}${growthNpat.toFixed(1)}% YoY). <br>
            Nhìn chung, doanh nghiệp đang duy trì xu hướng ${growthNpat > 0 ? 'tăng trưởng tích cực' : 'đi ngang/sụt giảm'} về mặt hiệu quả kinh doanh cốt lõi.
        `;
    }

    // 2. Commentary for ROE & EPS
    const epsEl = document.getElementById('analysis-text-eps-equity');
    if (epsEl && localJson?.ratios) {
        const roeArr = localJson.ratios.roe || [];
        const epsArr = localJson.data?.eps || [];
        
        if (roeArr.length > 0) {
            const lastRoe = (roeArr[roeArr.length - 1] * 100).toFixed(1);
            const lastEps = epsArr[epsArr.length - 1] ? formatNumber(epsArr[epsArr.length - 1]) : '-';
            
            epsEl.innerHTML = `
                Tỷ suất sinh lợi trên vốn chủ sở hữu (ROE) đạt <strong>${lastRoe}%</strong> ở năm gần nhất. <br>
                Thu nhập trên mỗi cổ phần (EPS) tương ứng là <strong>${lastEps} VND/CP</strong>. <br>
                Mức hiệu quả ROE ${parseFloat(lastRoe) > 15 ? 'ở mức cao (>15%), chứng tỏ khả năng sinh lời hiệu quả của nguồn vốn.' : 'ở mức trung bình thấp, cần theo dõi xu hướng tái cơ cấu tài sản.'}
            `;
        }
    } else {
        epsEl.innerHTML = `Hiệu quả sử dụng vốn chủ sở hữu (ROE) và EPS được tối ưu hóa dựa trên cơ cấu nợ và hiệu năng phân bổ vốn tài sản của doanh nghiệp.`;
    }

    // 3. Commentary for P/E & P/B valuation
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
                Hệ số P/E trailing hiện tại là <strong>${lastRatio.pe.toFixed(1)}x</strong>. <br>
                Trung vị lịch sử của doanh nghiệp là <strong>${peMedian ? peMedian.toFixed(1) + 'x' : '-'}</strong>. <br>
                Định giá P/E đang ${lastRatio.pe > peMedian ? `cao hơn trung vị (${diffPe}%), phản ánh kỳ vọng tăng trưởng lớn hoặc giá đang ở vùng đắt.` : `thấp hơn trung vị (${Math.abs(diffPe)}%), cho thấy biên an toàn định giá tương đối rẻ.`}
            `;
        }

        if (pbEl && lastRatio.pb) {
            const diffPb = ((lastRatio.pb - pbMedian) / pbMedian * 100).toFixed(1);
            pbEl.innerHTML = `
                Hệ số P/B trailing hiện tại là <strong>${lastRatio.pb.toFixed(2)}x</strong>. <br>
                Trung vị lịch sử của doanh nghiệp là <strong>${pbMedian ? pbMedian.toFixed(2) + 'x' : '-'}</strong>. <br>
                Vùng định giá P/B hiện tại ${lastRatio.pb > pbMedian ? `nằm trên trung vị lịch sử (${diffPb}%)` : `nằm dưới trung vị lịch sử (${Math.abs(diffPb)}%)`} là cơ sở để cân nhắc tích lũy.
            `;
        }
    }
}

