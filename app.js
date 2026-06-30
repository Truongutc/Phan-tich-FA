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
let chartCreditFundingGrowth = null;
let chartVolume = null;
let chartSpread = null;
let chartAsset = null;
let chartCost = null;

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
            const rq = localJson?.ratios_quarterly || {};
            const latest_idx = rq.quarters ? rq.quarters.length - 1 : -1;
            
            // 1. Tăng trưởng Quý gần nhất (YoY)
            let npatYoY = '-';
            const iq = localJson?.income_quarterly || [];
            if (iq.length >= 5) {
                const latest_q = iq[iq.length - 1];
                const yoy_q = iq[iq.length - 5];
                if (latest_q && yoy_q && yoy_q.npat > 0) {
                    npatYoY = ((latest_q.npat / yoy_q.npat - 1) * 100).toFixed(1) + '%';
                }
            }
            
            // 2. ROE Quý gần nhất (Năm hóa)
            let roeQ = '-';
            if (latest_idx !== -1 && rq.roe?.[latest_idx] !== undefined && rq.roe?.[latest_idx] !== null) {
                roeQ = (rq.roe[latest_idx] * 100).toFixed(1) + '%';
            }
            
            // 3. P/B Hiện tại
            const currentPB = latest.pb ? latest.pb.toFixed(2) + 'x' : '-';
            
            // 4. Tỷ lệ Nợ xấu Quý hiện tại
            let nplQ = '-';
            if (latest_idx !== -1 && rq.npl?.[latest_idx] !== undefined && rq.npl?.[latest_idx] !== null) {
                nplQ = rq.npl[latest_idx].toFixed(2) + '%';
            }
            
            // 5. Tăng trưởng Tín dụng Quý hiện tại (QoQ)
            let creditGrowthQ = '-';
            if (latest_idx !== -1 && rq.credit_growth?.[latest_idx] !== undefined && rq.credit_growth?.[latest_idx] !== null) {
                creditGrowthQ = rq.credit_growth[latest_idx].toFixed(2) + '%';
            }
            
            // 6. P/E Hiện tại
            const currentPE = latest.pe ? latest.pe.toFixed(1) + 'x' : '-';

            return [
                { label: 'Tăng trưởng LN Quý YoY', value: npatYoY, desc: 'Tăng trưởng LNST so cùng kỳ' },
                { label: 'ROE Quý gần nhất (năm hóa)', value: roeQ, desc: 'Tỷ suất LN/Vốn chủ quý gần nhất' },
                { label: 'P/B Hiện tại', value: currentPB, desc: 'Hệ số Giá/Sách trailing' },
                { label: 'Tỉ lệ Nợ xấu Quý hiện tại', value: nplQ, desc: 'Tỷ lệ nợ xấu NPL quý hiện tại' },
                { label: 'Tăng trưởng tín dụng YTD', value: creditGrowthQ, desc: 'So với đầu năm (cuối năm trước)' },
                { label: 'P/E Hiện tại', value: currentPE, desc: 'Hệ số Giá/Thu nhập trailing' }
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
    // Ngân hàng — tiếng Anh + tiếng Việt
    if (s.includes('bank') || s.includes('financial service') ||
        s.includes('ng\u00e2n h\u00e0ng') || s.includes('ngan hang')) return 'banks';
    // Thép & Vật liệu
    if (s.includes('basic resource') || s.includes('material') || s.includes('steel') ||
        s.includes('metal') || s.includes('mining') || s.includes('th\u00e9p') ||
        s.includes('v\u1eadt li\u1ec7u') || s.includes('khai kho\u00e1ng')) return 'materials';
    // Bất động sản
    if (s.includes('real estate') || s.includes('property') ||
        s.includes('b\u1ea5t \u0111\u1ed9ng s\u1ea3n')) return 'realestate';
    // Công nghệ
    if (s.includes('tech') || s.includes('software') || s.includes('it ') ||
        s.includes('c\u00f4ng ngh\u1ec7')) return 'technology';
    // Tiêu dùng / Bán lẻ
    if (s.includes('consumer') || s.includes('retail') || s.includes('food') ||
        s.includes('beverage') || s.includes('ti\u00eau d\u00f9ng') ||
        s.includes('b\u00e1n l\u1ebb')) return 'consumer';
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
    const [details, ratios, incomeYears, incomeQuarters] = await Promise.all([
        fetchVietcap(`/company/details?ticker=${ticker}`),
        fetchVietcap(`/company/${ticker}/statistics-financial`),
        fetchVietcap(`/company/${ticker}/financial-statement?section=INCOME_STATEMENT`),
        fetchVietcap(`/company/${ticker}/financial-statement?section=INCOME_STATEMENT&quarterly=true`),
    ]);
    // Ghép dữ liệu quý vào incomeYears để renderQuarterlyAndYTDEvaluation đọc được
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

    // ── Report Date ─────────────────────────────────────
    const reportDateEl = document.getElementById('report-date');
    if (reportDateEl) {
        if (localJson?.lastUpdated) {
            const parts = localJson.lastUpdated.split(' ');
            if (parts.length === 2) {
                const dateParts = parts[0].split('-');
                if (dateParts.length === 3) {
                    reportDateEl.innerHTML = `🕒 Thời gian lập báo cáo: <b>${dateParts[2]}/${dateParts[1]}/${dateParts[0]} ${parts[1]}</b>`;
                } else {
                    reportDateEl.innerHTML = `🕒 Thời gian lập báo cáo: <b>${localJson.lastUpdated}</b>`;
                }
            } else {
                reportDateEl.innerHTML = `🕒 Thời gian lập báo cáo: <b>${localJson.lastUpdated}</b>`;
            }
            reportDateEl.style.display = 'block';
        } else {
            reportDateEl.style.display = 'none';
        }
    }

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
    const bankKpiTitle = bankKpiCard.querySelector('h3');
    if (bankKpiTitle) {
        bankKpiTitle.textContent = cfg.label === 'Ngân Hàng' ? 'Chỉ số Ngân hàng then chốt' : 'Chỉ số Tài chính then chốt';
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
    renderCreditFundingGrowthChart(localJson, cfg);
    renderVolumeChart(localJson, sectorKey);
    renderSpreadChart(localJson, sectorKey);
    renderAssetChart(localJson, sectorKey);
    renderCostChart(localJson, sectorKey);

    // ── Earning Release & YTD Evaluation (New Component) ──
    renderQuarterlyAndYTDEvaluation(ticker, liveData, localJson, cfg);

    // ── Generate side-by-side chart commentaries ─────────
    generateChartCommentaries(ticker, annualYears, ttmQuarters, cfg, localJson);

    // ── Valuation scenarios ──────────────────────────────
    renderValuationScenarios(localJson?.valuation, currentPrice, details, latestRatio, cfg);
    renderValuationSnapshot(localJson);
    renderFinancialSnapshotTable(localJson, sectorKey);
    renderAssumptionsTable(localJson, sectorKey);
    renderResidualIncomeTable(localJson, sectorKey);
    renderPeerBenchmarkTable(localJson, sectorKey);
    renderNiiBreakdownChart(localJson, sectorKey);
    renderNplLlrChart(localJson, sectorKey);
    renderEarningAssetsChart(localJson, sectorKey);
    renderOperationCharts(localJson, sectorKey);

    // ── Qualitative sections ─────────────────────────────
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
    [chartRevNpat, chartEpsEquity, chartPE, chartPB, chartCreditFundingGrowth, chartVolume, chartSpread, chartAsset, chartCost].forEach(c => { if (c) c.destroy(); });
    chartRevNpat = chartEpsEquity = chartPE = chartPB = chartCreditFundingGrowth = chartVolume = chartSpread = chartAsset = chartCost = null;
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
// STEEL-SPECIFIC CHART RENDERERS
// ═══════════════════════════════════════════════════════════

function renderVolumeChart(localJson, sectorKey) {
    const card = document.getElementById('chart-volume-card');
    if (sectorKey !== 'materials' || !localJson?.quarterly?.hrcSales) { card.style.display = 'none'; return; }
    card.style.display = '';
    const q = localJson.quarterly;
    const labels = q.labels || [];
    const ctx = document.getElementById('volumeChart').getContext('2d');
    if (chartVolume) chartVolume.destroy();
    chartVolume = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'HRC (nghìn tấn)', data: q.hrcSales, borderColor: '#1F4E79', backgroundColor: 'rgba(31,78,121,0.1)', borderWidth: 2.5, pointRadius: 4, tension: 0.3, fill: true },
                { label: 'Thép XD (nghìn tấn)', data: q.xdSales, borderColor: '#E74C3C', borderWidth: 2, pointRadius: 4, tension: 0.3, borderDash: [5,3] }
            ]
        },
        options: { ...CHART_DEFAULTS, plugins: { ...CHART_DEFAULTS.plugins, legend: { ...CHART_DEFAULTS.plugins.legend, position: 'top' } } }
    });
    const analysisEl = document.getElementById('analysis-text-volume');
    if (analysisEl) {
        const lastHrc = q.hrcSales?.[q.hrcSales.length - 1] || 0;
        const prevHrc = q.hrcSales?.[q.hrcSales.length - 2] || 0;
        const lastXd = q.xdSales?.[q.xdSales.length - 1] || 0;
        const hrcGrowth = prevHrc ? ((lastHrc - prevHrc) / prevHrc * 100).toFixed(0) : 0;
        analysisEl.innerHTML = `HRC Q1/2026: <b>${lastHrc.toLocaleString('vi-VN')} nghìn tấn</b> (${hrcGrowth >= 0 ? '+' : ''}${hrcGrowth}% QoQ).<br/>Thép XD: <b>${lastXd.toLocaleString('vi-VN')} nghìn tấn</b>.<br/>DQ2 full năm 2026 giúp HRC tăng vọt từ 2025.`;
    }
}

function renderSpreadChart(localJson, sectorKey) {
    const card = document.getElementById('chart-spread-card');
    if (sectorKey !== 'materials' || !localJson?.quarterly?.spreadUsd) { card.style.display = 'none'; return; }
    card.style.display = '';
    const q = localJson.quarterly;
    const labels = q.labels || [];
    const ctx = document.getElementById('spreadChart').getContext('2d');
    if (chartSpread) chartSpread.destroy();
    chartSpread = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'Spread (USD/t)', data: q.spreadUsd, backgroundColor: '#2980B9', borderRadius: 3, order: 2, yAxisID: 'y' },
                { label: 'GP Margin (%)', data: q.gpMargin, type: 'line', borderColor: '#E67E22', backgroundColor: 'rgba(230,126,34,0.1)', borderWidth: 2.5, pointRadius: 4, tension: 0.3, fill: true, order: 1, yAxisID: 'y1' },
                { label: 'Giá HRC (USD/t)', data: q.hrcPrice, type: 'line', borderColor: '#7F8C8D', borderWidth: 1.5, pointRadius: 2, tension: 0.3, borderDash: [3,3], order: 3, yAxisID: 'y' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'USD/t', color: '#545f74', font: { size: 10 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'GP Margin %', color: '#545f74', font: { size: 10 } } }
            }
        }
    });
    const analysisEl = document.getElementById('analysis-text-spread');
    if (analysisEl) {
        const lastGpm = q.gpMargin?.[q.gpMargin.length - 1] || 0;
        const lastSpread = q.spreadUsd?.[q.spreadUsd.length - 1] || 0;
        analysisEl.innerHTML = `Spread Q1/2026: <b>~${lastSpread} USD/t</b>.<br/>GP Margin: <b>${lastGpm}%</b>.<br/>DQ2 full + giá HRC phục hồi giúp spread nở ra.`;
    }
}

function renderAssetChart(localJson, sectorKey) {
    const card = document.getElementById('chart-asset-card');
    if (sectorKey !== 'materials' || !localJson?.quarterly?.totalAssets) { card.style.display = 'none'; return; }
    card.style.display = '';
    const q = localJson.quarterly;
    const labels = q.labels || [];
    const ctx = document.getElementById('assetChart').getContext('2d');
    if (chartAsset) chartAsset.destroy();
    chartAsset = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'Tổng TS (tỷ)', data: q.totalAssets, borderColor: '#2E4057', backgroundColor: 'rgba(46,64,87,0.05)', borderWidth: 2.5, pointRadius: 4, tension: 0.3, fill: true },
                { label: 'Tồn kho (tỷ)', data: q.inventory, borderColor: '#E74C3C', borderWidth: 2, pointRadius: 3, tension: 0.3 },
                { label: 'Phải thu (tỷ)', data: q.receivables, borderColor: '#F39C12', borderWidth: 2, pointRadius: 3, tension: 0.3, borderDash: [5,3] },
                { label: 'Nợ vay (tỷ)', data: q.totalDebt, borderColor: '#27AE60', borderWidth: 2, pointRadius: 3, tension: 0.3, borderDash: [3,3] }
            ]
        },
        options: { ...CHART_DEFAULTS, plugins: { ...CHART_DEFAULTS.plugins, legend: { ...CHART_DEFAULTS.plugins.legend, position: 'top' } } }
    });
    const analysisEl = document.getElementById('analysis-text-asset');
    if (analysisEl) {
        const lastInv = q.inventory?.[q.inventory.length - 1] || 0;
        const lastAssets = q.totalAssets?.[q.totalAssets.length - 1] || 0;
        const invRatio = lastAssets > 0 ? (lastInv / lastAssets * 100).toFixed(0) : 0;
        analysisEl.innerHTML = `Tồn kho: <b>${lastInv.toLocaleString('vi-VN')} tỷ</b> (${invRatio}% tổng TS).<br/>DQ2 full giúp vòng quay tài sản cải thiện.`;
    }
}

function renderCostChart(localJson, sectorKey) {
    const card = document.getElementById('chart-cost-card');
    if (sectorKey !== 'materials' || !localJson?.quarterly?.sgkaRatio) { card.style.display = 'none'; return; }
    card.style.display = '';
    const q = localJson.quarterly;
    const labels = q.labels || [];
    const ctx = document.getElementById('costChart').getContext('2d');
    if (chartCost) chartCost.destroy();
    chartCost = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'SG&A/DT (%)', data: q.sgkaRatio, borderColor: '#8E44AD', backgroundColor: 'rgba(142,68,173,0.1)', borderWidth: 2.5, pointRadius: 4, tension: 0.3, fill: true, yAxisID: 'y' },
                { label: 'CP Bán hàng (tỷ)', data: q.sellExpense, borderColor: '#E67E22', borderWidth: 2, pointRadius: 3, tension: 0.3, yAxisID: 'y1' },
                { label: 'CP QLDN (tỷ)', data: q.adminExpense, borderColor: '#3498DB', borderWidth: 2, pointRadius: 3, tension: 0.3, borderDash: [5,3], yAxisID: 'y1' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { ...CHART_DEFAULTS.scales.x },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: '% Doanh thu', color: '#545f74', font: { size: 10 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Tỷ VND', color: '#545f74', font: { size: 10 } } }
            }
        }
    });
    const analysisEl = document.getElementById('analysis-text-cost');
    if (analysisEl) {
        const lastSgka = q.sgkaRatio?.[q.sgkaRatio.length - 1] || 0;
        analysisEl.innerHTML = `SG&A/DT Q1/2026: <b>${lastSgka}%</b>.<br/>DQ2 giúp tối ưu chi phí bán hàng và QLDN nhờ hiệu suất quy mô.`;
    }
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
    document.getElementById('comment-business').textContent  = data.businessModel || data.overall || 'Chưa có đánh giá mô hình kinh doanh.';
    document.getElementById('comment-financial').textContent = data.financialPerformance || data.financial || 'Chưa có đánh giá sức khỏe tài chính.';
    document.getElementById('comment-valuation').textContent = data.valuationText || data.valuation || 'Chưa có đánh giá định giá.';
}

function renderFinancialTable(annualYears, cfg, labels, localJson) {
    const table = document.getElementById('financial-table');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');
    if (!table || !thead || !tbody) return;

    // Build lists of years to display: historical years + 2026F + 2027F
    const mergedYears = [...labels];
    const forecastYears = [2026, 2027];
    forecastYears.forEach(fy => {
        if (!mergedYears.includes(fy.toString()) && !mergedYears.includes(fy.toString() + 'F')) {
            mergedYears.push(fy.toString() + 'F');
        }
    });

    thead.innerHTML = '<th>Chỉ tiêu</th>' + mergedYears.map(y => `<th>${y}</th>`).join('');
    tbody.innerHTML = '';

    const rows = [
        { label: cfg.incomeLabel + ' (tỷ)', field: cfg.incomeField, div: 1e9 },
        { label: cfg.income2Label + ' (tỷ)', field: cfg.incomeField2, div: 1e9 },
        { label: 'LNST (tỷ)', field: cfg.npat, div: 1e9, highlight: true },
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
                        if (row.field === cfg.incomeField) {
                            rawVal = fcData.revenue?.[fcIdx];
                        } else if (row.field === cfg.incomeField2) {
                            const yearsList = localJson?.years || fcData.years;
                            const yrIdx = yearsList.indexOf(numericYear);
                            if (yrIdx !== -1 && localJson?.toi) {
                                rawVal = localJson.toi[yrIdx];
                            } else {
                                rawVal = fcData.toi?.[fcIdx];
                            }
                        } else if (row.field === cfg.npat) {
                            rawVal = fcData.npat?.[fcIdx];
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
    let isrecs = liveData.incomeYears?.quarters || [];
    
    // ── Fallback: read from localJson.income_quarterly if live API has no quarters ──
    if (isrecs.length < 5 && localJson?.income_quarterly?.length >= 5) {
        // income_quarterly records have {yearReport, quarter, nii, npat} in billion VND
        // We map them to a format compatible with the existing logic below
        // but using the local fields directly (nii → cfg.incomeField, npat → cfg.npat)
        isrecs = localJson.income_quarterly.map(r => ({
            yearReport: r.yearReport,
            quarter: r.quarter,
            // Store raw values in billion VND — we'll use _local_ flag below
            _nii_ty: r.nii,
            _npat_ty: r.npat,
            // Fake the field names so generic code below still works (values × 1e9)
            [cfg.incomeField]: (r.nii || 0) * 1e9,
            [cfg.npat]: (r.npat || 0) * 1e9,
        }));
    }

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
    const sameQLastYear = quarters.find(q => q.yearReport === latestQ.yearReport - 1 && q.quarter === latestQ.quarter);
    
    // 1. Calculate growth metrics
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

    // 3. Extract quarterly ratios if available
    let ratioTextHTML = '';
    if (localJson?.ratios_quarterly) {
        const rq = localJson.ratios_quarterly;
        const idx = rq.quarters.indexOf(`${latestQ.yearReport}-Q${latestQ.quarter}`);
        if (idx !== -1) {
            const nimVal = rq.nim[idx] ? (rq.nim[idx] * 100).toFixed(2) + '%' : '-';
            const ldrVal = rq.ldr[idx] ? (rq.ldr[idx] * 100).toFixed(1) + '%' : '-';
            const casaVal = rq.casa[idx] ? (rq.casa[idx] * 100).toFixed(1) + '%' : '-';
            const nplVal = rq.npl[idx] ? rq.npl[idx].toFixed(2) + '%' : '-';
            
            ratioTextHTML = `
                <div class="q-ytd-item">
                    <div class="q-ytd-header">
                        <span class="q-ytd-title">Chỉ số Sức khỏe Quý gần nhất</span>
                        <span style="font-size:0.75rem;color:var(--text-dim)">Hệ số Quý tự tính toán</span>
                    </div>
                    <div class="q-ytd-metrics">
                        <div class="q-metric-row">
                            <span class="q-metric-label">${ticker === 'TCB' || ticker === 'MBB' || ticker === 'VPB' ? 'Tỷ lệ NIM Quý (năm hóa)' : 'Biên lợi nhuận gộp'}</span>
                            <span class="q-metric-value" style="color:var(--sector-color)">${nimVal}</span>
                        </div>
                        <div class="q-metric-row">
                            <span class="q-metric-label">${ticker === 'TCB' || ticker === 'MBB' || ticker === 'VPB' ? 'Tỷ lệ LDR Quý' : 'Nợ/Vốn chủ sở hữu'}</span>
                            <span class="q-metric-value">${ldrVal}</span>
                        </div>
                        <div class="q-metric-row">
                            <span class="q-metric-label">${ticker === 'TCB' || ticker === 'MBB' || ticker === 'VPB' ? 'Tỷ lệ CASA Quý' : 'Tỷ suất ROA Quý'}</span>
                            <span class="q-metric-value">${casaVal}</span>
                        </div>
                        <div class="q-metric-row">
                            <span class="q-metric-label">${ticker === 'TCB' || ticker === 'MBB' || ticker === 'VPB' ? 'Tỷ lệ Nợ xấu NPL Quý' : 'Vòng quay tổng tài sản'}</span>
                            <span class="q-metric-value" style="color:#ef4444">${nplVal}</span>
                        </div>
                    </div>
                </div>
            `;
        }
    }

    // 4. Generate summary texts
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
    
    // Inject structural bank quality commentary from Python json export
    if (localJson && localJson.analysis_comments) {
        const ac = localJson.analysis_comments;
        commentary += `<br><br><strong>🎯 Đánh giá chất lượng tài sản & nguồn gốc lợi nhuận thực tế:</strong><br>`;
        commentary += `• Tăng trưởng tín dụng Cho vay lũy kế đạt <strong>${ac.ytd_loans_growth >= 0 ? '+' : ''}${ac.ytd_loans_growth}% YTD</strong> so với đầu năm. Tăng trưởng huy động tiền gửi đạt <strong>${ac.ytd_dep_growth >= 0 ? '+' : ''}${ac.ytd_dep_growth}% YTD</strong>.<br>`;
        commentary += `• <strong>Chất lượng nguồn thu:</strong> ${ac.profit_source_comment}<br>`;
        commentary += `• <strong>Áp lực dự phòng:</strong> ${ac.provision_comment}`;
    }

    container.innerHTML = `
        <div class="quarterly-ytd-grid" style="grid-template-columns: repeat(auto-fit, minmax(320px, 1fr))">
            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">Doanh thu Quý gần nhất (${latestQ.quarter}Q/${latestQ.yearReport})</span>
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
                    <div class="q-metric-row">
                        <span class="q-metric-label">Lũy kế YTD cả năm</span>
                        <span class="q-metric-value">${formatNumber(ytdRev/1e9, 1)}</span>
                    </div>
                </div>
            </div>

            <div class="q-ytd-item">
                <div class="q-ytd-header">
                    <span class="q-ytd-title">Lợi nhuận Quý gần nhất</span>
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
                    <div class="q-metric-row">
                        <span class="q-metric-label">Lũy kế LNST YTD cả năm</span>
                        <span class="q-metric-value">${formatNumber(ytdNpat/1e9, 1)} (${formatGrowth(yoyYtdNpat)})</span>
                    </div>
                </div>
            </div>

            ${ratioTextHTML}
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

function renderCreditFundingGrowthChart(localJson, cfg) {
    const card = document.getElementById('chart-growth-card');
    const analysisEl = document.getElementById('analysis-text-growth');
    if (!card) return;

    if (cfg.label !== 'Ngân Hàng' || !localJson || !localJson.credit_funding_growth) {
        card.style.display = 'none';
        return;
    }

    card.style.display = 'flex';
    const cg = localJson.credit_funding_growth;
    const ctx = document.getElementById('creditFundingGrowthChart').getContext('2d');
    
    if (chartCreditFundingGrowth) chartCreditFundingGrowth.destroy();
    
    chartCreditFundingGrowth = new Chart(ctx, {
        type: 'line',
        data: {
            labels: cg.quarters,
            datasets: [
                {
                    label: 'Tăng trưởng Tín dụng QoQ (%)',
                    data: cg.credit_ytd,
                    borderColor: '#4472C4',
                    backgroundColor: 'rgba(68,114,196,0.1)',
                    borderWidth: 2.5,
                    pointRadius: 4,
                    tension: 0.2
                },
                {
                    label: 'Tăng trưởng Huy động QoQ (%)',
                    data: cg.funding_ytd,
                    borderColor: '#ED7D31',
                    backgroundColor: 'rgba(237,125,49,0.1)',
                    borderWidth: 2.5,
                    pointRadius: 4,
                    tension: 0.2
                }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: {
                    ...CHART_DEFAULTS.scales.x,
                    ticks: {
                        ...CHART_DEFAULTS.scales.x.ticks,
                        callback: function(val, index) {
                            return index % 2 === 0 ? this.getLabelForValue(val) : '';
                        }
                    }
                },
                y: { ...CHART_DEFAULTS.scales.y }
            }
        }
    });

    const lastCreditQoq = cg.credit_ytd[cg.credit_ytd.length - 1];
    const lastFundingQoq = cg.funding_ytd[cg.funding_ytd.length - 1];
    
    analysisEl.innerHTML = `
        Tăng trưởng tín dụng liên quý QoQ gần nhất đạt <strong>${lastCreditQoq >= 0 ? '+' : ''}${lastCreditQoq.toFixed(2)}%</strong>. <br>
        Tăng trưởng huy động liên quý QoQ tương ứng đạt <strong>${lastFundingQoq >= 0 ? '+' : ''}${lastFundingQoq.toFixed(2)}%</strong>. <br>
        Mức độ tương đồng tăng trưởng phản ánh sự kiểm soát nhịp nhàng thanh khoản và tỷ lệ LDR pháp lý.
    `;
}

// ═══════════════════════════════════════════════════════════
// FINANCIAL SNAPSHOT TABLE — mirrors PDF Page 1 summary table
// ═══════════════════════════════════════════════════════════
function renderFinancialSnapshotTable(localJson, sectorKey) {
    const card  = document.getElementById('web-financial-snapshot-card');
    const tbody = document.getElementById('web-financial-snapshot-tbody');
    if (!card || !tbody) return;
    const fs = localJson?.financial_snapshot;
    if (!fs || sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = '';
    const fcStyle = 'color:#60a5fa';
    const fmtN = (v, d=1) => v != null ? Number(v).toFixed(d) : '-';
    const rows = [
        { lbl: 'NII — Thu nhap lai thuan (ty VND)', arr: fs.nii,  fmt: (v)=>fmtN(v,0) },
        { lbl: 'TOI — Tong thu nhap HD (ty VND)',   arr: fs.toi,  fmt: (v)=>fmtN(v,0) },
        { lbl: 'NPAT — LN sau thue (ty VND)',        arr: fs.npat, fmt: (v)=>fmtN(v,0) },
        { lbl: 'NIM — Bien lai rong (%)',            arr: fs.nim,  fmt: (v)=>fmtN(v,2)+'%' },
        { lbl: 'ROE — Ty suat LN (%)',               arr: fs.roe,  fmt: (v)=>fmtN(v,1)+'%' },
        { lbl: 'LDR — Ty le tin dung/huy dong (%)', arr: fs.ldr,  fmt: (v)=>fmtN(v,1)+'%' },
        { lbl: 'NPL — Ty le no xau (%)',             arr: fs.npl,  fmt: (v)=>fmtN(v,2)+'%' },
    ];
    tbody.innerHTML = rows.map(r => {
        const cells = (r.arr || []).map((v, i) =>
            `<td ${i >= 3 ? `style="${fcStyle}"` : ''}>${r.fmt(v)}</td>`
        ).join('');
        return `<tr><td>${r.lbl}</td>${cells}</tr>`;
    }).join('');
}

// ═══════════════════════════════════════════════════════════
// VALUATION SNAPSHOT — fills the summary card at top
// ═══════════════════════════════════════════════════════════
function renderValuationSnapshot(localJson) {
    const val = localJson?.valuation;
    const currPrice = localJson?.currentPrice;
    if (!val || !currPrice) return;
    const fmt = (n) => n ? Math.round(n).toLocaleString('vi-VN') + ' d' : '-';
    const upside = val.upside ?? 0;
    const recText  = val.recommend ?? (upside >= 15 ? 'MUA' : upside < -5 ? 'BAN' : 'THEO DOI');
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
    if (el('snap-ev-price')) el('snap-ev-price').textContent = val.evEbitdaPrice ? Math.round(val.evEbitdaPrice).toLocaleString('vi-VN') + ' d' : '-';
    if (el('snap-pb-price')) el('snap-pb-price').textContent = val.pbPrice ? Math.round(val.pbPrice).toLocaleString('vi-VN') + ' d' : '-';
    if (el('snap-pb-attr')) el('snap-pb-attr').textContent = val.pbAttr ? Math.round(val.pbAttr).toLocaleString('vi-VN') + ' d' : '-';
    if (el('snap-pb-upper')) el('snap-pb-upper').textContent = val.pbUpper ? Math.round(val.pbUpper).toLocaleString('vi-VN') + ' d' : '-';
    if (el('snap-pe-price')) el('snap-pe-price').textContent = val.pePrice ? Math.round(val.pePrice).toLocaleString('vi-VN') + ' d' : '-';
}

// ═══════════════════════════════════════════════════════════
// ASSUMPTIONS TABLE
// ═══════════════════════════════════════════════════════════
function renderAssumptionsTable(localJson, sectorKey) {
    const card  = document.getElementById('web-assumptions-card');
    const tbody = document.getElementById('web-assumptions-tbody');
    if (!card || !tbody) return;
    const a = localJson?.assumptions;
    const ft = localJson?.forecast_text;
    if (!a || sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = '';
    const fmtPct = (v) => v != null ? Number(v).toFixed(1) + '%' : '-';
    const fmtPct2 = (v) => v != null ? Number(v).toFixed(2) + '%' : '-';
    const rows = [
        { label: 'Tang truong Tin dung', key: 'loans_growth', fmt: fmtPct },
        { label: 'Tang truong Huy dong', key: 'dep_growth',   fmt: fmtPct },
        { label: 'NIM — Bien lai rong',  key: 'nim',          fmt: fmtPct2 },
        { label: 'CIR — Ty le chi phi',  key: 'cir',          fmt: fmtPct },
        { label: 'CoC — Chi phi tin dung',key: 'coc',         fmt: fmtPct2 },
        { label: 'NPL — No xau muc tieu',key: 'npl',          fmt: fmtPct2 },
    ];
    tbody.innerHTML = rows.map(r => {
        const arr = a[r.key] || [];
        const cells = [0,1,2].map(i => `<td>${r.fmt(arr[i])}</td>`).join('');
        return `<tr><td>${r.label}</td>${cells}</tr>`;
    }).join('');
    if (ft) {
        const existingNote = card.querySelector('.forecast-note');
        if (existingNote) existingNote.remove();
        const note = document.createElement('div');
        note.className = 'forecast-note q-commentary-box';
        note.innerHTML = `<strong>Co so gia dinh:</strong> Tin dung 2026F +${ft.loans_g_26}%, Huy dong +${ft.dep_g_26}%, NIM ${ft.nim_26}% (${ft.nim_trend}), CoC ${ft.coc_26}% (${ft.coc_trend}), CIR ${ft.cir_26}%.`;
        card.appendChild(note);
    }
}

// ═══════════════════════════════════════════════════════════
// RESIDUAL INCOME TABLE — full RI model like PDF Page 5
// ═══════════════════════════════════════════════════════════
function renderResidualIncomeTable(localJson, sectorKey) {
    const card    = document.getElementById('web-detailed-ri-card');
    const tbody   = document.getElementById('web-ri-tbody');
    const summBox = document.getElementById('web-ri-summary-box');
    if (!card || !tbody) return;
    const val = localJson?.valuation;
    if (!val || sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = '';
    const fmt  = (v) => v != null ? Math.round(v).toLocaleString('vi-VN') : '-';
    const fmtD = (v, d=4) => v != null ? Number(v).toFixed(d) : '-';
    const bvps = val.bvpsBase || 0;
    const coe  = (val.COE || 0) / 100;
    const eps  = val.epsFc  || [];
    const ri   = val.riResults || [];
    const bvps1 = bvps;
    const bvps2 = bvps + (eps[0] || 0);
    const bvps3 = bvps + (eps[0] || 0) + (eps[1] || 0);
    const riRows = [
        ['EPS du phong (VND/CP)',        fmt(eps[0]),  fmt(eps[1]),  fmt(eps[2])],
        ['BVPS dau ky (VND/CP)',         fmt(bvps1),   fmt(bvps2),   fmt(bvps3)],
        ['Capital Charge = COE x BVPS',  fmt(bvps1*coe), fmt(bvps2*coe), fmt(bvps3*coe)],
        ['Loi nhuan thang du (RI)',       fmt(ri[0]),   fmt(ri[1]),   fmt(ri[2])],
        ['He so chiet khau',             fmtD(1/(1+coe)), fmtD(1/(1+coe)**2), fmtD(1/(1+coe)**3)],
        ['PV cua RI tung nam',           fmt((ri[0]||0)/(1+coe)), fmt((ri[1]||0)/(1+coe)**2), fmt((ri[2]||0)/(1+coe)**3)],
    ];
    tbody.innerHTML = riRows.map(([lbl, ...vals]) =>
        `<tr><td>${lbl}</td>${vals.map(v => `<td>${v}</td>`).join('')}</tr>`
    ).join('');
    if (summBox) {
        const up = val.upside ?? 0;
        summBox.innerHTML = `
            <strong>BVPS hien tai:</strong> ${fmt(bvps)} VND &nbsp;|&nbsp;
            <strong>Tong PV(RI) 3 nam:</strong> ${fmt(val.pvRi)} VND<br>
            <strong>PV Continuing Value (-50% bao thu):</strong> ${fmt(val.pvCv)} VND<br>
            <strong>Gia tri hop ly theo RI:</strong> <b>${fmt(val.riValue)} VND</b> &nbsp;|&nbsp;
            <strong>Theo P/B:</strong> <b>${fmt(val.pbValue)} VND</b><br>
            <strong>Gia muc tieu Weighted (50% RI + 50% P/B):</strong>
            <span style="color:#3b82f6;font-weight:800;font-size:1.05em"> ${fmt(val.weightedTarget)} VND/CP</span>
            <span style="color:${up >= 0 ? '#10b981' : '#ef4444'};font-weight:700">(${up >= 0 ? '+' : ''}${up.toFixed(1)}% upside)</span><br>
            <strong>Gia Bull (P/B target ${(val.pbTarget||0).toFixed(2)}x):</strong> ${fmt(val.pbTargetPrice)} VND &nbsp;|&nbsp;
            <strong>Gia Bear (P/B attractive ${(val.pbAttractive||0).toFixed(2)}x):</strong> ${fmt(val.pbAttractivePrice)} VND
        `;
    }
}

// ═══════════════════════════════════════════════════════════
// PEER BENCHMARK TABLE (Duy trì dữ liệu động từ API)
// ═══════════════════════════════════════════════════════════
async function renderPeerBenchmarkTable(localJson, sectorKey) {
    const card  = document.getElementById('web-peer-card');
    const tbody = document.getElementById('web-peer-tbody');
    if (!card || !tbody) return;

    if (sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = '';

    const currentTicker = localJson?.ticker ?? '';
    const fmtN = (v, d=2) => v != null ? v.toFixed(d) : '-';
    const fmtM = (v) => v != null ? Math.round(v).toLocaleString('vi-VN') : '-';

    try {
        // Fetch dynamically from our backend updated JSON file
        const resp = await fetch('data/peer_benchmark.json');
        if (!resp.ok) throw new Error("Failed to load peer_benchmark.json");
        const data = await resp.json();
        const peerList = data.peers || [];

        const avg = (key) => {
            const vals = peerList.map(p => p[key]).filter(v => v != null);
            return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
        };

        const avgNPL = avg('npl'), avgROE = avg('roe'), avgCG = avg('cg');

        // Update overall date label based on current ticker or database updated date
        const dateEl = document.getElementById('peer-update-date');
        const currentPeer = peerList.find(p => p.ticker === currentTicker);
        const updateDate = currentPeer?.date_updated || data._meta?.date || new Date().toISOString().split('T')[0];
        if (dateEl) {
            dateEl.innerHTML = `Số liệu tính toán tự động dựa trên thị giá đóng cửa ngày cập nhật mới nhất của <strong>${currentTicker || 'ngân hàng'}</strong>: <span style="color:#f59e0b; font-weight:bold">${updateDate}</span>`;
        }

        const avgRow = `<tr style="border-top:1.5px solid rgba(255,255,255,0.12);color:#f59e0b;font-weight:600">
            <td>⚡ Trung bình ngành</td>
            <td>${fmtN(avgNPL)}</td>
            <td>${fmtN(avg('nim'))}</td>
            <td>${fmtN(avg('casa'), 1)}</td>
            <td>${fmtN(avgROE, 1)}</td>
            <td>${fmtN(avg('cir'), 1)}</td>
            <td>${fmtN(avg('pb'))}</td>
            <td>${fmtN(avgCG, 1)}</td>
            <td>—</td>
        </tr>`;

        const rows = peerList.map(p => {
            const isCurrent = p.ticker === currentTicker;
            const rowStyle  = isCurrent ? 'background:rgba(59,130,246,0.10);font-weight:700;' : '';
            const nplColor  = p.npl > 3 ? '#ef4444' : p.npl > 2 ? '#f59e0b' : '#10b981';
            const pbColor   = p.pb  > 2 ? '#10b981' : p.pb  > 1 ? '' : '#ef4444';
            const cgColor   = p.cg >= avgCG ? '#10b981' : '#f59e0b';
            const roeColor  = p.roe >= avgROE ? '#10b981' : '';
            const dateStr   = p.date_updated ? `<br><span style="font-size:0.75em;color:#718096;font-weight:normal">Cập nhật: ${p.date_updated}</span>` : '';
            return `<tr style="${rowStyle}">
                <td>${isCurrent ? '👉 ' : ''}${p.ticker} — ${p.name}${dateStr}</td>
                <td style="color:${nplColor}">${fmtN(p.npl)}</td>
                <td>${fmtN(p.nim)}</td>
                <td>${fmtN(p.casa, 1)}</td>
                <td style="color:${roeColor}">${fmtN(p.roe, 1)}</td>
                <td>${fmtN(p.cir, 1)}</td>
                <td style="color:${pbColor}">${fmtN(p.pb)}</td>
                <td style="color:${cgColor}">${fmtN(p.cg, 1)}</td>
                <td>${fmtM(p.mcap)}</td>
            </tr>`;
        }).join('');

        tbody.innerHTML = avgRow + rows;
    } catch (e) {
        console.warn("Peer Benchmark data loading error, falling back to static presentation:", e);
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:#ef4444">Không thể tải dữ liệu So sánh Ngành động.</td></tr>`;
    }
}

// ═══════════════════════════════════════════════════════════
// NII & NON-II BREAKDOWN CHART (QUARTERLY STACKED BAR)
// ═══════════════════════════════════════════════════════════
let chartNiiBreakdown = null;
function renderNiiBreakdownChart(localJson, sectorKey) {
    const card = document.getElementById('chart-nii-card');
    const analysisEl = document.getElementById('analysis-text-nii');
    if (!card || !analysisEl) return;
    const data = localJson?.income_quarterly;
    if (!data || sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = 'flex';
    const ctx = document.getElementById('niiBreakdownChart').getContext('2d');
    if (chartNiiBreakdown) chartNiiBreakdown.destroy();
    const quarters = data.map(d => d.quarter);
    const nii = data.map(d => d.nii);
    const nonii = data.map(d => d.nonii);
    chartNiiBreakdown = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: quarters,
            datasets: [
                {
                    label: 'NII',
                    data: nii,
                    backgroundColor: '#3b82f6'
                },
                {
                    label: 'NonII',
                    data: nonii,
                    backgroundColor: '#10b981'
                }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    });
    const lastNii = nii[nii.length - 1] || 0;
    const lastNonii = nonii[nonii.length - 1] || 0;
    const noniiPct = (lastNonii / (lastNii + lastNonii) * 100).toFixed(1);
    analysisEl.innerHTML = `NII: <strong>${lastNii.toLocaleString('vi-VN')} ty</strong> | NonII: <strong>${lastNonii.toLocaleString('vi-VN')} ty</strong> (chiem ${noniiPct}%).`;
}

// ═══════════════════════════════════════════════════════════
// NPL % & LLR % CHART
// ═══════════════════════════════════════════════════════════
let chartNplLlr = null;
function renderNplLlrChart(localJson, sectorKey) {
    const card = document.getElementById('chart-npl-card');
    const analysisEl = document.getElementById('analysis-text-npl');
    if (!card || !analysisEl) return;
    const rq = localJson?.ratios_quarterly;
    if (!rq || sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = 'flex';
    const ctx = document.getElementById('nplLlrChart').getContext('2d');
    if (chartNplLlr) chartNplLlr.destroy();
    const quarters = rq.quarters;
    const npl = rq.npl || [];
    const llr = rq.llr || [];
    chartNplLlr = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: quarters,
            datasets: [
                {
                    type: 'bar',
                    label: 'NPL (%)',
                    data: npl,
                    backgroundColor: '#ef4444',
                    yAxisID: 'y'
                },
                {
                    type: 'line',
                    label: 'LLR (%)',
                    data: llr,
                    borderColor: '#f59e0b',
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                y: { title: { display: true, text: 'NPL (%)' } },
                y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'LLR (%)' } }
            }
        }
    });
    const lastNpl = npl[npl.length - 1] || 0;
    const lastLlr = llr[llr.length - 1] || 0;
    analysisEl.innerHTML = `NPL: <strong>${lastNpl.toFixed(2)}%</strong> | LLR: <strong>${lastLlr.toFixed(1)}%</strong>.`;
}

// ═══════════════════════════════════════════════════════════
// EARNING ASSETS STRUCTURE AREA CHART (highly contrasting colors)
// ═══════════════════════════════════════════════════════════
let chartEarningAssets = null;
function renderEarningAssetsChart(localJson, sectorKey) {
    const card = document.getElementById('chart-iea-card');
    const analysisEl = document.getElementById('analysis-text-iea');
    if (!card || !analysisEl) return;
    const ea = localJson?.earning_assets_quarterly || localJson?.earning_assets;
    if (!ea || sectorKey !== 'banks') { card.style.display = 'none'; return; }
    card.style.display = 'flex';
    const ctx = document.getElementById('earningAssetsChart').getContext('2d');
    if (chartEarningAssets) chartEarningAssets.destroy();
    chartEarningAssets = new Chart(ctx, {
        type: 'line',
        data: {
            labels: ea.quarters || ea.years,
            datasets: [
                { label: 'Inv Sec', data: ea.inv_sec, fill: true, backgroundColor: 'rgba(142, 68, 173, 0.8)' },
                { label: 'Loans', data: ea.loans, fill: true, backgroundColor: 'rgba(41, 128, 185, 0.8)' },
                { label: 'Bank Dep', data: ea.bank_dep, fill: true, backgroundColor: 'rgba(39, 174, 96, 0.8)' },
                { label: 'Cash & SBV', data: ea.cash_sbv, fill: true, backgroundColor: 'rgba(231, 76, 60, 0.8)' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            scales: {
                y: { stacked: true }
            }
        }
    });
    const lastLoans = ea.loans[ea.loans.length - 1] || 0;
    analysisEl.innerHTML = `Cho vay khách hàng quý gần nhất đạt <strong>${lastLoans.toLocaleString('vi-VN')} tỷ</strong>, là động lực phát triển cốt lõi.`;
}

// ── 4 NEW QUARTERLY BREAKDOWN CHARTS FOR BANKS ──────────────────────
let chartLoanIndustry = null;
let chartNplGroups = null;
let chartLoanTerms = null;
let chartDepositTypes = null;

function renderOperationCharts(localJson, sectorKey) {
    const card = document.getElementById('chart-operation-card');
    if (!card) return;
    if (sectorKey !== 'banks') { card.style.display = 'none'; return; }
    
    const indData = localJson?.loan_industry;
    const nplData = localJson?.npl_groups;
    const termData = localJson?.loan_terms;
    const depData = localJson?.deposit_types;
    
    if (!indData || !nplData || !termData || !depData) {
        card.style.display = 'none';
        return;
    }
    card.style.display = 'flex';
    
    // 1. Loan Industry Chart (highly contrasting colors)
    const ctxInd = document.getElementById('loanIndustryChart').getContext('2d');
    if (chartLoanIndustry) chartLoanIndustry.destroy();
    chartLoanIndustry = new Chart(ctxInd, {
        type: 'bar',
        data: {
            labels: indData.quarters,
            datasets: [
                { label: 'Bất động sản', data: indData.real_estate, backgroundColor: '#ED7D31' },
                { label: 'Cá nhân', data: indData.individuals, backgroundColor: '#2E75B6' },
                { label: 'Thương mại & DV', data: indData.wholesale_retail, backgroundColor: '#70AD47' },
                { label: 'Khác', data: indData.others, backgroundColor: '#FFC000' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                ...CHART_DEFAULTS.plugins,
                title: { display: true, text: 'Cơ cấu cho vay theo ngành (tỷ VND)' }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    });
    
    // 2. NPL Groups Chart (Group 1 removed, highly contrasting colors)
    const ctxNpl = document.getElementById('nplGroupsChart').getContext('2d');
    if (chartNplGroups) chartNplGroups.destroy();
    chartNplGroups = new Chart(ctxNpl, {
        type: 'bar',
        data: {
            labels: nplData.quarters,
            datasets: [
                { label: 'Nhóm 2 (Cần chú ý)', data: nplData.group2, backgroundColor: '#FFC000' },
                { label: 'Nhóm 3 (Dưới tiêu chuẩn)', data: nplData.group3, backgroundColor: '#ED7D31' },
                { label: 'Nhóm 4 (Nghi ngờ)', data: nplData.group4, backgroundColor: '#7030A0' },
                { label: 'Nhóm 5 (Mất vốn)', data: nplData.group5, backgroundColor: '#C00000' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                ...CHART_DEFAULTS.plugins,
                title: { display: true, text: 'Biến động nợ nhóm 2-5 (tỷ VND)' }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    });
    
    // 3. Loan Terms Chart (highly contrasting colors)
    const ctxTerm = document.getElementById('loanTermsChart').getContext('2d');
    if (chartLoanTerms) chartLoanTerms.destroy();
    chartLoanTerms = new Chart(ctxTerm, {
        type: 'bar',
        data: {
            labels: termData.quarters,
            datasets: [
                { label: 'Ngắn hạn', data: termData.short_term, backgroundColor: '#2E75B6' },
                { label: 'Trung hạn', data: termData.medium_term, backgroundColor: '#70AD47' },
                { label: 'Dài hạn', data: termData.long_term, backgroundColor: '#FFC000' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                ...CHART_DEFAULTS.plugins,
                title: { display: true, text: 'Cơ cấu kỳ hạn cho vay (tỷ VND)' }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    });
    
    // 4. Deposit Types Chart (highly contrasting colors)
    const ctxDep = document.getElementById('depositTypesChart').getContext('2d');
    if (chartDepositTypes) chartDepositTypes.destroy();
    chartDepositTypes = new Chart(ctxDep, {
        type: 'bar',
        data: {
            labels: depData.quarters,
            datasets: [
                { label: 'Không kỳ hạn (CASA)', data: depData.casa, backgroundColor: '#ED7D31' },
                { label: 'Có kỳ hạn', data: depData.term, backgroundColor: '#2E75B6' },
                { label: 'Ký quỹ & Khác', data: depData.others, backgroundColor: '#7F7F7F' }
            ]
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                ...CHART_DEFAULTS.plugins,
                title: { display: true, text: 'Cơ cấu loại tiền gửi (tỷ VND)' }
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true }
            }
        }
    });
}
