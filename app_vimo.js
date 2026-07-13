/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app_vimo.js (Phân tích Vĩ mô Kinh tế Việt Nam)
   Trang KHÔNG gắn với mã cổ phiếu — tải thẳng data/vimo.json khi mở
   trang, không có bước "chọn mã cổ phiếu" như các dashboard sector khác.
   ════════════════════════════════════════════════════════ */

'use strict';

const GROUP_LABELS = {
    growth: 'Tăng trưởng', inflation: 'Lạm phát', monetary: 'Tiền tệ & Lãi suất',
    trade: 'Thương mại & Vốn', fiscal: 'Tài khóa', labor: 'Lao động',
    external: 'Áp lực bên ngoài', market: 'Thị trường chứng khoán',
};
const GROUP_ORDER = ['growth', 'inflation', 'monetary', 'trade', 'fiscal', 'labor', 'external', 'market'];
const GROUP_ICONS = {
    growth: '📈', inflation: '💰', monetary: '🏦', trade: '🚢',
    fiscal: '🏛️', labor: '👷', external: '🌐', market: '📊',
};
const SOURCE_LABELS = {
    worldbank: 'World Bank API', imf: 'IMF DataMapper API', fred: 'FRED API',
    fx_api: 'exchangerate-api.com', pe_ratio_api: 'worldperatio.com',
    nso_scrape: 'nso.gov.vn (báo cáo quý, tự động)',
    nso_chart_embed: 'nso.gov.vn (biểu đồ tháng, tự động)',
    sbv_chart: 'sbv.gov.vn (biểu đồ, tự động)',
    sbv_table: 'sbv.gov.vn (bảng lãi suất, tự động)',
    vietnambiz: 'data.vietnambiz.vn (tự động)',
    bank_page: 'Trang NH chính thức (tự động)',
    news_rss: 'RSS tin tức CafeF/VietStock (tự động, chỉ khi có tin mới)',
    market_table: '24hmoney.vn (bảng đa ngân hàng, tự động)',
    '24hmoney_scrape': '24hmoney.vn (chỉ số P/E-P/B, tự động)',
    manual: 'Nghiên cứu thủ công',
};

const CHART_DEFAULTS = {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
        x: { ticks: { color: '#545f74', font: { size: 8 }, maxRotation: 45 }, grid: { display: false } },
        y: { ticks: { color: '#545f74', font: { size: 8 } }, grid: { color: 'rgba(255,255,255,0.04)' } },
    },
};

let chartInstances = [];

document.addEventListener('DOMContentLoaded', async () => {
    const data = await fetch('data/vimo.json').then(r => r.ok ? r.json() : null).catch(() => null);
    if (!data) {
        document.getElementById('indicator-groups-container').innerHTML =
            '<div class="loading-state card">Chưa có dữ liệu vĩ mô. Hãy chạy template_vimo.py hoặc GitHub Action "Cập nhật Vĩ mô".</div>';
        return;
    }
    renderHeader(data);
    renderVerdict(data.synthesis && data.synthesis.verdict, data.decision);
    renderScorecard(data.scorecard);
    renderDecision(data.decision, data.scorecard.total);
    renderSynthesis(data.synthesis);
    renderValuation(data.marketValuation);
    renderIndicatorGroups(data.indicators);
});

function renderSynthesis(synthesis) {
    if (!synthesis) return;
    const set = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text || '-'; };
    set('synthesis-overview', synthesis.overview);
    set('synthesis-economy', synthesis.economy_impact);
    set('synthesis-market', synthesis.market_impact);
    set('synthesis-watch', synthesis.watch_points);
}

// Đánh giá Tổng thể — 3 câu hỏi user luôn quan tâm: đang tốt lên/xấu đi (xu hướng so kỳ trước),
// bức tranh rõ ràng hay xám/hỗn hợp (mức đồng thuận giữa các chỉ báo), có phù hợp đầu tư không.
function renderVerdict(verdict, decision) {
    if (!verdict) return;
    const trendEl = document.getElementById('verdict-trend');
    const clarityEl = document.getElementById('verdict-clarity');
    const decisionEl = document.getElementById('verdict-decision');
    const detailEl = document.getElementById('verdict-detail');

    const trendColor = verdict.trend_arrow === '▲' ? '#10b981' : (verdict.trend_arrow === '▼' ? '#ef4444' : '#f59e0b');
    trendEl.textContent = `${verdict.trend_arrow} ${verdict.trend_label}`;
    trendEl.style.color = trendColor;

    const clarityColor = (verdict.clarity_label || '').includes('Sáng') ? '#10b981'
        : (verdict.clarity_label || '').includes('Tối') ? '#ef4444' : '#f59e0b';
    clarityEl.textContent = verdict.clarity_label || '-';
    clarityEl.style.color = clarityColor;

    if (decision) {
        const decisionColor = ['Bung vốn mạnh', 'Duy trì, chọn lọc', 'Mua từ từ, phân kỳ'].includes(decision.label) ? '#10b981' : '#ef4444';
        decisionEl.textContent = decision.label;
        decisionEl.style.color = decisionColor;
    }

    detailEl.textContent = `${verdict.trend_detail || ''} ${verdict.clarity_detail || ''}`.trim();
}

function renderHeader(data) {
    const btnPdf = document.getElementById('download-pdf');
    if (data.gdrivePdfUrl) {
        btnPdf.href = data.gdrivePdfUrl;
        btnPdf.classList.remove('hidden');
    }
    const lu = document.getElementById('last-updated');
    if (lu && data.lastUpdated) lu.textContent = `Cập nhật lần cuối: ${data.lastUpdated}`;
}

function formatNumber(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Number(num).toLocaleString('vi-VN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// ═══════════════════════════════════════════════════════════
// SCORECARD
// ═══════════════════════════════════════════════════════════
function renderScorecard(scorecard) {
    const grid = document.getElementById('scorecard-grid');
    const scoreColor = (s) => s > 0 ? '#10b981' : (s < 0 ? '#ef4444' : '#f59e0b');
    const scoreText = (s) => s > 0 ? '+1 Tốt' : (s < 0 ? '-1 Xấu' : '0 Trung tính');

    let html = Object.entries(scorecard.groups).map(([gname, g]) => `
        <div class="vimo-score-box">
            <span class="lbl">${gname}</span>
            <span class="val" style="color:${scoreColor(g.score)}">${scoreText(g.score)}</span>
            <div style="font-size:0.7em;color:var(--text-muted);margin-top:4px">${g.nVotes} chỉ báo</div>
        </div>
    `).join('');

    const totalColor = scoreColor(scorecard.total);
    html += `
        <div class="vimo-score-box" style="border:2px solid ${totalColor}">
            <span class="lbl">TỔNG SCORECARD</span>
            <span class="val" style="color:${totalColor}">${scorecard.total > 0 ? '+' : ''}${scorecard.total} / 5</span>
        </div>
    `;
    grid.innerHTML = html;
}

function renderDecision(decision, total) {
    const banner = document.getElementById('decision-banner');
    const label = document.getElementById('decision-label');
    const text = document.getElementById('decision-text');
    if (!decision) return;
    banner.style.display = 'block';
    const color = total > 0 ? '#10b981' : (total < 0 ? '#ef4444' : '#f59e0b');
    banner.style.borderLeftColor = color;
    banner.style.background = color + '15';
    label.textContent = `🎯 ${decision.label}`;
    label.style.color = color;
    text.textContent = decision.text;
}

// ═══════════════════════════════════════════════════════════
// MARKET VALUATION
// ═══════════════════════════════════════════════════════════
function renderValuation(val) {
    if (!val) return;
    document.getElementById('val-pe').textContent = val.pe ? `${formatNumber(val.pe)}x` : '-';
    const pbEl = document.getElementById('val-pb');
    if (pbEl) pbEl.textContent = val.pb ? `${formatNumber(val.pb)}x` : '-';
    document.getElementById('val-rf').textContent = val.rf ? `${(val.rf * 100).toFixed(2)}%` : '-';
    document.getElementById('val-erp').textContent = val.erp !== null && val.erp !== undefined ? `${(val.erp * 100).toFixed(2)}%` : '-';
    const labelEl = document.getElementById('val-label');
    labelEl.textContent = val.valuation_label || '-';
    labelEl.style.color = val.valuation_label === 'Rẻ/Hấp dẫn' ? '#10b981'
        : val.valuation_label === 'Đắt/Kém hấp dẫn' ? '#ef4444' : '#f59e0b';
}

// ═══════════════════════════════════════════════════════════
// INDICATOR GROUPS
// ═══════════════════════════════════════════════════════════
function renderIndicatorGroups(indicators) {
    chartInstances.forEach(c => c.destroy());
    chartInstances = [];

    const container = document.getElementById('indicator-groups-container');
    container.innerHTML = '';

    GROUP_ORDER.forEach(grp => {
        const entries = Object.entries(indicators).filter(([, ind]) => ind.group === grp);
        if (!entries.length) return;

        const section = document.createElement('div');
        section.innerHTML = `<div class="vimo-group-header"><h3>${GROUP_ICONS[grp] || ''} ${GROUP_LABELS[grp]}</h3></div>
            <div class="vimo-indicator-grid" id="grid-${grp}"></div>`;
        container.appendChild(section);
        const grid = section.querySelector(`#grid-${grp}`);

        entries.forEach(([key, ind]) => {
            const card = document.createElement('div');
            card.className = 'vimo-indicator-card';
            const t = ind.trend || {};
            const hasChart = (ind.series || []).filter(p => p.value !== null && p.value !== undefined).length >= 4;
            const judgColor = t.judgment_color || '#94a3b8';
            const canvasId = `chart-${key}`;

            card.innerHTML = `
                <div class="ind-header">
                    <span class="ind-name">${ind.label}</span>
                    ${t.judgment_label ? `<span class="ind-judgment" style="background:${judgColor}22;color:${judgColor}">${t.value_arrow || ''} ${t.judgment_label}</span>` : ''}
                </div>
                <div class="ind-value">${t.latest !== null && t.latest !== undefined ? formatNumber(t.latest) : '-'} <span style="font-size:0.5em;color:var(--text-muted)">${ind.unit}</span></div>
                <div class="ind-meta">Kỳ: ${t.latest_period || '—'} · Nguồn: ${SOURCE_LABELS[ind.autoSource] || ind.autoSource}</div>
                ${hasChart ? `<div class="ind-chart"><canvas id="${canvasId}"></canvas></div>` : ''}
                ${ind.impact ? `<div class="ind-note">${ind.impact}</div>` : ''}
                ${ind.note ? `<div class="ind-source-note">${ind.note}</div>` : ''}
            `;
            grid.appendChild(card);

            if (hasChart) {
                const valid = ind.series.filter(p => p.value !== null && p.value !== undefined);
                const ctx = card.querySelector(`#${canvasId}`);
                const improving = ind.goodDirection === 'higher'
                    ? valid[valid.length - 1].value >= valid[0].value
                    : valid[valid.length - 1].value <= valid[0].value;
                const color = improving ? '#10b981' : '#ef4444';
                const chart = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: valid.map(p => p.period),
                        datasets: [{
                            data: valid.map(p => p.value), borderColor: color,
                            backgroundColor: color + '15', fill: true, tension: 0.25, pointRadius: 2,
                        }],
                    },
                    options: CHART_DEFAULTS,
                });
                chartInstances.push(chart);
            }
        });

        if (grp === 'monetary') {
            renderInterbankCurveChart(grid, indicators);
            renderBankComparisonChart(grid, indicators);
        }
    });
}

// Đường cong lãi suất liên ngân hàng VNIBOR theo 7 kỳ hạn — khớp INTERBANK_TENOR_KEYS trong
// template_vimo.py (build_interbank_curve_chart()). Kiểu trình bày lấy cảm hứng từ chart VNIBOR
// đa kỳ hạn của vimo.cuthongthai.vn nhưng dùng dữ liệu tự cào từ sbv.gov.vn.
const INTERBANK_TENOR_KEYS = [
    ['interbank_rate_on', 'O/N'], ['interbank_rate_1w', '1 Tuần'], ['interbank_rate_2w', '2 Tuần'],
    ['interbank_rate_1m', '1 Tháng'], ['interbank_rate_3m', '3 Tháng'],
    ['interbank_rate_6m', '6 Tháng'], ['interbank_rate_9m', '9 Tháng'],
];

function renderInterbankCurveChart(grid, indicators) {
    const labels = [];
    const values = [];
    INTERBANK_TENOR_KEYS.forEach(([key, tenorLabel]) => {
        const series = (indicators[key] || {}).series || [];
        if (series.length) {
            labels.push(tenorLabel);
            values.push(series[series.length - 1].value);
        }
    });
    if (values.length < 2) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Đường cong lãi suất liên ngân hàng VNIBOR theo kỳ hạn</span></div>
        <div class="ind-chart" style="height:220px"><canvas id="chart-interbank-curve"></canvas></div>
        <div class="ind-note">Nguồn: sbv.gov.vn (bảng lãi suất BQ liên ngân hàng, tự động cập nhật).</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-interbank-curve');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [{
                data: values, borderColor: '#8b5cf6', backgroundColor: '#8b5cf615',
                fill: true, tension: 0.25, pointRadius: 3,
            }],
        },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: false } } },
    });
    chartInstances.push(chart);
}

// So sánh lãi suất huy động 12 tháng theo ngân hàng đại diện — CHART RIÊNG (cột so sánh nhiều
// ngân hàng cùng thời điểm), khác với chart đường thời gian của renderIndicatorGroups() ở trên.
// Khớp BANK_DEPOSIT_RATE_KEYS trong template_vimo.py (build_bank_comparison_chart()).
const BANK_DEPOSIT_RATE_KEYS = [
    ['deposit_rate_12m_vcb', 'VCB (lớn)'],
    ['deposit_rate_12m_ctg', 'VietinBank (lớn)'],
    ['deposit_rate_12m_nab', 'NamABank (nhỏ)'],
];

function renderBankComparisonChart(grid, indicators) {
    const labels = [];
    const values = [];
    BANK_DEPOSIT_RATE_KEYS.forEach(([key, bankLabel]) => {
        const series = (indicators[key] || {}).series || [];
        if (series.length) {
            labels.push(bankLabel);
            values.push(series[series.length - 1].value);
        }
    });
    if (!values.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📊 So sánh lãi suất huy động 12 tháng theo ngân hàng</span></div>
        <div class="ind-chart" style="height:220px"><canvas id="chart-bank-comparison"></canvas></div>
        <div class="ind-note">Đại diện nhóm lớn: VCB, VietinBank. Nhóm nhỏ: NamABank. Chưa có đại diện nhóm vừa (Techcombank/MBBank không có nguồn scrape ổn định — xem ghi chú từng chỉ báo).</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-bank-comparison');
    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: labels.map(l => l.includes('nhỏ') ? '#f59e0b' : '#3b82f6'),
            }],
        },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: false } } },
    });
    chartInstances.push(chart);
}
