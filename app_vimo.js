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
    renderScorecard(data.scorecard);
    renderDecision(data.decision, data.scorecard.total);
    renderValuation(data.marketValuation);
    renderIndicatorGroups(data.indicators);
});

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
    document.getElementById('val-pb').textContent = val.pb ? `${formatNumber(val.pb)}x` : '-';
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
                ${ind.note ? `<div class="ind-note">${ind.note}</div>` : ''}
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
    });
}
