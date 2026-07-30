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
    cafef_ajax: 'cafef.vn (khối ngoại HOSE, tự động)',
    vira: 'vira.org.vn (bản tin Kinh tế - Tài chính ngày, tự động)',
    derived: 'Tính từ chuỗi lũy kế đã có (phái sinh, không phải nguồn ngoài)',
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
    renderVnindexCompare(data.marketValuation, data.marketValuationHeadline, data.decision, data.decisionHeadline);
    renderIndicatorGroups(data.indicators);

    // File RIÊNG (không gộp vào vimo.json) — lịch sử P/E/P/B theo NGÀY ~17 năm (~4300 điểm/chỉ
    // số) từ Vietcap IQ, xem fetch_vietcap_index_valuation() trong fetch_macro_data.py. User
    // (2026-07-25) yêu cầu đưa lên web, đặt ngay dưới Scorecard, dạng ngang/rộng nhất có thể,
    // có nút xem toàn màn hình + khung thời gian lọc.
    const valHist = await fetch('data/vnindex_valuation_history.json').then(r => r.ok ? r.json() : null).catch(() => null);
    if (valHist) renderVnindexValuationHistory(valHist, data.marketValuation);
});

function renderSynthesis(synthesis) {
    if (!synthesis) return;
    const set = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text || '-'; };
    set('synthesis-overview', synthesis.overview);
    set('synthesis-market', synthesis.market_impact);
    set('synthesis-watch', synthesis.watch_points);

    // economy_impact giờ là list [{heading, text}] (không còn 1 chuỗi text duy nhất) — mỗi phần
    // render thành 1 khối có tiêu đề riêng rõ ràng, để biết ngay đoạn đang nói chủ đề gì.
    const econEl = document.getElementById('synthesis-economy');
    if (econEl) {
        const sections = synthesis.economy_impact;
        econEl.innerHTML = Array.isArray(sections)
            ? sections.map(s => `<div class="impact-block"><h5>${s.heading}</h5><p>${s.text}</p></div>`).join('')
            : (sections || '-');
    }
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
        // NÂNG CẤP 2026-07-25: khớp nhãn mới của calc_decision_matrix() (định giá dẫn dắt mức độ
        // giải ngân — "Mua mạnh"/"Mua tỷ trọng cao"/"Tăng tỷ trọng vừa phải"/"Nên mua vào"/"Duy
        // trì, chọn lọc"/"Giải ngân một phần" đều là mua/giữ, chỉ khác mức độ. "Mua tỷ trọng cao"
        // thiếu trong danh sách gốc (bug) — bổ sung cùng lúc thêm nhãn mới "Tăng tỷ trọng vừa
        // phải"/"Nên bán ra"/"Clear toàn bộ" (2 cái sau là bán, rơi vào nhánh đỏ mặc định).
        const BUY_HOLD_LABELS = ['Mua mạnh', 'Mua tỷ trọng cao', 'Tăng tỷ trọng vừa phải',
            'Nên mua vào', 'Duy trì, chọn lọc', 'Giải ngân một phần'];
        const decisionColor = BUY_HOLD_LABELS.includes(decision.label) ? '#10b981' : '#ef4444';
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

    const nGroups = Object.keys(scorecard.groups).length;
    let html = Object.entries(scorecard.groups).map(([gname, g]) => `
        <div class="vimo-score-box">
            <span class="lbl">${gname}</span>
            <span class="val" style="color:${scoreColor(g.score)}">${scoreText(g.score)}</span>
            <div style="font-size:0.7em;color:var(--text-muted);margin-top:4px">${g.nVotes} phiếu bầu</div>
            ${g.reason ? `<div style="font-size:0.68em;color:var(--text-muted);margin-top:4px;line-height:1.3">${g.reason}</div>` : ''}
        </div>
    `).join('');

    const totalColor = scoreColor(scorecard.total);
    html += `
        <div class="vimo-score-box" style="border:2px solid ${totalColor}">
            <span class="lbl">TỔNG SCORECARD</span>
            <span class="val" style="color:${totalColor}">${scorecard.total > 0 ? '+' : ''}${scorecard.total} / ${nGroups}</span>
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

    // Bù đắp rủi ro cổ phiếu so với gửi tiết kiệm/TPCP (user 2026-07-13: P/E-P/B suông không đủ,
    // cần biết có bù được rủi ro đầu tư cổ phiếu so với kênh an toàn hơn hay không).
    const rc = val.risk_compensation;
    const banner = document.getElementById('risk-comp-banner');
    if (banner && rc) {
        banner.style.display = 'block';
        const color = rc.color === 'good' ? '#10b981' : (rc.color === 'bad' ? '#ef4444' : '#f59e0b');
        banner.style.borderLeftColor = color;
        banner.style.background = color + '15';
        document.getElementById('risk-comp-label').textContent = `⚖️ Bù đắp rủi ro cổ phiếu: ${rc.label}`;
        document.getElementById('risk-comp-label').style.color = color;
        document.getElementById('risk-comp-text').textContent = rc.text;
    } else if (banner) {
        banner.style.display = 'none';
    }
}

// ═══════════════════════════════════════════════════════════
// SO SÁNH 2 QUYẾT ĐỊNH — headline (có VIN) vs ex-VIN (user 2026-07-25: "chia ra 2 quyết định:
// nếu nhìn vào VN-Index thì quyết định là gì... nếu nhìn theo VN-Index no VIN thì quyết định là gì").
// ═══════════════════════════════════════════════════════════
function renderVnindexCompare(valExvin, valHeadline, decisionExvin, decisionHeadline) {
    const card = document.getElementById('vnindex-compare-card');
    if (!valHeadline || !decisionHeadline) { card.style.display = 'none'; return; }
    card.style.display = '';

    const fmtX = (v) => v !== null && v !== undefined ? `${formatNumber(v)}x` : '-';
    const rows = [
        { label: 'VN-Index (headline, có VIN)', pe: valHeadline.pe, pb: valHeadline.pb, valLabel: valHeadline.valuation_label, decLabel: decisionHeadline.label },
        { label: 'VN-Index ex-VIN (loại VIC/VHM/VRE/VPL)', pe: valExvin.pe, pb: valExvin.pb, valLabel: valExvin.valuation_label, decLabel: decisionExvin.label },
    ];
    const valColor = (l) => l === 'Rẻ/Hấp dẫn' ? '#10b981' : l === 'Đắt/Kém hấp dẫn' ? '#ef4444' : '#f59e0b';

    document.getElementById('vnindex-compare-table').innerHTML = `
        <table style="width:100%;border-collapse:collapse;font-size:0.88em">
            <thead><tr style="border-bottom:1px solid var(--border-color,#1f2937)">
                <th style="text-align:left;padding:6px 8px">Góc nhìn</th>
                <th style="padding:6px 8px">P/E</th>
                <th style="padding:6px 8px">P/B</th>
                <th style="padding:6px 8px">Đánh giá định giá</th>
                <th style="padding:6px 8px">Khuyến nghị</th>
            </tr></thead>
            <tbody>
                ${rows.map(r => `
                    <tr>
                        <td style="padding:6px 8px">${r.label}</td>
                        <td style="text-align:center;padding:6px 8px">${fmtX(r.pe)}</td>
                        <td style="text-align:center;padding:6px 8px">${fmtX(r.pb)}</td>
                        <td style="text-align:center;padding:6px 8px;color:${valColor(r.valLabel)};font-weight:700">${r.valLabel || '-'}</td>
                        <td style="text-align:center;padding:6px 8px;font-weight:700">${r.decLabel || '-'}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    const warnEl = document.getElementById('vnindex-compare-warning');
    if (decisionExvin.label !== decisionHeadline.label) {
        warnEl.style.display = '';
        warnEl.textContent = `⚠ 2 góc nhìn cho khuyến nghị KHÁC NHAU — VIN (VIC/VHM/VRE/VPL) đang làm lệch kết luận định giá chung của thị trường một cách đáng kể.`;
    } else {
        warnEl.style.display = 'none';
    }
}

// ═══════════════════════════════════════════════════════════
// ĐỊNH GIÁ VN-INDEX THEO THỜI GIAN — P/E & P/B lịch sử ~17 năm (Vietcap IQ headline + GitHub
// ex-VIN) vs dải thống kê + ngưỡng hấp dẫn — 4 biểu đồ RIÊNG (headline/ex-VIN x P/E/P/B), mỗi
// biểu đồ có nút xem toàn màn hình + khung thời gian lọc, tooltip hiện giá trị theo ngày khi rê
// chuột (user 2026-07-25).
// ═══════════════════════════════════════════════════════════
const VALHIST_RANGES = [
    { key: '1Y', label: '1 năm', days: 365 },
    { key: '3Y', label: '3 năm', days: 365 * 3 },
    { key: '5Y', label: '5 năm', days: 365 * 5 },
    { key: '10Y', label: '10 năm', days: 365 * 10 },
    { key: 'ALL', label: 'Toàn bộ', days: null },
];
const VALHIST_MAX_POINTS = 1000; // giảm mẫu (decimate) khi khung thời gian dài để chart mượt

// 4 biểu đồ RIÊNG (user 2026-07-25): mỗi cái 1 canvas id + key dữ liệu riêng trong
// data/vnindex_valuation_history.json (pe/pe_exvin/pb/pb_exvin — xem
// update_vnindex_valuation_history() trong fetch_macro_data.py, ex-VIN có dải ±SD tự tính bằng
// statistics.mean/stdev vì GitHub ex-VIN không có sẵn như Vietcap).
const VALHIST_CHARTS_SPEC = [
    { dataKey: 'pe', canvasId: 'chart-vnindex-pe-history', unit: 'P/E', kind: 'pe' },
    { dataKey: 'pe_exvin', canvasId: 'chart-vnindex-pe-exvin-history', unit: 'P/E', kind: 'pe' },
    { dataKey: 'pb', canvasId: 'chart-vnindex-pb-history', unit: 'P/B', kind: 'pb' },
    { dataKey: 'pb_exvin', canvasId: 'chart-vnindex-pb-exvin-history', unit: 'P/B', kind: 'pb' },
];
let valHistCharts = {};
let valHistData = null; // {pe, pe_exvin, pb, pb_exvin} gốc, giữ lại để đổi khung thời gian không cần fetch lại

function decimate(arr, maxPoints) {
    if (arr.length <= maxPoints) return arr;
    const step = Math.ceil(arr.length / maxPoints);
    const out = [];
    for (let i = 0; i < arr.length; i += step) out.push(arr[i]);
    if (out[out.length - 1] !== arr[arr.length - 1]) out.push(arr[arr.length - 1]); // luôn giữ điểm mới nhất
    return out;
}

function renderVnindexValuationHistory(hist, marketValuation) {
    if (!hist || VALHIST_CHARTS_SPEC.every(s => !hist[s.dataKey])) return;
    valHistData = hist;
    document.getElementById('vnindex-valhist-card').style.display = '';

    // Nút khung thời gian
    const btnWrap = document.getElementById('vnindex-valhist-range-btns');
    btnWrap.innerHTML = VALHIST_RANGES.map((r, i) =>
        `<button class="vimo-range-btn${i === VALHIST_RANGES.length - 1 ? ' active' : ''}" data-range="${r.key}">${r.label}</button>`
    ).join('');
    btnWrap.querySelectorAll('.vimo-range-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            btnWrap.querySelectorAll('.vimo-range-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            drawValHistCharts(btn.dataset.range, marketValuation);
        });
    });

    // Nút toàn màn hình — Fullscreen API trên chính khung chứa chart. Icon/nhãn ĐỔI RÕ RÀNG giữa
    // "⛶ Mở rộng" và "✕ Đóng" theo trạng thái (user 2026-07-25: "có nút để close biểu đồ" — dùng
    // lại đúng 1 nút thay vì thêm nút riêng, nhưng phải rõ ràng là nút ĐÓNG khi đang toàn màn hình).
    document.querySelectorAll('#vnindex-valhist-card .vimo-fullscreen-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const el = document.getElementById(btn.dataset.target);
            if (!document.fullscreenElement) el.requestFullscreen?.();
            else document.exitFullscreen?.();
        });
    });
    document.addEventListener('fullscreenchange', () => {
        document.querySelectorAll('#vnindex-valhist-card .vimo-fullscreen-btn').forEach(btn => {
            const isFs = document.fullscreenElement && document.fullscreenElement.id === btn.dataset.target;
            btn.textContent = isFs ? '✕' : '⛶';
            btn.title = isFs ? 'Đóng toàn màn hình' : 'Xem toàn màn hình';
        });
        setTimeout(() => Object.values(valHistCharts).forEach(c => c?.resize()), 50);
    });

    drawValHistCharts('ALL', marketValuation);
}

function filterByRange(values, days) {
    if (!days) return values;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    const cutoffStr = cutoff.toISOString().slice(0, 10);
    return values.filter(p => p.date >= cutoffStr);
}

function drawValHistCharts(rangeKey, marketValuation) {
    const range = VALHIST_RANGES.find(r => r.key === rangeKey) || VALHIST_RANGES[VALHIST_RANGES.length - 1];
    const rf = marketValuation && marketValuation.rf;
    const capm = marketValuation && marketValuation.capm_valuation;

    const peExtra = [];
    if (rf) {
        peExtra.push({ label: `P/E hoà vốn (2×Rf ${(rf * 100).toFixed(2)}%)`, value: 1 / (2 * rf), color: '#8b5cf6' });
        peExtra.push({ label: `P/E trung tính (1.5×Rf ${(rf * 100).toFixed(2)}%)`, value: 1 / (1.5 * rf), color: '#f97316' });
    }
    const pbExtra = [];
    if (capm && capm.justified_pb) {
        pbExtra.push({ label: 'P/B hợp lý (CAPM)', value: capm.justified_pb, color: '#8b5cf6' });
    }

    VALHIST_CHARTS_SPEC.forEach(spec => {
        const data = valHistData[spec.dataKey];
        if (!data || !data.values) return;
        const filtered = decimate(filterByRange(data.values, range.days), VALHIST_MAX_POINTS);
        const extraLines = spec.kind === 'pe' ? peExtra : pbExtra;
        valHistCharts[spec.dataKey] = drawOneValHistChart(
            spec.canvasId, valHistCharts[spec.dataKey], filtered, data, spec.unit, '#3b82f6', extraLines);
    });
}

function drawOneValHistChart(canvasId, existingChart, points, bandData, unitLabel, lineColor, extraLines) {
    if (existingChart) existingChart.destroy();
    const labels = points.map(p => p.date);
    const bandSpecs = [
        ['average', 'Trung bình', '#f59e0b', [6, 3]],
        ['plusOneSD', '+1SD', '#ef4444', [2, 2]],
        ['minusOneSD', '-1SD', '#10b981', [2, 2]],
        ['plusTwoSD', '+2SD', '#ef4444', [1, 3]],
        ['minusTwoSD', '-2SD', '#10b981', [1, 3]],
    ];
    const datasets = [{
        label: `${unitLabel} VN-Index`, data: points.map(p => p.value),
        borderColor: lineColor, backgroundColor: lineColor + '10', fill: false,
        tension: 0, pointRadius: 0, borderWidth: 1.4,
    }];
    bandSpecs.forEach(([key, label, color, dash]) => {
        if (bandData[key] === undefined || bandData[key] === null) return;
        datasets.push({
            label: `${label} (${bandData[key].toFixed(2)})`, data: labels.map(() => bandData[key]),
            borderColor: color, borderDash: dash, borderWidth: 1, pointRadius: 0, fill: false,
        });
    });
    (extraLines || []).forEach(l => {
        datasets.push({
            label: `${l.label} (${l.value.toFixed(2)})`, data: labels.map(() => l.value),
            borderColor: l.color, borderWidth: 2, pointRadius: 0, fill: false,
        });
    });

    const ctx = document.getElementById(canvasId);
    return new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
            responsive: true, maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 9 }, color: '#8892a4' } },
                tooltip: {
                    callbacks: {
                        title: (items) => items[0] ? `Ngày: ${items[0].label}` : '',
                    },
                },
            },
            scales: {
                x: { ticks: { color: '#545f74', font: { size: 8 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 }, grid: { display: false } },
                y: { ticks: { color: '#545f74', font: { size: 9 } }, grid: { color: 'rgba(255,255,255,0.04)' } },
            },
        },
    });
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
            renderInterbank6mHistoryChart(grid, indicators);
            renderBondYieldHistoryChart(grid, indicators);
        }
        if (grp === 'growth') {
            renderStackedAreaChart(grid, indicators, {
                title: '🗺️ Cơ cấu GDP theo khu vực kinh tế (%)',
                keys: [
                    ['gdp_share_agri', 'Nông-Lâm-Thủy sản', '#10b981'],
                    ['gdp_share_industry', 'Công nghiệp-Xây dựng', '#3b82f6'],
                    ['gdp_share_services', 'Dịch vụ', '#f59e0b'],
                    ['gdp_share_tax', 'Thuế sản phẩm (ròng)', '#a78bfa'],
                ],
                canvasId: 'chart-gdp-structure',
                note: 'Nguồn: nso.gov.vn (Thông cáo báo chí KT-XH quý, tự động). Số liệu LŨY KẾ theo kỳ báo cáo (Q1/6 tháng/9 tháng/cả năm), không phải chuỗi quý độc lập.',
            });
        }
        if (grp === 'trade') {
            renderStackedAreaChart(grid, indicators, {
                title: '🗺️ Cơ cấu vốn đầu tư thực hiện toàn xã hội theo thành phần (%)',
                keys: [
                    ['investment_share_state', 'Nhà nước', '#3b82f6'],
                    ['investment_share_private', 'Ngoài Nhà nước (tư nhân)', '#10b981'],
                    ['investment_share_fdi', 'FDI', '#f59e0b'],
                ],
                canvasId: 'chart-investment-structure',
                note: 'Nguồn: nso.gov.vn (Thông cáo báo chí KT-XH quý, tự động). Số liệu LŨY KẾ theo kỳ báo cáo (Q1/6 tháng/9 tháng/cả năm), không phải chuỗi quý độc lập.',
            });
        }
    });
}

// So sánh 2 chuỗi period THEO THỜI GIAN THẬT — CẦN THIẾT cho các kỳ lũy kế kiểu Q1/H1/9M/FY (sort
// chuỗi mặc định cho ra "H1" < "Q1" < "FY" < "9M" theo abc, SAI thứ tự thời gian thật trong năm —
// phát hiện khi cơ cấu GDP có ≥2 điểm/năm, xem _period_sort_key() tương đương bên template_vimo.py).
const _PERIOD_SUB_RANK = { Q1: 1, H1: 2, "9M": 3, FY: 4 };
function _periodSortKey(period) {
    const m = /^(\d{4})-(Q1|H1|9M|FY)$/.exec(period);
    if (m) return [parseInt(m[1], 10), _PERIOD_SUB_RANK[m[2]]];
    return [period, 0]; // định dạng khác (ngày/tuần/tháng) — chuỗi ISO đã tự sort đúng theo abc
}
function _sortPeriods(periods) {
    return [...periods].sort((a, b) => {
        const ka = _periodSortKey(a), kb = _periodSortKey(b);
        if (ka[0] !== kb[0]) return ka[0] < kb[0] ? -1 : 1;
        return ka[1] - kb[1];
    });
}

// Biểu đồ miền (stacked area) DÙNG CHUNG cho cơ cấu GDP theo khu vực VÀ cơ cấu vốn đầu tư theo
// thành phần (user 2026-07-13) — mỗi kỳ báo cáo là 1 điểm trên trục X, các thành phần % cộng lại
// ~100%. Vẽ được ngay cả khi mới có 1 điểm (sẽ dài dần mỗi lần Action chạy, giống các chart khác).
function renderStackedAreaChart(grid, indicators, { title, keys, canvasId, note }) {
    const seriesByKey = keys.map(([key, label, color]) => [
        label, color, ((indicators[key] || {}).series || []).filter(p => p.value !== null && p.value !== undefined),
    ]);
    const allPeriods = _sortPeriods(new Set(seriesByKey.flatMap(([, , s]) => s.map(p => p.period))));
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">${title}</span></div>
        <div class="ind-chart" style="height:260px"><canvas id="${canvasId}"></canvas></div>
        <div class="ind-note">${note}</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector(`#${canvasId}`);
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: seriesByKey.map(([label, color, s]) => {
                const byPeriod = Object.fromEntries(s.map(p => [p.period, p.value]));
                return {
                    label, data: allPeriods.map(p => byPeriod[p] ?? null),
                    borderColor: color, backgroundColor: color + '55', fill: true,
                    tension: 0.15, pointRadius: 2, spanGaps: true,
                };
            }),
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 12 } } },
            scales: { ...CHART_DEFAULTS.scales, y: { ...CHART_DEFAULTS.scales.y, stacked: true, min: 0, max: 100 } },
        },
    });
    chartInstances.push(chart);
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

// Lãi suất liên ngân hàng — O/N, 1 tuần, 2 tuần, 1 tháng, 6 tháng THEO THỜI GIAN, cùng 1 chart
// đường nhiều dòng (không giới hạn số điểm tối thiểu như renderIndicatorGroups() ở trên) để càng
// nhiều Action chạy càng tích lũy được chuỗi dài, thay cho chart so sánh ngân hàng cũ (đã gỡ bỏ
// theo yêu cầu user). Theo yêu cầu user (2026-07-13): thêm O/N và 1 tháng vào chung biểu đồ 6
// tháng để so sánh nhiều kỳ hạn trên cùng 1 trục thời gian (giống kiểu trình bày tham khảo từ
// vimo.cuthongthai.vn). Thêm 1W/2W (2026-07-28, nguồn VIRA — xem fetch_vira_bulletin() trong
// fetch_macro_data.py): ON/1W/2W/1M giờ có chuỗi NGÀY thật (không còn snapshot theo tuần/tháng
// của SBV) nên đủ điểm để thấy xu hướng ngay; 6M vẫn thưa (nguồn SBV, tích lũy theo tuần).
const INTERBANK_HISTORY_TENORS = [
    ['interbank_rate_on', 'O/N', '#f59e0b'],
    ['interbank_rate_1w', '1 Tuần', '#ef4444'],
    ['interbank_rate_2w', '2 Tuần', '#10b981'],
    ['interbank_rate_1m', '1 Tháng', '#a78bfa'],
    ['interbank_rate_6m', '6 Tháng', '#3b82f6'],
];

// Khớp CHÍNH XÁC 2 định dạng period hợp lệ cho chart này: tuần "YYYY-Www" (SBV, dùng cho 6M/9M/3M)
// và ngày "YYYY-MM-DD" (VIRA, dùng cho ON/1W/2W/1M từ 2026-07-28) — KHÔNG khớp định dạng tháng cũ
// "YYYY-MM" (7 ký tự, không có cụm ngày thứ 2) vốn là điểm lũy kế/snapshot cũ còn sót lại trước khi
// đổi sang tuần, trộn chung sẽ khiến trục thời gian bị kéo phẳng sai (user 2026-07-25).
const INTERBANK_HISTORY_PERIOD_RE = /^\d{4}-(W\d{2}|\d{2}-\d{2})$/;

function renderInterbank6mHistoryChart(grid, indicators) {
    // Hợp nhất TOÀN BỘ period của cả 5 kỳ hạn thành 1 trục thời gian chung — hợp nhất (thay vì chỉ
    // lấy period của 1 kỳ hạn) để không mất điểm nếu có kỳ hạn nào lệch lịch sử. Khớp với
    // build_interbank_6m_history_chart() trong template_vimo.py (PDF).
    const seriesByTenor = INTERBANK_HISTORY_TENORS.map(([key, tenorLabel, color]) => [
        tenorLabel, color,
        ((indicators[key] || {}).series || []).filter(p => p.value !== null && p.value !== undefined && INTERBANK_HISTORY_PERIOD_RE.test(p.period)),
    ]);
    const allPeriods = [...new Set(seriesByTenor.flatMap(([, , s]) => s.map(p => p.period)))].sort();
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Lãi suất liên ngân hàng O/N, 1 tuần, 2 tuần, 1 tháng, 6 tháng theo thời gian</span></div>
        <div class="ind-chart" style="height:220px"><canvas id="chart-interbank-6m-history"></canvas></div>
        <div class="ind-note">Nguồn: O/N, 1 tuần, 2 tuần, 1 tháng — vira.org.vn (bản tin ngày, tự động, chuỗi theo NGÀY thật). 6 tháng — sbv.gov.vn (bảng lãi suất BQ liên ngân hàng, tự động, tích lũy theo tuần).</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-interbank-6m-history');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: seriesByTenor.map(([tenorLabel, color, s]) => {
                const byPeriod = Object.fromEntries(s.map(p => [p.period, p.value]));
                return {
                    label: tenorLabel, data: allPeriods.map(p => byPeriod[p] ?? null),
                    borderColor: color, backgroundColor: color + '15', fill: false,
                    tension: 0.25, pointRadius: 3, spanGaps: true,
                };
            }),
        },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: true, labels: { boxWidth: 12 } } } },
    });
    chartInstances.push(chart);
}

// Lợi suất TPCP thứ cấp 3Y/5Y/7Y/10Y/15Y theo thời gian — chỉ báo MỚI (2026-07-28, nguồn VIRA,
// xem fetch_vira_bulletin() trong fetch_macro_data.py), cùng kiểu trình bày với chart lãi suất
// liên ngân hàng ở trên để so sánh 2 đường cong chi phí vốn (liên ngân hàng vs TPCP Chính phủ).
const BOND_YIELD_TENORS = [
    ['govt_bond_yield_3y', '3 Năm', '#f59e0b'],
    ['govt_bond_yield_5y', '5 Năm', '#ef4444'],
    ['govt_bond_yield_7y', '7 Năm', '#10b981'],
    ['govt_bond_yield_10y', '10 Năm', '#a78bfa'],
    ['govt_bond_yield_15y', '15 Năm', '#3b82f6'],
];

function renderBondYieldHistoryChart(grid, indicators) {
    const seriesByTenor = BOND_YIELD_TENORS.map(([key, tenorLabel, color]) => [
        tenorLabel, color,
        ((indicators[key] || {}).series || []).filter(p => p.value !== null && p.value !== undefined),
    ]);
    const allPeriods = [...new Set(seriesByTenor.flatMap(([, , s]) => s.map(p => p.period)))].sort();
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Lợi suất TPCP thứ cấp 3-5-7-10-15 năm theo thời gian</span></div>
        <div class="ind-chart" style="height:220px"><canvas id="chart-bond-yield-history"></canvas></div>
        <div class="ind-note">Nguồn: vira.org.vn (bản tin Kinh tế - Tài chính ngày, tự động, chuỗi theo NGÀY thật). Lợi suất giao dịch thứ cấp, không phải lãi suất trúng thầu sơ cấp KBNN.</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-bond-yield-history');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: seriesByTenor.map(([tenorLabel, color, s]) => {
                const byPeriod = Object.fromEntries(s.map(p => [p.period, p.value]));
                return {
                    label: tenorLabel, data: allPeriods.map(p => byPeriod[p] ?? null),
                    borderColor: color, backgroundColor: color + '15', fill: false,
                    tension: 0.25, pointRadius: 3, spanGaps: true,
                };
            }),
        },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: true, labels: { boxWidth: 12 } } } },
    });
    chartInstances.push(chart);
}
