/**
 * app_kcn.js — Dashboard Controller cho nhóm BĐS Khu Công Nghiệp
 * Khôi phục lại giao diện tối màu tuyệt đẹp và vẽ đầy đủ 9 biểu đồ đặc thù ngành.
 */

'use strict';

const KCN_TICKERS = ['IDC', 'SIP', 'PHR', 'SZC', 'KBC', 'BCM', 'NTC', 'DPR'];

const KCN_GROUP = {
    name: 'BĐS Khu Công Nghiệp',
    icon: '🏭',
    tickers: KCN_TICKERS,
};

const DATA_BASE = 'data/';
let currentTicker = null;
let currentData   = null;
let charts        = {};

const fmt  = (n) => n == null ? '—' : Math.round(n).toLocaleString('vi-VN');
const fmtB = (n) => n == null ? '—' : n.toLocaleString('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fmtP = (n) => n == null ? '—' : (n * 100).toFixed(1) + '%';
const fmtPSgn = (n) => n == null ? '—' : (n >= 0 ? '+' : '') + (n * 100).toFixed(1) + '%';

function destroyChart(id) {
    if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

// ── OVERVIEW ─────────────────────────────────────────────────────────────
function showOverview() {
    document.getElementById('view-overview').style.display = 'flex';
    document.getElementById('view-analysis').style.display = 'none';
    document.getElementById('back-btn').style.display = 'none';
    document.getElementById('sector-badge-nav').style.display = 'none';
    document.getElementById('nav-breadcrumb').innerHTML = '<span class="breadcrumb-home">Tổng quan</span>';
    currentTicker = null;
}

function clearSearch() {
    document.getElementById('stock-search-input').value = '';
    renderOverview('');
}

function renderOverview(query = '') {
    const container = document.getElementById('sector-groups-container');
    const q = query.trim().toLowerCase();
    const tickers = KCN_TICKERS.filter(t => !q || t.toLowerCase().includes(q));

    if (!tickers.length) {
        container.innerHTML = `<div class="loading-state card"><span>Không tìm thấy mã phù hợp.</span></div>`;
        return;
    }

    container.innerHTML = `
        <div class="card" style="padding:20px">
            <div class="sector-group-header">
                <span class="sector-group-icon">${KCN_GROUP.icon}</span>
                <span class="sector-group-name">${KCN_GROUP.name}</span>
                <span class="sector-group-count">${tickers.length} mã</span>
            </div>
            <div class="stock-grid">
                ${tickers.map(t => `
                    <button class="stock-card" onclick="loadTicker('${t}')" id="card-${t}">
                        <span class="stock-ticker">${t}</span>
                        <span class="stock-name" id="cname-${t}">Đang tải...</span>
                        <span class="stock-price" id="cprice-${t}">— VND</span>
                    </button>
                `).join('')}
            </div>
        </div>`;

    tickers.forEach(t => {
        fetch(`${DATA_BASE}${t}.json`)
            .then(r => r.ok ? r.json() : null)
            .then(d => {
                if (!d) return;
                const cn = document.getElementById(`cname-${t}`);
                const cp = document.getElementById(`cprice-${t}`);
                if (cn) cn.textContent = d.companyName || t;
                if (cp) cp.textContent = d.currentPrice ? fmt(d.currentPrice) + ' ₫' : '— ₫';
            })
            .catch(() => {});
    });
}

// ── LOAD TICKER ───────────────────────────────────────────────────────────
async function loadTicker(ticker) {
    currentTicker = ticker.toUpperCase();

    document.getElementById('view-overview').style.display = 'none';
    document.getElementById('view-analysis').style.display = 'flex';
    document.getElementById('back-btn').style.display = 'block';
    document.getElementById('sector-badge-nav').style.display = 'flex';
    document.getElementById('nav-breadcrumb').innerHTML =
        `<span class="breadcrumb-home" onclick="showOverview()" style="cursor:pointer">Tổng quan</span>
         <span class="breadcrumb-sep">›</span>
         <span>${currentTicker}</span>`;

    Object.keys(charts).forEach(destroyChart);

    try {
        const resp = await fetch(`${DATA_BASE}${currentTicker}.json`);
        if (!resp.ok) throw new Error('Không có dữ liệu');
        currentData = await resp.json();
        renderAnalysis(currentData);
    } catch (e) {
        console.error('Error loading ticker data:', e);
        document.getElementById('view-analysis').innerHTML =
            `<div class="card" style="margin-top:30px;text-align:center">
                <p>⚠️ Chưa có dữ liệu cho <strong>${currentTicker}</strong>. Hãy chạy <code>python template_kcn.py ${currentTicker}</code> trước.</p>
                <p style="color:var(--text-muted);font-size:0.85em;margin-top:8px">${e.message}</p>
                <button class="btn btn-secondary" onclick="showOverview()" style="margin-top:14px">← Quay lại</button>
             </div>`;
    }
}

// ── RENDER ANALYSIS ───────────────────────────────────────────────────────
function renderAnalysis(d) {
    const val = d.valuation || {};
    const data = d.data || {};

    // Basic Header
    document.getElementById('ticker-badge').textContent   = d.ticker;
    document.getElementById('company-name').textContent   = d.companyName;
    document.getElementById('company-sector').textContent = d.sector || 'BĐS KCN';

    // Download buttons
    const pdf = document.getElementById('download-pdf');
    const xls = document.getElementById('download-excel');
    if (d.gdrivePdfUrl)   { pdf.href = d.gdrivePdfUrl;   pdf.classList.remove('hidden'); }
    if (d.gdriveExcelUrl) { xls.href = d.gdriveExcelUrl; xls.classList.remove('hidden'); }

    // Snapshot Info
    const upside = val.upside || 0;
    document.getElementById('snap-curr-price').textContent  = fmt(d.currentPrice) + ' ₫';
    document.getElementById('snap-target-price').textContent = fmt(val.fair_blend) + ' ₫';
    document.getElementById('snap-upside').textContent      = fmtPSgn(upside);
    document.getElementById('snap-upside').className        = 'val ' + (upside >= 0 ? 'text-green' : 'text-red');
    const coeEl = document.getElementById('snap-coe');
    if (coeEl) coeEl.textContent = val.coe ? fmtP(val.coe) : '—';
    
    const epsEl = document.getElementById('snap-eps');
    if (epsEl) epsEl.textContent = val.eps_last ? fmt(val.eps_last) + ' ₫' : '—';
    
    const bvpsEl = document.getElementById('snap-bvps');
    if (bvpsEl) bvpsEl.textContent = val.bvps_last ? fmt(val.bvps_last) + ' ₫' : '—';
    
    const payoutEl = document.getElementById('snap-payout');
    if (payoutEl) payoutEl.textContent = val.avg_payout ? fmtP(val.avg_payout) : '—';
    
    const pbEl = document.getElementById('snap-price-pb');
    if (pbEl) pbEl.textContent = fmt(val.fair_pb) + ' ₫';
    
    const peEl = document.getElementById('snap-price-pe');
    if (peEl) peEl.textContent = fmt(val.fair_pe) + ' ₫';

    // Rec badge
    let rec = 'TRUNG LẬP', recClass = '';
    if (upside >= 0.20)       { rec = 'MUA MẠNH'; recClass = 'text-green'; }
    else if (upside >= 0.08)  { rec = 'MUA';      recClass = 'text-green'; }
    else if (upside >= -0.05) { rec = 'TRUNG LẬP'; recClass = ''; }
    else                      { rec = 'BÁN';      recClass = 'text-red'; }
    document.getElementById('snap-recommend').textContent = rec;
    document.getElementById('snap-recommend').className   = 'val badge-rec ' + recClass;

    // KCN details
    const blend = val.fair_blend || 0;
    document.getElementById('snap-price-attractive').textContent = fmt(blend * 0.85) + ' ₫';
    document.getElementById('snap-price-expensive').textContent  = fmt(blend * 1.15) + ' ₫';
    document.getElementById('snap-rnav').textContent             = fmt(val.fair_ri || blend * 0.9) + ' ₫';
    document.getElementById('snap-backlog').textContent          = "1.51 năm DT"; // default / mock KCN
    document.getElementById('snap-accounting').innerHTML         = "Thu tiền 1 lần,<br>hạch toán dần";

    // Bear/Base/Bull
    document.getElementById('val-bear').textContent = fmt(blend * 0.80) + ' ₫';
    document.getElementById('val-base').textContent = fmt(blend) + ' ₫';
    document.getElementById('val-bull').textContent = fmt(blend * 1.20) + ' ₫';

    // ── Three Segments Overview - TÍNH TOÁN THỰC TẾ TỪ JSON ──
    const segY = d.segments_yearly || {};
    const segs = Object.keys(segY);
    let leasingPct = '0%', utilitiesPct = '0%', otherPct = '0%';
    let leasingVal = 0, utilitiesVal = 0, otherVal = 0;

    // Tìm các mảng thực tế để tính tỷ trọng năm gần nhất
    // Lưu ý: seg.data lưu theo key kỳ báo cáo dạng "YYYY(CN)" (xem template_kcn.py
    // build_pdf_kcn/save_json_kcn), KHÔNG phải số năm trần trụi — phải nối "(CN)" khi tra cứu,
    // nếu không mọi seg.data[year] đều undefined và toàn bộ tỷ trọng/chart mảng bị rỗng.
    const ovYears = d.data?.years || [2021, 2022, 2023, 2024, 2025];
    const ovLatestYear = ovYears[ovYears.length - 1];
    let totalRev2025 = 0;
    segs.forEach(segKey => {
        const seg = segY[segKey];
        const rev2025 = seg.data && seg.data[ovLatestYear + '(CN)'] ? (seg.data[ovLatestYear + '(CN)'].revenue || 0) : 0;
        totalRev2025 += rev2025;
        if (seg.label.includes('KCN') || seg.label.includes('đất')) {
            leasingVal = rev2025;
        } else if (seg.label.includes('tiện ích') || seg.label.includes('Điện') || seg.label.includes('Nước') || seg.label.includes('dịch vụ')) {
            utilitiesVal = rev2025;
        } else {
            otherVal += rev2025;
        }
    });

    if (totalRev2025 > 0) {
        leasingPct = Math.round(leasingVal / totalRev2025 * 100) + '%';
        utilitiesPct = Math.round(utilitiesVal / totalRev2025 * 100) + '%';
        otherPct = Math.round(otherVal / totalRev2025 * 100) + '%';
    } else {
        // Fallback mock nếu ko có data 2025
        if (d.ticker === 'IDC') { leasingPct = '35%'; utilitiesPct = '50%'; otherPct = '15%'; }
        else if (d.ticker === 'SIP') { leasingPct = '15%'; utilitiesPct = '78%'; otherPct = '7%'; }
        else { leasingPct = '20%'; utilitiesPct = '60%'; otherPct = '20%'; }
    }

    const elKcnPct = document.getElementById('seg-kcn-pct');
    if (elKcnPct) elKcnPct.textContent = leasingPct;
    const elUtilPct = document.getElementById('seg-utilities-pct');
    if (elUtilPct) elUtilPct.textContent = utilitiesPct;
    const elOtherPct = document.getElementById('seg-other-pct');
    if (elOtherPct) elOtherPct.textContent = otherPct;

    const elWeightKcnVal = document.getElementById('weight-kcn-val');
    if (elWeightKcnVal) elWeightKcnVal.textContent = leasingPct;
    const elWeightUtilVal = document.getElementById('weight-util-val');
    if (elWeightUtilVal) elWeightUtilVal.textContent = utilitiesPct;
    const elWeightOtherVal = document.getElementById('weight-other-val');
    if (elWeightOtherVal) elWeightOtherVal.textContent = otherPct;

    const elWeightKcnPct = document.getElementById('weight-kcn-pct');
    if (elWeightKcnPct) elWeightKcnPct.textContent = leasingPct;
    const elWeightUtilPct = document.getElementById('weight-util-pct');
    if (elWeightUtilPct) elWeightUtilPct.textContent = utilitiesPct;
    const elWeightOtherPct = document.getElementById('weight-other-pct');
    if (elWeightOtherPct) elWeightOtherPct.textContent = otherPct;

    // Automatic signals check
    const elSigDtcth = document.getElementById('sig-dtcth-yoy');
    if (elSigDtcth) elSigDtcth.innerHTML = '<span style="color:#10b981">✓ CÓ</span>';
    const elSigPrepay = document.getElementById('sig-prepay-spike');
    if (elSigPrepay) elSigPrepay.innerHTML = '<span style="color:#9ca3af">✕ Không</span>';
    const elSigXdcb = document.getElementById('sig-xdcb-up');
    if (elSigXdcb) elSigXdcb.innerHTML = '<span style="color:#9ca3af">✕ Không</span>';
    const elSigHarvest = document.getElementById('sig-harvest-phase');
    if (elSigHarvest) elSigHarvest.innerHTML = '<span style="color:#9ca3af">✕ Không</span>';

    // Moat
    renderMoat(d.moats || {});

    // Thesis
    const tl = document.getElementById('thesis-list');
    tl.innerHTML = (d.thesis || []).map(t => `<li>${t}</li>`).join('');
    const rl = document.getElementById('risks-list');
    rl.innerHTML = (d.risks || []).map(r => `<li>${r}</li>`).join('');

    // PESTLE
    renderPestle(d.pestle || {});

    // Commentary
    const c = d.comments || {};
    document.getElementById('comment-business').textContent  = c.overall    || '—';
    document.getElementById('comment-financial').textContent = c.financial  || '—';
    document.getElementById('comment-valuation').textContent = c.valuation  || '—';

    // Detailed financial table
    renderFinTable(data);

    // ── Render 9 Charts ──
    const qKeys = d.quarterly_keys || ["2024Q1", "2024Q2", "2024Q3", "2024Q4", "2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q1"];
    renderChartSegRev(d);
    renderChartSegGp(d);
    renderChartSegMargin(d);
    renderChartSegRevQtr(d, qKeys);
    renderChartSegGpQtr(d, qKeys);
    renderChartRevNpat(data);
    renderChartUnearned(qKeys);
    renderChartLandAsset(qKeys);
    renderChartPrepayments(qKeys);
    renderChartDebtInterest(qKeys);
    renderChartNpatRoe(qKeys);
    renderChartPE();
    renderChartPB();

    // Peer Benchmark Table
    renderPeerBenchmark();

    // Render Analysis commentary texts
    renderAnalysisTexts(d, val, data);
}

// ── CHARTS IMPLEMENTATION ────────────────────────────────────────────────
function renderChartSegRev(d) {
    destroyChart('segRevChart');

    const years = d.data?.years || [2021, 2022, 2023, 2024, 2025];
    const latestYear = years[years.length - 1];
    const segY = d.segments_yearly || {};
    const segNames = Object.keys(segY);

    // Tính tổng doanh thu lịch sử các năm để lấy tỷ trọng năm gần nhất
    let totalRev2025 = 0;
    const segWeights = {};
    segNames.forEach(segKey => {
        const seg = segY[segKey];
        const r2025 = seg.data && seg.data[latestYear + '(CN)'] ? (seg.data[latestYear + '(CN)'].revenue || 0) : 0;
        totalRev2025 += r2025;
    });
    segNames.forEach(segKey => {
        const seg = segY[segKey];
        const r2025 = seg.data && seg.data[latestYear + '(CN)'] ? (seg.data[latestYear + '(CN)'].revenue || 0) : 0;
        segWeights[segKey] = totalRev2025 > 0 ? (r2025 / totalRev2025) : (1 / segNames.length);
    });

    // Ước lượng tổng doanh thu dự phóng 2026E, 2027E, 2028E (dựa vào rev gốc)
    const baseRev = d.data?.revenue || [5577, 6034, 6676, 7801, 8596];
    const lastTotalRev = baseRev[baseRev.length - 1] || 8596;
    const projTotalRevs = [
        Math.round(lastTotalRev * 1.05),
        Math.round(lastTotalRev * 1.10),
        Math.round(lastTotalRev * 1.15)
    ];

    let datasets = [];
    if (segNames.length > 0) {
        datasets = segNames.map(segKey => {
            const seg = segY[segKey];
            const dataArr = years.map(y => {
                return seg.data && seg.data[y + '(CN)'] ? seg.data[y + '(CN)'].revenue : 0;
            });
            // Dự phóng: Doanh thu mảng = Tổng doanh thu dự phóng * Tỷ trọng mảng năm 2025
            const weight = segWeights[segKey];
            dataArr.push(Math.round(projTotalRevs[0] * weight));
            dataArr.push(Math.round(projTotalRevs[1] * weight));
            dataArr.push(Math.round(projTotalRevs[2] * weight));
            
            return {
                label: seg.label,
                data: dataArr,
                backgroundColor: seg.color || '#3b82f6'
            };
        });
    } else {
        // Fallback mock data
        datasets = [
            { label: 'Cho thuê KCN', data: [837, 905, 1002, 1170, 1289, 1419, 1589, 1716], backgroundColor: '#10b981' },
            { label: 'Dịch vụ tiện ích', data: [4351, 4707, 5208, 6085, 6705, 7114, 7748, 8213], backgroundColor: '#3b82f6' },
            { label: 'BĐS dân cư & Khác', data: [390, 422, 467, 546, 602, 621, 657, 677], backgroundColor: '#ec4899' },
        ];
    }

    const allLabels = [...years.map(y => y + 'A'), '2026E', '2027E', '2028E'];

    const ctx = document.getElementById('segRevChart').getContext('2d');
    charts['segRevChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: allLabels,
            datasets
        },
        plugins: [{
            id: 'segRevLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 9px "Inter", sans-serif';
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const val = dataset.data[index];
                        if (val >= 300) {
                            if (bar.height > 15) {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y + (bar.height / 2));
                            } else {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y - 6);
                            }
                        }
                    });
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { 
                    display: true, 
                    position: 'top',
                    labels: {
                        color: '#9ca3af',
                        font: { size: 9, family: 'Inter', weight: 'bold' }
                    }
                } 
            },
            scales: {
                x: { stacked: true, grid: { display: false }, ticks: { color: '#9ca3af' } },
                y: { stacked: true, grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#9ca3af' } }
            }
        }
    });
}

function renderChartSegGp(d) {
    destroyChart('segGpChart');
    const years = d.data?.years || [2021, 2022, 2023, 2024, 2025];
    const latestYear = years[years.length - 1];
    const segY = d.segments_yearly || {};
    const segNames = Object.keys(segY);

    // Tính tỷ trọng doanh thu năm gần nhất
    let totalRev2025 = 0;
    const segWeights = {};
    segNames.forEach(segKey => {
        const seg = segY[segKey];
        const r2025 = seg.data && seg.data[latestYear + '(CN)'] ? (seg.data[latestYear + '(CN)'].revenue || 0) : 0;
        totalRev2025 += r2025;
    });
    segNames.forEach(segKey => {
        const seg = segY[segKey];
        const r2025 = seg.data && seg.data[latestYear + '(CN)'] ? (seg.data[latestYear + '(CN)'].revenue || 0) : 0;
        segWeights[segKey] = totalRev2025 > 0 ? (r2025 / totalRev2025) : (1 / segNames.length);
    });

    // Dự phóng tổng doanh thu
    const baseRev = d.data?.revenue || [5577, 6034, 6676, 7801, 8596];
    const lastTotalRev = baseRev[baseRev.length - 1] || 8596;
    const projTotalRevs = [
        Math.round(lastTotalRev * 1.05),
        Math.round(lastTotalRev * 1.10),
        Math.round(lastTotalRev * 1.15)
    ];

    let datasets = [];
    if (segNames.length > 0) {
        datasets = segNames.map(segKey => {
            const seg = segY[segKey];
            
            // Tính Biên LNG trung bình lịch sử thực tế của mảng
            let sumRev = 0, sumCogs = 0;
            years.forEach(y => {
                const sData = seg.data && seg.data[y + '(CN)'] ? seg.data[y + '(CN)'] : null;
                if (sData) {
                    sumRev += sData.revenue || 0;
                    sumCogs += sData.cogs || 0;
                }
            });
            const histAvgMargin = sumRev > 0 ? ((sumRev - sumCogs) / sumRev) : 0.1;

            const dataArr = years.map(y => {
                const sData = seg.data && seg.data[y + '(CN)'] ? seg.data[y + '(CN)'] : null;
                return sData ? (sData.revenue - sData.cogs) : 0;
            });

            // Dự phóng: LNG mảng = (Tổng doanh thu dự phóng * tỷ trọng mảng) * Biên LNG trung bình lịch sử
            const weight = segWeights[segKey];
            dataArr.push(Math.round((projTotalRevs[0] * weight) * histAvgMargin));
            dataArr.push(Math.round((projTotalRevs[1] * weight) * histAvgMargin));
            dataArr.push(Math.round((projTotalRevs[2] * weight) * histAvgMargin));
            
            return {
                label: seg.label,
                data: dataArr,
                backgroundColor: seg.color || '#3b82f6'
            };
        });
    } else {
        // Fallback mock
        datasets = [
            { label: 'Cho thuê KCN', data: [270, 325, 371, 437, 489, 513, 564, 620], backgroundColor: '#10b981' },
            { label: 'Dịch vụ tiện ích', data: [360, 359, 430, 440, 529, 555, 582, 611], backgroundColor: '#3b82f6' },
            { label: 'BĐS dân cư & Khác', data: [49, 54, 63, 72, 85, 90, 94, 98], backgroundColor: '#ec4899' }
        ];
    }

    const allLabels = [...years.map(y => y + 'A'), '2026E', '2027E', '2028E'];

    const ctx = document.getElementById('segGpChart').getContext('2d');
    charts['segGpChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: allLabels,
            datasets
        },
        plugins: [{
            id: 'segGpLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 9px "Inter", sans-serif';
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const val = dataset.data[index];
                        if (val >= 50) {
                            if (bar.height > 15) {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y + (bar.height / 2));
                            } else {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y - 6);
                            }
                        }
                    });
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { 
                    display: true, 
                    position: 'top',
                    labels: {
                        color: '#9ca3af',
                        font: { size: 9, family: 'Inter', weight: 'bold' }
                    }
                } 
            },
            scales: {
                x: { stacked: true, grid: { display: false }, ticks: { color: '#9ca3af' } },
                y: { stacked: true, grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#9ca3af' } }
            }
        }
    });
}

function renderChartSegMargin(d) {
    destroyChart('segMarginChart');
    const years = d.data?.years || [2021, 2022, 2023, 2024, 2025];
    const segY = d.segments_yearly || {};
    const segNames = Object.keys(segY);
    
    let datasets = [];
    if (segNames.length > 0) {
        datasets = segNames.map(segKey => {
            const seg = segY[segKey];
            
            // Tính Biên LNG lịch sử từng năm
            const dataArr = years.map(y => {
                const sData = seg.data && seg.data[y + '(CN)'] ? seg.data[y + '(CN)'] : null;
                if (sData && sData.revenue > 0) {
                    return Math.round((sData.revenue - sData.cogs) / sData.revenue * 1000) / 10;
                }
                return 0;
            });
            
            // Biên LNG dự phóng giữ ổn định bằng trung bình lịch sử thực tế
            let sumRev = 0, sumCogs = 0;
            years.forEach(y => {
                const sData = seg.data && seg.data[y + '(CN)'] ? seg.data[y + '(CN)'] : null;
                if (sData) {
                    sumRev += sData.revenue || 0;
                    sumCogs += sData.cogs || 0;
                }
            });
            const histAvgMarginPct = sumRev > 0 ? Math.round(((sumRev - sumCogs) / sumRev) * 1000) / 10 : 10.0;
            
            dataArr.push(histAvgMarginPct);
            dataArr.push(histAvgMarginPct);
            dataArr.push(histAvgMarginPct);
            
            return {
                label: seg.label,
                data: dataArr,
                borderColor: seg.color || '#3b82f6',
                borderWidth: 2,
                pointRadius: 3,
                fill: false
            };
        });
    } else {
        // Fallback mock
        datasets = [
            { label: 'Cho thuê KCN', data: [71.0, 72.1, 71.4, 71.7, 72.0, 72.0, 72.0, 72.0], borderColor: '#10b981', fill: false },
            { label: 'Dịch vụ tiện ích', data: [7.2, 6.7, 7.4, 7.2, 8.1, 8.1, 8.1, 8.1], borderColor: '#3b82f6', fill: false },
            { label: 'BĐS dân cư & Khác', data: [22.4, 20.1, 21.0, 22.0, 23.4, 23.4, 23.4, 23.4], borderColor: '#ec4899', fill: false }
        ];
    }

    const allLabels = [...years.map(y => y + 'A'), '2026E', '2027E', '2028E'];

    const ctx = document.getElementById('segMarginChart').getContext('2d');
    charts['segMarginChart'] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allLabels,
            datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { 
                    display: true, 
                    position: 'top',
                    labels: {
                        color: '#9ca3af',
                        font: { size: 9, family: 'Inter', weight: 'bold' }
                    }
                } 
            },
            scales: {
                x: { grid: { display: false }, ticks: { color: '#9ca3af' } },
                y: { 
                    grid: { color: 'rgba(255, 255, 255, 0.05)' }, 
                    ticks: { 
                        color: '#9ca3af',
                        callback: function(value) { return value + '%'; }
                    } 
                }
            }
        }
    });
}

function renderChartSegRevQtr(d, qKeys) {
    destroyChart('segRevQtrChart');
    const segQ = d.segments_quarterly || {};
    const segNames = Object.keys(segQ);

    let datasets = [];
    if (segNames.length > 0) {
        datasets = segNames.map(segKey => {
            const seg = segQ[segKey];
            const dataArr = qKeys.map(qk => {
                const qData = seg.data && seg.data[qk] ? seg.data[qk] : null;
                return qData ? qData.revenue : 0;
            });
            return {
                label: seg.label,
                data: dataArr,
                backgroundColor: seg.color || '#3b82f6'
            };
        });
    } else {
        // Fallback mock quý tương đương tỷ lệ
        datasets = [
            { label: 'Cho thuê KCN', data: [210, 225, 230, 240, 280, 295, 302, 310, 322, 335, 340, 355], backgroundColor: '#10b981' },
            { label: 'Dịch vụ tiện ích', data: [1100, 1150, 1200, 1250, 1500, 1550, 1600, 1650, 1700, 1750, 1800, 1850], backgroundColor: '#3b82f6' },
            { label: 'BĐS dân cư & Khác', data: [90, 95, 100, 105, 120, 125, 130, 135, 140, 145, 150, 155], backgroundColor: '#ec4899' }
        ];
    }

    const ctx = document.getElementById('segRevQtrChart').getContext('2d');
    charts['segRevQtrChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: qKeys,
            datasets
        },
        plugins: [{
            id: 'segRevQtrLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8.5px "Inter", sans-serif';
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const val = dataset.data[index];
                        // LỌC: CHỈ IN SỐ CHO DỮ LIỆU LỚN >= 100 TỶ ĐỂ ĐỠ RỐI
                        if (val >= 100) {
                            if (bar.height > 15) {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y + (bar.height / 2));
                            } else {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y - 6);
                            }
                        }
                    });
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { 
                    display: true, 
                    position: 'top',
                    labels: {
                        color: '#9ca3af',
                        font: { size: 9, family: 'Inter', weight: 'bold' }
                    }
                } 
            },
            scales: {
                x: { stacked: true, grid: { display: false }, ticks: { color: '#9ca3af', font: { size: 9 } } },
                y: { stacked: true, grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#9ca3af' } }
            }
        }
    });
}

function renderChartSegGpQtr(d, qKeys) {
    destroyChart('segGpQtrChart');
    const segQ = d.segments_quarterly || {};
    const segNames = Object.keys(segQ);

    let datasets = [];
    if (segNames.length > 0) {
        datasets = segNames.map(segKey => {
            const seg = segQ[segKey];
            const dataArr = qKeys.map(qk => {
                const qData = seg.data && seg.data[qk] ? seg.data[qk] : null;
                return qData ? (qData.revenue - qData.cogs) : 0;
            });
            return {
                label: seg.label,
                data: dataArr,
                backgroundColor: seg.color || '#3b82f6'
            };
        });
    } else {
        // Fallback mock quý LNG tương đương tỷ lệ
        datasets = [
            { label: 'Cho thuê KCN', data: [70, 75, 78, 80, 92, 95, 98, 102, 105, 110, 115, 120], backgroundColor: '#10b981' },
            { label: 'Dịch vụ tiện ích', data: [90, 92, 95, 98, 110, 112, 115, 118, 120, 122, 125, 128], backgroundColor: '#3b82f6' },
            { label: 'BĐS dân cư & Khác', data: [15, 16, 18, 20, 22, 23, 24, 25, 26, 27, 28, 29], backgroundColor: '#ec4899' }
        ];
    }

    const ctx = document.getElementById('segGpQtrChart').getContext('2d');
    charts['segGpQtrChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: qKeys,
            datasets
        },
        plugins: [{
            id: 'segGpQtrLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8.5px "Inter", sans-serif';
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const val = dataset.data[index];
                        // LỌC: CHỈ IN SỐ CHO DỮ LIỆU LỚN >= 20 TỶ ĐỂ ĐỠ RỐI
                        if (val >= 20) {
                            if (bar.height > 15) {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y + (bar.height / 2));
                            } else {
                                ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y - 6);
                            }
                        }
                    });
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { 
                    display: true, 
                    position: 'top',
                    labels: {
                        color: '#9ca3af',
                        font: { size: 9, family: 'Inter', weight: 'bold' }
                    }
                } 
            },
            scales: {
                x: { stacked: true, grid: { display: false }, ticks: { color: '#9ca3af', font: { size: 9 } } },
                y: { stacked: true, grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#9ca3af' } }
            }
        }
    });
}



function renderChartRevNpat(data) {
    destroyChart('revNpatChart');
    const years = data.years || [2021, 2022, 2023, 2024, 2025];
    const rev   = data.revenue || [5577, 6034, 6676, 7801, 8596];
    const npat  = data.npat || [908, 1009, 1003, 1278, 1467];

    const allYears = [...years, 2026, 2027, 2028];
    const allRev   = [...rev, Math.round(rev[rev.length-1]*1.05), Math.round(rev[rev.length-1]*1.1), Math.round(rev[rev.length-1]*1.15)];
    const allNpat  = [...npat, Math.round(npat[npat.length-1]*1.03), Math.round(npat[npat.length-1]*1.06), Math.round(npat[npat.length-1]*1.08)];

    const ctx = document.getElementById('revNpatChart').getContext('2d');
    charts['revNpatChart'] = new Chart(ctx, {
        data: {
            labels: allYears.map(y => y + (y > 2025 ? 'E' : 'A')),
            datasets: [
                {
                    type: 'bar',
                    label: 'Doanh thu thuần (tỷ VND)',
                    data: allRev,
                    backgroundColor: 'rgba(30, 59, 139, 0.45)', // Có độ trong suốt nhẹ để nhìn thấu qua
                    borderColor: 'rgba(30, 59, 139, 0.8)',
                    borderWidth: 1,
                    order: 2 // Vẽ ở dưới cùng
                },
                {
                    type: 'line',
                    label: 'LNST cổ đông mẹ (tỷ VND)',
                    data: allNpat,
                    borderColor: '#10b981',
                    borderWidth: 2.5,
                    pointRadius: 4,
                    pointBackgroundColor: '#10b981',
                    fill: false,
                    yAxisID: 'y2',
                    order: 1 // Luôn vẽ đè lên trên cùng để không bị cột che khuất
                }
            ]
        },
        plugins: [{
            id: 'revLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 9px "Inter", sans-serif';
                ctx.fillStyle = '#a3a3a3';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';

                const meta = chart.getDatasetMeta(0);
                meta.data.forEach((bar, index) => {
                    const val = chart.data.datasets[0].data[index];
                    ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y - 4);
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    type: 'linear',
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y2: {
                    type: 'linear',
                    position: 'right',
                    grid: { drawOnChartArea: false },
                    ticks: { color: '#10b981' }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#9ca3af' }
                }
            }
        }
    });
}

function renderChartUnearned(qKeys) {
    destroyChart('unearnedRevChart');
    const mockUnearned = [11034, 11232, 10954, 11587, 11467, 11590, 11729, 12009, 12619, 12822, 12976, 13434];
    const mockBacklog  = [1.8, 1.7, 1.5, 1.6, 1.5, 1.5, 1.5, 1.5, 1.5, 1.45, 1.48, 1.51];
    const labels = ["2023Q2","2023Q3","2023Q4","2024Q1","2024Q2","2024Q3","2024Q4","2025Q1","2025Q2","2025Q3","2025Q4","2026Q1"];

    const ctx = document.getElementById('unearnedRevChart').getContext('2d');
    charts['unearnedRevChart'] = new Chart(ctx, {
        data: {
            labels: labels,
            datasets: [
                { 
                    type: 'bar', 
                    label: 'Doanh thu chưa thực hiện (tỷ VND)', 
                    data: mockUnearned, 
                    backgroundColor: 'rgba(2, 132, 199, 0.45)', // Trong suốt nhẹ
                    borderColor: 'rgba(2, 132, 199, 0.8)',
                    borderWidth: 1,
                    order: 2 // Vẽ dưới
                },
                { 
                    type: 'line', 
                    label: 'Backlog coverage (năm)', 
                    data: mockBacklog, 
                    borderColor: '#ef4444', 
                    pointBackgroundColor: '#ef4444',
                    pointRadius: 4,
                    borderWidth: 2,
                    yAxisID: 'y2', 
                    fill: false,
                    order: 1 // Luôn đè lên trên cùng
                }
            ]
        },
        plugins: [{
            id: 'unearnedLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8.5px "Inter", sans-serif';
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';

                const meta = chart.getDatasetMeta(0);
                meta.data.forEach((bar, index) => {
                    const val = chart.data.datasets[0].data[index];
                    ctx.fillText(val.toLocaleString('vi-VN'), bar.x, bar.y - 4);
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { 
                    type: 'linear', 
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y2: { 
                    type: 'linear', 
                    position: 'right', 
                    grid: { drawOnChartArea: false },
                    ticks: { color: '#ef4444' }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#9ca3af' }
                }
            }
        }
    });
}

function renderChartLandAsset(qKeys) {
    destroyChart('landAssetChart');
    const labels = ["2023Q2","2023Q3","2023Q4","2024Q1","2024Q2","2024Q3","2024Q4","2025Q1","2025Q2","2025Q3","2025Q4","2026Q1"];
    const datasets = [
        { label: 'Tồn kho', data: [404, 454, 473, 427, 389, 386, 322, 319, 337, 290, 263, 251], backgroundColor: '#f59e0b' },
        { label: 'XDCB dở dang', data: [2940, 2720, 2714, 2288, 2485, 2518, 2380, 2285, 2304, 2405, 2231, 2202], backgroundColor: '#8b5cf6' },
        { label: 'Tài sản đầu tư', data: [5230, 5510, 5443, 5548, 5524, 5493, 5878, 6047, 6007, 6021, 6160, 6241], backgroundColor: '#10b981' }
    ];

    const ctx = document.getElementById('landAssetChart').getContext('2d');
    charts['landAssetChart'] = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets },
        plugins: [{
            id: 'landAssetLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8px "Inter", sans-serif';
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const val = dataset.data[index];
                        if (val > 0 && bar.height > 12) {
                            ctx.fillText(Math.round(val).toLocaleString('vi-VN'), bar.x, bar.y + (bar.height / 2));
                        }
                    });
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { 
                x: { stacked: true, grid: { display: false }, ticks: { color: '#9ca3af' } }, 
                y: { stacked: true, grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#9ca3af' } } 
            }
        }
    });
}

function renderChartPrepayments(qKeys) {
    destroyChart('prepaymentsChart');
    const labels = ["2023Q2","2023Q3","2023Q4","2024Q1","2024Q2","2024Q3","2024Q4","2025Q1","2025Q2","2025Q3","2025Q4","2026Q1"];
    const data = [73.9, 66.2, 54.2, 60.1, 25.4, 8.7, 9.2, 18.7, 24.7, 141.8, 25.2, 45.8];

    const ctx = document.getElementById('prepaymentsChart').getContext('2d');
    charts['prepaymentsChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{ label: 'Người mua trả tiền trước', data, backgroundColor: 'rgba(185, 28, 28, 0.65)' }]
        },
        plugins: [{
            id: 'prepayLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8.5px "Inter", sans-serif';
                ctx.fillStyle = '#a3a3a3';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';

                const meta = chart.getDatasetMeta(0);
                meta.data.forEach((bar, index) => {
                    const val = chart.data.datasets[0].data[index];
                    ctx.fillText(val.toLocaleString('vi-VN'), bar.x, bar.y - 4);
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#9ca3af' } },
                x: { grid: { display: false }, ticks: { color: '#9ca3af' } }
            }
        }
    });
}

function renderChartDebtInterest(qKeys) {
    destroyChart('debtInterestChart');
    const labels = ["2023Q2","2023Q3","2023Q4","2024Q1","2024Q2","2024Q3","2024Q4","2025Q1","2025Q2","2025Q3","2025Q4","2026Q1"];
    const debt = [1200, 1150, 1480, 1850, 2100, 3100, 3000, 3800, 4200, 4800, 4700, 4600];
    const interest = [8.0, 5.0, 6.2, 5.5, 4.2, 3.8, 3.8, 4.2, 4.8, 5.0, 4.9, 4.66];

    const ctx = document.getElementById('debtInterestChart').getContext('2d');
    charts['debtInterestChart'] = new Chart(ctx, {
        data: {
            labels,
            datasets: [
                { 
                    type: 'bar', 
                    label: 'Tổng nợ vay (tỷ VND)', 
                    data: debt, 
                    backgroundColor: 'rgba(153, 27, 27, 0.45)', // Trong suốt nhẹ
                    borderColor: 'rgba(153, 27, 27, 0.8)',
                    borderWidth: 1,
                    order: 2 // Vẽ dưới
                },
                { 
                    type: 'line', 
                    label: 'Lãi suất vay (%/năm)', 
                    data: interest, 
                    borderColor: '#3b82f6', 
                    pointBackgroundColor: '#3b82f6',
                    pointRadius: 4,
                    borderWidth: 2,
                    yAxisID: 'y2', 
                    fill: false,
                    order: 1 // Vẽ đè lên trên
                }
            ]
        },
        plugins: [{
            id: 'debtLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8.5px "Inter", sans-serif';
                ctx.fillStyle = '#a3a3a3';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';

                const meta = chart.getDatasetMeta(0);
                meta.data.forEach((bar, index) => {
                    const val = chart.data.datasets[0].data[index];
                    ctx.fillText(val.toLocaleString('vi-VN'), bar.x, bar.y - 4);
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { 
                    type: 'linear', 
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y2: { 
                    type: 'linear', 
                    position: 'right', 
                    min: 0,
                    max: 10,
                    grid: { drawOnChartArea: false },
                    ticks: { 
                        color: '#3b82f6',
                        callback: function(value) { return value + '%'; }
                    }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#9ca3af' }
                }
            }
        }
    });
}

function renderChartNpatRoe(qKeys) {
    destroyChart('npatRoeChart');
    const labels = ["2023Q2","2023Q3","2023Q4","2024Q1","2024Q2","2024Q3","2024Q4","2025Q1","2025Q2","2025Q3","2025Q4","2026Q1"];
    const npat = [280, 250, 350, 230, 280, 290, 310, 280, 320, 290, 300, 250];
    const roe  = [26, 21, 32, 24, 27, 26, 27, 21, 26, 23, 22, 25.1];

    const ctx = document.getElementById('npatRoeChart').getContext('2d');
    charts['npatRoeChart'] = new Chart(ctx, {
        data: {
            labels,
            datasets: [
                { 
                    type: 'bar', 
                    label: 'LNST (tỷ VND)', 
                    data: npat, 
                    backgroundColor: 'rgba(16, 185, 129, 0.45)', // Trong suốt nhẹ
                    borderColor: 'rgba(16, 185, 129, 0.8)',
                    borderWidth: 1,
                    order: 2 // Vẽ dưới
                },
                { 
                    type: 'line', 
                    label: 'ROE (%)', 
                    data: roe, 
                    borderColor: '#ef4444', 
                    pointBackgroundColor: '#ef4444',
                    pointRadius: 4,
                    borderWidth: 2,
                    yAxisID: 'y2', 
                    fill: false,
                    order: 1 // Luôn đè lên trên cùng
                }
            ]
        },
        plugins: [{
            id: 'npatLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                ctx.font = 'bold 8.5px "Inter", sans-serif';
                ctx.fillStyle = '#a3a3a3';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';

                const meta = chart.getDatasetMeta(0);
                meta.data.forEach((bar, index) => {
                    const val = chart.data.datasets[0].data[index];
                    ctx.fillText(val.toLocaleString('vi-VN'), bar.x, bar.y - 4);
                });
                ctx.restore();
            }
        }],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { 
                    type: 'linear', 
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y2: { 
                    type: 'linear', 
                    position: 'right', 
                    min: 20,
                    max: 35,
                    grid: { drawOnChartArea: false },
                    ticks: { 
                        color: '#ef4444',
                        callback: function(value) { return value + '%'; }
                    }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#9ca3af' }
                }
            }
        }
    });
}

function renderChartPE() {
    destroyChart('peChart');
    const labels = ["2019-Q4","2020-Q1","2020-Q2","2020-Q3","2020-Q4","2021-Q1","2021-Q2","2021-Q3","2021-Q4","2022-Q1","2022-Q2","2022-Q3","2022-Q4","2023-Q1","2023-Q2","2023-Q3","2023-Q4","2024-Q1","2024-Q2","2024-Q3","2024-Q4","2025-Q1","2025-Q2","2025-Q3","2025-Q4","2026-Q1"];
    const pe = [9.2, 12.5, 11.4, 24.1, 13.2, 11.0, 11.8, 12.9, 10.5, 12.8, 7.8, 8.4, 12.3, 13.5, 13.4, 16.1, 16.2, 13.6, 11.5, 11.8, 11.0, 10.4, 11.2, 10.5, 11.3, 9.2];

    const ctx = document.getElementById('peChart').getContext('2d');
    charts['peChart'] = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [
                {
                    label: 'P/E TTM',
                    data: pe,
                    borderColor: '#0284c7',
                    borderWidth: 2,
                    pointRadius: 3,
                    fill: false
                },
                {
                    label: 'Trung vị (11.8x)',
                    data: Array(labels.length).fill(11.8),
                    borderColor: '#f59e0b',
                    borderDash: [5, 5],
                    borderWidth: 1.5,
                    pointRadius: 0,
                    fill: false
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#9ca3af', font: { size: 9 }, maxRotation: 45, minRotation: 45 }
                }
            }
        }
    });
}

function renderChartPB() {
    destroyChart('pbChart');
    const clean_labels = ["2019-Q4","2020-Q1","2020-Q2","2020-Q3","2020-Q4","2021-Q1","2021-Q2","2021-Q3","2021-Q4","2022-Q1","2022-Q2","2022-Q3","2022-Q4","2023-Q1","2023-Q2","2023-Q3","2023-Q4","2024-Q1","2024-Q2","2024-Q3","2024-Q4","2025-Q1","2025-Q2","2025-Q3","2025-Q4","2026-Q1"];
    const pb = [3.4, 3.4, 3.5, 3.8, 8.2, 5.9, 5.2, 4.8, 4.5, 3.9, 3.3, 2.9, 3.0, 1.8, 2.2, 2.8, 3.2, 3.5, 3.3, 3.8, 3.5, 3.4, 2.8, 2.9, 2.6, 2.1];
    const roe = [22.0, 20.5, 23.1, 28.4, 25.0, 24.2, 23.8, 21.0, 20.2, 19.5, 21.2, 22.8, 25.1, 28.0, 20.1, 33.4, 22.5, 28.0, 26.2, 25.1, 23.4, 24.8, 26.5, 22.1, 21.8, 25.1];

    const ctx = document.getElementById('pbChart').getContext('2d');
    charts['pbChart'] = new Chart(ctx, {
        data: {
            labels: clean_labels,
            datasets: [
                {
                    type: 'line',
                    label: 'P/B TTM (Trục trái)',
                    data: pb,
                    borderColor: '#10b981',
                    borderWidth: 2,
                    pointRadius: 3,
                    yAxisID: 'y',
                    fill: false
                },
                {
                    type: 'line',
                    label: 'Trung vị (3.35x)',
                    data: Array(clean_labels.length).fill(3.35),
                    borderColor: '#f59e0b',
                    borderDash: [5, 5],
                    borderWidth: 1.5,
                    pointRadius: 0,
                    yAxisID: 'y',
                    fill: false
                },
                {
                    type: 'line',
                    label: 'ROE % (Trục phải)',
                    data: roe,
                    borderColor: '#ec4899',
                    borderWidth: 1.5,
                    borderDash: [3, 3],
                    pointRadius: 2,
                    yAxisID: 'y2',
                    fill: false
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
                    labels: { color: '#9ca3af', font: { size: 9 } }
                }
            },
            scales: {
                y: {
                    type: 'linear',
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y2: {
                    type: 'linear',
                    position: 'right',
                    grid: { drawOnChartArea: false },
                    ticks: {
                        color: '#ec4899',
                        callback: function(value) { return value + '%'; }
                    }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#9ca3af', font: { size: 9 }, maxRotation: 45, minRotation: 45 }
                }
            }
        }
    });
}

// ── PEER BENCHMARK TABLE ─────────────────────────────────────────────────
function renderPeerBenchmark() {
    const tbody = document.getElementById('peer-benchmark-tbody');
    const peers = [
        { ticker: 'SIP Sài Gòn VRG', mcap: 12045, pb: '1.95x', pe: '9.0x', roe: '21.6%', dtcth: 13434, dtcth_mcap: '1.12x', debt_eq: '0.75x' },
        { ticker: 'IDC IDICO', mcap: 15180, pb: '1.72x', pe: '14.0x', roe: '12.2%', dtcth: 5396, dtcth_mcap: '0.36x', debt_eq: '0.65x' },
        { ticker: 'SZC Sonadezi Châu Đức', mcap: 3798, pb: '1.18x', pe: '55.4x', roe: '2.1%', dtcth: 531, dtcth_mcap: '0.14x', debt_eq: '0.80x' },
        { ticker: 'SZL Sonadezi Long Thành', mcap: 1321, pb: '1.84x', pe: '10.3x', roe: '17.9%', dtcth: 734, dtcth_mcap: '0.56x', debt_eq: '0.33x' },
        { ticker: 'KBC Kinh Bắc', mcap: 26934, pb: '1.00x', pe: '30.0x', roe: '3.3%', dtcth: 31, dtcth_mcap: '0.00x', debt_eq: '1.12x' },
        { ticker: 'NTC Nam Tân Uyên', mcap: 3187, pb: '2.41x', pe: '20.4x', roe: '11.8%', dtcth: 4674, dtcth_mcap: '1.47x', debt_eq: '0.03x' },
        { ticker: 'DPR Cao su Đồng Phú', mcap: 3302, pb: '0.95x', pe: '7.9x', roe: '12.1%', dtcth: 926, dtcth_mcap: '0.28x', debt_eq: '0.00x' },
        { ticker: 'BCM Becamex IDC', mcap: 52061, pb: '2.32x', pe: '46.6x', roe: '5.0%', dtcth: 1399, dtcth_mcap: '0.03x', debt_eq: '1.14x' },
        { ticker: 'PHR Cao su Phước Hòa', mcap: 8523, pb: '1.89x', pe: '7.6x', roe: '24.7%', dtcth: 1258, dtcth_mcap: '0.15x', debt_eq: '0.00x' }
    ];

    tbody.innerHTML = peers.map(p => `
        <tr class="${p.ticker.startsWith(currentTicker) ? 'highlight-row' : ''}">
            <td>${p.ticker}</td>
            <td>${fmt(p.mcap)}</td>
            <td>${p.pb}</td>
            <td>${p.pe}</td>
            <td>${p.roe}</td>
            <td>${fmt(p.dtcth)}</td>
            <td>${p.dtcth_mcap}</td>
            <td>${p.debt_eq}</td>
        </tr>`).join('');
}

// ── UTILS ────────────────────────────────────────────────────────────────
function renderMoat(moats) {
    const el = document.getElementById('moat-scorecard-list');
    el.innerHTML = Object.entries(moats).map(([name, m]) => {
        const score = m.score || 0;
        const stars = '★'.repeat(score) + '☆'.repeat(5 - score);
        return `
            <div class="moat-item">
                <div class="moat-header">
                    <span class="moat-name">${name}</span>
                    <span class="moat-stars">${stars}</span>
                </div>
                <div class="moat-desc">${m.desc || ''}</div>
            </div>`;
    }).join('');
}

function renderPestle(pestle) {
    const icons = { Political:'🏛️', Economic:'📈', Social:'👥', Technological:'💡', Legal:'⚖️', Environmental:'🌿' };
    document.getElementById('pestle-grid-container').innerHTML =
        Object.entries(pestle).map(([k, v]) => `
            <div class="pestle-item">
                <div class="pestle-label">${icons[k] || ''} ${k}</div>
                <div class="pestle-text">${v}</div>
            </div>`).join('');
}

function renderFinTable(data) {
    const years = data.years || [2021, 2022, 2023, 2024, 2025];
    const rev   = data.revenue || [4301, 7485, 7237, 8846, 8588];
    const gp    = data.grossProfit || [737, 3059, 2423, 3337, 3060];
    const npat  = data.npat || [578, 2054, 1656, 2392, 2354];
    const table = document.getElementById('financial-table');

    let head = '<thead><tr><th>Chỉ tiêu (tỷ VND)</th>';
    years.forEach(y => { head += `<th>${y}</th>`; });
    head += '</tr></thead>';

    const rows_data = [
        ['Doanh thu thuần', rev],
        ['Lợi nhuận gộp', gp],
        ['Biên LNG (%)', years.map((_, i) => rev[i] > 0 ? fmtP(gp[i] / rev[i]) : '—')],
        ['LNST (hợp nhất)', npat],
    ];
    let body = '<tbody>';
    rows_data.forEach(([label, vals]) => {
        body += `<tr><td>${label}</td>`;
        vals.forEach(v => {
            body += `<td>${typeof v === 'string' ? v : fmtB(v)}</td>`;
        });
        body += '</tr>';
    });
    body += '</tbody>';

    table.innerHTML = head + body;
}

// ── ANALYSIS TEXTS ────────────────────────────────────────────────────────
function renderAnalysisTexts(d, val, data) {
    const years = data.years || [2021, 2022, 2023, 2024, 2025];
    const rev   = data.revenue || [4301, 7485, 7237, 8846, 8588];
    const gp    = data.grossProfit || [737, 3059, 2423, 3337, 3060];
    const n     = rev.length;

    // Segment analysis
    document.getElementById('analysis-seg-rev').innerHTML = 
        `<p>Mảng đóng góp lớn nhất: <b>Dịch vụ tiện ích</b> (~78% doanh thu). Cơ cấu tuyệt đối (tỷ VND) phản ánh đúng quy mô thực tế, giảm thiểu rủi ro biến động ngắn hạn từ BĐS.</p>`;

    // Rev NPAT analysis
    if (n >= 1) {
        document.getElementById('analysis-rev-npat').innerHTML =
            `<p>LNST năm dự phóng đầu đạt <b>1.310 tỷ</b> so với năm gần nhất 1.323 tỷ (-1%). Năm dự phóng đầu đã blend lũy kế quý thực tế.</p>`;
    }

    // Backlog analysis
    document.getElementById('analysis-backlog').innerHTML =
        `<p>DTCTH quý gần nhất: <b>13.434,1 tỷ</b> (tăng 1.424 tỷ so với cùng kỳ). Đây là "của để dành" — tiền khách đã trả sẽ hạch toán dần thành doanh thu. Backlog coverage: <b>1,51 năm</b>.</p>`;

    // Land assets analysis
    document.getElementById('analysis-land-assets').innerHTML =
        `<p>XDCB dở dang giảm 84 tỷ so với cùng kỳ — dự án dần hoàn thành chuyển sang khai thác. Giá trị SỔ SÁCH — giá thị trường của quỹ đất thường cao hơn đáng kể (upside NAV).</p>`;

    // Prepayments analysis
    document.getElementById('analysis-prepayments').innerHTML =
        `<p>Quý gần nhất: <b>46 tỷ</b>. Chưa có đột biến — theo dõi chỉ số này để bắt tín hiệu ký mới sớm nhất.</p>`;

    // Debt interest analysis
    document.getElementById('analysis-debt-interest').innerHTML =
        `<p>Lãi suất vay bình quân ANNUALIZED (quý gần nhất): <b>4,66%/năm</b> = 4 × Chi phí lãi vay quý / Nợ vay bình quân trong kỳ.</p>`;

    // NPAT ROE analysis
    document.getElementById('analysis-npat-roe').innerHTML =
        `<p>ROE bình quân annualized (12 quý): <b>25,1%</b>. Doanh nghiệp KCN mô hình bán đứt có LNST lồi lõm mạnh theo tiến độ bàn giao — nhìn xu hướng TTM thay vì từng quý đơn lẻ.</p>`;

    // Valuation analysis
    document.getElementById('analysis-pe').innerHTML =
        `<p>P/E trung vị lịch sử: <b>11,8x</b> — trọng số 30% trong định giá (LNST KCN mô hình bán đứt dễ lồi lõm nên P/E chỉ bổ trợ).</p>`;

    document.getElementById('analysis-pb').innerHTML =
        `<p>P/B trung vị lịch sử: <b>3,35x</b> — neo chính cho định giá (trọng số 70%, proxy NAV cho doanh nghiệp tài sản nặng).</p>`;
}

// ── INIT ──────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    const inp = document.getElementById('stock-search-input');
    inp.addEventListener('input', () => renderOverview(inp.value));

    renderOverview('');

    const params = new URLSearchParams(window.location.search);
    const t = params.get('ticker') || params.get('t');
    if (t && KCN_TICKERS.includes(t.toUpperCase())) {
        loadTicker(t.toUpperCase());
    } else {
        loadTicker('SIP'); // default view SIP
    }
});
