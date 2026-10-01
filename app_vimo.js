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
    demographics: 'Dân số', bank_alm: 'Rủi ro hệ thống ngân hàng (ALM)',
};
const GROUP_ORDER = ['growth', 'inflation', 'monetary', 'trade', 'fiscal', 'labor', 'external', 'market', 'demographics', 'bank_alm'];
const GROUP_ICONS = {
    growth: '📈', inflation: '💰', monetary: '🏦', trade: '🚢',
    fiscal: '🏛️', labor: '👷', external: '🌐', market: '📊',
    demographics: '👥', bank_alm: '🏦',
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
    '40yo': '40yo.vn (tổng hợp IMF/World Bank/ADB, tự động)',
};

const CHART_DEFAULTS = {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
        x: { ticks: { color: '#9aa5bd', font: { size: 9 }, maxRotation: 45, autoSkip: true, maxTicksLimit: 8 }, grid: { display: false } },
        y: { ticks: { color: '#9aa5bd', font: { size: 9 } }, grid: { color: 'rgba(255,255,255,0.04)' } },
    },
};

// Cấu hình chartjs-plugin-datalabels dùng chung cho các chart NHIỀU đường + NHIỀU điểm/ngày
// (interbank/bond yield history) — chỉ hiện số ở điểm CUỐI mỗi đường (giá trị mới nhất) để tránh
// rối mắt khi có 15-20+ điểm x nhiều kỳ hạn chồng lên nhau; vẫn cho biết "số giá trị" ngay trên
// biểu đồ như yêu cầu, thay vì phải rê chuột xem tooltip.
function _endpointDatalabelsConfig(decimals) {
    return {
        display: (ctx) => ctx.dataset.data.slice(ctx.dataIndex + 1).every(v => v === null || v === undefined)
            && (ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] !== undefined),
        color: (ctx) => ctx.dataset.borderColor, font: { size: 9, weight: '600' },
        anchor: 'end', align: 'right', offset: 4, clip: false,
        formatter: (v) => (v === null || v === undefined) ? '' : v.toFixed(decimals),
    };
}

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
    renderMonitoringTable(data.monitoringTable);
    renderSynthesis(data.synthesis);
    renderValuation(data.marketValuation);
    renderVnindexCompare(data.marketValuation, data.marketValuationHeadline, data.decisionExvin, data.decisionHeadline, data.decision);
    renderIndicatorGroups(data.indicators);
    renderInternationalSection(data.indicators);
    // PHẢI gọi SAU renderIndicatorGroups() — hàm đó destroy() TOÀN BỘ chartInstances hiện có ở
    // đầu (dọn dẹp cho lần render riêng của nó), nên nếu gọi renderMacroOverview() trước đó thì
    // biểu đồ vừa tạo sẽ bị destroy() ngay sau, canvas về trạng thái rỗng dù không có lỗi console
    // nào (Chart.js destroy() im lặng) — đã xác nhận qua Playwright (canvas kẹt ở 300x150 mặc
    // định, Chart.getChart() trả null) trước khi đổi thứ tự gọi.
    renderMacroOverview(data.macroOverview);
    // Cùng lý do thứ tự gọi như renderMacroOverview() ở trên (SAU renderIndicatorGroups()) — 2
    // chart trong mục này dùng _renderGenericIndicatorCard(), cùng cơ chế chartInstances.
    renderBankingSystemRiskSection(data.bankingSystemRisk, data.indicators);
    renderCreditDepositStructure(data.bankingSystemRisk && data.bankingSystemRisk.creditDepositStructure);
    renderMaturityStructure(data.bankingSystemRisk && data.bankingSystemRisk.maturityStructure);
    renderFxPressureCard(data.indicators);

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

// Mục RIÊNG trong "🧭 Tổng hợp Phân tích Đa Chỉ số — Bức tranh Tổng thể" (user 2026-09-19): đánh
// giá rủi ro lãi suất + thanh khoản của TOÀN NGÀNH ngân hàng (tổng hợp có trọng số theo quy mô từ
// 26 ngân hàng niêm yết/UPCoM, xem bank_system_risk.py) — số liệu risk (bankingSystemRisk, kỳ MỚI
// NHẤT) + 2 chart LỊCH SỬ THEO QUÝ ngay dưới, dùng LẠI đúng series đã có sẵn ở
// indicators.bank_alm_system_ir_risk_ratio/_liquidity_risk_ratio (KHÔNG tính lại ở JS, tránh lệch
// với PDF/Excel). PHẢI gọi SAU renderIndicatorGroups() — hàm đó destroy() toàn bộ chartInstances
// hiện có ở đầu, xem ghi chú tại nơi gọi trong DOMContentLoaded.
function renderBankingSystemRiskSection(risk, indicators) {
    const section = document.getElementById('synthesis-banking-risk-section');
    if (!section) return;
    if (!risk) { section.style.display = 'none'; return; }
    section.style.display = '';

    const set = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text || '-'; };
    set('synthesis-banking-risk-summary', risk.summaryText);
    _renderBankingAssessment(risk.assessment);

    const cov = risk.coverage || {};
    const ir = risk.interestRateRisk || {};
    const liq = risk.liquidityRisk || {};
    const pct = (v, d = 2) => (v === null || v === undefined) ? 'N/A' : `${(v * 100).toFixed(d)}%`;
    const coverageDetail = `${cov.nBanksReported ?? 0}/${cov.nBanksTotal ?? 0} đã công bố, `
        + `${cov.nBanksPatched ?? 0} vá từ kỳ trước, ${cov.nBanksMissing ?? 0} chưa có`
        + (cov.assetsCoveragePct != null ? ` (${cov.assetsCoveragePct.toFixed(0)}% tổng tài sản)` : '');
    const stats = [
        ['Kỳ tổng hợp', risk.asOf],
        ['Độ phủ dữ liệu', coverageDetail],
        ['Gap ròng lãi suất ≤1 năm / Tổng TS', pct(ir.netGapRatio)],
        ['Mức phân tán lãi suất (không bù trừ giữa các NH)', pct(ir.dispersionGapRatio)],
        ['NH lệch lãi suất nhiều nhất', ir.worstBank ? `${ir.worstBank.ticker} (${(ir.worstBank.ratio * 100).toFixed(1)}%)` : 'N/A'],
        ['Liquid Assets / Tổng TS', pct(liq.liquidAssetsRatio, 1)],
        ['Che phủ nếu rút -10% tiền gửi', (liq.depositRunCoverageByStress || {})['-10%'] != null
            ? `${((liq.depositRunCoverageByStress['-10%']) * 100).toFixed(0)}%` : 'N/A'],
        ['NH thanh khoản yếu nhất', liq.weakestBank ? `${liq.weakestBank.ticker} (che phủ ${(liq.weakestBank.coverage * 100).toFixed(0)}%)` : 'N/A'],
    ];
    // Cau truc ky han nguon von he thong (xem "Danh gia rui ro thanh khoan cau truc he thong.docx",
    // user 2026-09-21) - them vao CUNG danh sach stats tren, khong tach khoi rieng.
    const sf = risk.structuralFunding || {};
    if (sf.rolloverDependency12m != null) {
        stats.push(
            ['Rollover Dependency 12 tháng', pct(sf.rolloverDependency12m, 1)],
            ['Long-term Funding Coverage', pct(sf.longTermFundingCoverage, 1)],
            ['NH phụ thuộc rollover nhiều nhất', sf.mostDependentBank
                ? `${sf.mostDependentBank.ticker} (${(sf.mostDependentBank.rollover_dependency_12m * 100).toFixed(0)}%)` : 'N/A'],
        );
    }
    // Do phu du lieu RIENG cho cau truc ky han (user 2026-09-21, sau khi phat hien Rollover
    // Dependency he thong bi sai lech chi vi thieu du lieu — xem classify_structural_funding_phase()
    // trong bank_system_risk.py) — hien ro "kỳ nào, bao nhieu/26 ngan hang" de biet CAN backfill
    // them ngan hang nao moi du dai dien toan he thong, khong chi tin vao 1 con so % dep.
    if (sf.nBanksIncluded != null) {
        const missingList = (sf.missingTickers && sf.missingTickers.length) ? sf.missingTickers.join(', ') : 'không có';
        stats.push(['Dữ liệu quý gần nhất (cấu trúc kỳ hạn)',
            `Quý ${risk.asOf} — ${sf.nBanksIncluded}/26 ngân hàng (${(sf.coveragePct ?? 0).toFixed(0)}% tổng tài sản). `
            + `Còn thiếu: ${missingList}`]);
    }
    const statsGrid = document.getElementById('banking-risk-stats-grid');
    if (statsGrid) {
        statsGrid.innerHTML = stats.map(([lbl, val]) => {
            const isWide = lbl.startsWith('Dữ liệu quý gần nhất');
            return `
            <div class="vimo-indicator-card" ${isWide ? 'style="grid-column:1/-1"' : ''}>
                <div class="ind-header"><span class="ind-name">${lbl}</span></div>
                <div class="ind-value" style="font-size:${isWide ? '0.85em' : '1em'};line-height:1.4">${val}</div>
            </div>`;
        }).join('');
    }
    set('synthesis-banking-risk-missing', cov.missingTickers && cov.missingTickers.length
        ? `Chưa có dữ liệu: ${cov.missingTickers.join(', ')}` : '');
    const phaseEl = document.getElementById('synthesis-banking-risk-phase');
    if (phaseEl) {
        phaseEl.textContent = sf.phase
            ? `${sf.phase.phaseLabel} (dựa trên ${sf.phase.periodsUsed.join(', ')})` : '';
    }

    const chartsGrid = document.getElementById('banking-risk-charts-grid');
    if (chartsGrid) {
        chartsGrid.innerHTML = '';
        // SUA (user 2026-09-26): xep 4 chart thanh luoi 2x2 - HANG TREN 2 chi bao "cao hon = an toan hon"
        // (xanh) canh nhau, HANG DUOI 2 chi bao "cao hon = rui ro hon" (do) canh nhau, de doc chieu
        // tot/xau khong can doi chieu tung the. Phan loai theo goodDirection cua chinh chi bao (khong
        // hard-code theo ten), ten chi bao da noi ro thuoc mang nao (thanh khoan/cau truc ky han/lai suat).
        const _KEYS = ['bank_alm_system_liquidity_risk_ratio', 'bank_alm_system_long_term_funding_coverage',
                       'bank_alm_system_rollover_dependency_12m', 'bank_alm_system_ir_risk_ratio'];
        const _valid = _KEYS.filter((k) => indicators && indicators[k]);
        const _rows = [
            {title: '🟢 Chỉ báo CAO HƠN = AN TOÀN HƠN (càng cao càng tốt)', color: '#10b981',
             keys: _valid.filter((k) => indicators[k].goodDirection !== 'lower')},
            {title: '🔴 Chỉ báo CAO HƠN = RỦI RO HƠN (càng thấp càng tốt)', color: '#ef4444',
             keys: _valid.filter((k) => indicators[k].goodDirection === 'lower')},
        ];
        _rows.forEach((r) => {
            if (!r.keys.length) return;
            const rowEl = document.createElement('div');
            rowEl.innerHTML = `<div class="bank-chart-row-title" style="color:${r.color}">${r.title}</div>
                <div class="bank-chart-grid-2"></div>`;
            chartsGrid.appendChild(rowEl);
            const g = rowEl.querySelector('.bank-chart-grid-2');
            r.keys.forEach((key) => _renderBankingRiskChartCard(g, key, indicators[key]));
        });
    }
}

// Khoi DANH GIA trang thai toan he thong (user 2026-09-26): tinh trang chung + tung mang (thanh khoan /
// lai suat / co cau ky han) voi MOI Y 1 DONG RIENG + rui ro dang co + dieu can nho. Du lieu tu
// risk.assessment (bank_system_risk.build_system_assessment) - khong tinh lai o JS.
function _renderBankingAssessment(a) {
    const el = document.getElementById('banking-risk-assessment');
    if (!el) return;
    if (!a) { el.innerHTML = ''; return; }
    const COL = {0: '#10b981', 1: '#f59e0b', 2: '#ef4444'};
    const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const ov = a.overall || {};
    const c0 = COL[ov.level] || '#94a3b8';
    let html = `<div class="bank-assess-overall" style="background:${c0}22;color:${c0}">
        Tình trạng chung: ${esc(ov.label || '')} <div style="font-weight:500;color:var(--text-main,#f3f4f6);margin-top:2px;font-size:0.9em">${esc(ov.headline || '')}</div></div>`;
    html += '<div class="bank-assess-grid">' + (a.areas || []).map((ar) => {
        const c = COL[ar.level] || '#94a3b8';
        return `<div class="bank-assess-card">
            <h5><span>${ar.icon || ''} ${esc(ar.title)}</span><span class="lvl" style="background:${c}22;color:${c}">${esc(ar.label)}</span></h5>
            <ul>${(ar.points || []).map((p) => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
    }).join('') + '</div>';
    if ((a.risks || []).length) {
        html += `<div class="bank-assess-box"><h5>⚠️ Rủi ro đang có</h5><ul class="bank-assess-list">${a.risks.map((p) => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
    }
    if ((a.takeaways || []).length) {
        html += `<div class="bank-assess-box"><h5>📌 Điều cần nhớ về hệ thống ngân hàng (${esc(a.asOf || '')})</h5><ul class="bank-assess-list">${a.takeaways.map((p) => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
    }
    el.innerHTML = html;
}

// Chart chuyen dung cho muc "Rui ro he thong ngan hang" (khac _renderGenericIndicatorCard o cho
// dung mau CO DINH theo ban chat chi bao, khong phai theo xu huong tang/giam trong khung hien thi -
// xem ghi chu o renderBankingSystemRiskSection). Cung hien badge chieu rui ro + canh bao coverage
// thap NGAY TREN CARD (user 2026-09-23: doc chart mot minh de hieu nham "khong con rui ro" khi diem
// moi nhat thuc ra mau qua nho, chua dai dien toan he thong - truoc day chi co canh bao nay trong
// doan van rieng, de bi bo qua khi chi nhin chart).
function _renderBankingRiskChartCard(grid, key, ind) {
    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    const t = ind.trend || {};
    const valid = (ind.series || []).filter(p => p.value !== null && p.value !== undefined);
    const hasChart = valid.length >= 2;
    // SUA (user 2026-10-01, phat hien qua kiem tra duplicate canvas id khi them card FX): 4 key nay
    // co group="bank_alm" (nam trong GROUP_ORDER) nen DA render 1 lan trong renderIndicatorGroups()
    // generic, roi lai render THEM o day voi CUNG prefix "chart-" -> 2 <canvas> trung id trong DOM
    // (HTML khong hop le, co the gay loi ngoai y muon cho code khac dung getElementById theo id nay).
    // Doi prefix rieng, giong cach da sua cho card "Áp lực Ngoại tệ" (idPrefix 'fxchart-').
    const canvasId = `bankriskchart-${key}`;
    const isHigherBad = ind.goodDirection === 'lower';
    const dirColor = isHigherBad ? '#ef4444' : '#10b981';
    const dirBadge = isHigherBad ? '▲ Cao hơn = rủi ro cao hơn' : '▲ Cao hơn = an toàn hơn';

    const latestPoint = valid[valid.length - 1];
    const lowCovWarning = (latestPoint && latestPoint.coverage_pct != null && latestPoint.coverage_pct < 60)
        ? `<div class="ind-note" style="color:#f59e0b">⚠️ Điểm mới nhất chỉ đại diện ${latestPoint.coverage_pct.toFixed(0)}% tổng tài sản hệ thống — CHƯA đủ để kết luận xu hướng.</div>`
        : '';

    card.innerHTML = `
        <div class="ind-header">
            <span class="ind-name">${ind.label}</span>
            <span class="ind-judgment" style="background:${dirColor}22;color:${dirColor}">${dirBadge}</span>
        </div>
        <div class="ind-value">${t.latest !== null && t.latest !== undefined ? formatNumber(t.latest) : '-'} <span style="font-size:0.5em;color:var(--text-muted)">${ind.unit}</span></div>
        <div class="ind-meta">Kỳ: ${t.latest_period ? _periodToDisplayLabel(t.latest_period) : '—'} · Nguồn: ${SOURCE_LABELS[ind.autoSource] || ind.autoSource}</div>
        ${hasChart ? `<div class="ind-chart"><canvas id="${canvasId}"></canvas></div>` : ''}
        ${lowCovWarning}
        ${ind.impact ? `<div class="ind-note">${ind.impact}</div>` : ''}
        ${ind.note ? `<div class="ind-source-note">${ind.note}</div>` : ''}
    `;
    grid.appendChild(card);

    if (hasChart) {
        const ctx = card.querySelector(`#${canvasId}`);
        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: valid.map(p => _periodToDisplayLabel(p.period)),
                datasets: [{
                    data: valid.map(p => p.value), borderColor: dirColor,
                    backgroundColor: dirColor + '15', fill: true, tension: 0.25, pointRadius: 2,
                }],
            },
            options: CHART_DEFAULTS,
        });
        chartInstances.push(chart);
    }
}

// ═══════════════════════════════════════════════════════════
// CƠ CẤU TÍN DỤNG & HUY ĐỘNG TOÀN NGÀNH NGÂN HÀNG (user 2026-09-28) — tổng hợp trực tiếp từ BCTC
// 26 ngân hàng niêm yết/UPCoM (bank_system_risk.build_bank_credit_deposit_system_series, KHÔNG
// cần OCR như ALM gap kỳ hạn), đóng gói vào bankingSystemRisk.creditDepositStructure =
// {periods, nBanks, creditComposition:{loans,tpdn}, depositComposition:{customerDeposits,bonds,
// tctdDeposits,kbnnCounted}}. User yêu cầu ĐÚNG dạng "biểu đồ miền" (stacked area mượt, có chấm
// tròn trong legend) như ảnh mẫu gửi kèm — 1 cụm giá trị tuyệt đối (tổng cột TĂNG theo quy mô
// thật, không bó ở 100%) + 1 cụm theo % cơ cấu (0-100%, tính lại từ giá trị tuyệt đối phía trên).
// ═══════════════════════════════════════════════════════════
function renderCreditDepositStructure(cds) {
    const card = document.getElementById('credit-deposit-structure-card');
    if (!card) return;
    if (!cds || !cds.periods || !cds.periods.length) { card.style.display = 'none'; return; }
    card.style.display = '';

    const nMax = Math.max(...cds.nBanks);
    const nMin = Math.min(...cds.nBanks);
    document.getElementById('credit-deposit-structure-title').textContent =
        `📅 ${cds.periods[0]} — ${cds.periods[cds.periods.length - 1]} `
        + (nMin === nMax ? `(${nMax}/26 ngân hàng có dữ liệu mọi quý)`
                          : `(số ngân hàng có dữ liệu mỗi quý: ${nMin}–${nMax}/26)`);

    // THEM 'equity' (user 2026-09-30): VCSH la von KHONG co ky han (khong ai "rut" duoc nhu tien
    // gui) - tang lon nghia la co them 1 lop dem von ben vung, xem duoc quy mo VCSH bien dong ra
    // sao NGAY canh cac thanh phan huy dong khac. LUU Y: tu day chart nay la "Nguon von" (huy dong
    // + VCSH), KHAC "Tong huy dong (mau so LDR theo TT22/26)" dung o cac cho khac (LDR/GAP tin
    // dung-huy dong KHONG gom VCSH) - da sua tieu de + chu thich duoi chart de ro rang, tranh nham.
    const DEPOSIT_SERIES = [
        { key: 'customerDeposits', label: 'Tiền gửi khách hàng', color: '#10b981' },
        { key: 'bonds', label: 'Giấy tờ có giá phát hành', color: '#a78bfa' },
        { key: 'tctdDeposits', label: 'Tiền gửi TCTD khác', color: '#3b82f6' },
        { key: 'kbnnCounted', label: 'KBNN (tính theo TT26)', color: '#f59e0b' },
        { key: 'equity', label: 'Vốn chủ sở hữu (VCSH)', color: '#ef4444' },
    ];

    // SUA (user 2026-09-30): bỏ 2 biểu đồ "cơ cấu tín dụng" (Cho vay KH vs TPDN) — TPDN quá nhỏ so
    // Cho vay KH nên chart gần như vô nghĩa (thấy 1 màu). Thay bằng 2 biểu đồ Tổng tín dụng vs
    // Tổng huy động (+ 1 đường nét đứt LDR) — 1 bản huy động THƯỜNG (mẫu số LDR theo TT22/26),
    // 1 bản CỘNG THÊM VCSH (đúng yêu cầu "tính VCSH vào tổng huy động thôi").
    const totalDepositPlusEquity = cds.totalDeposit.map((v, i) => (v ?? 0) + (cds.totalEquity[i] ?? 0));
    const ldrWithEquity = cds.totalCredit.map((v, i) => totalDepositPlusEquity[i] ? (v / totalDepositPlusEquity[i] * 100) : null);
    _renderCreditFundingLdrChart('chart-credit-structure-abs', cds.periods, cds.totalCredit, cds.totalDeposit, cds.ldrSystem, 'Tổng huy động (theo TT22/26)');
    _renderCreditFundingLdrChart('chart-credit-structure-pct', cds.periods, cds.totalCredit, totalDepositPlusEquity, ldrWithEquity, 'Tổng huy động + VCSH');
    _renderAreaCompositionChart('chart-deposit-structure-abs', cds.periods, DEPOSIT_SERIES, cds.depositComposition, false);
    _renderAreaCompositionChart('chart-deposit-structure-pct', cds.periods, DEPOSIT_SERIES, cds.depositComposition, true);
    // THEM (user 2026-09-28): "vẽ thêm cái biểu đồ tăng trưởng tín dụng và tăng trưởng huy động
    // theo số liệu 26 bank" — 4 đường: tín dụng/huy động YoY (nét liền) + tín dụng/huy động YTD
    // (nét đứt, so cuối năm trước) trên CÙNG 1 chart — 2 khái niệm đã đối chiếu với số SBV công
    // bố ở các lượt trước (headline SBV thường trích YTD, không phải YoY), để cạnh nhau cho dễ so.
    // THEM depositGrowthYoyNarrow/YtdNarrow (user 2026-09-30): "Huy động" YoY/YTD 2 dong tren la
    // dinh nghia RONG (TT22/26 - gom ca tien gui TCTD khac + trai phieu, tang RAT NHANH nam 2025
    // nen keo sat tin dung) - khac han so "huy dong" HEP (chi tien gui KH) ma VBMA/bao chi hay
    // dung, gay hieu lam "gap khong khop" khi doi chieu 2 chart. Them 2 dong nay (mau cam, khop
    // mau "huy dong" tren chart VBMA quoc gia) de so TRUC TIEP ca 2 dinh nghia tren CUNG 1 chart.
    if (cds.creditGrowthYoy) {
        _renderGrowthComparisonChart('chart-bank-credit-deposit-growth', cds.periods, [
            { key: 'creditGrowthYoy', label: 'Tín dụng YoY', color: '#3b82f6', dash: false },
            { key: 'depositGrowthYoy', label: 'Huy động YoY (rộng — TT22/26)', color: '#10b981', dash: false },
            { key: 'depositGrowthYoyNarrow', label: 'Huy động YoY (hẹp — chỉ tiền gửi KH, khớp VBMA)', color: '#f59e0b', dash: false },
            { key: 'creditGrowthYtd', label: 'Tín dụng YTD (so cuối năm trước)', color: '#3b82f6', dash: true },
            { key: 'depositGrowthYtd', label: 'Huy động YTD (rộng — TT22/26)', color: '#10b981', dash: true },
            { key: 'depositGrowthYtdNarrow', label: 'Huy động YTD (hẹp — chỉ tiền gửi KH, khớp VBMA)', color: '#f59e0b', dash: true },
        ], cds);
    }
}

function _renderGrowthComparisonChart(canvasId, periods, seriesDefs, cds) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const datasets = seriesDefs.map(s => ({
        label: s.label, data: cds[s.key],
        borderColor: s.color, backgroundColor: s.color + '15', fill: false,
        borderDash: s.dash ? [6, 4] : [], borderWidth: s.dash ? 1.5 : 2.5,
        tension: 0.25, pointRadius: s.dash ? 0 : 3, pointBackgroundColor: s.color, spanGaps: true,
        // Chi hien nhan o DIEM CUOI (giong _endpointDatalabelsConfig da dung cho cac chart nhieu
        // duong khac) - 4 duong x 10 quy hien het se roi, gia tri chi tiet tung quy da co san trong
        // sheet Excel LDR_TongHop_HeThong cho nguoi can xem day du.
        datalabels: _endpointDatalabelsConfig(1),
    }));
    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels: periods, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 12, font: { size: 10 } } } },
            scales: {
                // SUA (user 2026-09-30, phat hien qua nham lan doc sai ky khi so voi chart VBMA
                // khac): maxTicksLimit ke thua tu CHART_DEFAULTS (=8) VAN gioi han so nhan hien du
                // autoSkip=false - voi 10 ky (2024-Q1..2026-Q2) bi RUT xuong chi con 4 nhan
                // (2024-Q1/Q3, 2025-Q3, 2026-Q1), khien dinh Q4 (dung ky "cuoi nam") KHONG co nhan
                // truc X rieng, de nham thanh dang nhin vao ky khac. Ghi de maxTicksLimit = so ky
                // thuc te de LUON hien DU ca 10 nhan, khong bi rut gon.
                // SUA (user 2026-09-30): maxRotation/autoSkip/maxTicksLimit la thuoc tinh cua
                // "ticks" (long BEN TRONG ticks:{...}), KHONG PHAI thuoc tinh cap scale nhu viet
                // truoc do - spread CHART_DEFAULTS.scales.x roi ghi de o CAP SCALE la NO-OP hoan
                // toan (ticks long ben trong van giu nguyen autoSkip:true, maxTicksLimit:8 cu), day
                // la nguyen nhan cac chart quy chi hien 5/10 nhan du code "tuong nhu" da doi.
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: '%', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM (user 2026-09-30): "biểu đồ cơ cấu nguồn vốn và cơ cấu tài sản theo các kỳ hạn ... để biết
// tài sản đang cơ cấu lệch về kỳ dài hạn hơn nguồn vốn nên áp lực rolling tiền gửi tăng" — 2 biểu
// đồ miền xếp lớp (TOÀN HỆ THỐNG) theo TỪNG bucket kỳ hạn, tái dùng _renderAreaCompositionChart
// (đã dùng cho cơ cấu huy động). Gộp 2 bucket "quá hạn" (qua_han_tren_3t/qua_han_den_3t, thường
// rất nhỏ) thành 1 "Quá hạn" cho gọn — KHÁC maturityStructure.assetsByBucket/liabilitiesByBucket ở
// vimo.json (giữ nguyên 7 bucket gốc, không gộp, để backend/Excel tra cứu chi tiết được).
function renderMaturityStructure(ms) {
    const card = document.getElementById('maturity-structure-card');
    if (!card) return;
    if (!ms || !ms.periods || !ms.periods.length) { card.style.display = 'none'; return; }
    card.style.display = '';

    const nMax = Math.max(...ms.nBanks);
    const nMin = Math.min(...ms.nBanks);
    document.getElementById('maturity-structure-title').textContent =
        `📅 ${ms.periods[0]} — ${ms.periods[ms.periods.length - 1]} `
        + (nMin === nMax ? `(${nMax}/26 ngân hàng có đủ dữ liệu thuyết minh kỳ hạn)`
                          : `(số ngân hàng có dữ liệu mỗi quý: ${nMin}–${nMax}/26 — phụ thuộc OCR thuyết minh, độ phủ thấp hơn hẳn các biểu đồ Tín dụng/Huy động khác)`);

    const MATURITY_SERIES = [
        { key: 'qua_han', label: 'Quá hạn', color: '#ef4444' },
        { key: 'den_1_thang', label: 'Đến 1 tháng', color: '#f59e0b' },
        { key: 'tu_1_3_thang', label: '1-3 tháng', color: '#eab308' },
        { key: 'tu_3_12_thang', label: '3-12 tháng', color: '#3b82f6' },
        { key: 'tu_1_5_nam', label: '1-5 năm', color: '#8b5cf6' },
        { key: 'tren_5_nam', label: '>5 năm', color: '#10b981' },
    ];
    const _mergeOverdue = (byBucket) => {
        const qh = ms.periods.map((_, i) => (byBucket['qua_han_tren_3t'][i] || 0) + (byBucket['qua_han_den_3t'][i] || 0));
        return { qua_han: qh, den_1_thang: byBucket['den_1_thang'], tu_1_3_thang: byBucket['tu_1_3_thang'],
                 tu_3_12_thang: byBucket['tu_3_12_thang'], tu_1_5_nam: byBucket['tu_1_5_nam'], tren_5_nam: byBucket['tren_5_nam'] };
    };
    const assetsData = _mergeOverdue(ms.assetsByBucket);
    const liabData = _mergeOverdue(ms.liabilitiesByBucket);
    _renderAreaCompositionChart('chart-maturity-assets-abs', ms.periods, MATURITY_SERIES, assetsData, false);
    _renderAreaCompositionChart('chart-maturity-assets-pct', ms.periods, MATURITY_SERIES, assetsData, true);
    _renderAreaCompositionChart('chart-maturity-liab-abs', ms.periods, MATURITY_SERIES, liabData, false);
    _renderAreaCompositionChart('chart-maturity-liab-pct', ms.periods, MATURITY_SERIES, liabData, true);

    // THEM (user 2026-09-30): "vẽ biểu đồ miền tỉ trọng 100%... theo 3 loại kỳ hạn: ngắn hạn (gồm
    // quá hạn + toàn bộ tiền kỳ hạn ≤12 tháng), trung hạn (1-5 năm), dài hạn (>5 năm)" — gộp tiếp
    // 6 bucket ở trên xuống còn 3 nhóm lớn, CHỈ vẽ bản % (100% stacked) vì đó là điều user yêu cầu
    // cụ thể — đọc trực tiếp lệch kỳ hạn tài sản/nguồn vốn ở mức tổng quan nhất, không cần soi 6
    // bucket chi tiết.
    const MATURITY_SERIES_3 = [
        { key: 'ngan_han', label: 'Ngắn hạn (≤12 tháng, gồm quá hạn)', color: '#ef4444' },
        { key: 'trung_han', label: 'Trung hạn (1-5 năm)', color: '#8b5cf6' },
        { key: 'dai_han', label: 'Dài hạn (>5 năm)', color: '#10b981' },
    ];
    const _mergeTo3 = (byBucket) => {
        const ngan = ms.periods.map((_, i) => (byBucket['qua_han_tren_3t'][i] || 0) + (byBucket['qua_han_den_3t'][i] || 0)
            + (byBucket['den_1_thang'][i] || 0) + (byBucket['tu_1_3_thang'][i] || 0) + (byBucket['tu_3_12_thang'][i] || 0));
        return { ngan_han: ngan, trung_han: byBucket['tu_1_5_nam'], dai_han: byBucket['tren_5_nam'] };
    };
    _renderAreaCompositionChart('chart-maturity-assets-3buckets-pct', ms.periods, MATURITY_SERIES_3, _mergeTo3(ms.assetsByBucket), true);
    _renderAreaCompositionChart('chart-maturity-liab-3buckets-pct', ms.periods, MATURITY_SERIES_3, _mergeTo3(ms.liabilitiesByBucket), true);
}

// THEM (user 2026-10-01, "card Áp lực Ngoại tệ" — khung Cầu/Cung/Đối chiếu BOP/Thị trường/Phản
// ứng NHNN, KHÔNG gộp thành 1 "FX stress score"): curate các chỉ báo macro ĐÃ CÓ SẴN (rải rác
// trong các nhóm "trade"/"external"/"monetary" chung) vào 1 card riêng theo đúng 5 lớp user yêu
// cầu — KHÔNG tính toán gì mới ở đây, chỉ tổ chức lại cách hiển thị những gì đã có + đánh dấu rõ
// phần CHƯA CÓ dữ liệu (NEER/REER, Errors & Omissions/Overall Balance, kiều hối/du lịch, FDI
// XNK tách riêng) để biết chính xác bổ sung vào ĐÚNG LỚP nào sau này, không phải thiết kế lại.
const FX_PRESSURE_LAYERS = [
    {
        id: 'demand', title: '① Cầu ngoại tệ (Potential USD Demand)',
        keys: ['import_growth_customs', 'import_growth_customs_mom'],
        missing: ['Dịch vụ nhập/trả lợi nhuận FDI/trả nợ nước ngoài — hiện chỉ có số THUẦN ở lớp "Đối chiếu BOP" bên dưới, chưa tách riêng chiều "chi/trả" (cần nguồn SBV/IMF chi tiết hơn)'],
    },
    {
        id: 'supply', title: '② Cung ngoại tệ (Potential USD Supply)',
        keys: ['export_growth_customs', 'export_growth_customs_mom', 'fdi_disbursed', 'fdi_registered_usd_bn', 'trade_balance'],
        missing: ['Kiều hối, doanh thu du lịch quốc tế — chưa có (nên lấy từ BOP/NSO theo đúng khuyến nghị, KHÔNG lấy số báo chí theo địa phương)'],
    },
    {
        id: 'bop', title: '③ Đối chiếu BOP (Current Account + Financial Account)',
        keys: ['bop_goods', 'bop_services', 'bop_primary_income', 'bop_secondary_income',
               'finacc_fdi_assets', 'finacc_fdi_liabilities', 'finacc_portfolio_assets', 'finacc_portfolio_liabilities',
               'finacc_other_assets', 'finacc_other_liabilities'],
        missing: ['Errors & Omissions, Overall Balance, Δ Dự trữ (reconciliation) — CHƯA CÓ, nguồn hiện tại (40yo.vn) không có 2 dòng này cho Việt Nam; cần tích hợp mới IMF SDMX/BOP hoặc trang BOP quý của SBV'],
    },
    {
        id: 'market', title: '④ Áp lực thị trường',
        keys: ['usdvnd', 'usdvnd_monthly_avg', 'usdvnd_growth_mom', 'usdvnd_growth_yoy',
               'usdvnd_vcb_sell_daily', 'usd_cho_den_sell_daily', 'usd_cho_den_vcb_gap', 'usd_cho_den_vcb_gap_pct',
               'interbank_rate_on', 'fed_funds_rate', 'vnd_usd_rate_spread_on'],
        missing: ['NEER/REER (BIS, theo tháng) — CHƯA CÓ, cần scraper mới cho CSV effective exchange rate của BIS'],
    },
    {
        id: 'response', title: '⑤ Phản ứng NHNN',
        keys: ['forex_reserves_monthly', 'forex_reserves_sdr', 'omo_rate_7d', 'tin_phieu_outstanding_balance', 'tin_phieu_net_operation'],
        missing: [],
    },
];

function renderFxPressureCard(indicators) {
    const card = document.getElementById('fx-pressure-card');
    const body = document.getElementById('fx-pressure-body');
    if (!card || !body) return;

    const layersWithData = FX_PRESSURE_LAYERS.map(layer => ({
        ...layer, validKeys: layer.keys.filter(k => indicators[k]),
    }));
    if (!layersWithData.some(l => l.validKeys.length)) { card.style.display = 'none'; return; }
    card.style.display = '';

    body.innerHTML = layersWithData.map(layer => {
        if (!layer.validKeys.length && !layer.missing.length) return '';
        const missingHtml = layer.missing.length
            ? `<p class="ind-source-note" style="margin:-4px 0 12px 4px">⏳ Chưa có dữ liệu: ${layer.missing.join(' · ')}</p>` : '';
        return `<div class="vimo-group-header"><h3>${layer.title}</h3></div>
            <div class="vimo-indicator-grid" id="fxgrid-${layer.id}"></div>${missingHtml}`;
    }).join('');

    layersWithData.forEach(layer => {
        const grid = document.getElementById(`fxgrid-${layer.id}`);
        if (!grid) return;
        layer.validKeys.forEach(k => _renderGenericIndicatorCard(grid, k, indicators[k], 'fxchart-'));
    });
}

// Tổng tín dụng vs Tổng huy động (2 miền, KHÔNG xếp lớp — 2 đại lượng độc lập so cạnh nhau, không
// phải 2 thành phần cộng thành 1 tổng) + 1 đường nét đứt thể hiện LDR (hoặc tỷ lệ tương đương khi
// mẫu số có cộng thêm VCSH) — trục phải riêng cho %. Thay cho biểu đồ "cơ cấu tín dụng" cũ (Cho
// vay KH vs TPDN, user 2026-09-30 chỉ ra TPDN quá nhỏ nên chart gần như vô nghĩa).
function _renderCreditFundingLdrChart(canvasId, periods, creditArr, fundingArr, ratioArr, fundingLabel) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const levelDatalabels = (color) => ({
        display: (ctx) => ctx.dataset.data.slice(ctx.dataIndex + 1).every(v => v === null || v === undefined)
            && (ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] !== undefined),
        color, font: { size: 9, weight: '700' }, anchor: 'end', align: 'right', offset: 4, clip: false,
        formatter: (v) => v.toLocaleString('vi-VN', { maximumFractionDigits: 0 }),
    });
    const datasets = [
        {
            label: 'Tổng tín dụng', data: creditArr, yAxisID: 'y',
            borderColor: '#3b82f6', backgroundColor: '#3b82f620', fill: true, tension: 0.3,
            pointRadius: 3, pointBackgroundColor: '#3b82f6', borderWidth: 2, spanGaps: true,
            datalabels: levelDatalabels('#3b82f6'),
        },
        {
            label: fundingLabel, data: fundingArr, yAxisID: 'y',
            borderColor: '#10b981', backgroundColor: '#10b98120', fill: true, tension: 0.3,
            pointRadius: 3, pointBackgroundColor: '#10b981', borderWidth: 2, spanGaps: true,
            datalabels: levelDatalabels('#10b981'),
        },
        {
            label: 'LDR (%)', data: ratioArr, yAxisID: 'y1',
            borderColor: '#f59e0b', borderDash: [6, 4], borderWidth: 2, pointRadius: 0, fill: false,
            tension: 0.25, spanGaps: true, datalabels: _endpointDatalabelsConfig(1),
        },
    ];
    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels: periods, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 10 } } } },
            scales: {
                // SUA (user 2026-09-30): maxRotation/autoSkip/maxTicksLimit la thuoc tinh cua
                // "ticks" (long BEN TRONG ticks:{...}), KHONG PHAI thuoc tinh cap scale nhu viet
                // truoc do - spread CHART_DEFAULTS.scales.x roi ghi de o CAP SCALE la NO-OP hoan
                // toan (ticks long ben trong van giu nguyen autoSkip:true, maxTicksLimit:8 cu), day
                // la nguyen nhan cac chart quy chi hien 5/10 nhan du code "tuong nhu" da doi.
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     title: { display: true, text: 'Tỷ đồng', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: '%', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

function _renderAreaCompositionChart(canvasId, periods, seriesDefs, compositionData, pctMode) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const totals = periods.map((_, i) => seriesDefs.reduce((sum, s) => sum + (compositionData[s.key][i] || 0), 0));
    const dataFor = (key) => pctMode
        ? periods.map((_, i) => totals[i] ? (compositionData[key][i] / totals[i] * 100) : null)
        : compositionData[key];

    // SUA (user 2026-09-28: "tôi nhìn chả biết cái nào tăng giảm" — biểu đồ miền xếp lớp trước đó
    // KHÔNG hiện số nào ngoài tổng, không đọc được từng thành phần đang tăng/giảm ra sao) — hiện
    // GIÁ TRỊ RIÊNG của TỪNG lớp tại MỌI điểm, đặt giữa đúng dải màu của lớp đó (anchor/align
    // 'center' — vị trí tự nhiên nhất cho biểu đồ miền xếp lớp, khác biểu đồ cột/đường thường).
    const datasets = seriesDefs.map(s => ({
        label: s.label, data: dataFor(s.key),
        borderColor: s.color, backgroundColor: s.color + '70', fill: true,
        tension: 0.3, pointRadius: 3, pointBackgroundColor: s.color, borderWidth: 2, spanGaps: true,
        datalabels: {
            // An nhan khi lop QUA MONG (< 2% tong ky do, vd TPDN gan 0 giai doan dau) - nhan cua 1
            // lop mong se choang len nhan lop ben canh (da thay qua screenshot thuc te), thay vi co
            // nghia ("0" lap lai nhieu lan). Van hien du trong hinh dang mien + chu giai, chi bo
            // qua so cho diem qua nho khong doc duoc.
            display: (ctx) => {
                const v = ctx.dataset.data[ctx.dataIndex];
                if (v === null || v === undefined) return false;
                // pctMode: v DA la % (0-100) nen ty trong = v/100; abs mode: ty trong = v/tong tuyet doi.
                const t = totals[ctx.dataIndex] || 0;
                const share = pctMode ? (v / 100) : (t ? v / t : 0);
                return share >= 0.02;
            },
            anchor: 'center', align: 'center', color: '#f8fafc', font: { size: 8, weight: '700' },
            formatter: (v) => pctMode ? `${v.toFixed(1)}%` : Math.round(v).toLocaleString('vi-VN'),
        },
    }));
    // Nhãn TỔNG THÊM (chỉ ở cụm giá trị tuyệt đối, không cần ở % vì luôn ~100%) trên đỉnh lớp cuối
    // cùng — đúng yêu cầu "có sự tăng về giá trị tổng" nhìn thấy được ngay, không phải suy ra từ
    // mắt. chartjs-plugin-datalabels cho phép NHIỀU nhãn/dataset qua khoá con "labels" (mỗi khoá
    // là 1 nhãn độc lập, merge với config gốc) — dùng đúng API này thay vì gán chồng đè 1 object.
    if (!pctMode) {
        const lastDataset = datasets[datasets.length - 1];
        const valueCfg = lastDataset.datalabels;
        lastDataset.datalabels = {
            labels: {
                value: valueCfg,
                total: {
                    anchor: 'end', align: 'top', color: '#e5e9f0', font: { size: 9, weight: '700' },
                    formatter: (v, ctx) => {
                        const total = ctx.chart.data.datasets.reduce((sum, d) => sum + (d.data[ctx.dataIndex] || 0), 0);
                        return 'Tổng ' + total.toLocaleString('vi-VN', { maximumFractionDigits: 0 });
                    },
                },
            },
        };
    }

    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels: periods, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: true, position: 'top',
                          labels: { boxWidth: 8, usePointStyle: true, pointStyle: 'circle', font: { size: 10 } } },
            },
            scales: {
                // SUA (user 2026-09-30): maxRotation/autoSkip/maxTicksLimit la thuoc tinh cua
                // "ticks" (long BEN TRONG ticks:{...}), KHONG PHAI thuoc tinh cap scale nhu viet
                // truoc do - spread CHART_DEFAULTS.scales.x roi ghi de o CAP SCALE la NO-OP hoan
                // toan (ticks long ben trong van giu nguyen autoSkip:true, maxTicksLimit:8 cu), day
                // la nguyen nhan cac chart quy chi hien 5/10 nhan du code "tuong nhu" da doi.
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, min: 0, ...(pctMode ? { max: 100 } : {}),
                     title: { display: true, text: pctMode ? '%' : 'Tỷ đồng', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// Đánh giá Tổng thể — 3 câu hỏi user luôn quan tâm: đang tốt lên/xấu đi (xu hướng so kỳ trước),
// bức tranh rõ ràng hay xám/hỗn hợp (mức đồng thuận giữa các chỉ báo), có phù hợp đầu tư không.
function renderVerdict(verdict, decision) {
    if (!verdict) return;
    const clarityEl = document.getElementById('verdict-clarity');
    const decisionEl = document.getElementById('verdict-decision');
    const detailEl = document.getElementById('verdict-detail');

    // 2 dòng ĐỘC LẬP so tuần trước / so tháng trước (user 2026-08-01) — mỗi dòng tự tính
    // tốt lên/xấu đi riêng theo đúng mốc của nó, xem trend_week/trend_month trong
    // calc_overall_verdict() (template_vimo.py).
    const _trendColor = (arrow) => arrow === '▲' ? '#10b981' : (arrow === '▼' ? '#ef4444' : '#f59e0b');
    [['verdict-trend-week', verdict.trend_week], ['verdict-trend-month', verdict.trend_month]].forEach(([id, t]) => {
        const el = document.getElementById(id);
        if (!el || !t) return;
        el.textContent = `${t.arrow} ${t.label}`;
        el.style.color = _trendColor(t.arrow);
    });

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
            ${g.reason ? `<div style="font-size:0.68em;color:var(--text-muted);margin-top:4px;line-height:1.4;text-align:left">${String(g.reason).replace(/;\s+/g, ';<br>')}</div>` : ''}
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

// Bảng giám sát chỉ số vĩ mô hàng tháng (heatmap, user 2026-08-03) — table dựng sẵn trong
// _build_monitoring_table() (template_vimo.py): mỗi hàng 1 chỉ báo THEO THÁNG, cột là N tháng gần
// nhất. Màu ô = thang RIÊNG của từng hàng (min-max trong cửa sổ hiển thị, không so giữa các hàng
// khác nhau vì đơn vị/độ lớn khác nhau) — đảo chiều theo goodDirection để "xanh" LUÔN nghĩa là
// tốt hơn dù chỉ báo higher-is-good hay lower-is-good.
function _periodToShortLabel(period) {
    const m = /^(\d{4})-(\d{2})$/.exec(period);
    if (!m) return period;
    return `T${+m[2]}-${m[1].slice(2)}`;
}

function renderMonitoringTable(table) {
    const card = document.getElementById('monitoring-table-card');
    if (!table || !table.rows || !table.rows.length) return;
    card.style.display = 'block';

    const el = document.getElementById('monitoring-table');
    const periods = table.periods;
    const thead = `<thead><tr><th>Chỉ báo</th>${periods.map(p => `<th>${_periodToShortLabel(p)}</th>`).join('')}</tr></thead>`;

    const tbody = table.rows.map(row => {
        // colorMin/colorMax = min-max của TOÀN BỘ LỊCH SỬ chỉ báo (tính sẵn trong
        // _build_monitoring_table, template_vimo.py) — KHÔNG dùng min-max của riêng N tháng đang
        // hiển thị, tránh bóp méo màu khi cửa sổ hiển thị vô tình chỉ toàn giá trị đã cao/thấp
        // sẵn (user 2026-08-03: tín dụng 18,23% vẫn cao so lịch sử nhưng bị tô đỏ vì cửa sổ hiện
        // tại chỉ có 18-22%).
        const lo = row.colorMin, hi = row.colorMax;
        const alt = new Set(row.altSourceIdx || []);
        const cells = row.values.map((v, ci) => {
            if (v === null || v === undefined) return `<td class="na">—</td>`;
            let g = hi === lo ? 0.5 : (v - lo) / (hi - lo);
            if (row.goodDirection === 'lower') g = 1 - g;
            const bg = _heatmapColor(g);
            const star = alt.has(ci) ? '<sup title="Nguồn phụ: Tổng cục Thống kê (số Hải quan chưa công bố)">*</sup>' : '';
            return `<td style="background:${bg}">${formatNumber(v)}${row.unit === '%' ? '%' : ''}${star}</td>`;
        }).join('');
        return `<tr><th title="${row.key}">${row.label}</th>${cells}</tr>`;
    }).join('');

    el.innerHTML = thead + `<tbody>${tbody}</tbody>`;
    const anyAlt = table.rows.some(r => (r.altSourceIdx || []).length);
    let noteEl = document.getElementById('monitoring-table-altnote');
    if (anyAlt) {
        if (!noteEl) { noteEl = document.createElement('p'); noteEl.id = 'monitoring-table-altnote'; noteEl.className = 'ind-source-note'; el.insertAdjacentElement('afterend', noteEl); }
        noteEl.textContent = '* Ô có dấu * lấy từ nguồn phụ (Tổng cục Thống kê) vì số liệu Hải quan tháng đó chưa được công bố; các ô "—" còn lại là kỳ nguồn chưa công bố hoặc nguồn không có số liệu.';
    } else if (noteEl) { noteEl.remove(); }
}

function _heatmapColor(g) {
    // g=0 -> đỏ (#ef4444), g=1 -> xanh (#10b981) — lerp RGB tuyến tính. Alpha: 0.55 (gốc) -> 0.85
    // (2026-08-07) -> 1.0 (2026-08-08, user vẫn thấy tối trên điện thoại) — ĐỦ 1.0 nghĩa là màu
    // nền ô = ĐÚNG màu đỏ/xanh gốc, không còn bị pha loãng bởi nền thẻ tối phía sau nữa (alpha <1
    // luôn bị nền tối đằng sau ăn bớt độ sáng, thấy rõ hơn trên màn hình điện thoại độ sáng/tương
    // phản thấp hơn desktop). Chữ trong ô vẫn cố định màu TỐI (color:#0b1220 trong CSS) nên nền
    // càng đặc màu càng dễ đọc, không có lý do giữ alpha <1 nữa.
    const red = [239, 68, 68], green = [16, 185, 129];
    const r = Math.round(red[0] + (green[0] - red[0]) * g);
    const gr = Math.round(red[1] + (green[1] - red[1]) * g);
    const b = Math.round(red[2] + (green[2] - red[2]) * g);
    return `rgba(${r},${gr},${b},1)`;
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
// SUA (user 2026-09-28): decisionExvin/decisionHeadline ở đây là 2 khuyến nghị ĐỨNG RIÊNG (mỗi góc
// nhìn tính độc lập, KHÔNG kết hợp) — chỉ để so sánh/minh hoạ chênh lệch. Quyết định CHÍNH hiển thị
// ở "Đánh giá Tổng thể" (data.decision) đã KẾT HỢP cả 2 (mua tỷ trọng cao cần CẢ 2 cùng xác nhận
// rẻ, chỉ cần 1 trong 2 báo đắt là đủ để cảnh báo bán/giảm tỷ trọng) — truyền thêm decisionCombined
// để ghi rõ trong cảnh báo lệch, tránh gây hiểu nhầm với 2 khuyến nghị đứng riêng trong bảng.
// ═══════════════════════════════════════════════════════════
function renderVnindexCompare(valExvin, valHeadline, decisionExvin, decisionHeadline, decisionCombined) {
    const card = document.getElementById('vnindex-compare-card');
    if (!valHeadline || !decisionHeadline || !decisionExvin) { card.style.display = 'none'; return; }
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
                <th style="padding:6px 8px">Khuyến nghị (đứng riêng)</th>
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
        warnEl.textContent = `⚠ 2 góc nhìn ĐỨNG RIÊNG cho khuyến nghị KHÁC NHAU — VIN (VIC/VHM/VRE/VPL) đang làm lệch kết luận định giá chung của thị trường một cách đáng kể. Khuyến nghị CHÍNH (kết hợp cả 2, xem mục "Đánh giá Tổng thể" phía trên)${decisionCombined ? `: ${decisionCombined.label}` : ''} — mua tỷ trọng cao/mạnh cần CẢ 2 góc nhìn cùng xác nhận rẻ, chỉ cần 1 trong 2 báo đắt là đủ để cảnh báo giảm tỷ trọng.`;
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
                x: { ticks: { color: '#9aa5bd', font: { size: 9 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 }, grid: { display: false } },
                y: { ticks: { color: '#9aa5bd', font: { size: 9 } }, grid: { color: 'rgba(255,255,255,0.04)' } },
            },
        },
    });
}

// ═══════════════════════════════════════════════════════════
// INDICATOR GROUPS
// ═══════════════════════════════════════════════════════════
// Render 1 card chỉ báo GENERIC (giá trị mới nhất + đánh giá tốt/xấu + sparkline nếu ≥4 điểm) —
// dùng chung cho cả nhóm chỉ báo Việt Nam (renderIndicatorGroups) VÀ cụm Dữ liệu Quốc tế
// (renderInternationalSection) để không lặp lại logic.
function _renderGenericIndicatorCard(grid, key, ind, idPrefix = 'chart-') {
    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    const t = ind.trend || {};
    const nValid = (ind.series || []).filter(p => p.value !== null && p.value !== undefined).length;
    // Hạ ngưỡng 4 -> 2 (user 2026-08-08: muốn "chart hóa" Lãi suất iPower TCBS/Tăng trưởng huy
    // động vốn dù mới có 1-2 điểm) — 2 điểm đã đủ để so sánh xu hướng tăng/giảm giữa 2 lần cập
    // nhật gần nhất, không cần chờ tích lũy đủ 4 điểm mới hiện.
    const hasChart = nValid >= 2;
    const judgColor = t.judgment_color || '#94a3b8';
    // idPrefix (user 2026-10-01, card "Áp lực Ngoại tệ"): chỉ báo như usdvnd/fed_funds_rate/
    // trade_balance ĐÃ render 1 lần trong renderIndicatorGroups (nhóm "external"/"trade" riêng) —
    // card mới tái dùng CÙNG indicator nhưng CẦN canvas id KHÁC, không thì 2 <canvas> trùng id
    // trong DOM (Chart.js/getElementById chỉ thấy cái đầu tiên, cái sau vẽ lên canvas rỗng/lỗi).
    const canvasId = `${idPrefix}${key}`;

    // Chuỗi nhiều điểm (vd lạm phát cơ bản backfill từ 2020 = ~78 điểm) khiến nhãn trục X xoay
    // 45° bị chật/tràn ra ngoài card 320px mặc định (user 2026-08-01: "biểu đồ hẹp quá... rộng
    // ngang thêm") — card nào ≥20 điểm được nới rộng gấp đôi (span 2 cột) để có chỗ cho nhãn.
    if (nValid >= 20) card.style.gridColumn = 'span 2';

    card.innerHTML = `
        <div class="ind-header">
            <span class="ind-name">${ind.label}</span>
            ${t.judgment_label ? `<span class="ind-judgment" style="background:${judgColor}22;color:${judgColor}">${t.value_arrow || ''} ${t.judgment_label}</span>` : ''}
        </div>
        <div class="ind-value">${t.latest !== null && t.latest !== undefined ? formatNumber(t.latest) : '-'} <span style="font-size:0.5em;color:var(--text-muted)">${ind.unit}</span></div>
        <div class="ind-meta">Kỳ: ${t.latest_period ? _periodToDisplayLabel(t.latest_period) : '—'} · Nguồn: ${SOURCE_LABELS[ind.autoSource] || ind.autoSource}</div>
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
                labels: valid.map(p => _periodToDisplayLabel(p.period)),
                datasets: [{
                    data: valid.map(p => p.value), borderColor: color,
                    backgroundColor: color + '15', fill: true, tension: 0.25, pointRadius: 2,
                }],
            },
            options: CHART_DEFAULTS,
        });
        chartInstances.push(chart);
    }
}

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

        entries.forEach(([key, ind]) => _renderGenericIndicatorCard(grid, key, ind));

        if (grp === 'monetary') {
            renderInterbankCurveChart(grid, indicators);
            renderInterbank6mHistoryChart(grid, indicators);
            renderBondYieldHistoryChart(grid, indicators);
            renderOmoHistoryChart(grid, indicators);
            renderTinPhieuHistoryChart(grid, indicators);
            // Tăng trưởng tín dụng, cung tiền M2 & huy động cùng 1 chart (user 2026-08-01/08-03,
            // đối chiếu ảnh tham khảo từ vbma.org.vn/vi/market-data/money-supply +
            // data.vietnambiz.vn/currency-interest-rate) — credit_growth_yoy_monthly/
            // deposit_growth_yoy_monthly (phái sinh, xem _add_credit_derived_indicators) + m2_growth
            // (đã có sẵn, VBMA) đều CÙNG PHƯƠNG PHÁP YoY thật nên so sánh trực tiếp được.
            renderMultiTenorHistoryChart(grid, indicators, [
                ['credit_growth_yoy_monthly', 'Tăng trưởng tín dụng (YoY)', '#3b82f6'],
                ['m2_growth', 'Tăng trưởng cung tiền M2 (YoY)', '#a78bfa'],
                ['deposit_growth_yoy_monthly', 'Tăng trưởng huy động (YoY)', '#f59e0b'],
            ], 'chart-credit-money-supply', '📈 Tăng trưởng tín dụng, cung tiền M2 & huy động theo tháng (so cùng kỳ năm trước)',
            'Nguồn: vbma.org.vn (cung tiền M2 trực tiếp; tín dụng/huy động phái sinh từ dư nợ/tiền gửi tuyệt đối) — cả 3 đều YoY thật theo tháng, không phải so với đầu năm.', null);

            // Bản "so cuối năm trước" (YTD, reset mỗi tháng 1) của CÙNG 3 chỉ báo trên — user
            // (2026-08-08): muốn xem diễn biến TRONG NĂM rõ hơn, vì có giai đoạn cùng kỳ năm
            // trước tăng mạnh khiến YoY hiện tại trông thấp đi không phản ánh đúng xu hướng năm
            // nay. Cả 3 đều phái sinh từ mức tuyệt đối (credit/deposit/m2_balance_total), xem
            // _ytd_from_level_series() trong template_vimo.py.
            renderMultiTenorHistoryChart(grid, indicators, [
                ['credit_growth_ytd_monthly', 'Tăng trưởng tín dụng (YTD)', '#3b82f6'],
                ['m2_growth_ytd_monthly', 'Tăng trưởng cung tiền M2 (YTD)', '#a78bfa'],
                ['deposit_growth_ytd_monthly', 'Tăng trưởng huy động (YTD)', '#f59e0b'],
            ], 'chart-credit-money-supply-ytd', '📈 Tăng trưởng tín dụng, cung tiền M2 & huy động theo tháng (so cuối năm trước)',
            'Nguồn: vbma.org.vn (phái sinh từ mức tuyệt đối tín dụng/M2/huy động) — cả 3 đều so với mốc 31/12 năm trước, RESET về gần 0% mỗi tháng 1 rồi cộng dồn tới tháng 12, không phải so cùng kỳ.', null);
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
            renderGdpUseContributionChart(grid, indicators);
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
            renderTradeBalanceMonthlyChart(grid, indicators);
        }
    });
}

// So sánh 2 chuỗi period THEO THỜI GIAN THẬT — quy hết về 1 mốc thời gian (ms) để so sánh được
// GIỮA NHIỀU ĐỊNH DẠNG KHÁC NHAU trộn chung 1 trục (vd chart lãi suất liên ngân hàng: ON/1W/2W/1M
// dạng NGÀY "YYYY-MM-DD" từ VIRA + 6M dạng TUẦN "YYYY-Www" từ SBV — user 2026-07-30 phát hiện
// "2026-W30"/"2026-W31" (thực ra RƠI VÀO GIỮA khoảng ngày đã có, ~20/7-2/8) bị chuỗi sort mặc định
// đẩy ra CUỐI trục vì 'W' > các chữ số, tạo khoảng trống giả ở bên phải). Cũng xử lý luôn Q1/H1/
// 9M/FY (sort chuỗi mặc định ra "H1" < "Q1" < "FY" < "9M" theo abc, sai thứ tự thời gian thật).
function _periodSortKey(period) {
    let m;
    if ((m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(period))) {
        return Date.UTC(+m[1], +m[2] - 1, +m[3]);
    }
    if ((m = /^(\d{4})-W(\d{2})$/.exec(period))) {
        // Monday cua ISO week do (thuat toan chuan: 4/1 luon nam trong tuan 1)
        const jan4 = Date.UTC(+m[1], 0, 4);
        const jan4Dow = (new Date(jan4).getUTCDay() + 6) % 7; // Mon=0..Sun=6
        const week1Monday = jan4 - jan4Dow * 86400000;
        return week1Monday + (+m[2] - 1) * 7 * 86400000;
    }
    if ((m = /^(\d{4})-(Q1|H1|9M|FY)$/.exec(period))) {
        const monthEnd = { Q1: 2, H1: 5, "9M": 8, FY: 11 }[m[2]]; // thang cuoi ky (0-based)
        return Date.UTC(+m[1], monthEnd + 1, 0); // ngay cuoi thang do
    }
    if ((m = /^(\d{4})-(\d{2})$/.exec(period))) {
        return Date.UTC(+m[1], +m[2] - 1, 1);
    }
    if ((m = /^(\d{4})$/.exec(period))) {
        return Date.UTC(+m[1], 0, 1);
    }
    return NaN; // dinh dang la -> giu nguyen vi tri gap duoc (Array.sort coi NaN so sanh khong on dinh, chap nhan)
}
function _sortPeriods(periods) {
    return [...periods].sort((a, b) => _periodSortKey(a) - _periodSortKey(b));
}

// Chuyển period dạng TUẦN "YYYY-Wnn" thành NGÀY THẬT (Chủ nhật — ngày cuối ISO week đó) dạng
// "YYYY-MM-DD" — user (2026-08-09): "dữ liệu biểu đồ đang thể hiện W, hãy thể hiện rõ ngày tháng
// nào, nếu là tuần thì lấy giá trị ngày cuối cùng của tuần". Các định dạng khác (tháng/quý/ngày)
// giữ nguyên. Dùng CHUNG thuật toán tính Thứ Hai đầu tuần với _periodSortKey() ở trên (4/1 luôn
// nằm trong tuần 1 theo chuẩn ISO) rồi cộng thêm 6 ngày ra Chủ nhật.
function _periodToDisplayLabel(period) {
    const m = /^(\d{4})-W(\d{2})$/.exec(period);
    if (!m) return period;
    const jan4 = Date.UTC(+m[1], 0, 4);
    const jan4Dow = (new Date(jan4).getUTCDay() + 6) % 7;
    const weekMonday = jan4 - jan4Dow * 86400000 + (+m[2] - 1) * 7 * 86400000;
    const d = new Date(weekMonday + 6 * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
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
            plugins: {
                legend: { display: true, labels: { boxWidth: 12 } },
                datalabels: {
                    display: true, color: '#e5e9f0', font: { size: 8, weight: '600' },
                    anchor: 'center', align: 'center',
                    formatter: (v) => (v === null || v === undefined) ? '' : v.toFixed(1),
                },
            },
            scales: { ...CHART_DEFAULTS.scales, y: { ...CHART_DEFAULTS.scales.y, stacked: true, min: 0, max: 100 } },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// Đóng góp (ĐIỂM %) của từng cấu phần SỬ DỤNG (đầu tư/tiêu dùng tư nhân/tiêu dùng chính phủ) vào
// tăng trưởng GDP chung theo NĂM — user (2026-08-09) hỏi "tạo biểu đồ tỷ trọng GDP theo cấu phần
// sử dụng". KHÔNG dùng renderStackedAreaChart() ở trên (giả định cố định trục Y 0-100%, hợp cho
// %GDP theo khu vực/thành phần vốn đầu tư — 2 chart đã có) vì đây là ĐIỂM % ĐÓNG GÓP VÀO TĂNG
// TRƯỞNG (3 cấu phần cộng lại ≈ GDP Growth cùng năm, thường chỉ ~5-10, không phải 100) — dùng
// trục Y tự co giãn thay vì cố định 0-100 để không bị dồn cụm đáy biểu đồ.
function renderGdpUseContributionChart(grid, indicators) {
    const keys = [
        ['gdp_use_contrib_investment', 'Đầu tư', '#3b82f6'],
        ['gdp_use_contrib_private_consumption', 'Tiêu dùng tư nhân', '#10b981'],
        ['gdp_use_contrib_public_consumption', 'Tiêu dùng chính phủ', '#f59e0b'],
    ];
    const seriesByKey = keys.map(([key, label, color]) => [
        label, color, ((indicators[key] || {}).series || []).filter(p => p.value !== null && p.value !== undefined),
    ]);
    const allPeriods = _sortPeriods(new Set(seriesByKey.flatMap(([, , s]) => s.map(p => p.period))));
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">🗺️ Đóng góp vào tăng trưởng GDP theo cấu phần sử dụng (điểm %, theo năm)</span></div>
        <div class="ind-chart" style="height:260px"><canvas id="chart-gdp-use-contribution"></canvas></div>
        <div class="ind-note">Nguồn: aric.adb.org (ADB/CEIC), theo NĂM (trễ ~1 năm). 3 cấu phần CỘNG LẠI xấp xỉ bằng GDP Growth cùng năm (phần chênh nhỏ là xuất khẩu ròng/tồn kho, ARIC không tách riêng) — ĐÂY LÀ ĐIỂM % ĐÓNG GÓP VÀO TĂNG TRƯỞNG, KHÔNG PHẢI %GDP tuyệt đối (Việt Nam không công bố %GDP theo cấu phần sử dụng).</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-gdp-use-contribution');
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
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true,
                     title: { display: true, text: 'Điểm % đóng góp', color: '#9aa5bd', font: { size: 9 } } },
            },
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
        <div class="ind-chart" style="height:320px"><canvas id="chart-interbank-curve"></canvas></div>
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
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: true, color: '#c9d2e3', font: { size: 9, weight: '600' },
                    anchor: 'end', align: 'top',
                    formatter: (v) => (v === null || v === undefined) ? '' : v.toFixed(2),
                },
            },
        },
        plugins: [ChartDataLabels],
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
    const allPeriods = _sortPeriods(new Set(seriesByTenor.flatMap(([, , s]) => s.map(p => p.period))));
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Lãi suất liên ngân hàng O/N, 1 tuần, 2 tuần, 1 tháng, 6 tháng theo thời gian</span></div>
        <div class="ind-chart" style="height:460px"><canvas id="chart-interbank-6m-history"></canvas></div>
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
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: true, labels: { boxWidth: 12 } },
                datalabels: _endpointDatalabelsConfig(2),
            },
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 10 } } },
        },
        plugins: [ChartDataLabels],
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
    const allPeriods = _sortPeriods(new Set(seriesByTenor.flatMap(([, , s]) => s.map(p => p.period))));
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Lợi suất TPCP thứ cấp 3-5-7-10-15 năm theo thời gian</span></div>
        <div class="ind-chart" style="height:320px"><canvas id="chart-bond-yield-history"></canvas></div>
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
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: true, labels: { boxWidth: 12 } },
                datalabels: _endpointDatalabelsConfig(2),
            },
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 10 } } },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// Số dư OMO đang lưu hành (TỒN KHO, kênh cầm cố) + bơm/hút ròng OMO (DÒNG CHẢY/ngày) THEO THỜI
// GIAN, cùng 1 chart — user (2026-08-01): "cái dữ liệu này quan trọng mà" + muốn xem được vài
// tháng gần đây thay vì chỉ vài ngày (xem backfill lịch sử VIRA ~120 ngày trong
// fetch_vira_bulletin()). 2 chỉ báo lệch quy mô rất nhiều (tồn kho ~hàng trăm nghìn tỷ, dòng chảy
// ~vài nghìn tỷ/ngày) nên dùng 2 TRỤC Y riêng (line tồn kho bên trái, bar dòng chảy bên phải, màu
// theo dấu bơm/hút) — kiểu trình bày "stock vs flow" kinh điển, không cần thêm plugin ngoài
// datalabels đã có.
function renderOmoHistoryChart(grid, indicators) {
    const outstanding = ((indicators.omo_outstanding_balance || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const net = ((indicators.omo_net_operation || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const allPeriods = _sortPeriods(new Set([...outstanding, ...net].map(p => p.period)));
    if (!allPeriods.length) return;

    const outstandingByPeriod = Object.fromEntries(outstanding.map(p => [p.period, p.value]));
    const netByPeriod = Object.fromEntries(net.map(p => [p.period, p.value]));

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Số dư OMO đang lưu hành (kênh cầm cố) & Bơm/hút ròng OMO theo thời gian</span></div>
        <div class="ind-chart" style="height:320px"><canvas id="chart-omo-history"></canvas></div>
        <div class="ind-note">Nguồn: vira.org.vn (bản tin Kinh tế - Tài chính ngày, tự động, chuỗi theo NGÀY thật). Đường (trục trái) = tồn kho lưu hành; cột (trục phải) = dòng chảy ròng/ngày (xanh = bơm ròng, đỏ = hút ròng).</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-omo-history');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: [
                {
                    type: 'bar', label: 'Bơm/hút ròng OMO (phải)', yAxisID: 'y1',
                    data: allPeriods.map(p => netByPeriod[p] ?? null),
                    backgroundColor: allPeriods.map(p => (netByPeriod[p] ?? 0) >= 0 ? '#10b98188' : '#ef444488'),
                    borderWidth: 0, order: 2,
                    datalabels: { ..._endpointDatalabelsConfig(0), color: '#c9d2e3' },
                },
                {
                    type: 'line', label: 'Số dư lưu hành (trái)', yAxisID: 'y',
                    data: allPeriods.map(p => outstandingByPeriod[p] ?? null),
                    borderColor: '#3b82f6', backgroundColor: '#3b82f615', fill: true,
                    tension: 0.25, pointRadius: 2, spanGaps: true, order: 1,
                    datalabels: _endpointDatalabelsConfig(0),
                },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: true, labels: { boxWidth: 12 } },
            },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 10 } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'Tồn kho (tỷ đồng)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'Bơm/hút ròng (tỷ đồng/ngày)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// Kim ngạch xuất khẩu + nhập khẩu (2 đường, cùng đơn vị tỷ USD) và cán cân thương mại THEO TỪNG
// THÁNG RIÊNG LẺ (cột, xanh = thặng dư, đỏ = thâm hụt) — user (2026-08-08): "bổ sung dữ liệu xuất
// khẩu nhập khẩu các tháng... thêm biểu đồ cán cân xuất nhập khẩu theo tháng" (khác trade_balance
// hiện có trên trang chỉ là số LŨY KẾ theo quý từ NSO). Nguồn export_value_monthly/import_value_
// monthly/trade_balance_monthly (Hải quan, xem load_customs_xnk_local() trong fetch_macro_data.py)
// ĐÃ CÓ SẴN 144 điểm/chỉ báo nhưng CHƯA từng được vẽ chart — dữ liệu này CHỈ cập nhật khi chạy
// pipeline THỦ CÔNG trên máy có sẵn thư mục Excel Hải quan (GitHub Action không có, xem comment
// CUSTOMS_XNK_FOLDER), nên chuỗi có thể trễ vài tháng so với hiện tại — không phải lỗi hiển thị.
function renderTradeBalanceMonthlyChart(grid, indicators) {
    const exportSeries = ((indicators.export_value_monthly || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const importSeries = ((indicators.import_value_monthly || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const balanceSeries = ((indicators.trade_balance_monthly || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const allPeriods = _sortPeriods(new Set([...exportSeries, ...importSeries, ...balanceSeries].map(p => p.period)));
    if (!allPeriods.length) return;

    const exportByPeriod = Object.fromEntries(exportSeries.map(p => [p.period, p.value]));
    const importByPeriod = Object.fromEntries(importSeries.map(p => [p.period, p.value]));
    const balanceByPeriod = Object.fromEntries(balanceSeries.map(p => [p.period, p.value]));

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Xuất khẩu, nhập khẩu & cán cân thương mại theo từng tháng</span></div>
        <div class="ind-chart" style="height:320px"><canvas id="chart-trade-balance-monthly"></canvas></div>
        <div class="ind-note">Nguồn: Tổng cục Hải quan (file "Trị giá xuất/nhập khẩu sơ bộ các tháng", cập nhật THỦ CÔNG — không tự động qua GitHub Action nên có thể trễ vài tháng). Cột (trục phải) = cán cân THÁNG ĐÓ (xanh = xuất siêu, đỏ = nhập siêu); đường (trục trái) = kim ngạch xuất/nhập khẩu tuyệt đối — 2 trục lệch quy mô (kim ngạch ~30-60 tỷ USD, cán cân ~vài tỷ USD) nên tách riêng để nhìn rõ cả hai.</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-trade-balance-monthly');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: [
                {
                    type: 'bar', label: 'Cán cân thương mại (phải)', yAxisID: 'y1',
                    data: allPeriods.map(p => balanceByPeriod[p] ?? null),
                    backgroundColor: allPeriods.map(p => (balanceByPeriod[p] ?? 0) >= 0 ? '#10b98188' : '#ef444488'),
                    borderWidth: 0, order: 3,
                },
                {
                    type: 'line', label: 'Xuất khẩu (trái)', yAxisID: 'y',
                    data: allPeriods.map(p => exportByPeriod[p] ?? null),
                    borderColor: '#10b981', backgroundColor: '#10b98115', fill: false,
                    tension: 0.15, pointRadius: 1, spanGaps: true, order: 1,
                },
                {
                    type: 'line', label: 'Nhập khẩu (trái)', yAxisID: 'y',
                    data: allPeriods.map(p => importByPeriod[p] ?? null),
                    borderColor: '#ef4444', backgroundColor: '#ef444415', fill: false,
                    tension: 0.15, pointRadius: 1, spanGaps: true, order: 2,
                },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 12 } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     title: { display: true, text: 'Kim ngạch XK/NK (tỷ USD)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'Cán cân thương mại (tỷ USD)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    });
    chartInstances.push(chart);
}

// Số dư tín phiếu NHNN đang lưu hành (TỒN KHO, kênh HÚT — đối lập OMO là kênh BƠM) + bơm/hút ròng
// RIÊNG kênh tín phiếu (DÒNG CHẢY/ngày, suy ra từ chênh lệch tồn kho — xem
// _add_tin_phieu_net_operation trong template_vimo.py) — cùng kiểu trình bày "stock vs flow" như
// renderOmoHistoryChart ở trên (user 2026-08-08: muốn 1 biểu đồ Tbill tương tự OMO để thấy trực
// quan kênh này đang "tắt" — NHNN không chào thầu tín phiếu nào từ 30/10/2025, series này bằng 0
// xuyên suốt từ mốc đó tới nay, KHÔNG phải thiếu dữ liệu).
function renderTinPhieuHistoryChart(grid, indicators) {
    const outstanding = ((indicators.tin_phieu_outstanding_balance || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const net = ((indicators.tin_phieu_net_operation || {}).series || [])
        .filter(p => p.value !== null && p.value !== undefined);
    const allPeriods = _sortPeriods(new Set([...outstanding, ...net].map(p => p.period)));
    if (!allPeriods.length) return;

    const outstandingByPeriod = Object.fromEntries(outstanding.map(p => [p.period, p.value]));
    const netByPeriod = Object.fromEntries(net.map(p => [p.period, p.value]));

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">📈 Số dư tín phiếu NHNN đang lưu hành & Bơm/hút ròng qua tín phiếu theo thời gian</span></div>
        <div class="ind-chart" style="height:320px"><canvas id="chart-tin-phieu-history"></canvas></div>
        <div class="ind-note">Nguồn: vira.org.vn (bản tin Kinh tế - Tài chính ngày, tự động, chuỗi theo NGÀY thật). Đường (trục trái) = tồn kho lưu hành; cột (trục phải) = dòng chảy ròng/ngày (xanh = hút ròng thêm, đỏ = bơm ròng trả lại/đáo hạn nhiều hơn phát hành). Kênh HÚT thanh khoản, đối lập OMO (kênh BƠM) — bằng 0 xuyên suốt từ 30/10/2025 (lần chào bán tín phiếu gần nhất) tới nay vì NHNN không dùng kênh này, không phải thiếu dữ liệu.</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-tin-phieu-history');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: [
                {
                    type: 'bar', label: 'Bơm/hút ròng tín phiếu (phải)', yAxisID: 'y1',
                    data: allPeriods.map(p => netByPeriod[p] ?? null),
                    backgroundColor: allPeriods.map(p => (netByPeriod[p] ?? 0) >= 0 ? '#10b98188' : '#ef444488'),
                    borderWidth: 0, order: 2,
                    datalabels: { ..._endpointDatalabelsConfig(0), color: '#c9d2e3' },
                },
                {
                    type: 'line', label: 'Số dư lưu hành (trái)', yAxisID: 'y',
                    data: allPeriods.map(p => outstandingByPeriod[p] ?? null),
                    borderColor: '#f59e0b', backgroundColor: '#f59e0b15', fill: true,
                    tension: 0.1, pointRadius: 1, spanGaps: true, order: 1,
                    datalabels: _endpointDatalabelsConfig(0),
                },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 12 } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'Tồn kho (tỷ đồng)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'Bơm/hút ròng (tỷ đồng)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// Phiên bản TỔNG QUÁT của renderInterbank6mHistoryChart/renderBondYieldHistoryChart ở trên — dùng
// cho đường cong lợi suất TPCP Mỹ (renderInternationalSection). 2 hàm VN phía trên GIỮ NGUYÊN
// không đụng vào (đã ổn định, tránh rủi ro regression); hàm này viết mới để dùng lại logic hợp
// nhất period + autoSkip trục X + datalabels điểm cuối cho các chart nhiều kỳ hạn khác trong
// tương lai mà không phải chép lại.
function renderMultiTenorHistoryChart(grid, indicators, tenors, canvasId, title, note, periodFilterRe) {
    const seriesByTenor = tenors.map(([key, tenorLabel, color]) => [
        tenorLabel, color,
        ((indicators[key] || {}).series || []).filter(p =>
            p.value !== null && p.value !== undefined && (!periodFilterRe || periodFilterRe.test(p.period))),
    ]);
    const allPeriods = _sortPeriods(new Set(seriesByTenor.flatMap(([, , s]) => s.map(p => p.period))));
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">${title}</span></div>
        <div class="ind-chart" style="height:320px"><canvas id="${canvasId}"></canvas></div>
        <div class="ind-note">${note}</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector(`#${canvasId}`);
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
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: true, labels: { boxWidth: 12 } },
                datalabels: _endpointDatalabelsConfig(2),
            },
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 10 } } },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

const US_YIELD_TENORS = [
    ['us_yield_3m', '3 Tháng', '#f59e0b'], ['us_yield_1y', '1 Năm', '#ef4444'],
    ['us_yield_2y', '2 Năm', '#10b981'], ['us_yield_5y', '5 Năm', '#a78bfa'],
    ['us_10y_yield', '10 Năm', '#3b82f6'], ['us_yield_30y', '30 Năm', '#ec4899'],
];

// Spread 10Y-2Y / 10Y-3M — 2 chỉ báo đảo ngược đường cong lợi suất kinh điển (âm = đảo ngược,
// cảnh báo suy thoái Mỹ), tính sẵn trong template_vimo.py (_add_us_yield_curve_spread, phái sinh
// KHÔNG lưu vimo_raw.json). Vẽ thêm 1 đường "0" phẳng làm mốc tham chiếu — không cần plugin
// annotation, chỉ cần 1 dataset toàn giá trị 0 vẽ nét đứt màu xám.
function renderUsYieldSpreadChart(grid, indicators) {
    const SPREAD_KEYS = [
        ['us_yield_spread_10y2y', '10 Năm - 2 Năm', '#3b82f6'],
        ['us_yield_spread_10y3m', '10 Năm - 3 Tháng', '#f59e0b'],
    ];
    const seriesByKey = SPREAD_KEYS.map(([key, label, color]) => [
        label, color, ((indicators[key] || {}).series || []).filter(p => p.value !== null && p.value !== undefined),
    ]);
    const allPeriods = _sortPeriods(new Set(seriesByKey.flatMap(([, , s]) => s.map(p => p.period))));
    if (!allPeriods.length) return;

    const card = document.createElement('div');
    card.className = 'vimo-indicator-card';
    card.style.gridColumn = '1 / -1';
    card.innerHTML = `
        <div class="ind-header"><span class="ind-name">⚠️ Spread lợi suất TPCP Mỹ (10Y-2Y, 10Y-3M) — cảnh báo đảo ngược đường cong</span></div>
        <div class="ind-chart" style="height:320px"><canvas id="chart-us-yield-spread"></canvas></div>
        <div class="ind-note">Nguồn: fred.stlouisfed.org (tính từ lợi suất TPCP các kỳ hạn, hàng ngày). Spread ÂM (đường xuống dưới mốc "0" nét đứt) = đường cong ĐẢO NGƯỢC — tín hiệu cảnh báo suy thoái Mỹ được thị trường toàn cầu theo dõi sát nhất.</div>
    `;
    grid.appendChild(card);

    const ctx = card.querySelector('#chart-us-yield-spread');
    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allPeriods,
            datasets: [
                ...seriesByKey.map(([label, color, s]) => {
                    const byPeriod = Object.fromEntries(s.map(p => [p.period, p.value]));
                    return {
                        label, data: allPeriods.map(p => byPeriod[p] ?? null),
                        borderColor: color, backgroundColor: color + '15', fill: false,
                        tension: 0.25, pointRadius: 3, spanGaps: true,
                    };
                }),
                {
                    label: 'Mốc 0 (ranh giới đảo ngược)', data: allPeriods.map(() => 0),
                    borderColor: '#94a3b8', borderDash: [6, 4], pointRadius: 0, borderWidth: 1.5, fill: false,
                },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: { display: true, labels: { boxWidth: 12 } },
                datalabels: _endpointDatalabelsConfig(2),
            },
            scales: { ...CHART_DEFAULTS.scales, x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 10 } } },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// ═══════════════════════════════════════════════════════════
// DỮ LIỆU QUỐC TẾ — Mỹ / Châu Âu / Trung Quốc (user 2026-07-31: đối chiếu vĩ mô VN với 3 thị
// trường lớn). Đặt Ở CUỐI TRANG, cụm riêng theo thị trường — KHÔNG trộn vào GROUP_ORDER/
// renderIndicatorGroups() ở trên (group intl_us/intl_eu/intl_cn cố tình KHÔNG có trong
// GROUP_ORDER nên vòng lặp renderIndicatorGroups() tự động bỏ qua, không cần thay đổi gì ở đó).
// ═══════════════════════════════════════════════════════════
const INTL_MARKETS = [
    ['intl_us', '🇺🇸 Mỹ (Hoa Kỳ)'],
    ['intl_eu', '🇪🇺 Châu Âu (Eurozone)'],
    ['intl_uk', '🇬🇧 Anh (Vương quốc Anh)'],
    ['intl_jp', '🇯🇵 Nhật Bản'],
    ['intl_cn', '🇨🇳 Trung Quốc'],
];

function renderInternationalSection(indicators) {
    const container = document.getElementById('international-section-container');
    container.innerHTML = '';

    const hasAnyIntlData = INTL_MARKETS.some(([grp]) =>
        Object.values(indicators).some(ind => ind.group === grp));
    if (!hasAnyIntlData) return;

    const header = document.createElement('div');
    header.innerHTML = `<div class="vimo-group-header"><h3>🌍 Dữ liệu Quốc tế — Đối chiếu xu hướng chính sách tiền tệ & vĩ mô toàn cầu</h3></div>`;
    container.appendChild(header);

    INTL_MARKETS.forEach(([grp, label]) => {
        const entries = Object.entries(indicators).filter(([, ind]) => ind.group === grp);
        if (!entries.length) return;

        const section = document.createElement('div');
        section.innerHTML = `<div class="vimo-group-header"><h3>${label}</h3></div>
            <div class="vimo-indicator-grid" id="grid-${grp}"></div>`;
        container.appendChild(section);
        const grid = section.querySelector(`#grid-${grp}`);

        entries.forEach(([key, ind]) => _renderGenericIndicatorCard(grid, key, ind));

        if (grp === 'intl_us') {
            renderMultiTenorHistoryChart(grid, indicators, US_YIELD_TENORS, 'chart-us-yield-curve',
                '📈 Đường cong lợi suất TPCP Mỹ theo thời gian (3 tháng — 30 năm)',
                'Nguồn: fred.stlouisfed.org (Treasury Constant Maturity Rate, theo ngày).', null);
            renderUsYieldSpreadChart(grid, indicators);
        }
    });
}

// ═══════════════════════════════════════════════════════════
// BIỂU ĐỒ TỔNG QUAN VĨ MÔ — đặt NGAY TRÊN ĐẦU trang (user 2026-08-08, kèm ảnh mẫu dashboard
// "VĨ MÔ VIỆT NAM THÁNG 7/2026") để thấy diễn biến chung mà không phải kéo xem nhiều chart bên
// dưới. Dữ liệu do _build_macro_overview() (template_vimo.py) đóng gói sẵn thành {currentYear,
// recentFullYear}, mỗi năm là list {key, label, unit, values[]} theo đúng thứ tự tháng 1..N.
//
// NÂNG CẤP (user 2026-09-28, đối chiếu lại với ảnh mẫu — bản cũ dồn 7 chỉ báo + 3 trục vào 1
// biểu đồ duy nhất, "quá sơ sài và đọc khó hiểu"): 3 thay đổi chính, bám sát đúng cách trình bày
// của ảnh mẫu thay vì chỉ vẽ 1 line chart phẳng:
//   1. Dải thẻ KPI phía trên — đọc THẲNG giá trị mới nhất từng chỉ báo (không cần dò biểu đồ),
//      giống các ô số bên phải ảnh mẫu ("SẢN XUẤT CÔNG NGHIỆP 14,5%"...).
//   2. Tách riêng nhóm % tăng trưởng (IIP/Bán lẻ/XK/NK/CPI, chung 1 trục) khỏi nhóm vốn giải
//      ngân lũy kế (FDI tỷ USD / ĐT công nghìn tỷ, 2 trục riêng) — 2 biểu đồ NHỎ dễ đọc thay vì
//      1 biểu đồ 3 trục chồng chéo.
//   3. Nhãn giá trị tại từng điểm (khi năm hiện tại còn ít tháng, giống ảnh mẫu ghi số dưới mỗi
//      tháng) + đường nét đứt "bình quân lũy kế từ đầu năm" cho nhóm %, tương ứng đúng cặp
//      "Tháng X/2026" (nét liền) / "Bình quân NT/2026" (nét đứt) trong ảnh mẫu — GSO cũng công
//      bố song song 2 số này (tăng trưởng của riêng tháng đó vs bình quân N tháng đầu năm).
// ═══════════════════════════════════════════════════════════
const MACRO_OVERVIEW_COLORS = {
    iip_growth: '#3b82f6',
    retail_sales_growth: '#ec4899',
    export_growth_customs: '#10b981',
    import_growth_customs: '#ef4444',
    cpi_yoy: '#f59e0b',
    fdi_disbursed: '#a78bfa',
    public_investment_disbursement_value: '#22d3ee',
};

function renderMacroOverview(macroOverview) {
    if (!macroOverview) return;
    const card = document.getElementById('macro-overview-card');
    if (!card) return;

    const current = macroOverview.currentYear;
    const recent = macroOverview.recentFullYear;
    const hasCurrent = current && current.series && current.series.length && current.monthsShown;
    const hasRecent = recent && recent.series && recent.series.length;
    if (!hasCurrent && !hasRecent) return;
    card.style.display = '';

    const titleCurrent = document.getElementById('macro-overview-title-current');
    const titleRecent = document.getElementById('macro-overview-title-recent');
    if (hasCurrent) {
        titleCurrent.textContent = `📅 Năm ${current.year} — số liệu mới nhất: tháng ${current.monthsShown}/${current.year} (đến hết tháng ${current.monthsShown})`;
        _renderMacroOverviewKpiStrip('macro-overview-kpi-current', current);
        _renderMacroOverviewChartPair('current', current);
    }
    if (hasRecent) {
        titleRecent.textContent = `📅 Đối chiếu: năm ${recent.year} (đã hoàn tất 12 tháng)`;
        _renderMacroOverviewChartPair('recent', recent);
    }
}

// Dải thẻ KPI: DIỄN HỌA giá trị THÁNG MỚI NHẤT của từng chỉ báo ra dạng số dễ đọc (đúng như ảnh
// mẫu — mỗi ô chỉ "kể lại" 1 số đã có trên chart cho dễ nhìn, KHÔNG PHẢI so sánh mới).
// SỬA (user 2026-09-28, xem lại đúng ảnh mẫu): bản trước hiện "▲ +0.5 so T7" — dễ hiểu NHẦM thành
// "so với tháng trước" trong khi bản chất % ở đây là tăng trưởng YoY (so với CÙNG KỲ năm trước),
// giá trị tháng này cao/thấp hơn giá trị tháng trước (đều là số YoY) không phải là 1 khái niệm "so
// tháng trước" theo nghĩa thông thường — BỎ delta này. Thay bằng đúng cặp số ảnh mẫu dùng: giá trị
// THÁNG (khớp đường nét liền trên chart) + BÌNH QUÂN LŨY KẾ TỪ ĐẦU NĂM đến đúng tháng đó (khớp
// đường nét đứt) — vd ảnh mẫu "BÁN LẺ...13,1% / 7T: 13,1%". Chỉ áp dụng cho chỉ báo %; chỉ báo lũy
// kế tuyệt đối (FDI/ĐT công) không có khái niệm "bình quân" tương đương nên chỉ ghi rõ "lũy kế từ
// đầu năm" để không ai hiểu lầm đó là YoY.
function _renderMacroOverviewKpiStrip(containerId, yearData) {
    const el = document.getElementById(containerId);
    if (!el) return;
    const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    el.innerHTML = yearData.series.map(row => {
        const color = MACRO_OVERVIEW_COLORS[row.key] || '#9aa5bd';
        const vals = row.values;
        let lastIdx = -1;
        for (let i = vals.length - 1; i >= 0; i--) {
            if (vals[i] !== null && vals[i] !== undefined) { lastIdx = i; break; }
        }
        if (lastIdx < 0) return '';
        const latest = vals[lastIdx];
        const unitSuffix = row.unit === '%' ? '%' : ` ${row.unit}`;
        const metaHtml = row.unit === '%'
            ? `Tháng ${lastIdx + 1}/${yearData.year} (YoY) · BQ ${lastIdx + 1}T: ${_cumulativeAvg(vals)[lastIdx].toFixed(1)}%`
            : `Tháng ${lastIdx + 1}/${yearData.year} · lũy kế từ đầu năm`;
        return `
            <div class="vimo-macro-kpi-card" style="border-left-color:${color}">
                <span class="vimo-macro-kpi-label">${esc(row.label)}</span>
                <span class="vimo-macro-kpi-value" style="color:${color}">${latest.toFixed(1)}<small>${esc(unitSuffix)}</small></span>
                <span class="vimo-macro-kpi-meta">${metaHtml}</span>
            </div>`;
    }).join('');
}

// Bình quân CỘNG DỒN từ tháng 1 đến đúng tháng đó (vd tháng 5 = trung bình giá trị tháng 1..5) —
// đúng cách GSO công bố song song "riêng tháng"/"bình quân N tháng đầu năm", dùng làm đường nét
// đứt đối chiếu trong _renderMacroPctChart (giống cặp nét liền/nét đứt trong ảnh mẫu).
function _cumulativeAvg(values) {
    const out = [];
    let sum = 0, n = 0;
    for (const v of values) {
        if (v !== null && v !== undefined) { sum += v; n += 1; }
        out.push(n ? sum / n : null);
    }
    return out;
}

function _renderMacroOverviewChartPair(prefix, yearData) {
    const labels = Array.from({ length: yearData.monthsShown }, (_, i) => `T${i + 1}`);
    const pctSeries = yearData.series.filter(r => r.unit === '%');
    const absSeries = yearData.series.filter(r => r.unit !== '%');

    const pctWrap = document.getElementById(`macro-overview-${prefix}-pct-wrap`);
    const absWrap = document.getElementById(`macro-overview-${prefix}-abs-wrap`);
    if (pctWrap) pctWrap.style.display = pctSeries.length ? '' : 'none';
    if (absWrap) absWrap.style.display = absSeries.length ? '' : 'none';

    if (pctSeries.length) {
        _renderMacroPctChart(`chart-macro-overview-${prefix}-pct`, labels, pctSeries, yearData.monthsShown);
    }
    if (absSeries.length) {
        _renderMacroAbsChart(`chart-macro-overview-${prefix}-abs`, labels, absSeries);
    }
}

function _renderMacroPctChart(canvasId, labels, series, monthsShown) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    // Năm hiện tại thường mới có vài tháng (ít điểm) -> hiện nhãn giá trị ở MỌI điểm, giống ảnh
    // mẫu; năm đối chiếu đủ 12 tháng x nhiều đường dễ rối -> chỉ hiện nhãn ở điểm CUỐI (quy ước
    // _endpointDatalabelsConfig đã dùng cho các chart nhiều đường/nhiều điểm khác trong file này).
    const showAllPoints = monthsShown <= 9;
    const datasets = [];
    series.forEach(row => {
        const color = MACRO_OVERVIEW_COLORS[row.key] || '#9aa5bd';
        datasets.push({
            label: row.label, data: row.values,
            borderColor: color, backgroundColor: color + '20', pointBackgroundColor: color,
            fill: false, tension: 0.25, pointRadius: 3, borderWidth: 2.5, spanGaps: true,
            datalabels: showAllPoints ? {
                display: (ctx) => ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] !== undefined,
                color, align: 'top', anchor: 'end', offset: 3, font: { size: 9, weight: '700' },
                formatter: (v) => v.toFixed(1),
            } : _endpointDatalabelsConfig(1),
        });
        // Đường nét đứt "bình quân lũy kế từ đầu năm" — chỉ hiện nhãn ở điểm cuối (giá trị bình
        // quân tính đến tháng mới nhất), tránh chồng lấn với nhãn của đường nét liền ở trên.
        datasets.push({
            label: `${row.label} (bình quân lũy kế)`, data: _cumulativeAvg(row.values),
            borderColor: color, borderDash: [5, 4], borderWidth: 1.5, pointRadius: 0,
            fill: false, tension: 0.25, spanGaps: true,
            datalabels: { ..._endpointDatalabelsConfig(1), font: { size: 8, style: 'italic' } },
        });
    });

    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: {
                legend: {
                    display: true, labels: { boxWidth: 11, font: { size: 9 },
                        // Chỉ liệt kê tên chỉ báo (đường nét liền) trong chú giải — đường bình quân
                        // dùng CHUNG màu + kiểu nét đứt đã ghi rõ trong phụ đề phía trên biểu đồ,
                        // liệt kê thêm ở đây sẽ nhân đôi số dòng chú giải không cần thiết.
                        filter: (item) => !item.text.includes('(bình quân lũy kế)') },
                },
            },
            scales: {
                ...CHART_DEFAULTS.scales,
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: labels.length } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: '% YoY', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

function _renderMacroAbsChart(canvasId, labels, series) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const hasFdi = series.some(r => r.key === 'fdi_disbursed');
    const hasPublicInv = series.some(r => r.key === 'public_investment_disbursement_value');

    const datasets = series.map(row => {
        const color = MACRO_OVERVIEW_COLORS[row.key] || '#9aa5bd';
        const yAxisID = row.key === 'fdi_disbursed' ? 'y1' : 'y2';
        return {
            label: row.label, data: row.values, yAxisID,
            borderColor: color, backgroundColor: color + '20', pointBackgroundColor: color,
            fill: false, tension: 0.25, pointRadius: 3, borderWidth: 2.5, spanGaps: true,
            // Chỉ 2 đường tối đa (FDI + ĐT công) nên hiện nhãn ở MỌI điểm luôn an toàn (không rối
            // như biểu đồ % có tới 5 đường), giống các con số lũy kế ghi dưới mỗi tháng trong ảnh mẫu.
            datalabels: {
                display: (ctx) => ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] !== undefined,
                color, align: 'top', anchor: 'end', offset: 3, font: { size: 9, weight: '700' },
                formatter: (v) => v.toFixed(1),
            },
        };
    });

    const scales = { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: labels.length } } };
    if (hasFdi) {
        scales.y1 = { ...CHART_DEFAULTS.scales.y, position: 'left',
                      title: { display: true, text: 'FDI giải ngân (tỷ USD)', color: '#9aa5bd', font: { size: 9 } } };
    }
    if (hasPublicInv) {
        scales.y2 = { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'ĐT công giải ngân (nghìn tỷ đồng)', color: '#9aa5bd', font: { size: 9 } } };
    }

    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 11, font: { size: 9 } } } },
            scales,
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}
