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

// THEM 2026-10-10 (user: "tất cả các ghi chú kí hiệu biểu đồ cho chữ rõ hơn đi, nhìn mờ quá, để
// màu trắng cho rõ" — hầu hết chart KHÔNG set color riêng cho legend.labels, nên rơi về màu mặc
// định của Chart.js (#666, xám tối) — trên nền tối #0b1220 của dashboard thì gần như không đọc
// được/mờ. Set 1 LẦN DUY NHẤT ở global default thay vì sửa tay từng chart (hàng chục chỗ) — mọi
// legend không ghi đè riêng sẽ tự dùng màu này.
Chart.defaults.color = '#ffffff';

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

// SUA 2026-10-09 (user: "cột bị nhỏ đi, có vẻ khung biểu đồ bị lấn ra" ở tab Báo cáo — đo thử
// width thật của từng cột bằng Chart.getChart().getDatasetMeta() thì các cột RỘNG BẰNG NHAU tuyệt
// đối, không có lỗi kích thước; vấn đề thực sự là align:'right' của _endpointDatalabelsConfig đẩy
// nhãn số RA NGOÀI mép phải vùng vẽ — chart có trục phải (y1) thì nhãn đè lên nhãn trục làm cột
// cuối trông như bị "lấn"; chart KHÔNG có trục phải thì nhãn bị cắt mất hẳn ra ngoài canvas, không
// hiện số như mong muốn. Dùng align:'top' thay vì 'right' — nhãn nổi NGAY TRÊN điểm cuối, luôn nằm
// trong vùng vẽ bất kể chart có trục phụ hay không.
function _endpointAboveLabelConfig(decimals, suffix) {
    return {
        display: (ctx) => ctx.dataset.data.slice(ctx.dataIndex + 1).every(v => v === null || v === undefined)
            && (ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] !== undefined),
        color: (ctx) => ctx.dataset.borderColor, font: { size: 10, weight: '700' },
        anchor: 'end', align: 'top', offset: 6, clip: false,
        formatter: (v) => (v === null || v === undefined) ? '' : (v >= 0 ? '+' : '') + v.toFixed(decimals) + (suffix || ''),
    };
}

let chartInstances = [];

// Cac bang heatmap (.monitoring-table-scroll, overflow-x:auto) chi cuon ngang duoc qua scrollbar/
// Shift+wheel theo mac dinh trinh duyet - user (2026-10-01) khong nhan ra co the cuon vi scrollbar
// OS/trinh duyet thuong rat manh/an tren nen toi. Doi lan cuon chuot DOC (deltaY) binh thuong thanh
// cuon NGANG khi hover cac bang nay, giong UX chuan cua bang rong (vd Google Sheets/Notion).
document.addEventListener('wheel', (e) => {
    const scroller = e.target.closest('.monitoring-table-scroll');
    if (!scroller || scroller.scrollWidth <= scroller.clientWidth) return;
    if (Math.abs(e.deltaY) <= Math.abs(e.deltaX)) return; // da la cuon ngang (trackpad) -> de trinh duyet tu xu ly
    e.preventDefault();
    scroller.scrollLeft += e.deltaY;
}, { passive: false });

// THEM 2026-10-05 (user: "làm dạng tab này để dễ xem và bấm vào cái nào cần") — gom 19 khối trang
// thành 6 tab theo chủ đề (xem data-tab trong vimo.html). Tab nằm trên hash (#tab=...) để chia sẻ được.
function initVimoTabs() {
    const buttons = document.querySelectorAll('[data-tab-btn]');
    if (!buttons.length) return;
    const show = (tab) => {
        const valid = [...buttons].some(b => b.dataset.tabBtn === tab) ? tab : 'report';
        buttons.forEach(b => b.classList.toggle('active', b.dataset.tabBtn === valid));
        document.querySelectorAll('[data-tab]').forEach(el => {
            el.classList.toggle('vimo-tab-hidden', el.dataset.tab !== valid);
        });
        history.replaceState(null, '', '#tab=' + valid);
        setTimeout(() => {
            chartInstances.forEach(c => c.resize());
            // SUA 2026-10-08 (user: "tôi mở biểu đồ đỡ phải cuộn lại") — khi vừa chuyển sang tab
            // này, các bảng/chart cuộn-ngang (.monitoring-table-scroll) ĐANG ẨN lúc trang load lần
            // đầu nên scrollLeft=scrollWidth lúc đó là no-op (phần tử display:none, scrollWidth=0).
            // Chạy lại NGAY KHI tab thật sự hiện ra (sau resize, DOM đã có kích thước thật).
            document.querySelectorAll(`[data-tab="${valid}"] .monitoring-table-scroll`).forEach(el => { el.scrollLeft = el.scrollWidth; });
        }, 0);
    };
    buttons.forEach(b => b.addEventListener('click', () => show(b.dataset.tabBtn)));
    const fromHash = (location.hash.match(/tab=([a-z]+)/) || [])[1];
    show(fromHash || 'report');
}

document.addEventListener('DOMContentLoaded', async () => {
    initVimoTabs();
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
    renderGdpSectorTable(data.gdpSectorTable);
    renderCpiGroupTable(data.cpiGroupTable);
    renderExportCommodityTable(data.exportCommodityTable);
    renderImportCommodityTable(data.importCommodityTable);
    renderExportPriceTable(data.exportPriceTable);
    renderImportPriceTable(data.importPriceTable);
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
    renderFxPressureSignalsChart(data.indicators);
    renderFxPressureSignalsMonthlyChart(data.indicators);
    renderFxYieldDiffChart(data.indicators, data.usMacro);
    renderDepositRateChart(data.indicators);
    renderDepositRateCakeChart(data.indicators);
    renderUsMacro(data.usMacro);
    renderVnReport(data.vnReport);
    renderFxBalanceOverviewChart(data.indicators);
    renderFxSupplyDemandTotalChart(data.indicators);
    renderFxSupplyDemandChart(data.indicators);
    renderFxFinancialAccountChart(data.indicators);
    renderFxRateGapChart(data.indicators);

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

    // SUA 2026-10-02 (Thông tư 50/2026/TT-NHNN THAY HẲN Thông tư 22/2019 — user chụp Điều 12 Mục 4):
    // "Tổng tiền gửi" (D) ở TT50 đã TỰ GỘP VCSH vào công thức chính thức (khoản 4g/h/i) — khác TT22
    // cũ (D KHÔNG gồm VCSH). 'tctdDeposits' (gộp thô) đổi thành 'netTctd' (chỉ phần RÒNG DƯƠNG mỗi
    // bank — TT50 khoản 4đ) — xem _ldr_components() (bank_system_risk.py). KHÔNG đưa 'tpdnDeduction'
    // (ÂM, TT50 khoản 5b trừ TPDN khỏi D) vào biểu đồ miền này — scale Y cố định min:0 (stacked)
    // không vẽ được giá trị âm đúng cách; TPDN thường nhỏ so tổng huy động nên bỏ qua trên CHART,
    // vẫn tính ĐÚNG trong totalDeposit/LDR thật (xem ghi chú dưới chart).
    const DEPOSIT_SERIES = [
        { key: 'customerDeposits', label: 'Tiền gửi khách hàng (đã trừ ký quỹ/vốn CD)', color: '#10b981' },
        { key: 'bonds', label: 'Giấy tờ có giá phát hành', color: '#a78bfa' },
        { key: 'netTctd', label: 'Vị thế liên NH RÒNG (chỉ tính nếu dương)', color: '#3b82f6' },
        { key: 'kbnnCounted', label: 'KBNN (tính theo tỷ lệ lộ trình)', color: '#f59e0b' },
        { key: 'equity', label: 'Vốn chủ sở hữu (VCSH)', color: '#ef4444' },
    ];

    // SUA 2026-10-03 (user: "cái tính theo TT50 thì tính đủ vẫn là ròng dương nhé, còn biểu đồ trừ
    // VCSH đi thì tính như cách cũ của tôi nhé, để tôi so ngang được xem LDR như nào") — phát hiện
    // "KHÔNG gồm VCSH" (bản cũ: TT50 trừ thẳng equity ra) làm LDR nhảy từ ~85% lên ~94,6%, chủ yếu
    // do đổi "Tiền gửi TCTD khác" từ GỘP THÔ sang RÒNG DƯƠNG (TT50) làm mẫu số giảm RẤT NHIỀU (vd
    // Q2-2026 toàn hệ thống: gộp thô 2.875.874 tỷ vs ròng 376.575 tỷ, chênh 2.499.299 tỷ) — KHÔNG
    // phải do VCSH. User muốn bên phải dùng ĐÚNG công thức CŨ của họ (_ldr_components_old trong
    // bank_system_risk.py: gộp thô TCTD, KHÔNG VCSH, KHÔNG trừ TPDN khỏi D) để so ngang 2 công thức,
    // bên trái GIỮ NGUYÊN TT50 (ròng dương + VCSH, không đổi).
    _renderCreditFundingLdrChart('chart-credit-structure-abs', cds.periods, cds.totalCredit, cds.totalDeposit, cds.ldrSystem, 'Tổng huy động (TT50, đã gồm VCSH)');
    _renderCreditFundingLdrChart('chart-credit-structure-pct', cds.periods, cds.totalCredit, cds.totalDepositOld, cds.ldrSystemOld, 'Tổng huy động (công thức cũ — gộp TCTD, không VCSH)');
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
            { key: 'depositGrowthYoy', label: 'Huy động YoY (rộng — TT50/2026)', color: '#10b981', dash: false },
            { key: 'depositGrowthYoyNarrow', label: 'Huy động YoY (hẹp — chỉ tiền gửi KH, khớp VBMA)', color: '#f59e0b', dash: false },
            { key: 'creditGrowthYtd', label: 'Tín dụng YTD (so cuối năm trước)', color: '#3b82f6', dash: true },
            { key: 'depositGrowthYtd', label: 'Huy động YTD (rộng — TT50/2026)', color: '#10b981', dash: true },
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
// phần CHƯA CÓ dữ liệu để biết chính xác bổ sung vào ĐÚNG LỚP nào sau này, không phải thiết kế
// lại — ĐÃ LẤP: NEER/REER (Darvas/Bruegel, BIS/IMF.STA:EER không có Việt Nam), Errors &
// Omissions/Overall Balance/Δ Dự trữ (NHNN BOP quý), FDI XNK tách Domestic/FDI (lớp ⑥ riêng,
// NSO), Kiều hối đúng nghĩa — CHỈ PHẠM VI TP.HCM (kieu_hoi_hcm, dulieukinhte.com/NHNN Chi nhánh
// Khu vực 2 — IMF BOP SDMX vẫn trả 0 dữ liệu nên dùng nguồn này thay thế). CÒN THIẾU: Kiều hối
// toàn quốc, doanh thu du lịch quốc tế tách riêng.
// SUA 2026-10-03 (user: "nhập khẩu ròng mà có áp lực tỷ giá đâu... xuất khẩu ròng mà tỷ giá lên cao
// kinh khủng" — chỉ ra Cầu/Cung dựa nhiều vào XNK hàng hóa là proxy KÉM cho áp lực thực, vì nhiều
// dòng XNK FDI không thực sự "qua" hệ thống NHTM VN [thanh toán qua công ty mẹ/tài khoản offshore],
// và dòng VỐN có thể bù/đảo ngược hoàn toàn chiều vãng lai — ví dụ thật Q2-2026: hàng hóa NHẬP SIÊU
// -5.497tr USD nhưng Cán cân tổng thể CHỈ +81tr USD gần như cân bằng, vì Cán cân tài chính +1.899 +
// Lỗi&Sai sót +9.283 bù hết). ĐỔI THỨ TỰ: đưa ③④ CŨ (Đối chiếu BOP/Áp lực thị trường — 2 tín hiệu
// THỰC, không suy luận từ dòng kế toán XNK) lên ①② MỚI, đẩy Cầu/Cung (dựa XNK) xuống ③④ kèm cảnh
// báo rõ KHÔNG phải áp lực thực — giống cách lớp ⑥ Trade Structure đã cảnh báo cho FDI trade.
const FX_PRESSURE_LAYERS = [
    {
        id: 'bop', title: '① Đối chiếu BOP — TÍN HIỆU THỰC (phần NHNN thực sự phải giải quyết bằng dự trữ)',
        keys: ['bop_sbv_current_account', 'bop_sbv_financial_account', 'bop_sbv_external_debt_net',
               'bop_sbv_errors_omissions', 'bop_sbv_overall_balance', 'bop_sbv_reserve_assets_change'],
        missing: [],
    },
    {
        id: 'market', title: '② Áp lực thị trường — TÍN HIỆU THỰC (quan sát trực tiếp, không suy luận từ dòng kế toán)',
        keys: ['usdvnd', 'usdvnd_monthly_avg', 'usdvnd_growth_mom', 'usdvnd_growth_yoy',
               'usdvnd_vcb_sell_daily', 'usd_cho_den_sell_daily', 'usd_cho_den_vcb_gap', 'usd_cho_den_vcb_gap_pct',
               'interbank_rate_on', 'fed_funds_rate', 'vnd_usd_rate_spread_on',
               // THEM 2026-10-09 (user hỏi về "Swap Interest Rate Curve" làm bằng chứng lãi suất
               // VN khó hạ — nguồn thật VBMA FX Swap Curve bị khóa sau login hội viên, KHÔNG có
               // API công khai — xem ghi chú _add_vnd_usd_swap_proxy_gap trong template_vimo.py)
               // — PROXY từ chênh lệch lãi suất liên ngân hàng VND-USD CÙNG ngày, CÙNG nguồn VIRA,
               // 4 kỳ hạn ngắn khớp đúng phần đầu đường cong swap gốc (ON/1W/2W/1M).
               'interbank_usd_on', 'interbank_usd_1w', 'interbank_usd_2w', 'interbank_usd_1m',
               'interbank_vnd_usd_gap_on', 'interbank_vnd_usd_gap_1w', 'interbank_vnd_usd_gap_2w', 'interbank_vnd_usd_gap_1m',
               'darvas_neer_vn', 'darvas_reer_vn'],
        missing: [],
    },
    {
        id: 'demand', title: '③ Cầu ngoại tệ (cơ cấu dòng vãng lai — KHÔNG phải áp lực tỷ giá thực, xem ghi chú)',
        keys: ['import_growth_customs', 'import_growth_customs_mom',
               'bop_sbv_services_import', 'bop_sbv_investment_income_paid', 'bop_sbv_secondary_income_paid',
               'bop_sbv_fdi_assets_bop', 'bop_sbv_portfolio_assets_bop'],
        warning: 'Đây là PHÂN RÃ cán cân vãng lai (kế toán ghi nhận khi hàng hóa/dịch vụ đổi chủ), KHÔNG PHẢI đo lường tiền USD thực sự chảy qua hệ thống ngân hàng VN — nhiều khoản NK (nhất là của DN FDI) thanh toán qua công ty mẹ/tài khoản nước ngoài, không cần mua USD trong nước. Xem lớp ① Đối chiếu BOP (Cán cân tổng thể) và ② Áp lực thị trường để có tín hiệu áp lực THỰC.',
        missing: ['Trả nợ gốc nước ngoài TÁCH RIÊNG khỏi rút vốn mới — hiện chỉ có số RÒNG (external_debt_net ở lớp "Đối chiếu BOP"); lợi nhuận FDI chuyển ra TÁCH RIÊNG khỏi tổng Thu nhập đầu tư — NHNN BOP không tách, không nên tự gắn nhãn "FDI profit remittance" cho investment_income_paid (rộng hơn)'],
    },
    {
        id: 'supply', title: '④ Cung ngoại tệ (cơ cấu dòng vãng lai — KHÔNG phải áp lực tỷ giá thực, xem ghi chú)',
        keys: ['export_growth_customs', 'export_growth_customs_mom', 'fdi_disbursed', 'fdi_registered_usd_bn', 'trade_balance',
               'bop_sbv_services_export', 'bop_sbv_investment_income_received', 'bop_sbv_secondary_income_received',
               'bop_sbv_fdi_liabilities_bop', 'bop_sbv_portfolio_liabilities_bop', 'kieu_hoi_hcm'],
        warning: 'Cùng lý do với lớp ③ Cầu — nhiều khoản XK (nhất là DN FDI) giữ ngoại tệ ở tài khoản nước ngoài/chuyển thẳng về công ty mẹ, không BÁN lại USD cho NHTM trong nước. Dòng VỐN (FDI, đầu tư gián tiếp, vay nợ — xem lớp ① Đối chiếu BOP) có thể LỚN HƠN và NGƯỢC HƯỚNG với cán cân vãng lai, nên vãng lai dương KHÔNG đồng nghĩa dư cung USD thực tế (ví dụ: xuất siêu nhưng tỷ giá vẫn tăng mạnh nếu dòng vốn rút ra đủ lớn).',
        missing: ['Kiều hối ĐÚNG NGHĨA chỉ lấp được PHẠM VI TP.HCM (kieu_hoi_hcm, NHNN Chi nhánh Khu vực 2) — chưa có số toàn quốc; doanh thu du lịch quốc tế tách riêng — chưa có'],
    },
    {
        id: 'response', title: '⑤ Phản ứng NHNN',
        keys: ['forex_reserves_monthly', 'forex_reserves_sdr', 'omo_rate_7d', 'tin_phieu_outstanding_balance', 'tin_phieu_net_operation'],
        missing: [],
    },
    {
        // THEM (user 2026-10-01): FDI XNK tách Domestic/FDI — user NHẤN MẠNH đây là "TRADE
        // STRUCTURE", KHÔNG PHẢI "FX flow" (DN FDI có thể dùng vốn/giữ doanh thu offshore, không
        // chắc USD thực sự qua hệ thống NHTM VN) — cố tình KHÔNG gộp vào lớp ①②③ ở trên, tách
        // thành lớp riêng để không bị đọc nhầm là 1 proxy cung/cầu USD.
        id: 'trade_structure', title: '⑥ Cơ cấu Thương mại theo Khu vực DN (Trade Structure — KHÔNG phải dòng ngoại tệ thực)',
        keys: ['export_domestic_usd_bn', 'export_fdi_usd_bn', 'import_domestic_usd_bn', 'import_fdi_usd_bn',
               'fdi_trade_balance', 'domestic_trade_balance',
               'export_monthly_total', 'export_monthly_domestic', 'export_monthly_fdi', 'export_fdi_share_pct',
               'import_monthly_total', 'import_monthly_domestic', 'import_monthly_fdi', 'import_fdi_share_pct'],
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
        const warningHtml = layer.warning
            ? `<p class="ind-source-note" style="margin:4px 0 12px 4px;color:#f59e0b;font-weight:600">⚠️ ${layer.warning}</p>` : '';
        const missingHtml = layer.missing.length
            ? `<p class="ind-source-note" style="margin:-4px 0 12px 4px">⏳ Chưa có dữ liệu: ${layer.missing.join(' · ')}</p>` : '';
        return `<div class="vimo-group-header"><h3>${layer.title}</h3></div>${warningHtml}
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

// THEM (user 2026-10-01): "biểu đồ tỷ giá chợ đen bán ra và tỷ giá VCB bán ra cùng với gap tỷ
// giá" — 2 đường mức (VND, trục trái) + 1 đường gap nét đứt (VND, trục phải, khác thang đo vì
// gap nhỏ hơn mức tỷ giá rất nhiều lần). Khớp theo NGÀY có CẢ 2 nguồn (dùng chính periods của
// usd_cho_den_vcb_gap làm trục X — gap chỉ tính được ở ngày có đủ cả 2 phía, xem fetch_chogia_
// usd_cho_den/fetch_vcb_usd_sell_rate trong fetch_macro_data.py).
function renderFxRateGapChart(indicators) {
    const canvas = document.getElementById('chart-fx-vcb-cho-den-gap');
    if (!canvas) return;
    const gap = indicators['usd_cho_den_vcb_gap'];
    const vcb = indicators['usdvnd_vcb_sell_daily'];
    const choDen = indicators['usd_cho_den_sell_daily'];
    const card = document.getElementById('fx-rate-gap-chart-card');
    if (!gap || !vcb || !choDen || !gap.series.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    const periods = gap.series.map(p => p.period);
    const vcbByPeriod = Object.fromEntries(vcb.series.map(p => [p.period, p.value]));
    const choDenByPeriod = Object.fromEntries(choDen.series.map(p => [p.period, p.value]));
    const vcbArr = periods.map(p => vcbByPeriod[p] ?? null);
    const choDenArr = periods.map(p => choDenByPeriod[p] ?? null);
    const gapArr = gap.series.map(p => p.value);

    const levelDatalabels = (color) => ({
        display: (ctx) => ctx.dataset.data.slice(ctx.dataIndex + 1).every(v => v === null || v === undefined)
            && (ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] !== undefined),
        color, font: { size: 9, weight: '700' }, anchor: 'end', align: 'right', offset: 4, clip: false,
        formatter: (v) => v.toLocaleString('vi-VN', { maximumFractionDigits: 0 }),
    });
    const chart = new Chart(canvas, {
        type: 'line',
        data: {
            labels: periods,
            datasets: [
                { label: 'VCB — bán ra', data: vcbArr, yAxisID: 'y', borderColor: '#10b981', backgroundColor: '#10b98120',
                  fill: false, tension: 0.2, pointRadius: 2, pointBackgroundColor: '#10b981', borderWidth: 2, spanGaps: true,
                  datalabels: levelDatalabels('#10b981') },
                { label: 'Chợ đen — bán ra', data: choDenArr, yAxisID: 'y', borderColor: '#ef4444', backgroundColor: '#ef444420',
                  fill: false, tension: 0.2, pointRadius: 2, pointBackgroundColor: '#ef4444', borderWidth: 2, spanGaps: true,
                  datalabels: levelDatalabels('#ef4444') },
                { label: 'Gap (Chợ đen − VCB)', data: gapArr, yAxisID: 'y1', borderColor: '#f59e0b',
                  borderDash: [6, 4], borderWidth: 2, pointRadius: 0, fill: false, tension: 0.25, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(0) },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 10 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 45, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'Gap (VND)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM (user 2026-10-03): "tạo biểu đồ để đánh giá mức độ áp lực tỷ giá" — SAU KHI user chỉ ra
// Cầu/Cung (dựa nhiều vào XNK hàng hóa) là proxy KÉM cho áp lực thực (ví dụ thật Q2-2026: hàng hóa
// NHẬP SIÊU -5.497tr USD nhưng Cán cân tổng thể chỉ +81tr USD gần như cân bằng — vì Cán cân tài
// chính +1.899 và Lỗi&Sai sót +9.283 đã bù hết phần vãng lai âm). 2 TÍN HIỆU THỰC hơn: (a) Cán cân
// tổng thể (BOP, phần NHNN thực sự phải giải quyết bằng dự trữ) và (b) USD/VND thị trường TĂNG/GIẢM
// thật (không suy luận từ dòng kế toán). User đã CHỐT: dashboard nhiều đường ĐỘC LẬP, KHÔNG gộp
// thành 1 điểm số/thang điểm (giữ đúng nguyên tắc đã có từ renderFxSupplyDemandChart bên dưới).
function renderFxPressureSignalsChart(indicators) {
    const canvas = document.getElementById('chart-fx-pressure-signals');
    const card = document.getElementById('fx-pressure-signals-chart-card');
    if (!canvas) return;
    const ob = indicators['bop_sbv_overall_balance'];
    const usdvndYoy = indicators['usdvnd_growth_yoy'];
    const usdvndLevel = indicators['usdvnd_monthly_avg'];
    if (!ob || !ob.series.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    // SUA 2026-10-03 (user: "chỉ cần dữ liệu cảnh báo áp lực tỷ giá từ 2020 tới nay thôi, đừng cố
    // vẽ nhiều quá làm gì cho lãng phí") — dữ liệu BOP đã backfill tới 1996 (xem fetch_imf_bop_
    // vietnam_history, fetch_macro_data.py) vẫn GIỮ ĐẦY ĐỦ trong vimo_raw.json, CHỈ giới hạn PHẦN
    // VẼ ở đây từ 2020 trở đi — đủ cho mục đích xem áp lực gần đây, không cần kéo dài tới 1996.
    const obSeries = ob.series.filter(p => p.period >= '2020-Q1');
    const periods = obSeries.map(p => p.period);
    const obArr = obSeries.map(p => p.value);
    // Quy đổi USD/VND tăng trưởng YoY + MỨC thực tế (theo THÁNG) về cuối mỗi quý (tháng 3/6/9/12)
    // để so cùng trục X với BOP (theo QUÝ) — chỉ để VẼ CẠNH NHAU, không tính toán gộp gì cả.
    const QUARTER_END_MONTH = { '1': '03', '2': '06', '3': '09', '4': '12' };
    const usdvndByMonth = usdvndYoy ? Object.fromEntries(usdvndYoy.series.map(p => [p.period, p.value])) : {};
    const usdvndLevelByMonth = usdvndLevel ? Object.fromEntries(usdvndLevel.series.map(p => [p.period, p.value])) : {};
    const usdvndArr = periods.map(period => {
        const [year, q] = period.split('-Q');
        return usdvndByMonth[`${year}-${QUARTER_END_MONTH[q]}`] ?? null;
    });
    // THEM (user 2026-10-03): "áp thêm cho tôi đường tỷ giá thực USD/VND nhé" — MỨC tỷ giá thực tế
    // (VND), KHÁC hẳn đường %YoY đã có — cần TRỤC RIÊNG thứ 3 (y2) vì đơn vị/độ lớn khác hoàn toàn
    // (VND ~24.000-27.000 vs %YoY vài điểm % vs BOP vài nghìn-chục nghìn triệu USD).
    const usdvndLevelArr = periods.map(period => {
        const [year, q] = period.split('-Q');
        return usdvndLevelByMonth[`${year}-${QUARTER_END_MONTH[q]}`] ?? null;
    });

    const chart = new Chart(canvas, {
        type: 'bar',
        data: {
            labels: periods,
            datasets: [
                { type: 'bar', label: 'Cán cân tổng thể (BOP, triệu USD)', data: obArr, yAxisID: 'y',
                  backgroundColor: obArr.map(v => v >= 0 ? '#10b98180' : '#ef444480'),
                  borderColor: obArr.map(v => v >= 0 ? '#10b981' : '#ef4444'), borderWidth: 1.5,
                  datalabels: { color: '#e5e9f0', font: { size: 9, weight: '700' }, anchor: 'end',
                                align: (ctx) => (ctx.dataset.data[ctx.dataIndex] >= 0 ? 'end' : 'start'),
                                formatter: (v) => v.toLocaleString('vi-VN', { maximumFractionDigits: 0 }) } },
                { type: 'line', label: 'USD/VND tăng/giảm YoY tại cuối quý (%, thị trường thực)', data: usdvndArr,
                  yAxisID: 'y1', borderColor: '#f59e0b', borderWidth: 2.5, borderDash: [6, 4],
                  pointRadius: 3, pointBackgroundColor: '#f59e0b', fill: false, tension: 0.2, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(1) },
                { type: 'line', label: 'USD/VND mức thực tế tại cuối quý (VND, bình quân tháng)', data: usdvndLevelArr,
                  yAxisID: 'y2', borderColor: '#60a5fa', borderWidth: 2, pointRadius: 2,
                  pointBackgroundColor: '#60a5fa', fill: false, tension: 0.2, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(0) },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0 } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     title: { display: true, text: 'Triệu USD', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: '% YoY', color: '#9aa5bd', font: { size: 9 } } },
                y2: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM (user 2026-10-03): "chỉ vẽ thêm biểu đồ khác theo tháng nhé, chứ không chuyển từ biểu đồ
// quý kia sang tháng, sẽ dễ bị sai" — biểu đồ RIÊNG, KHÔNG đụng vào renderFxPressureSignalsChart ở
// trên (vẫn giữ nguyên quý). Dùng 2 chỉ báo CÓ SẴN THEO THÁNG THẬT (không resample/nội suy từ quý
// sang tháng): NEER Darvas (tỷ giá hiệu lực danh nghĩa, phản ánh VND so với CẢ RỔ đối tác thương
// mại, không chỉ riêng USD) + USD/VND tăng/giảm YoY. CHỦ ĐỘNG KHÔNG dùng forex_reserves_monthly
// (phát hiện 2026-10-03: điểm mới nhất 2026-06 = 86318 — SAI ĐƠN VỊ so với các điểm khác đều ~80-86
// TỶ USD, lỗi nằm ở DỮ LIỆU GỐC 40yo.vn [JSON thô, không qua regex parse nào ở code mình] — không
// tự đoán/sửa, chỉ loại khỏi chart này cho tới khi xác minh lại được).
// Chuỗi kỳ theo tuần ISO ('2026-W40') -> ngày thứ Sáu của tuần đó (ngày chạy cập nhật), để hiển thị theo ngày.
function _weekPeriodToDate(p) {
    const m = /^(\d{4})-W(\d{2})$/.exec(p);
    if (!m) return p;
    const year = +m[1], week = +m[2];
    const jan4 = new Date(Date.UTC(year, 0, 4));
    const mondayW1 = new Date(jan4); mondayW1.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() + 6) % 7));
    const friday = new Date(mondayW1); friday.setUTCDate(mondayW1.getUTCDate() + (week - 1) * 7 + 4);
    return friday.toISOString().slice(0, 10);
}

// THEM 2026-10-05 (user: "thêm biểu đồ lãi suất của Cake để xem lãi suất xu hướng thực") — mức lãi suất 12
// tháng theo kênh (snapshot mới nhất) + chuỗi thị trường cao nhất/bình quân. Kênh Cake và TCBS iPower chỉ
// có 1-2 điểm vì nguồn chỉ ghi khi giá đổi, nên không vẽ xu hướng cho 2 kênh này.
function renderDepositRateChart(indicators) {
    const card = document.getElementById('deposit-rate-card');
    const cLevels = document.getElementById('chart-deposit-rate-levels');
    const cTrend = document.getElementById('chart-deposit-rate-trend');
    if (!cLevels || !cTrend) return;
    const KEYS = [
        ['deposit_rate_12m_vcb', 'Vietcombank'], ['deposit_rate_12m_ctg', 'VietinBank'], ['deposit_rate_12m_nab', 'Nam A Bank'],
        ['deposit_rate_12m_market_avg', 'Thị trường (bình quân)'], ['deposit_rate_12m_market_max', 'Thị trường (cao nhất)'],
        ['deposit_rate_tcbs_ipower_max', 'TCBS iPower (kênh số)'], ['deposit_rate_cake_max', 'Cake (cơ bản + ưu đãi)'],
    ];
    const latest = KEYS.map(([k, lbl]) => {
        const sr = indicators[k] && indicators[k].series;
        return sr && sr.length ? { label: lbl, value: sr[sr.length - 1].value } : null;
    }).filter(Boolean);
    if (!latest.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    chartInstances.push(new Chart(cLevels, {
        type: 'bar',
        data: { labels: latest.map(x => x.label),
                datasets: [{ label: 'Lãi suất 12 tháng (%)', data: latest.map(x => x.value),
                             backgroundColor: latest.map(x => x.label.startsWith('Cake') || x.label.startsWith('TCBS') ? 'rgba(249,115,22,0.75)' : 'rgba(96,165,250,0.7)'),
                             borderWidth: 0 }] },
        options: { ...CHART_DEFAULTS,
                   plugins: { legend: { display: false } },
                   scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 25, autoSkip: false } },
                             y: { ...CHART_DEFAULTS.scales.y, beginAtZero: true } } },
        plugins: [ChartDataLabels],
    }));

    const mx = indicators['deposit_rate_12m_market_max'], av = indicators['deposit_rate_12m_market_avg'];
    const periods = mx ? mx.series.map(p => p.period) : [];
    const avBy = av ? Object.fromEntries(av.series.map(p => [p.period, p.value])) : {};
    const labels = periods.map(_weekPeriodToDate);
    chartInstances.push(new Chart(cTrend, {
        type: 'line',
        data: { labels,
                datasets: [
                    { label: 'Thị trường — cao nhất', data: mx.series.map(p => p.value), borderColor: '#f97316', backgroundColor: '#f97316',
                      borderWidth: 2, pointRadius: 3, tension: 0.2, datalabels: _endpointDatalabelsConfig(2) },
                    { label: 'Thị trường — bình quân', data: periods.map(p => avBy[p] ?? null), borderColor: '#60a5fa', backgroundColor: '#60a5fa',
                      borderWidth: 2, pointRadius: 3, tension: 0.2, spanGaps: true, datalabels: _endpointDatalabelsConfig(2) },
                ] },
        options: { ...CHART_DEFAULTS,
                   plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
                   scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 8 } },
                             y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: '%/năm', color: '#9aa5bd', font: { size: 9 } } } } },
        plugins: [ChartDataLabels],
    }));
}

// THEM 2026-10-05 (user): biểu đồ riêng chỉ có lãi suất Cake 12 tháng + ưu đãi, ghi theo tuần.
function renderDepositRateCakeChart(indicators) {
    const card = document.getElementById('deposit-rate-cake-card');
    const canvas = document.getElementById('chart-deposit-rate-cake');
    const sr = indicators['deposit_rate_cake_max'] && indicators['deposit_rate_cake_max'].series;
    if (!canvas || !sr || !sr.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';
    chartInstances.push(new Chart(canvas, {
        type: 'line',
        data: { labels: sr.map(p => p.period),
                datasets: [{ label: 'Cake: lãi 12 tháng + ưu đãi cao nhất (%/năm)', data: sr.map(p => p.value),
                             borderColor: '#f97316', backgroundColor: '#f97316', borderWidth: 2.5, pointRadius: 4,
                             tension: 0, stepped: true, datalabels: _endpointDatalabelsConfig(2) }] },
        options: { ...CHART_DEFAULTS,
                   plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
                   scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 8 } },
                             y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: '%/năm', color: '#9aa5bd', font: { size: 9 } } } } },
        plugins: [ChartDataLabels],
    }));
}

// THEM 2026-10-07: tab "Kinh tế Mỹ" — theo mô-tuýp cắt lớp: (1) headline, (2) composition,
// (3) breadth, (4) hàng hóa vs dịch vụ, (5) pipeline PPI/nhập khẩu → CPI, (6) tiêu dùng thực.
// Dữ liệu đã tính sẵn trong vimo.json → usMacro (us_macro_analysis.py), không tính lại ở web.
// Bản đồ nhiệt: màu theo YoY (đỏ = tăng nhanh, xanh = thấp/giảm). Quý cuối có thể chưa đủ 3 tháng.
// SUA 2026-10-07 (user: "cái nào tăng mạnh thì đỏ, giảm thì xanh dương nhạt, tăng thấp thì
// trắng, chuyển màu mượt từ đỏ sang trắng sang xanh dương") — thang PHÂN KỲ liên tục quanh 0%,
// không chia bậc cố định (bậc cũ 0/1/3/6 làm 2 giá trị gần nhau nhảy màu đột ngột). Mốc bão hoà
// ±US_HEATMAP_SCALE_PCT (giá trị vượt mốc vẫn giữ màu đậm nhất, không đậm hơn).
const US_HEATMAP_SCALE_PCT = 15;
function _usHeatmapColor(v) {
    if (v === null || v === undefined) return 'transparent';
    const white = [255, 255, 255];
    const red = [239, 68, 68], blue = [59, 130, 246];
    const target = v >= 0 ? red : blue;
    const t = Math.min(Math.abs(v) / US_HEATMAP_SCALE_PCT, 1);
    const mix = (w, c) => Math.round(w + (c - w) * t);
    return `rgba(${mix(white[0], target[0])},${mix(white[1], target[1])},${mix(white[2], target[2])},1)`;
}
// SUA 2026-10-07 (user: "Mỹ công bố CPI theo tháng, chỉnh lại được không") — đổi cột từ quý sang
// tháng, khớp đúng tần suất BLS công bố. ~140 cột nên cần cuộn ngang — dùng class
// "monitoring-table-scroll" để có UX cuộn ngang bằng lăn chuột (xem listener 'wheel' đầu file).
function usMacroHeatmap(hm) {
    if (!hm || !hm.months) return '';
    const head = hm.months.map(m => `<th>${m}</th>`).join('');
    const body = hm.rows.map(r => `<tr><th style="text-align:left;white-space:nowrap">${r.label}</th>` +
        r.values.map(v => `<td style="background:${_usHeatmapColor(v)};color:#0b1220;text-align:center;min-width:44px">${v === null || v === undefined ? '—' : v.toFixed(1)}</td>`).join('') + '</tr>').join('');
    return `<div class="monitoring-table-scroll" style="overflow-x:auto"><table class="monitoring-table"><thead><tr><th style="text-align:left">Nhóm</th>${head}</tr></thead><tbody>${body}</tbody></table></div>
            <p class="ind-source-note">Cột là tháng (đúng tần suất BLS công bố), hàng là nhóm; đỏ = YoY tăng mạnh, trắng = quanh 0%, xanh dương = giảm (màu bão hoà ở ±${US_HEATMAP_SCALE_PCT}%). Kéo/lăn chuột ngang để xem lịch sử. Trọng số đóng góp chưa có ở bảng này nên không cộng các ô thành CPI (xem mục 2c cho phần đóng góp theo trọng số).</p>`;
}

// THEM 2026-10-08 (user: "quan trọng nhất là mục so sánh CPI này thì so với tháng liền trước để
// xem biến động ngắn xu hướng như nào sẽ đúng hơn" — kiểm chứng bằng số liệu thật: CPI MoM Jun
// -0.42% kéo momentum 3 tháng xuống 0.18% dù Jul/Aug đã tăng tốc lại +0.07%→+0.4%, YoY/3M/6M năm
// hóa đều "san phẳng" nên KHÔNG bắt được pha đảo chiều này) — bảng + biểu đồ MoM từng tháng riêng,
// để tự nhìn xu hướng ngắn hạn thay vì chỉ tin 1 con số đã gộp.
function usMomTable(mh, f) {
    if (!mh || !mh.periods || !mh.periods.length) return '';
    const sign = v => (v === null || v === undefined) ? '' : (v >= 0 ? 'color:#ef4444' : 'color:#60a5fa');
    const cols = mh.periods.map((p, i) => `<th style="text-align:center">${p}</th>`).join('');
    const cpiRow = mh.cpi_mom.map(v => `<td style="text-align:center;${sign(v)}">${v === null ? '—' : (v >= 0 ? '+' : '') + f(v) + '%'}</td>`).join('');
    const coreRow = mh.core_mom.map(v => `<td style="text-align:center;${sign(v)}">${v === null ? '—' : (v >= 0 ? '+' : '') + f(v) + '%'}</td>`).join('');
    return `<div style="overflow-x:auto"><table class="monitoring-table"><thead><tr>
        <th style="text-align:left">MoM (SA)</th>${cols}
    </tr></thead><tbody>
        <tr><th style="text-align:left">CPI toàn phần</th>${cpiRow}</tr>
        <tr><th style="text-align:left">CPI lõi</th>${coreRow}</tr>
    </tbody></table></div>`;
}

// SUA 2026-10-07 (user: "nhìn chả hiểu gì... tôi muốn food làm CPI tăng bao nhiêu %, học phí tăng
// rất cao nhưng đóng góp ít vì tiêu dùng ít, năng lượng tăng nhẹ nhưng đóng góp nhiều vì tiêu dùng
// nhiều") — bỏ 3M/6M/Pressure (gây rối, không phải điều user hỏi), CHỈ giữ đúng 3 cái cần: Trọng
// số (đại diện "tiêu dùng bao nhiêu"), YoY (bản thân nhóm tự tăng bao nhiêu), Đóng góp (nhóm đó
// LÀM CPI đổi bao nhiêu điểm %) — 2 cột YoY và Đóng góp đặt CẠNH NHAU để thấy ngay sự khác biệt.
function usContribSnapshotTable(ct) {
    if (!ct || !ct.snapshot) return '';
    const f = (v, d = 2) => (v === null || v === undefined) ? '—' : Number(v).toFixed(d);
    const sign = v => (v === null || v === undefined) ? '' : (v >= 0 ? 'color:#ef4444' : 'color:#60a5fa');
    // THEM 2026-10-08 (user gửi tài liệu "Contribution change": "Energy contribution +0.3pp →
    // +0.8pp => đỏ; Shelter +1.2pp → +1.0pp => xanh" — biết nhóm nào đang TĂNG áp lực, không chỉ
    // mức đóng góp hiện tại).
    // SUA 2026-10-08 (user: "đã bảo phần này là thay đổi so với tháng trước, cứ đi so 3 tháng
    // trước thì sao mà đúng được" — đổi mốc so sánh từ 3 tháng xuống ĐÚNG 1 THÁNG TRƯỚC, khớp
    // us_macro_analysis.py:_cpi_contributions bản mới (contribution_chg_1m/contribution_1m_ago).
    const rows = ct.snapshot.map(r => `<tr>
        <th style="text-align:left;white-space:nowrap">${r.label}</th>
        <td style="color:#e5e7eb">${f(r.weight_pct, 1)}%</td>
        <td style="${sign(r.yoy)}">${f(r.yoy)}% <span class="ind-source-note">(bản thân nhóm tự tăng)</span></td>
        <td style="${sign(r.contribution)};font-weight:700">${f(r.contribution)}pp <span class="ind-source-note">(làm CPI đổi)</span></td>
        <td style="${sign(r.contribution_chg_1m)}">${r.contribution_chg_1m === null ? '—' : (r.contribution_chg_1m >= 0 ? '+' : '') + f(r.contribution_chg_1m) + 'pp'} <span class="ind-source-note">(tháng trước: ${f(r.contribution_1m_ago)}pp)</span></td>
    </tr>`).join('');
    return `<div style="overflow-x:auto"><table class="monitoring-table"><thead><tr>
        <th style="text-align:left">Nhóm (kỳ ${ct.snapshot_period})</th><th>Trọng số<br>(tiêu dùng chiếm)</th><th>YoY</th><th>Đóng góp</th><th>Thay đổi vs tháng trước</th>
    </tr></thead><tbody>${rows}</tbody></table></div>
    <p class="ind-source-note">Đóng góp = Trọng số × YoY. Nhóm trọng số lớn (vd Nhà ở ~35%) chỉ cần tăng nhẹ đã đóng góp nhiều; nhóm trọng số nhỏ (vd Giáo dục ~5,7%) dù tự tăng rất cao vẫn đóng góp ít — vì phần chi tiêu của người Mỹ dành cho nhóm đó nhỏ. Cột cuối: đóng góp đang TĂNG (đỏ, áp lực lên CPI từ nhóm này đang nặng thêm) hay GIẢM (xanh, đang hạ nhiệt) so với ĐÚNG 1 tháng trước — cửa sổ ngắn nhất để bắt biến động mới nhất.</p>`;
}

// SUA 2026-10-07 — biểu đồ cột NGANG cho 1 kỳ gần nhất, sắp theo |đóng góp| giảm dần, trả lời
// trực tiếp câu hỏi "CPI tăng X% tháng này thì cái gì gây ra, bao nhiêu điểm mỗi cái" — rõ hơn
// biểu đồ miền 10 năm (quá nhiều đường chồng lên nhau, khó đọc cho 1 kỳ cụ thể).
// SUA 2026-10-08 (user: "sao lại có cái này nhỉ tự nhiên nó lệch hẳn dòng nhìn xấu thế" — nhãn
// trục X ghép CẢ tên nhóm DÀI (vd "Giao thông (trừ xăng dầu — xem Năng lượng)") lẫn số liệu, với
// maxRotation:30 + autoSkip:false thì 2 nhãn dài liền kề cùng phải xuống 3 dòng, dòng cuối tràn ra
// ngoài khung canvas, đè lên bảng bên dưới — rút ngắn CHỈ nhãn trên trục X (bảng dưới vẫn giữ tên
// đầy đủ, đã có ô cảnh báo giải thích "Giao thông" ở trên rồi nên không cần lặp lại hết trên biểu đồ).
const _shortChartLabel = l => l.length > 16 ? l.slice(0, 14) + '…' : l;

function usContribBarChart(ct) {
    const canvas = document.getElementById('chart-us-contrib-bar');
    if (!canvas || !ct || !ct.snapshot) return;
    const rows = ct.snapshot.filter(r => r.contribution !== null);
    const colors = rows.map(r => r.contribution >= 0 ? 'rgba(239,68,68,0.8)' : 'rgba(96,165,250,0.8)');
    chartInstances.push(new Chart(canvas, {
        type: 'bar',
        data: { labels: rows.map(r => `${_shortChartLabel(r.label)}\n(YoY ${r.yoy >= 0 ? '+' : ''}${r.yoy.toFixed(1)}%)`),
                datasets: [{ label: 'Đóng góp vào CPI YoY (điểm %)', data: rows.map(r => r.contribution), backgroundColor: colors,
                             datalabels: { display: true, color: '#fff', anchor: 'end', align: 'top', offset: 2, clip: false,
                                           formatter: v => (v >= 0 ? '+' : '') + v.toFixed(2) + 'pp', font: { size: 10, weight: '600' } } }] },
        options: { ...CHART_DEFAULTS,
                   plugins: { legend: { display: false } },
                   scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, autoSkip: false, maxRotation: 30, font: { size: 8 } } },
                             y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Điểm % đóng góp vào CPI YoY', color: '#9aa5bd', font: { size: 9 } } } } },
        plugins: [ChartDataLabels],
    }));
}

// THEM 2026-10-08 (user gửi tài liệu đối chiếu BLS Table 6: "contribution YoY lớn không đồng
// nghĩa đang tăng mạnh — Nhà ở lớn chủ yếu vì trọng số, Energy/gasoline mới là cú kéo MoM mạnh
// nhất") — bảng SONG SONG với usContribSnapshotTable nhưng dùng MoM thay YoY, trả lời câu hỏi
// KHÁC: "CPI vừa tăng/giảm vì gì" (flow, bắt turning point) thay vì "CPI đang ở mức nào" (stock).
function usContribMomTable(ct, f) {
    if (!ct || !ct.mom_snapshot) return '';
    const sign = v => (v === null || v === undefined) ? '' : (v >= 0 ? 'color:#ef4444' : 'color:#60a5fa');
    const rows = ct.mom_snapshot.map(r => `<tr>
        <th style="text-align:left;white-space:nowrap">${r.label}</th>
        <td style="color:#e5e7eb">${r.weight_pct === null ? '—' : f(r.weight_pct, 1) + '%'}</td>
        <td style="${sign(r.mom)}">${r.mom === null ? '—' : (r.mom >= 0 ? '+' : '') + f(r.mom) + '%'} <span class="ind-source-note">(bản thân nhóm tự tăng/giảm tháng này)</span></td>
        <td style="${sign(r.contribution)};font-weight:700">${r.contribution === null ? '—' : (r.contribution >= 0 ? '+' : '') + f(r.contribution) + 'pp'} <span class="ind-source-note">(làm CPI đổi tháng này)</span></td>
    </tr>`).join('');
    return `<div style="overflow-x:auto"><table class="monitoring-table"><thead><tr>
        <th style="text-align:left">Nhóm (kỳ ${ct.snapshot_period})</th><th>Trọng số</th><th>MoM</th><th>Đóng góp MoM</th>
    </tr></thead><tbody>${rows}</tbody></table></div>`;
}

function usContribMomBarChart(ct) {
    const canvas = document.getElementById('chart-us-contrib-mom-bar');
    if (!canvas || !ct || !ct.mom_snapshot) return;
    const rows = ct.mom_snapshot.filter(r => r.contribution !== null);
    const colors = rows.map(r => r.contribution >= 0 ? 'rgba(239,68,68,0.8)' : 'rgba(96,165,250,0.8)');
    chartInstances.push(new Chart(canvas, {
        type: 'bar',
        data: { labels: rows.map(r => r.mom !== null ? `${_shortChartLabel(r.label)}\n(MoM ${r.mom >= 0 ? '+' : ''}${r.mom.toFixed(2)}%)` : _shortChartLabel(r.label)),
                datasets: [{ label: 'Đóng góp vào CPI MoM (điểm %)', data: rows.map(r => r.contribution), backgroundColor: colors,
                             datalabels: { display: true, color: '#fff', anchor: 'end', align: 'top', offset: 2, clip: false,
                                           formatter: v => (v >= 0 ? '+' : '') + v.toFixed(3) + 'pp', font: { size: 10, weight: '600' } } }] },
        options: { ...CHART_DEFAULTS,
                   plugins: { legend: { display: false } },
                   scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, autoSkip: false, maxRotation: 30, font: { size: 8 } } },
                             y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Điểm % đóng góp vào CPI MoM', color: '#9aa5bd', font: { size: 9 } } } } },
        plugins: [ChartDataLabels],
    }));
}

// THEM 2026-10-08 (user: "lấy dữ liệu bảng cân đối Fed, QE/QT, thu hẹp/mở rộng bảng cân đối") —
// bảng cân đối Fed (H.4.1, FRED mirror) KHÔNG chấm điểm "nới/thắt" — chỉ liệt kê số liệu thật +
// % thay đổi 6/12 tháng, đúng nguyên tắc không gộp nhiều tín hiệu thành 1 điểm số (đã chốt ở
// phần vĩ mô VN) — người đọc tự kết luận QE hay QT từ số liệu, không bị áp đặt.
function usFedLiquidityCard(liq, f) {
    const T = v => (v === null || v === undefined) ? '—' : (v / 1e6).toFixed(2) + 'T'; // triệu USD -> nghìn tỷ USD
    const pct = v => (v === null || v === undefined) ? '—' : (v >= 0 ? '+' : '') + v.toFixed(2) + '%';
    const sign = v => (v === null || v === undefined) ? 'color:#e5e7eb' : (v >= 0 ? 'color:#ef4444' : 'color:#60a5fa');
    const rows = [
        ['usm_fed_assets', 'Tổng tài sản'], ['usm_fed_treasury', '— Trái phiếu Chính phủ (Treasury)'],
        ['usm_fed_mbs', '— MBS (trái phiếu BĐS)'], ['usm_fed_reserves', 'Dự trữ ngân hàng tại Fed'],
        ['usm_fed_rrp', 'Reverse Repo (RRP)'], ['usm_fed_tga', 'Tài khoản Treasury (TGA)'],
        ['usm_fed_net_liquidity', 'Net Liquidity (= Tài sản − RRP − TGA)'],
    ];
    const body = rows.map(([k, lbl]) => `<tr>
        <th style="text-align:left;white-space:nowrap">${lbl}</th>
        <td style="color:#e5e7eb;font-weight:600">${T(liq.latest_values[k])}</td>
        <td style="${sign(liq.changes[k].chg_pct_6m)}">${pct(liq.changes[k].chg_pct_6m)}</td>
        <td style="${sign(liq.changes[k].chg_pct_12m)}">${pct(liq.changes[k].chg_pct_12m)}</td>
    </tr>`).join('');
    return `
        <div class="us-kpi-grid">
            <div class="us-kpi-tile"><div class="us-kpi-label">Tổng tài sản Fed</div><div class="us-kpi-value">${T(liq.latest_values.usm_fed_assets)}</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">So với đỉnh (${liq.peak_period})</div><div class="us-kpi-value">${pct(liq.assets_vs_peak_pct)}</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">Net Liquidity</div><div class="us-kpi-value">${T(liq.latest_values.usm_fed_net_liquidity)}</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">ECB tổng tài sản</div><div class="us-kpi-value">${liq.ecb_latest ? (liq.ecb_latest.value / 1e6).toFixed(2) + 'T€' : '—'}</div></div>
        </div>
        <div style="overflow-x:auto"><table class="monitoring-table"><thead><tr>
            <th style="text-align:left">Hạng mục (kỳ ${liq.latest})</th><th>Giá trị</th><th>6 tháng</th><th>12 tháng</th>
        </tr></thead><tbody>${body}</tbody></table></div>
        <div class="monitoring-table-scroll" style="overflow-x:auto">
            <div style="width:${Math.max(1100, liq.periods.length * 10)}px">
                <div class="bank-chart-grid-2">
                    <div class="ind-chart" style="height:300px"><canvas id="chart-us-fed-assets"></canvas></div>
                    <div class="ind-chart" style="height:300px"><canvas id="chart-us-fed-drains"></canvas></div>
                </div>
                <div class="ind-chart" style="height:280px;margin-top:10px"><canvas id="chart-us-fed-netliq"></canvas></div>
                <div class="ind-chart" style="height:260px;margin-top:10px"><canvas id="chart-us-fed-impulse"></canvas></div>
            </div>
        </div>
        <p class="ind-source-note">Nguồn: Fed H.4.1 qua FRED (cập nhật hàng tuần, lấy trung bình tháng). "Net Liquidity" là cách giới phân tích thị trường hay dùng (Tổng tài sản − RRP − TGA), KHÔNG phải định nghĩa chính thức của Fed — RRP và TGA là 2 "bể chứa" hút tiền ra khỏi hệ thống ngân hàng, dù Fed không đổi tổng tài sản. Treasury holdings đang TĂNG trong khi MBS vẫn giảm — không phải thuần QE hay thuần QT. "Liquidity Impulse" (biểu đồ cuối) = thay đổi HÀNG THÁNG của (Dự trữ ngân hàng − RRP − TGA) — khác Net Liquidity ở chỗ dùng Dự trữ (tiền thực sự nằm trong hệ thống ngân hàng) thay vì Tổng tài sản Fed, và nhìn vào TỐC ĐỘ thay đổi (dương = đang bơm ròng vào hệ thống tháng đó, âm = đang rút ròng) thay vì MỨC tuyệt đối.</p>`;
}

// THEM 2026-10-08 (user: "tiếp tục triển khai theo ma trận đã bàn" — Growth/Labor/Treasury&Credit/
// USD trong "US Macro Liquidity Matrix" user gửi). KHÔNG chấm điểm gộp — chỉ bảng/biểu đồ số liệu
// thô độc lập, giống cách làm card Fed liquidity ở trên. "Capital Flows" (TIC) không có trên FRED
// nên bỏ qua, xem ghi chú cuối renderUsMacro.
function usGrowthCard(g, f) {
    const gdp = g.gdp, ip = g.indpro;
    return `
        <div class="us-kpi-grid">
            ${gdp ? `<div class="us-kpi-tile"><div class="us-kpi-label">GDP thực YoY (${gdp.latest})</div><div class="us-kpi-value">${f(gdp.yoy)}%</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">GDP thực QoQ năm hóa</div><div class="us-kpi-value">${f(gdp.qoq_annualized)}%</div></div>` : ''}
            ${ip ? `<div class="us-kpi-tile"><div class="us-kpi-label">SX công nghiệp YoY (${ip.latest})</div><div class="us-kpi-value">${f(ip.yoy)}%</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">SX công nghiệp 3M năm hóa</div><div class="us-kpi-value">${f(ip.mom_3m_ann)}%</div></div>` : ''}
        </div>
        <div class="bank-chart-grid-2">
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-gdp"></canvas></div>
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-indpro"></canvas></div>
        </div>
        <p class="ind-source-note">Nguồn: FRED — GDPC1 (GDP thực, theo quý, SAAR), INDPRO (sản xuất công nghiệp, theo tháng). QoQ năm hóa = tăng trưởng quý so quý trước, quy ra tốc độ năm (cách BEA công bố GDP Mỹ chính thức).</p>`;
}

function usLaborCard(l, f) {
    return `
        <div class="us-kpi-grid">
            <div class="us-kpi-tile"><div class="us-kpi-label">Việc làm phi NN thêm (${l.latest_period})</div><div class="us-kpi-value">${f(l.payrolls_mom, 0)}k</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">— TB 3 tháng</div><div class="us-kpi-value">${f(l.payrolls_mom_3m_avg, 0)}k</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">Thất nghiệp</div><div class="us-kpi-value">${f(l.unemployment_latest)}%</div></div>
            ${l.openings ? `<div class="us-kpi-tile"><div class="us-kpi-label">JOLTS vị trí tuyển (${l.openings.period})</div><div class="us-kpi-value">${f(l.openings.latest / 1000, 2)}tr</div></div>` : ''}
            ${l.participation ? `<div class="us-kpi-tile"><div class="us-kpi-label">Tỷ lệ tham gia LLLĐ</div><div class="us-kpi-value">${f(l.participation.latest)}%</div></div>` : ''}
            <div class="us-kpi-tile"><div class="us-kpi-label">Lương bình quân giờ YoY</div><div class="us-kpi-value">${f(l.earnings_yoy)}%</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">— Lương THỰC YoY (trừ CPI)</div><div class="us-kpi-value">${l.real_wage_yoy === null ? '—' : f(l.real_wage_yoy) + '%'}</div></div>
            ${l.claims ? `<div class="us-kpi-tile"><div class="us-kpi-label">Trợ cấp TN lần đầu/tuần (TB ${l.claims.period})</div><div class="us-kpi-value">${f(l.claims.latest / 1000, 0)}k</div></div>` : ''}
        </div>
        <div class="bank-chart-grid-2">
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-payrolls"></canvas></div>
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-labor-other"></canvas></div>
        </div>
        <p class="ind-source-note">Nguồn: FRED — PAYEMS (việc làm phi NN), ICSA (trợ cấp TN lần đầu, TB tuần→tháng), JTSJOL (JOLTS), CIVPART, CES0500000003 (lương bình quân giờ), UNRATE. "Lương thực" = lương bình quân giờ YoY trừ CPI YoY — lương danh nghĩa tăng nhưng nếu thấp hơn lạm phát thì sức mua thực vẫn giảm.</p>`;
}

function usTreasuryCreditCard(tc, f) {
    const v = tc.latest_values, c = tc.changes;
    const bp = x => (x === null || x === undefined) ? '—' : (x >= 0 ? '+' : '') + Math.round(x * 100) + 'bp';
    const rows = [
        ['usm_yield_2y', 'Lợi suất 2 năm'], ['usm_yield_10y', 'Lợi suất 10 năm'],
        ['usm_spread_10y_2y', 'Chênh lệch 10Y-2Y'], ['usm_real_yield_10y', 'Lợi suất thực TIPS 10 năm'],
        ['usm_breakeven_10y', 'Lạm phát kỳ vọng hòa vốn 10 năm'],
        ['usm_hy_oas', 'Chênh lệch tín dụng High Yield (OAS)'], ['usm_ig_oas', 'Chênh lệch tín dụng Investment Grade (OAS)'],
    ];
    const body = rows.map(([k, lbl]) => `<tr>
        <th style="text-align:left;white-space:nowrap">${lbl}</th>
        <td style="color:#e5e7eb;font-weight:600">${f(v[k])}%</td>
        <td style="color:${c[k].chg_6m >= 0 ? '#ef4444' : '#60a5fa'}">${bp(c[k].chg_6m)}</td>
        <td style="color:${c[k].chg_12m >= 0 ? '#ef4444' : '#60a5fa'}">${bp(c[k].chg_12m)}</td>
    </tr>`).join('');
    return `
        <div style="overflow-x:auto"><table class="monitoring-table"><thead><tr>
            <th style="text-align:left">Hạng mục (kỳ ${tc.latest})</th><th>Giá trị</th><th>6 tháng</th><th>12 tháng</th>
        </tr></thead><tbody>${body}</tbody></table></div>
        <div class="bank-chart-grid-2" style="margin-top:10px">
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-curve"></canvas></div>
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-credit"></canvas></div>
        </div>
        <p class="ind-source-note">Nguồn: FRED — DGS2/DGS10 (lợi suất danh nghĩa), T10Y2Y (đường cong đảo khi &lt;0 — tín hiệu suy thoái kinh điển), DFII10 (lợi suất thực TIPS), T10YIE (breakeven = kỳ vọng lạm phát thị trường trái phiếu định giá), BAMLH0A0HYM2/BAMLC0A0CM (OAS — spread tín dụng rộng ra khi thị trường lo ngại rủi ro vỡ nợ doanh nghiệp). ⚠ 2 chuỗi OAS chỉ có từ 2023-10 trên FRED — do ICE giới hạn cấp phép, FRED tự ghi rõ "chỉ giữ 3 năm dữ liệu gần nhất", KHÔNG phải lỗi tải.</p>`;
}

function usCapitalFlowsCard(cf, f) {
    return `
        <div class="us-kpi-grid">
            <div class="us-kpi-tile"><div class="us-kpi-label">Tổng nước ngoài nắm giữ (${cf.latest})</div><div class="us-kpi-value">${f(cf.total_latest / 1000, 2)}T$</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">— 6 tháng</div><div class="us-kpi-value">${f(cf.total_chg_pct_6m)}%</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">— 12 tháng</div><div class="us-kpi-value">${f(cf.total_chg_pct_12m)}%</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">Khối chính thức/NHTW</div><div class="us-kpi-value">${cf.official_latest === null || cf.official_latest === undefined ? '—' : f(cf.official_latest / 1000, 2) + 'T$'}</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">Khối tư nhân (= Tổng − Chính thức)</div><div class="us-kpi-value">${cf.private_latest === null || cf.private_latest === undefined ? '—' : f(cf.private_latest / 1000, 2) + 'T$'}</div></div>
        </div>
        <div class="ind-chart" style="height:280px"><canvas id="chart-us-tic"></canvas></div>
        <p class="ind-source-note">Nguồn: Bộ Tài chính Mỹ — TIC (Treasury International Capital System), báo cáo "Major Foreign Holders of Treasury Securities" (ticdata.treasury.gov) — KHÔNG có trên FRED, độ trễ công bố ~2 tháng. "Khối chính thức" = NHTW/chính phủ nước ngoài (phản ánh hành vi dự trữ ngoại hối quốc gia — Trung Quốc/Nhật giảm nắm giữ thường được đọc là tín hiệu địa chính trị/đa dạng hóa dự trữ); "Khối tư nhân" = quỹ đầu tư/doanh nghiệp/cá nhân nước ngoài, mang tính đầu cơ/tìm lợi suất nhiều hơn.</p>`;
}

function usUsdCard(u, f) {
    return `
        <div class="us-kpi-grid">
            <div class="us-kpi-tile"><div class="us-kpi-label">Chỉ số USD Broad (${u.latest})</div><div class="us-kpi-value">${f(u.latest_value)}</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">6 tháng</div><div class="us-kpi-value">${f(u.chg_pct_6m)}%</div></div>
            <div class="us-kpi-tile"><div class="us-kpi-label">12 tháng</div><div class="us-kpi-value">${f(u.chg_pct_12m)}%</div></div>
        </div>
        <div class="ind-chart" style="height:260px"><canvas id="chart-us-dxy"></canvas></div>
        <p class="ind-source-note">Nguồn: FRED — DTWEXBGS (Trade Weighted US Dollar Index: Broad, Goods and Services, 2006=100). USD mạnh lên thường gây áp lực giảm giá hàng hóa định giá bằng USD (dầu, vàng) và hút vốn khỏi thị trường mới nổi (ảnh hưởng tỷ giá VND gián tiếp).</p>`;
}

// THEM 2026-10-08 (user: "không nên làm chấm điểm gộp... làm đánh giá từng chỉ tiêu để đánh giá
// trạng thái") — badge trạng thái RULE-BASED (us_macro_analysis.py:_assess_states), ĐỘC LẬP theo
// từng nhóm, KHÔNG cộng dồn thành điểm tổng. Màu chỉ phản ánh trạng thái CỦA RIÊNG nhóm đó.
function usStateBadge(s) {
    if (!s) return '';
    const colors = { good: '#10b981', neutral: '#9aa5bd', warn: '#f97316', bad: '#ef4444' };
    const c = colors[s.color] || '#9aa5bd';
    return `<div style="background:${c}1a;border:1px solid ${c};border-radius:8px;padding:8px 12px;margin-bottom:12px">
        <span style="color:${c};font-weight:700">● ${s.label}</span>
        <div style="color:#cbd5e1;font-size:0.85em;margin-top:3px">${s.detail}</div>
    </div>`;
}

// THEM 2026-10-08 (user: "có phần 1 khu vực để đánh giá tổng quan chưa, tôi muốn có phần đó để
// đọc xem vĩ mô đang có vấn đề gì, trạng thái ra sao, tôi cũng không thấy bổ sung thêm các chart
// minh họa, hoặc bảng dữ liệu minh họa") — 1 bảng tổng quan Ở ĐẦU tab, liệt kê lại 7 trạng thái
// (usm.states, đã có sẵn ở từng thẻ chi tiết bên dưới) CẠNH NHAU để quét nhanh toàn cảnh + 1
// sparkline nhỏ minh họa xu hướng gần đây mỗi nhóm. CHỈ LIỆT KÊ/LỌC lại các trạng thái ĐỘC LẬP đã
// có — KHÔNG tính thêm bất kỳ con số tổng hợp/điểm số mới nào (đúng nguyên tắc đã chốt).
function usOverviewSection(usm, f) {
    // SUA 2026-10-08 (user gửi tài liệu "3 câu hỏi độc lập" + tách Fed Policy/Balance Sheet, tách
    // mức lạm phát/độ dai dẳng) — 9 dòng thay vì 7, khớp us_macro_analysis.py:_assess_states bản
    // mới (fed_policy và fed_balance_sheet giờ là 2 state riêng; inflation_persistence tách khỏi
    // inflation).
    // THEM 2026-10-08 (user gửi tài liệu "Turning point": "không chỉ hỏi hiện tại tốt/xấu, mà hỏi
    // đang tốt lên hay xấu đi" — mỗi dòng thêm `trendData` = ĐÚNG mảng dữ liệu dùng cho sparkline
    // (xem mkSpark bên dưới), so giá trị mới nhất với 3 kỳ trước để ra mũi tên ↑/↓/→. CHỈ LÀ MŨI
    // TÊN MÔ TẢ (số đang tăng/giảm), KHÔNG tự gán "tăng = tốt" hay "giảm = xấu" — vì nhiều chỉ số
    // (USD, lợi suất, dòng vốn...) không có 1 chiều "tốt/xấu" cố định, tùy bối cảnh — người đọc tự
    // kết hợp với badge trạng thái (đã có) để luận ra tốt lên hay xấu đi.
    const rows = [
        { key: 'growth', label: 'Tăng trưởng (GDP)', metric: usm.growth && usm.growth.gdp ? `GDP QoQ năm hóa: ${f(usm.growth.gdp.qoq_annualized)}%` : '—', spark: 'spark-growth',
          trendData: usm.growth && usm.growth.gdp ? usm.growth.history.gdp.qoq_ann : null },
        { key: 'labor', label: 'Lao động', metric: usm.labor ? `Việc làm thêm TB 3T: ${f(usm.labor.payrolls_mom_3m_avg, 0)}k` : '—', spark: 'spark-labor',
          trendData: usm.labor ? usm.labor.history.payrolls_mom_3m_avg : null },
        { key: 'inflation', label: 'Lạm phát — mức độ (CPI)', metric: `CPI YoY: ${f(usm.headline.cpi_yoy)}%`, spark: 'spark-inflation',
          trendData: usm.history.cpi_yoy },
        { key: 'inflation_persistence', label: 'Lạm phát — độ dai dẳng/lan truyền', metric: usm.breadth ? `${f(usm.breadth.pct_gt_3, 0)}% nhóm CPI tăng &gt;3%` : '—', spark: 'spark-persistence',
          trendData: usm.history.breadth_gt3_pct },
        { key: 'fed_policy', label: 'Fed — chính sách lãi suất', metric: usm.rates.usm_fed_funds ? `Lãi suất quỹ LB: ${f(usm.rates.usm_fed_funds.latest)}%` : '—', spark: 'spark-fed',
          trendData: usm.rates.usm_fed_funds && usm.rates.usm_fed_funds.history ? usm.rates.usm_fed_funds.history.values : null },
        { key: 'fed_balance_sheet', label: 'Fed — bảng cân đối (QE/QT)', metric: usm.liquidity ? `Liquidity Impulse: ${f(usm.liquidity.liquidity_impulse_latest / 1000, 1)} tỷ$` : '—', spark: 'spark-fedbs',
          trendData: usm.liquidity ? usm.liquidity.history.liquidity_impulse : null },
        { key: 'treasury_credit', label: 'Lợi suất & tín dụng', metric: usm.treasury_credit ? `10Y-2Y: ${f(usm.treasury_credit.latest_values.usm_spread_10y_2y)}%` : '—', spark: 'spark-credit',
          trendData: usm.treasury_credit ? usm.treasury_credit.history.spread_10y_2y : null },
        { key: 'usd', label: 'USD', metric: usm.usd ? `Chỉ số Broad: ${f(usm.usd.latest_value)}` : '—', spark: 'spark-usd',
          trendData: usm.usd ? usm.usd.history.values : null },
        { key: 'capital_flows', label: 'Capital Flows (TIC)', metric: usm.capital_flows ? `Tổng NN nắm giữ: ${f(usm.capital_flows.total_latest / 1000, 2)}T$` : '—', spark: 'spark-capflows',
          trendData: usm.capital_flows ? usm.capital_flows.history.total : null },
    ];
    const colors = { good: '#10b981', neutral: '#9aa5bd', warn: '#f97316', bad: '#ef4444' };
    const trendArrow = (arr, n = 4) => {
        if (!arr || arr.length < n) return '';
        const last = arr[arr.length - 1], prev = arr[arr.length - n];
        if (last === null || last === undefined || prev === null || prev === undefined) return '';
        const diff = last - prev;
        if (Math.abs(diff) < Math.abs(last || 1) * 0.005) return '→';
        return diff > 0 ? '↑' : '↓';
    };
    // THEM 2026-10-08 (user: "3 câu hỏi độc lập: A. Kinh tế khỏe/yếu? B. Tiền tệ nới/thắt? C. Môi
    // trường có thuận lợi cho tài sản rủi ro không? — 3 cái có thể cho 3 kết quả khác nhau") — 3
    // đoạn văn NGẮN, ĐỘC LẬP (us_macro_analysis.py:_build_synthesis — ghép câu rule-based từ các
    // state đã có, KHÔNG phải điểm số), để RIÊNG, KHÔNG gộp lại thành 1 kết luận chung.
    const synth = usm.synthesis;
    // THEM 2026-10-08 (user: "tạo cho tôi 1 vị trí để đánh giá tổng quát lại toàn bộ các chỉ số vĩ
    // mô đang nói lên xấu hay tốt, và câu kết luận là có thuận cho đầu tư hay không" — KHÁC "chấm
    // điểm gộp" đã từ chối: đây là 1 ĐOẠN VĂN kết luận rule-based (us_macro_analysis.py:
    // _build_synthesis -> "overall"), chỉ xem 3 nhóm A/B/C bên dưới đang nghiêng tốt/xấu/hỗn hợp
    // rồi viết câu kết luận tương ứng — KHÔNG cộng điểm số nào. Đặt NỔI BẬT ở đầu, A/B/C bên dưới
    // giải thích "vì sao" ra kết luận đó.
    const overallHtml = synth && synth.overall ? `
        <div style="background:${colors[synth.overall.color]}1a;border:2px solid ${colors[synth.overall.color]};border-radius:10px;padding:14px 16px;margin-bottom:16px">
            <div style="font-size:0.75em;color:#9aa5bd;text-transform:uppercase;letter-spacing:0.5px">Kết luận tổng quát — có thuận lợi cho đầu tư không?</div>
            <div style="color:${colors[synth.overall.color]};font-weight:800;font-size:1.25em;margin-top:4px">${synth.overall.label}</div>
            <div style="color:#e5e7eb;font-size:0.88em;margin-top:6px;line-height:1.5">${synth.overall.text}</div>
        </div>` : '';
    const synthHtml = synth ? `
        ${overallHtml}
        <div class="us-kpi-grid" style="grid-template-columns:repeat(auto-fit,minmax(280px,1fr))">
            <div class="us-kpi-tile" style="text-align:left"><div class="us-kpi-label">A. Nền kinh tế đang khỏe hay yếu?</div><div style="color:#e5e7eb;font-size:0.88em;margin-top:4px;line-height:1.5">${synth.economic}</div></div>
            <div class="us-kpi-tile" style="text-align:left"><div class="us-kpi-label">B. Chính sách tiền tệ đang nới hay thắt?</div><div style="color:#e5e7eb;font-size:0.88em;margin-top:4px;line-height:1.5">${synth.monetary}</div></div>
            <div class="us-kpi-tile" style="text-align:left"><div class="us-kpi-label">C. Môi trường có thuận lợi cho tài sản rủi ro?</div><div style="color:#e5e7eb;font-size:0.88em;margin-top:4px;line-height:1.5">${synth.investment}</div></div>
        </div>
        <p class="ind-source-note">3 câu hỏi để RIÊNG vì có thể cho 3 kết quả khác nhau — vd kinh tế vẫn khỏe nhưng Fed vẫn phải thắt chặt vì lạm phát, không có nghĩa "tốt" ở câu A thì "tốt" luôn ở câu B/C. "Kết luận tổng quát" ở trên là ĐỌC TỔNG HỢP có chủ đích từ 3 câu này, KHÔNG phải cộng điểm số.</p>` : '';
    const body = rows.map(r => {
        const s = usm.states && usm.states[r.key];
        const c = s ? (colors[s.color] || '#9aa5bd') : '#9aa5bd';
        const arrow = trendArrow(r.trendData);
        return `<tr>
            <th style="text-align:left;white-space:nowrap">${r.label}</th>
            <td style="color:${c};font-weight:700;white-space:nowrap">● ${s ? s.label : '—'}</td>
            <td style="color:#e5e7eb;white-space:nowrap">${r.metric}${arrow ? ` <span style="color:#9aa5bd;font-weight:700" title="So với ~3 kỳ gần nhất trước đó">${arrow}</span>` : ''}</td>
            <td style="width:110px"><div style="width:100px;height:32px"><canvas id="${r.spark}"></canvas></div></td>
        </tr>`;
    }).join('');
    const watch = rows.filter(r => usm.states && usm.states[r.key] && ['warn', 'bad'].includes(usm.states[r.key].color));
    const watchHtml = watch.length ? `
        <p style="margin-top:14px;margin-bottom:4px;font-weight:600;color:#f97316">⚠ Cần chú ý (${watch.length}/${rows.length} nhóm):</p>
        <ul class="ind-source-note" style="line-height:1.8">${watch.map(r => `<li><b>${r.label}</b>: ${usm.states[r.key].label} — ${usm.states[r.key].detail}</li>`).join('')}</ul>`
        : `<p style="margin-top:14px;color:#10b981">✓ Không nhóm nào đang ở trạng thái cảnh báo theo ngưỡng rule-based hiện tại.</p>`;
    return `
        ${synthHtml}
        <div style="overflow-x:auto;margin-top:14px"><table class="monitoring-table"><thead><tr>
            <th style="text-align:left">Nhóm</th><th>Trạng thái</th><th>Chỉ số chính</th><th>Xu hướng gần đây</th>
        </tr></thead><tbody>${body}</tbody></table></div>
        ${watchHtml}
        <p class="ind-source-note">Bảng là LIỆT KÊ lại ${rows.length} trạng thái độc lập đã đánh giá chi tiết ở các mục bên dưới — KHÔNG cộng dồn/tính điểm tổng. Mỗi nhóm đứng riêng, tự đọc theo đúng bối cảnh của nó (vd USD mạnh không "tốt" hay "xấu" per se, chỉ là 1 sự kiện cần biết). Mũi tên cạnh "Chỉ số chính" chỉ MÔ TẢ số liệu đang tăng (↑)/giảm (↓)/đi ngang (→) so ~3 kỳ gần nhất trước đó — KHÔNG tự gán tăng=tốt hay giảm=xấu, vì nhiều chỉ số (USD, lợi suất...) không có 1 chiều tốt/xấu cố định.</p>`;
}

function renderUsMacro(usm) {
    const box = document.getElementById('us-macro-container');
    if (!box || !usm) return;
    box.style.display = '';
    const f = (v, d = 2) => (v === null || v === undefined) ? '—' : Number(v).toFixed(d);
    const h = usm.headline, rc = usm.real_consumption, rates = usm.rates;
    // SUA 2026-10-09 (user: "ở tab kinh tế mỹ cũng cho copy đi, gộp các mục to to vào nhé, chứ để
    // mỗi cái bé tí vào thì không đẹp" — tab này có ~18 card rời rạc (Tổng quan, Tóm tắt, 1, 2,
    // 2b...2e, 3-6, 7-12...), gộp theo CHỦ ĐỀ thành 7 nhóm lớn: Tổng quan / Lạm phát-Headline / Lạm
    // phát-Đóng góp / Lạm phát-Lan tỏa&Pipeline / Thanh khoản Fed&Lãi suất / Tăng trưởng&Lao động /
    // Lợi suất-Tín dụng&Dòng vốn — mỗi nhóm 1 nút copy ảnh (dùng chung saveVnReportSection() đã viết
    // cho tab Báo cáo), tránh vừa quá nhiều nút bé vừa tránh gộp hết thành 1 ảnh siêu dài (dễ bị
    // nén mờ khi chia sẻ, xem bài học ở tab Báo cáo).
    let usSectionSeq = 0;
    const group = (title, bodyHtml) => {
        if (!bodyHtml) return '';
        usSectionSeq += 1;
        const id = `us-report-section-${usSectionSeq}`;
        return `<div class="card margin-top-20" id="${id}">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px">
                <h3 class="border-blue" style="margin:0">${title}</h3>
                <button class="vn-report-no-capture" onclick="saveVnReportSection('${id}', this)"
                    title="Lưu/Copy riêng mục này thành ảnh"
                    style="flex-shrink:0;background:#0ea5e9;border:none;color:#fff;font-weight:600;font-size:0.78em;padding:6px 10px;border-radius:7px;cursor:pointer">
                    📷 Copy ảnh mục này
                </button>
            </div>
            ${bodyHtml}
        </div>`;
    };
    const sub = (title, body) => `<h4 style="margin:18px 0 8px;border-top:1px solid rgba(255,255,255,0.08);padding-top:14px">${title}</h4>${body}`;
    const card = (title, body) => body ? `<div>${sub(title, body)}</div>` : '';
    const kpi = (lbl, val) => `<div class="us-kpi-tile"><div class="us-kpi-label">${lbl}</div><div class="us-kpi-value">${val}</div></div>`;
    // SUA 2026-10-08 (user gửi tài liệu: "Fed không điều hành theo CPI 2%, Fed nhắm mục tiêu PCE" —
    // thêm PCE/PCE lõi YoY, trước đây usm_pce_price đã fetch nhưng CHƯA từng tính YoY/hiển thị ở
    // đâu; cũng thêm CPI MoM/lõi MoM — trước đây chỉ có YoY/3M/6M năm hóa, thiếu đúng "biến động
    // mới nhất so tháng liền trước" mà user yêu cầu).
    const kpis = [
        kpi('CPI YoY (%)', f(h.cpi_yoy)), kpi('CPI MoM (%)', f(h.cpi_mom)),
        kpi('CPI lõi YoY (%)', f(h.core_yoy)), kpi('CPI lõi MoM (%)', f(h.core_mom)),
        kpi('CPI 3 tháng, năm hóa (%)', f(h.cpi_3m_ann)), kpi('CPI 6 tháng, năm hóa (%)', f(h.cpi_6m_ann)),
        kpi('PCE YoY (%, Fed ưu tiên)', f(h.pce_yoy)), kpi('PCE lõi YoY (%, Fed ưu tiên)', f(h.pce_core_yoy)),
        kpi('CPI hàng hóa YoY (%)', f(h.goods_yoy)), kpi('CPI dịch vụ YoY (%)', f(h.services_yoy)),
        kpi('Bán lẻ danh nghĩa YoY (%)', f(rc.retail_nominal_yoy)), kpi('Chi tiêu thực PCE YoY (%)', f(rc.pce_real_yoy)),
        kpi('Lãi suất quỹ liên bang (%)', f(rates.usm_fed_funds && rates.usm_fed_funds.latest)),
        kpi('Lợi suất 10 năm (%)', f(rates.usm_yield_10y && rates.usm_yield_10y.latest)),
        kpi('Thất nghiệp (%)', f(rates.usm_unemployment && rates.usm_unemployment.latest)),
    ].join('');
    box.innerHTML = `
      ${group('📊 Tổng quan — toàn bộ ma trận vĩ mô Mỹ (9 nhóm)', usOverviewSection(usm, f))}
      ${group('🇺🇸 Lạm phát Mỹ — Headline & Cấu phần', `
      ${card('Tóm tắt — lạm phát Mỹ tới ' + h.latest, `<div class="us-kpi-grid">${kpis}</div>
            <p class="ind-source-note">Nguồn: FRED (BLS, BEA, Fed). Số liệu gốc theo tháng, từ 2015. "Năm hóa" tính từ MoM đã điều chỉnh mùa vụ. ⚠ Fed KHÔNG điều hành theo CPI — mục tiêu lạm phát 2% chính thức của Fed tính theo PCE (ưu tiên PCE lõi), CPI chỉ là thước đo tham chiếu phổ biến hơn với công chúng. Dữ liệu PCE công bố TRỄ hơn CPI ~2-4 tuần (kỳ PCE mới nhất: ${h.pce_latest || '—'}).</p>`)}
      ${card('1. Headline & lõi — CPI tăng hay giảm, lõi có dai dẳng không', `${usStateBadge(usm.states && usm.states.inflation)}${usStateBadge(usm.states && usm.states.inflation_persistence)}
            <p style="font-weight:600;margin:4px 0 8px">MoM theo từng tháng — xem xu hướng NGẮN HẠN (YoY/năm hóa có thể "san phẳng" 1 tháng vừa đảo chiều)</p>
            <div class="ind-chart" style="height:220px"><canvas id="chart-us-mom"></canvas></div>
            ${usMomTable(h.mom_history, f)}
            <div class="ind-chart" style="height:300px;margin-top:14px"><canvas id="chart-us-headline"></canvas></div>`)}
      ${card('2. Cấu phần — nhóm nào kéo CPI (YoY hiện tại, %)', `<div class="ind-chart" style="height:300px"><canvas id="chart-us-groups"></canvas></div>
            <p class="ind-source-note">Đây là mức tự tăng/giảm (YoY) của riêng từng nhóm — CHƯA nhân trọng số. Xem mục 2c để biết mỗi nhóm LÀM CPI đổi bao nhiêu điểm % (nhóm trọng số nhỏ dù tự tăng cao vẫn đóng góp ít).</p>`)}
      ${card('2b. Bản đồ nhiệt theo tháng — cơ cấu CPI biến động thế nào (YoY, %)', usMacroHeatmap(usm.heatmap))}`)}
      ${group('🇺🇸 Lạm phát Mỹ — Đóng góp & Turning point', `
      ${card('2c. CPI ĐANG Ở MỨC NÀO (YoY, tích lũy 12 tháng) — tại tháng ' + usm.contributions.snapshot_period, `
            <p style="background:rgba(249,115,22,0.12);border:1px solid rgba(249,115,22,0.4);border-radius:8px;padding:8px 12px;font-size:0.85em;margin:0 0 10px">
                ⚠ <b>"Giao thông" ở BẢNG NÀY khác "Giao thông" ở bảng nhiệt 2b phía trên</b> (YoY ~${f((usm.contributions.snapshot.find(r => r.key === 'usm_cpi_transport') || {}).yoy)}% ở đây so với ~${f(usm.groups.find(g => g.key === 'usm_cpi_transport').yoy)}% ở bảng nhiệt) — xăng dầu đã TÁCH RA, cộng gộp vào "Năng lượng" (YoY ~${f((usm.contributions.snapshot.find(r => r.key === 'usm_cpi_energy') || {}).yoy)}%, nay KHỚP ĐÚNG bảng nhiệt 2b) để không đếm trùng khi cộng thành đóng góp.
            </p>
            <div class="ind-chart" style="height:320px"><canvas id="chart-us-contrib-bar"></canvas></div>
            ${usContribSnapshotTable(usm.contributions)}
            <p class="ind-source-note">Trọng số lấy 1 lần từ BLS (${usm.contributions.weights_source}) — www.bls.gov chặn fetch tự động (403) nên KHÔNG tự cập nhật theo lịch. ⚠ Bảng này trả lời "CPI 3,4% hiện tại được TẠO NÊN từ đâu" (cộng dồn 12 tháng) — KHÔNG phải "CPI tháng này vừa tăng/giảm vì gì". Đóng góp YoY lớn (vd Nhà ở) có thể chỉ do TRỌNG SỐ lớn, không có nghĩa nhóm đó đang là động lực tăng tốc — xem mục 2c-mom bên dưới.</p>`)}
      ${card('2c-mom. CPI VỪA TĂNG/GIẢM VÌ GÌ (MoM, chỉ 1 tháng) — tháng ' + usm.contributions.snapshot_period, `
            <p style="background:rgba(96,165,250,0.12);border:1px solid rgba(96,165,250,0.4);border-radius:8px;padding:8px 12px;font-size:0.85em;margin:0 0 10px">
                💡 Khác câu hỏi ở mục 2c (CPI đang ở mức nào, cộng dồn 12 tháng) — bảng này trả lời "cú tăng/giảm MỚI NHẤT của CPI đến từ đâu", dùng để bắt turning point. Vd thực tế: Nhà ở đóng góp YoY lớn nhất (2c) nhưng ở ĐÂY lại không phải động lực MoM mạnh nhất — vì bản thân Nhà ở chỉ tăng nhẹ, đóng góp lớn chủ yếu do trọng số ~35%; Năng lượng/xăng dầu mới là cú kéo MoM mạnh nhất tháng này dù trọng số nhỏ hơn nhiều.
            </p>
            <div class="ind-chart" style="height:320px"><canvas id="chart-us-contrib-mom-bar"></canvas></div>
            ${usContribMomTable(usm.contributions, f)}
            <p class="ind-source-note">Đóng góp MoM = Trọng số × MoM. Cộng đủ 9 nhóm + phần dư = ĐÚNG CPI MoM ${f(usm.contributions.cpi_mom_actual)}% tháng ${usm.contributions.snapshot_period}.</p>`)}
      ${card('2d. Đóng góp theo thời gian (biểu đồ cột chồng, 2016 tới nay)', `
            <div class="monitoring-table-scroll" style="overflow-x:auto">
                <div style="width:${Math.max(1100, usm.contributions.periods.length * 14 + 40)}px"><div class="ind-chart" style="height:420px"><canvas id="chart-us-contrib"></canvas></div></div>
            </div>
            <p class="ind-source-note">Cộng 9 nhóm + phần dư (Fuel oil + mục nhỏ chưa gán) = ĐÚNG CPI YoY. Kéo/lăn chuột ngang để xem lịch sử — mặc định hiện tháng gần nhất. Trục Y dùng CHUNG cho cả giai đoạn (kể cả đỉnh lạm phát 2021-2022) nên các tháng gần đây (CPI thấp hơn nhiều) nhìn cột bị "lùn" đi — xem biểu đồ 2e bên dưới để phóng to riêng giai đoạn gần nhất.</p>`)}
      ${card('2e. Đóng góp 24 tháng gần nhất (phóng to, trục Y riêng)', `
            <div class="ind-chart" style="height:380px"><canvas id="chart-us-contrib-recent"></canvas></div>
            <p class="ind-source-note">Giống hệt dữ liệu ở mục 2d, chỉ CẮT RIÊNG 24 tháng gần nhất và để Chart.js tự giãn trục Y theo đúng biên độ của riêng giai đoạn này — dễ đọc độ cao từng cấu phần hơn khi không bị đỉnh lạm phát 2021-2022 "đè" cho thấp xuống.</p>`)}`)}
      ${group('🇺🇸 Lạm phát Mỹ — Độ lan tỏa & Pipeline', `
      ${card('3. Độ lan tỏa — bao nhiêu nhóm đang tăng nhanh', `<div class="ind-chart" style="height:260px"><canvas id="chart-us-breadth"></canvas></div>
            <p class="ind-source-note">Hiện: ${f(usm.breadth.pct_gt_3, 0)}% nhóm có YoY &gt; 3%; ${f(usm.breadth.pct_gt_5, 0)}% nhóm &gt; 5%; ${f(usm.breadth.pct_rising_mom, 0)}% nhóm đang tăng MoM.</p>`)}
      ${card('4. Hàng hóa vs dịch vụ — dịch vụ bền, hàng hóa biến động', `<div class="ind-chart" style="height:260px"><canvas id="chart-us-goods-services"></canvas></div>`)}
      ${card('5. Pipeline — giá sản xuất, nhập khẩu, dầu → CPI', `<div class="bank-chart-grid-2">
            <div class="ind-chart" style="height:280px"><canvas id="chart-us-pipeline"></canvas></div>
            <div class="ind-chart" style="height:280px"><canvas id="chart-us-leadlag"></canvas></div></div>
            <p class="ind-source-note">Trái: PPI cầu cuối cùng, giá nhập khẩu, dầu WTI (YoY). Phải: tương quan MoM giữa PPI (hoặc giá nhập khẩu) ở tháng t−L và CPI hàng hóa ở tháng t. Tương quan KHÔNG phải nhân quả.</p>`)}
      ${card('5b. Crack spread — biên lọc dầu: diesel, xăng, 3-2-1 ($/thùng)', `<div class="ind-chart" style="height:280px"><canvas id="chart-us-crack"></canvas></div>
            <p class="ind-source-note">Crack = giá sản phẩm (Vịnh Mexico, EIA) × 42 − WTI. Crack 3-2-1 = (2 × xăng + diesel)/3 − WTI. Hiện: diesel ${f(usm.cracks.latest.diesel_crack, 1)} $/thùng, 3-2-1 ${f(usm.cracks.latest.crack_321, 1)} $/thùng (${usm.cracks.latest.period}). Crack cao cho thấy lọc dầu đang bán sản phẩm đắt hơn nhiều so với giá dầu thô — đúng kiểu cú sốc diesel mà bài viết nêu.</p>`)}
      ${card('5c. Lan truyền lạm phát — cú sốc đang cô lập hay đang truyền đi?', `<div class="bank-chart-grid-2">
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-transmission-oil"></canvas></div>
            <div class="ind-chart" style="height:260px"><canvas id="chart-us-transmission-wage"></canvas></div></div>
            <p class="ind-source-note">Trái: tương quan MoM giữa giá dầu WTI ở tháng t−L và PPI cầu cuối ở tháng t (tầng ĐẦU chuỗi truyền dẫn, trước khi vào PPI/CPI). Phải: tương quan MoM giữa lương bình quân giờ ở tháng t−L và CPI dịch vụ ở tháng t (tín hiệu "wage-price spiral" nếu có — lưu ý CPI dịch vụ ở đây gồm cả Nhà ở, vốn chiếm tỷ trọng rất lớn). Tương quan KHÔNG phải nhân quả — chỉ cho thấy độ trễ/độ mạnh liên hệ giữa các tầng để tự đánh giá cú sốc có đang lan rộng không (xem thêm badge "Lạm phát — độ dai dẳng" ở mục 1).</p>`)}
      ${card('6. Tiêu dùng danh nghĩa vs thực — tăng trưởng bán lẻ có phải do giá không', `<div class="ind-chart" style="height:280px"><canvas id="chart-us-real"></canvas></div>
            <p class="ind-source-note">Bán lẻ danh nghĩa YoY trừ CPI YoY ≈ tăng trưởng thực (xấp xỉ). Chi tiêu thực PCE hiện ${f(rc.pce_real_yoy)}% so với danh nghĩa ${f(rc.pce_nominal_yoy)}%.</p>`)}`)}
      ${group('🏦 Thanh khoản Fed & Lãi suất', (usm.liquidity ? card('7. Bảng cân đối Fed — QE/QT & thanh khoản hệ thống', usStateBadge(usm.states && usm.states.fed_balance_sheet) + usFedLiquidityCard(usm.liquidity, f)) : '')
          + card('Lãi suất & việc làm (tóm tắt nhanh)', `${usStateBadge(usm.states && usm.states.fed_policy)}<ul class="ind-source-note" style="line-height:1.8">${Object.values(rates).map(r => `<li>${r.label}: <b>${f(r.latest)}</b> (${r.period})</li>`).join('')}</ul>`))}
      ${group('📈 Tăng trưởng & Lao động', (usm.growth ? card('8. Tăng trưởng — GDP thực & sản xuất công nghiệp', usStateBadge(usm.states && usm.states.growth) + usGrowthCard(usm.growth, f)) : '')
          + (usm.labor ? card('9. Lao động — việc làm, thất nghiệp, lương thực', usStateBadge(usm.states && usm.states.labor) + usLaborCard(usm.labor, f)) : ''))}
      ${group('💵 Lợi suất, Tín dụng & Dòng vốn quốc tế', (usm.treasury_credit ? card('10. Lợi suất & tín dụng — đường cong, lãi suất thực, spread rủi ro', usStateBadge(usm.states && usm.states.treasury_credit) + usTreasuryCreditCard(usm.treasury_credit, f)) : '')
          + (usm.usd ? card('11. USD — chỉ số USD trọng số thương mại', usStateBadge(usm.states && usm.states.usd) + usUsdCard(usm.usd, f)) : '')
          + (usm.capital_flows ? card('12. Capital Flows — nước ngoài nắm giữ Treasury Mỹ (TIC)', usStateBadge(usm.states && usm.states.capital_flows) + usCapitalFlowsCard(usm.capital_flows, f)) : ''))}
    `;

    // Cuộn-sang-phải cho các bảng .monitoring-table-scroll trong tab này được xử lý ở
    // initVimoTabs() (chạy ĐÚNG lúc tab hiện ra — ở đây container còn display:none lúc trang mới
    // load lần đầu nên scrollWidth đọc ra 0, gán lúc này vô nghĩa).

    const hist = usm.history, ph = usm.pipeline;
    const line = (label, data, color, dash) => ({ label, data, borderColor: color, backgroundColor: color, borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true, borderDash: dash || [] });
    const mk = (id, cfg) => { const c = document.getElementById(id); if (c) chartInstances.push(new Chart(c, cfg)); };
    // Sparkline nhỏ cho bảng tổng quan — không trục, không legend, chỉ minh họa xu hướng gần đây.
    const mkSpark = (id, data, color) => mk(id, { type: 'line', data: { labels: data.map((_, i) => i), datasets: [
        { data, borderColor: color, borderWidth: 1.5, pointRadius: 0, tension: 0.3, fill: false }] },
        options: { responsive: true, maintainAspectRatio: false, animation: false,
                   plugins: { legend: { display: false }, tooltip: { enabled: false } },
                   scales: { x: { display: false }, y: { display: false } } } });

    if (usm.growth && usm.growth.gdp) mkSpark('spark-growth', usm.growth.history.gdp.qoq_ann.slice(-24), '#60a5fa');
    if (usm.labor) mkSpark('spark-labor', usm.labor.history.payrolls_mom.slice(-24), '#a78bfa');
    mkSpark('spark-inflation', hist.cpi_yoy.slice(-24), '#f59e0b');
    mkSpark('spark-persistence', hist.breadth_gt3_pct.slice(-24), '#f97316');
    if (usm.rates.usm_fed_funds && usm.rates.usm_fed_funds.history) mkSpark('spark-fed', usm.rates.usm_fed_funds.history.values.slice(-24), '#10b981');
    if (usm.liquidity) mkSpark('spark-fedbs', usm.liquidity.history.liquidity_impulse.slice(-24), '#a78bfa');
    if (usm.treasury_credit) mkSpark('spark-credit', usm.treasury_credit.history.spread_10y_2y.slice(-24), '#ef4444');
    if (usm.usd) mkSpark('spark-usd', usm.usd.history.values.slice(-24), '#60a5fa');
    if (usm.capital_flows) mkSpark('spark-capflows', usm.capital_flows.history.total.slice(-24), '#10b981');
    const legend = { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } };
    const ax = { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 10 } };

    // SUA 2026-10-08 (user: "thêm đường line PPI đi, để xem có dẫn truyền sang chi phí sản xuất
    // không" — ph.ppi_hist dùng CHUNG mảng periods với hist (cùng biến `periods` gốc bên Python)
    // nên ghép thẳng được, không lệch trục thời gian).
    mk('chart-us-headline', { type: 'line', data: { labels: hist.periods, datasets: [
        line('CPI toàn phần YoY (%)', hist.cpi_yoy, '#60a5fa'),
        line('CPI lõi YoY (%)', hist.core_yoy, '#f59e0b', [5, 4]),
        line('PPI cầu cuối YoY (%)', ph.ppi_hist.ppi_yoy, '#a78bfa', [2, 2])] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });

    // SUA 2026-10-08 (user: "làm dạng 2 line đi, nhìn [cột+đường] này sao rõ được" — cột đỏ/xanh
    // lẫn với đường lõi khó so sánh xu hướng) — đổi cả 2 sang line để so trực tiếp CPI toàn phần
    // vs lõi, thấy ngay 2 tháng gần nhất có đang tăng tốc trở lại hay không (vd case thực tế: Jun
    // -0.42% → Jul +0.07% → Aug +0.4%).
    if (h.mom_history) mk('chart-us-mom', { type: 'line', data: { labels: h.mom_history.periods, datasets: [
        { ...line('CPI MoM (%)', h.mom_history.cpi_mom, '#60a5fa'), pointRadius: 3 },
        { ...line('CPI lõi MoM (%)', h.mom_history.core_mom, '#f59e0b', [5, 4]), pointRadius: 3 }] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false } }, y: CHART_DEFAULTS.scales.y } } });

    const gl = usm.groups.filter(g => g.yoy !== null).sort((a, b) => b.yoy - a.yoy);
    // SUA 2026-10-08 (user: "thêm con số vào để tôi nhìn giá trị cho rõ") — thêm datalabels (đặt ở
    // DATASET, không phải options.plugins — đặt ở plugins global từng không hiện, xem bài học ở
    // usContribBarChart).
    mk('chart-us-groups', { type: 'bar', data: { labels: gl.map(g => g.label), datasets: [{
        label: 'YoY (%)', data: gl.map(g => g.yoy), backgroundColor: gl.map(g => g.yoy > 3 ? 'rgba(239,68,68,0.75)' : 'rgba(96,165,250,0.7)'),
        datalabels: { display: true, color: '#fff', anchor: 'end', align: 'top', offset: 2, clip: false,
                      formatter: v => (v >= 0 ? '+' : '') + v.toFixed(1) + '%', font: { size: 10, weight: '600' } } }] },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: false } }, scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, autoSkip: false, maxRotation: 30 } }, y: CHART_DEFAULTS.scales.y } },
        plugins: [ChartDataLabels] });

    (function () {
        const ct = usm.contributions;
        usContribBarChart(ct);
        usContribMomBarChart(ct);
        const colors = { usm_cpi_food: '#60a5fa', usm_cpi_energy: '#f97316', usm_cpi_shelter: '#a78bfa',
            usm_cpi_transport: '#ef4444', usm_cpi_medical: '#10b981', usm_cpi_apparel: '#eab308',
            usm_cpi_recreation: '#ec4899', usm_cpi_education_comm: '#14b8a6', usm_cpi_other: '#94a3b8',
            residual: 'rgba(148,163,184,0.35)' };
        // SUA 2026-10-07 (user: "tôi muốn biểu đồ CỘT CHỒNG thể hiện CPI tăng là do đâu, CPI +3%
        // thì 3% đó do từng cấu phần làm tăng bao nhiêu điểm") — đổi line/area-stack sang bar
        // stack: mỗi cột 1 tháng, tổng chiều cao cột = ĐÚNG CPI YoY tháng đó, rõ ràng hơn kiểu
        // miền/đường (dễ đọc nhầm thành YoY riêng từng nhóm, không phải phần đóng góp).
        // SUA 2026-10-08 (user: "cột tháng gần nhất nhìn mảnh, khó nhìn, có vẻ khung biểu đồ che
        // khuất 1 phần, hãy làm cột tháng gần nhất to x2" — ĐÃ THỬ 3 CÁCH trước khi ra cách này:
        // (1) barThickness dạng hàm scriptable, (2) barThickness dạng mảng theo index — CẢ 2 đều
        // KHÔNG hoạt động (Chart.js chỉ tính barThickness 1 LẦN cho cả category scale lúc layout).
        // (3) ghi đè el.width trong hook afterDatasetsUpdate — hoạt động lúc tạo biểu đồ, nhưng bị
        // GHI ĐÈ LẠI mỗi khi chart.resize() chạy (initVimoTabs() gọi resize() khi chuyển tab active,
        // resize() update() lại layout nên tính lại width đồng đều, mất hiệu lực ghi đè trước đó —
        // kiểm chứng bằng Playwright thấy width TĂNG đồng đều ở MỌI cột, không chỉ cột cuối).
        // CÁCH DÙNG: vẽ ĐÈ trực tiếp lên canvas ở hook afterDatasetsDraw (chạy SAU MỖI lần Chart.js
        // tự vẽ, kể cả sau resize) — tô 1 hình chữ nhật rộng hơn tại ĐÚNG vị trí Y đã tính (el.y/
        // el.base đã đúng theo stacking, chỉ cần vẽ RỘNG hơn theo X) — không phụ thuộc Chart.js có
        // cho ghi đè width hay không.
        const widenLastBar = (extra) => ({
            id: 'widenLastBar',
            afterDatasetsDraw(chart) {
                const ctx = chart.ctx;
                const lastIdx = chart.data.labels.length - 1;
                chart.data.datasets.forEach((ds, dsIdx) => {
                    const meta = chart.getDatasetMeta(dsIdx);
                    if (meta.hidden) return;
                    const el = meta.data[lastIdx];
                    if (!el || el.width === undefined) return;
                    const w = el.width + extra;
                    const top = Math.min(el.y, el.base);
                    const h = Math.abs(el.base - el.y);
                    if (h <= 0) return;
                    ctx.save();
                    ctx.fillStyle = ds.backgroundColor;
                    ctx.fillRect(el.x - w / 2, top, w, h);
                    ctx.restore();
                });
            },
        });
        mk('chart-us-contrib', { type: 'bar', data: { labels: ct.periods, datasets: ct.rows.map(r => ({
            label: r.label + (r.weight_pct !== null ? ` (${r.weight_pct.toFixed(1)}%)` : ''),
            data: r.values, backgroundColor: colors[r.key] || '#999', borderWidth: 0 })) },
            options: { ...CHART_DEFAULTS, plugins: legend, interaction: { mode: 'index' }, layout: { padding: { right: 16 } },
                       scales: { x: { ...ax, stacked: true }, y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Điểm % đóng góp vào CPI YoY', color: '#9aa5bd', font: { size: 9 } } } } },
            // SUA 2026-10-08 (user: "vẽ số tháng 7 và tháng 8 liền nhau à, tôi thấy dính vào nhau
            // thế... co nhỏ cái cột tháng gần nhất lại" — +10px (nửa mỗi bên +5px) trên khoảng cách
            // tâm cột chỉ ~14px (127 cột / 1818px) khiến rìa cột cuối ĐÈ qua rìa cột liền trước —
            // giảm xuống +4 để đủ rộng hơn rõ rệt nhưng không chạm cột bên cạnh.
            plugins: [widenLastBar(4)] });

        // THEM 2026-10-08 (user: "số tháng gần nhất bị bé quá, cột không rộng như bình thường nên
        // khó nhìn") — cùng dữ liệu, cắt 24 tháng gần nhất, để Chart.js tự giãn trục Y RIÊNG cho
        // đoạn này (không bị đỉnh lạm phát 2021-2022 kéo giãn trục chung làm cột gần đây bị "lùn").
        const n24 = Math.min(24, ct.periods.length);
        mk('chart-us-contrib-recent', { type: 'bar', data: { labels: ct.periods.slice(-n24), datasets: ct.rows.map(r => ({
            label: r.label + (r.weight_pct !== null ? ` (${r.weight_pct.toFixed(1)}%)` : ''),
            data: r.values.slice(-n24), backgroundColor: colors[r.key] || '#999', borderWidth: 0 })) },
            options: { ...CHART_DEFAULTS, plugins: legend, interaction: { mode: 'index' }, layout: { padding: { right: 16 } },
                       scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false }, stacked: true },
                                 y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Điểm % đóng góp vào CPI YoY', color: '#9aa5bd', font: { size: 9 } } } } },
            plugins: [widenLastBar(24)] });
    })();
    mk('chart-us-breadth', { type: 'line', data: { labels: hist.periods, datasets: [
        line('% nhóm CPI có YoY > 3%', hist.breadth_gt3_pct, '#a78bfa')] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: { ...CHART_DEFAULTS.scales.y, min: 0, max: 100 } } } });

    mk('chart-us-goods-services', { type: 'line', data: { labels: hist.periods, datasets: [
        line('CPI hàng hóa YoY (%)', hist.goods_yoy, '#10b981'),
        line('CPI dịch vụ YoY (%)', hist.services_yoy, '#60a5fa')] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });

    mk('chart-us-pipeline', { type: 'line', data: { labels: ph.ppi_hist.periods, datasets: [
        line('PPI cầu cuối YoY (%)', ph.ppi_hist.ppi_yoy, '#f97316'),
        line('Giá nhập khẩu YoY (%)', ph.ppi_hist.import_yoy, '#60a5fa'),
        { ...line('Dầu WTI YoY (%, trục phải)', ph.ppi_hist.oil_yoy, '#a78bfa', [4, 3]), yAxisID: 'y1' }] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y, y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false } } } } });

    mk('chart-us-leadlag', { type: 'bar', data: { labels: ph.ppi_to_cpi_goods_corr.map(x => 'trễ ' + x.lag_months + ' tháng'), datasets: [
        { label: 'PPI → CPI hàng hóa (tương quan MoM)', data: ph.ppi_to_cpi_goods_corr.map(x => x.corr), backgroundColor: 'rgba(249,115,22,0.75)' },
        { label: 'Giá nhập khẩu → CPI hàng hóa', data: ph.import_to_cpi_goods_corr.map(x => x.corr), backgroundColor: 'rgba(96,165,250,0.75)' }] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: CHART_DEFAULTS.scales.x, y: { ...CHART_DEFAULTS.scales.y, min: -1, max: 1 } } } });

    // THEM 2026-10-08 (user gửi tài liệu "Inflation Transmission" — thêm 2 tầng truyền dẫn SỚM/
    // MUỘN hơn tầng PPI→CPI hàng hóa đã có ở trên).
    if (ph.oil_to_ppi_corr) mk('chart-us-transmission-oil', { type: 'bar', data: { labels: ph.oil_to_ppi_corr.map(x => 'trễ ' + x.lag_months + ' tháng'), datasets: [
        { label: 'Dầu WTI → PPI cầu cuối (tương quan MoM)', data: ph.oil_to_ppi_corr.map(x => x.corr), backgroundColor: 'rgba(167,139,250,0.75)' }] },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: false } }, scales: { x: CHART_DEFAULTS.scales.x, y: { ...CHART_DEFAULTS.scales.y, min: -1, max: 1 } } } });

    if (ph.wage_to_services_corr) mk('chart-us-transmission-wage', { type: 'bar', data: { labels: ph.wage_to_services_corr.map(x => 'trễ ' + x.lag_months + ' tháng'), datasets: [
        { label: 'Lương → CPI dịch vụ (tương quan MoM)', data: ph.wage_to_services_corr.map(x => x.corr), backgroundColor: 'rgba(236,72,153,0.75)' }] },
        options: { ...CHART_DEFAULTS, plugins: { legend: { display: false } }, scales: { x: CHART_DEFAULTS.scales.x, y: { ...CHART_DEFAULTS.scales.y, min: -1, max: 1 } } } });

    mk('chart-us-crack', { type: 'line', data: { labels: usm.cracks.periods, datasets: [
        line('Crack diesel ($/thùng)', usm.cracks.diesel_crack, '#f97316'),
        line('Crack xăng ($/thùng)', usm.cracks.gasoline_crack, '#60a5fa'),
        line('Crack 3-2-1 ($/thùng)', usm.cracks.crack_321, '#10b981', [5, 4])] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });
    // SUA 2026-10-07 (user: "biểu đồ này nhiễu giai đoạn covid quá") — cú sốc bán lẻ 2020-2021 (đáy
    // -20%, đỉnh +51%) kéo giãn trục Y làm phần 2023+ bị dẹt khó đọc; giới hạn hiển thị từ 2023-Q1.
    const rcIdx0 = rc.hist.periods.findIndex(p => p >= '2023-01');
    const rcPeriods = rc.hist.periods.slice(rcIdx0);
    mk('chart-us-real', { type: 'line', data: { labels: rcPeriods, datasets: [
        line('Bán lẻ danh nghĩa YoY (%)', rc.hist.retail_nominal_yoy.slice(rcIdx0), '#60a5fa'),
        line('CPI YoY (%)', rc.hist.cpi_yoy.slice(rcIdx0), '#f59e0b', [5, 4])] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });

    if (usm.liquidity) (function () {
        const liq = usm.liquidity, h = liq.history;
        const T = arr => arr.map(v => v === null ? null : v / 1e6); // trieu USD -> nghin ty USD
        mk('chart-us-fed-assets', { type: 'line', data: { labels: h.periods, datasets: [
            line('Tổng tài sản Fed (T$)', T(h.assets), '#60a5fa'),
            line('— Treasury (T$)', T(h.treasury), '#10b981'),
            line('— MBS (T$)', T(h.mbs), '#f59e0b', [5, 4])] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Nghìn tỷ USD', color: '#9aa5bd', font: { size: 9 } } } } } });

        mk('chart-us-fed-drains', { type: 'line', data: { labels: h.periods, datasets: [
            line('Dự trữ ngân hàng (T$)', T(h.reserves), '#a78bfa'),
            line('RRP (T$)', T(h.rrp), '#ef4444'),
            line('TGA (T$)', T(h.tga), '#f97316', [5, 4])] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Nghìn tỷ USD', color: '#9aa5bd', font: { size: 9 } } } } } });

        mk('chart-us-fed-netliq', { type: 'line', data: { labels: h.periods, datasets: [
            line('Net Liquidity Fed (T$, trục trái)', T(h.net_liquidity), '#60a5fa')] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Nghìn tỷ USD', color: '#9aa5bd', font: { size: 9 } } } } } });

        // THEM 2026-10-08 (user gửi tài liệu "Liquidity Impulse = ΔReserves−ΔTGA−ΔRRP") — biểu đồ
        // cột MoM, dương (bơm ròng) = xanh, âm (rút ròng) = đỏ.
        const Bn = v => v === null ? null : v / 1000; // trieu USD -> ty USD
        mk('chart-us-fed-impulse', { type: 'bar', data: { labels: h.periods, datasets: [
            { label: 'Liquidity Impulse (tỷ$/tháng)', data: h.liquidity_impulse.map(Bn),
              backgroundColor: h.liquidity_impulse.map(v => v === null ? '#999' : (v >= 0 ? 'rgba(16,185,129,0.75)' : 'rgba(239,68,68,0.75)')) }] },
            options: { ...CHART_DEFAULTS, plugins: { legend: { display: false } }, scales: { x: ax, y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Tỷ USD/tháng', color: '#9aa5bd', font: { size: 9 } } } } } });
    })();

    // SUA 2026-10-08 (cùng pattern đã sửa cho biểu đồ bán lẻ/tiêu dùng thực: "biểu đồ này nhiễu
    // giai đoạn covid quá") — cú sốc 2020 (GDP QoQ năm hóa -30%/+35%, thất nghiệp 14,8%) kéo giãn
    // trục Y làm cả 5 năm gần nhất bị dẹt khó đọc; cắt hiển thị từ 2021-01, giữ nguyên dữ liệu gốc.
    const cov0 = arr => { const i = arr.findIndex(p => p >= '2021-01'); return i < 0 ? 0 : i; };
    if (usm.growth) (function () {
        const gh = usm.growth.history;
        const gi = cov0(gh.gdp.periods), ii = cov0(gh.indpro.periods);
        mk('chart-us-gdp', { type: 'line', data: { labels: gh.gdp.periods.slice(gi), datasets: [
            line('GDP thực YoY (%)', gh.gdp.yoy.slice(gi), '#60a5fa'),
            line('GDP thực QoQ năm hóa (%)', gh.gdp.qoq_ann.slice(gi), '#f59e0b', [5, 4])] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });
        mk('chart-us-indpro', { type: 'line', data: { labels: gh.indpro.periods.slice(ii), datasets: [
            line('Sản xuất công nghiệp YoY (%)', gh.indpro.yoy.slice(ii), '#10b981')] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });
    })();

    if (usm.labor) (function () {
        const lh = usm.labor.history;
        const li = cov0(lh.periods);
        const periods = lh.periods.slice(li);
        mk('chart-us-payrolls', { type: 'bar', data: { labels: periods, datasets: [
            { label: 'Việc làm phi NN thêm/tháng (nghìn)', data: lh.payrolls_mom.slice(li), backgroundColor: lh.payrolls_mom.slice(li).map(v => v >= 0 ? 'rgba(16,185,129,0.75)' : 'rgba(239,68,68,0.75)') },
            { ...line('TB 3 tháng (nghìn)', lh.payrolls_mom_3m_avg.slice(li), '#f59e0b'), type: 'line' }] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });
        mk('chart-us-labor-other', { type: 'line', data: { labels: periods, datasets: [
            line('Thất nghiệp (%)', lh.unemployment.slice(li), '#a78bfa'),
            { ...line('Lương THỰC YoY (%, trục phải)', lh.real_wage_yoy.slice(li), '#60a5fa', [4, 3]), yAxisID: 'y1' }] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y, y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false } } } } });
    })();

    if (usm.treasury_credit) (function () {
        const th = usm.treasury_credit.history;
        mk('chart-us-curve', { type: 'line', data: { labels: th.periods, datasets: [
            line('Lợi suất 2 năm (%)', th.yield_2y, '#60a5fa'),
            line('Lợi suất 10 năm (%)', th.yield_10y, '#f59e0b'),
            { ...line('10Y-2Y (%, trục phải)', th.spread_10y_2y, '#10b981', [4, 3]), yAxisID: 'y1' }] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y, y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false } } } } });
        // SUA 2026-10-08 (user: "đây là dữ liệu gì, không hiểu nó nói lên điều gì, vẽ lại chỉ từ
        // thời điểm có dữ liệu, biểu đồ có nhiều vị trí trống quá" — 2 chuỗi OAS chỉ có từ 2023-10
        // trên FRED (ICE giới hạn cấp phép, xem ind-source-note bên dưới) nhưng trục X dùng chung
        // th.periods bắt đầu 2015-01 của 2Y/10Y nên 2/3 biểu đồ trống) — cắt riêng từ index đầu
        // tiên hy_oas có giá trị, không dùng chung trục với biểu đồ lợi suất ở trên.
        const hyIdx0 = th.hy_oas.findIndex(v => v !== null);
        const creditPeriods = hyIdx0 >= 0 ? th.periods.slice(hyIdx0) : th.periods;
        mk('chart-us-credit', { type: 'line', data: { labels: creditPeriods, datasets: [
            line('High Yield OAS (%) — spread tín dụng rác, rộng ra khi lo ngại vỡ nợ DN', hyIdx0 >= 0 ? th.hy_oas.slice(hyIdx0) : th.hy_oas, '#ef4444'),
            { ...line('Investment Grade OAS (%, trục phải) — spread tín dụng DN xếp hạng cao', hyIdx0 >= 0 ? th.ig_oas.slice(hyIdx0) : th.ig_oas, '#60a5fa', [4, 3]), yAxisID: 'y1' }] },
            options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 8 } }, y: CHART_DEFAULTS.scales.y, y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false } } } } });
    })();

    if (usm.usd) mk('chart-us-dxy', { type: 'line', data: { labels: usm.usd.history.periods, datasets: [
        line('Chỉ số USD Broad (DTWEXBGS)', usm.usd.history.values, '#60a5fa')] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });

    if (usm.capital_flows) mk('chart-us-tic', { type: 'line', data: { labels: usm.capital_flows.history.periods, datasets: [
        line('Tổng nước ngoài nắm giữ (tỷ $)', usm.capital_flows.history.total, '#60a5fa'),
        line('Khối chính thức/NHTW (tỷ $)', usm.capital_flows.history.official, '#f59e0b'),
        line('Khối tư nhân (tỷ $)', usm.capital_flows.history.private, '#10b981')] },
        options: { ...CHART_DEFAULTS, plugins: legend, scales: { x: ax, y: CHART_DEFAULTS.scales.y } } });

    // SUA 2026-10-09 (user: "dữ liệu bảng và biểu đồ CPI này đang hiện ở các tháng trước, với bảng
    // biểu đồ dài kỳ thì sẽ hiện ở tháng gần nhất nhé" — initVimoTabs() CHỈ cuộn các
    // .monitoring-table-scroll về mép phải (tháng mới nhất) khi NGƯỜI DÙNG BẤM chuyển tab, lúc đó
    // nội dung đã render xong. Nếu vào tab "us" NGAY từ URL hash (#tab=us) hoặc link chia sẻ,
    // show('us') chạy lúc container CÒN TRỐNG (chưa fetch xong dữ liệu) — cuộn lúc đó vô nghĩa
    // (scrollWidth=0), và render sau đó không ai gọi lại cuộn nữa, cố định kẹt ở mép trái (tháng
    // CŨ NHẤT). Tự áp lại y hệt logic cuộn phải ở đây, NGAY SAU KHI render xong — chỉ cần container
    // đang hiển thị thật (không bị .vimo-tab-hidden) thì cuộn có tác dụng ngay, không phải đợi
    // click tab nữa. Gọi 2 lần (ngay + sau 300ms): đo thực tế mục "2d" (div rộng tính bằng JS theo
    // số tháng) có scrollWidth TĂNG THÊM sau khi Chart.js hoàn tất layout canvas (lần đầu đo được
    // scrollWidth NHỎ HƠN giá trị cuối, set scrollLeft bị "chốt" ở mức cũ, không tự bắt kịp
    // scrollWidth lớn hơn sau đó) — lần gọi thứ 2 bắt đúng kích thước cuối cùng.
    const scrollTablesToLatest = () => {
        if (!box.classList.contains('vimo-tab-hidden')) {
            box.querySelectorAll('.monitoring-table-scroll').forEach(el => { el.scrollLeft = el.scrollWidth; });
        }
    };
    setTimeout(scrollTablesToLatest, 0);
    setTimeout(scrollTablesToLatest, 300);
}

function renderFxPressureSignalsMonthlyChart(indicators) {
    const canvas = document.getElementById('chart-fx-pressure-signals-monthly');
    const card = document.getElementById('fx-pressure-signals-monthly-chart-card');
    if (!canvas) return;
    const neer = indicators['darvas_neer_vn'];
    const usdvndYoy = indicators['usdvnd_growth_yoy'];
    const usdvndLevel = indicators['usdvnd_monthly_avg'];
    if (!neer || !neer.series.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    // SUA 2026-10-03 (user: "chỉ cần từ 2020 tới nay thôi, đừng vẽ nhiều quá") — NEER có từ 1993,
    // giới hạn vẽ từ 2020 để thấy rõ giai đoạn 2021-2022 (tỷ giá lên cao kinh khủng cuối 2022) mà
    // không kéo dài lãng phí.
    const neerRecent = neer.series.filter(p => p.period >= '2020-01');
    const periods = neerRecent.map(p => p.period);
    const neerArr = neerRecent.map(p => p.value);
    const usdvndByMonth = usdvndYoy ? Object.fromEntries(usdvndYoy.series.map(p => [p.period, p.value])) : {};
    const usdvndArr = periods.map(p => usdvndByMonth[p] ?? null);
    // THEM (user 2026-10-03): "áp thêm cho tôi đường tỷ giá thực USD/VND nhé" — MỨC thực tế (VND),
    // trục RIÊNG thứ 3 (y2) vì đơn vị khác hẳn NEER (index) và %YoY.
    const usdvndLevelByMonth = usdvndLevel ? Object.fromEntries(usdvndLevel.series.map(p => [p.period, p.value])) : {};
    const usdvndLevelArr = periods.map(p => usdvndLevelByMonth[p] ?? null);

    const chart = new Chart(canvas, {
        type: 'line',
        data: {
            labels: periods,
            datasets: [
                { label: 'NEER Việt Nam (Darvas/Bruegel, index)', data: neerArr, yAxisID: 'y',
                  borderColor: '#60a5fa', backgroundColor: '#60a5fa15', fill: false,
                  tension: 0.2, pointRadius: 0, borderWidth: 2, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(1) },
                { label: 'USD/VND tăng/giảm YoY (%, thị trường thực)', data: usdvndArr, yAxisID: 'y1',
                  borderColor: '#f59e0b', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0,
                  fill: false, tension: 0.2, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(1) },
                { label: 'USD/VND mức thực tế (VND, bình quân tháng)', data: usdvndLevelArr, yAxisID: 'y2',
                  borderColor: '#10b981', borderWidth: 2, pointRadius: 0,
                  fill: false, tension: 0.2, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(0) },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 45, autoSkip: true, maxTicksLimit: 18 } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     title: { display: true, text: 'NEER (index)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: '% YoY', color: '#9aa5bd', font: { size: 9 } } },
                y2: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM (user 2026-10-09): "biểu đồ tương quan lợi suất và tỷ giá: lợi suất TPCP 10 năm VN và Mỹ
// theo line, chênh lệch VN-Mỹ theo cột, tỷ giá USD/VND theo line nét đứt" — 3 nguồn TẦN SUẤT khác
// nhau (VN bond: NGÀY từ VIRA, chỉ từ ~2026-04; US bond: THÁNG từ usMacro/FRED; USD/VND: THÁNG) —
// gộp VN về THÁNG (trung bình) để ghép chung 1 trục X, lấy PERIODS của VN bond làm trục (ngắn nhất,
// chỉ có từ khi VIRA được thêm) thay vì nội suy ngược cho giai đoạn không có dữ liệu VN.
function renderFxYieldDiffChart(indicators, usm) {
    const canvas = document.getElementById('chart-fx-yield-diff');
    const card = document.getElementById('fx-yield-diff-chart-card');
    if (!canvas) return;
    const vnBond = indicators['govt_bond_yield_10y'];
    const usHist = usm && usm.treasury_credit && usm.treasury_credit.history;
    const usdvnd = indicators['usdvnd_monthly_avg'];
    if (!vnBond || !vnBond.series.length || !usHist || !usdvnd || !usdvnd.series.length) {
        if (card) card.style.display = 'none';
        return;
    }
    if (card) card.style.display = '';

    const vnMonthly = {};
    vnBond.series.forEach(p => { (vnMonthly[p.period.slice(0, 7)] = vnMonthly[p.period.slice(0, 7)] || []).push(p.value); });
    const vnByMonth = Object.fromEntries(Object.entries(vnMonthly).map(([m, vs]) => [m, vs.reduce((a, b) => a + b, 0) / vs.length]));
    const usByMonth = Object.fromEntries(usHist.periods.map((p, i) => [p, usHist.yield_10y[i]]));
    const usdvndByMonth = Object.fromEntries(usdvnd.series.map(p => [p.period, p.value]));

    const periods = Object.keys(vnByMonth).sort();
    const vnArr = periods.map(p => Math.round(vnByMonth[p] * 100) / 100);
    const usArr = periods.map(p => usByMonth[p] ?? null);
    const diffArr = periods.map((p, i) => (vnArr[i] !== null && usArr[i] !== null) ? Math.round((vnArr[i] - usArr[i]) * 100) / 100 : null);
    const usdvndArr = periods.map(p => usdvndByMonth[p] ?? null);

    const chart = new Chart(canvas, {
        type: 'bar',
        data: {
            labels: periods,
            datasets: [
                { label: 'Chênh lệch VN − Mỹ (điểm %, trục phải trong)', data: diffArr, yAxisID: 'y1',
                  backgroundColor: diffArr.map(v => v === null ? '#999' : (v >= 0 ? 'rgba(239,68,68,0.55)' : 'rgba(96,165,250,0.55)')), order: 3 },
                { type: 'line', label: 'Lợi suất TPCP 10 năm Việt Nam (%)', data: vnArr, yAxisID: 'y',
                  borderColor: '#f59e0b', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true, order: 1 },
                { type: 'line', label: 'Lợi suất TPCP 10 năm Mỹ (%)', data: usArr, yAxisID: 'y',
                  borderColor: '#60a5fa', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true, order: 1 },
                { type: 'line', label: 'USD/VND (trục phải ngoài)', data: usdvndArr, yAxisID: 'y2',
                  borderColor: '#10b981', borderWidth: 2, borderDash: [6, 4], pointRadius: 0, tension: 0.2, spanGaps: true, order: 2 },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left', title: { display: true, text: '% (lợi suất 10 năm)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false }, title: { display: true, text: 'Điểm % chênh lệch', color: '#9aa5bd', font: { size: 9 } } },
                y2: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false }, title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    });
    chartInstances.push(chart);
}

// THEM (user 2026-10-01): "biểu đồ đo lường cung cầu ngoại tệ" — các cấu phần Cầu (đỏ)/Cung
// (xanh) từ BOP quý NHNN, CÙNG đơn vị (triệu USD) nên so được trực tiếp — CHỈ vẽ nhiều đường
// cạnh nhau (KHÔNG cộng dồn/trừ thành 1 điểm số, đúng yêu cầu giữ Cầu/Cung là 2 lớp ĐỘC LẬP).
// kieu_hoi_hcm (thêm sau, nguồn NHNN Chi nhánh Khu vực 2 — CHỈ PHẠM VI TP.HCM) CỐ TÌNH không đưa
// vào renderFxSupplyDemandTotalChart() bên dưới: quan hệ chính xác giữa nó và
// bop_sbv_secondary_income_received (toàn quốc) chưa được NHNN đối chiếu công khai — có thể
// trùng lặp 1 phần (kiều hối HCM vốn là 1 phần của chuyển giao vãng lai toàn quốc), cộng vào
// Tổng Cung sẽ double-count mà không chắc mức độ — chỉ vẽ cạnh nhau ở đây để tham khảo/đối chiếu.
// SUA 2026-10-03 (user: "phần chỗ này theo tôi tính ròng đi... vì NXK lớn quá nhưng tính ròng ra
// thì khéo cũng chỉ ngang với các phần khác thôi thì sẽ dễ nhìn hơn. thêm cho tôi đường line tỷ
// giá thực tế usd/vnd nữa để xem trực quan") — ĐỔI từ 8-9 đường GỘP (xuất/nhập riêng, hàng hóa gộp
// quá lớn so dịch vụ/thu nhập đầu tư/chuyển giao vãng lai, lệch thang nhìn không rõ) sang 4 đường
// RÒNG (Xuất-Nhập mỗi cặp) — CÙNG LÀ 4 "cán cân con" CHUẨN của Cán cân vãng lai (hàng hóa+dịch
// vụ+thu nhập đầu tư+chuyển giao vãng lai = Current Account, giống cách NHNN tự trình bày), KHÔNG
// VI PHẠM nguyên tắc "không gộp thành 1 điểm áp lực" (vẫn 4 đường ĐỘC LẬP, chỉ đổi cách tính TỪNG
// đường từ gộp sang ròng, không cộng 4 đường lại). Thêm USD/VND thực tế (đường tham chiếu, trục
// phải) để so trực quan biến động tỷ giá với các cán cân ròng — bỏ kieu_hoi_hcm khỏi chart NÀY (số
// CHỈ CÓ phía cung, không có cặp để tính ròng, xem card riêng ở lớp ④ Cung ngoại tệ).
function renderFxSupplyDemandChart(indicators) {
    const canvas = document.getElementById('chart-fx-supply-demand');
    const card = document.getElementById('fx-supply-demand-chart-card');
    if (!canvas) return;
    const NET_SERIES = [
        { exportKey: 'bop_sbv_goods_export', importKey: 'bop_sbv_goods_import', label: 'Cán cân hàng hóa (ròng)', color: '#10b981' },
        { exportKey: 'bop_sbv_services_export', importKey: 'bop_sbv_services_import', label: 'Cán cân dịch vụ (ròng)', color: '#60a5fa' },
        { exportKey: 'bop_sbv_investment_income_received', importKey: 'bop_sbv_investment_income_paid', label: 'Cán cân thu nhập đầu tư (ròng)', color: '#f59e0b' },
        { exportKey: 'bop_sbv_secondary_income_received', importKey: 'bop_sbv_secondary_income_paid', label: 'Cán cân chuyển giao vãng lai (ròng)', color: '#a78bfa' },
    ];
    const validSeries = NET_SERIES.filter(s => indicators[s.exportKey] && indicators[s.importKey]
        && indicators[s.exportKey].series.length && indicators[s.importKey].series.length);
    if (!validSeries.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    // Gioi han hien thi tu 2020 (nhu renderFxPressureSignalsChart) - BOP da duoc backfill tu IMF
    // ve 1996-Q1 nen neu khong gioi han, nhan truc X se qua nhieu (>100 ky) gay chong chat/roi mat.
    const periods = Array.from(new Set(validSeries.flatMap(s =>
        indicators[s.exportKey].series.map(p => p.period)))).filter(p => p >= '2020-Q1').sort();
    const datasets = validSeries.map(s => {
        const expByPeriod = Object.fromEntries(indicators[s.exportKey].series.map(p => [p.period, p.value]));
        const impByPeriod = Object.fromEntries(indicators[s.importKey].series.map(p => [p.period, p.value]));
        const data = periods.map(p => {
            const e = expByPeriod[p], i = impByPeriod[p];
            return (e === null || e === undefined || i === null || i === undefined) ? null : e - i;
        });
        return {
            label: s.label, data, yAxisID: 'y',
            borderColor: s.color, backgroundColor: s.color + '15', fill: false,
            borderWidth: 2, tension: 0.25, pointRadius: 3, pointBackgroundColor: s.color, spanGaps: true,
            datalabels: _endpointDatalabelsConfig(0),
        };
    });
    // USD/VND thực tế (tham chiếu, trục phải) — quy về cuối quý để so cùng trục X với BOP (quý).
    const usdvnd = indicators['usdvnd_monthly_avg'];
    if (usdvnd && usdvnd.series.length) {
        const QUARTER_END_MONTH = { '1': '03', '2': '06', '3': '09', '4': '12' };
        const usdvndByMonth = Object.fromEntries(usdvnd.series.map(p => [p.period, p.value]));
        const usdvndArr = periods.map(period => {
            const [year, q] = period.split('-Q');
            if (!q) return null;
            return usdvndByMonth[`${year}-${QUARTER_END_MONTH[q]}`] ?? null;
        });
        if (usdvndArr.some(v => v !== null)) {
            datasets.push({
                label: 'USD/VND thực tế (bình quân tháng, cuối quý)', data: usdvndArr, yAxisID: 'y1',
                borderColor: '#ef4444', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0,
                fill: false, tension: 0.2, spanGaps: true,
                datalabels: _endpointDatalabelsConfig(0),
            });
        }
    }
    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels: periods, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     title: { display: true, text: 'Triệu USD (ròng)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM 2026-10-10 (user hỏi "các yếu tố trong cán cân tài chính là gì" rồi yêu cầu "tạo thêm biểu
// đồ về các yếu tố cán cân vốn [[cán cân tài chính]] tác động lên FDI" — vẽ các cấu phần CHÍNH của
// Cán cân tài chính (FDI/Đầu tư gián tiếp/Vay nợ nước ngoài), đối chiếu tỷ giá, CÙNG KIỂU với
// renderFxSupplyDemandChart (4 cấu phần vãng lai) để so sánh trực quan 2 loại cán cân.
// KHÁC với cán cân vãng lai (phải LẤY Xuất trừ Nhập vì NHNN báo Thu/Chi là 2 số GỘP dương riêng
// biệt): mỗi dòng Tài sản/Nợ trong Cán cân tài chính NHNN đã tự báo cáo dưới dạng SỐ RÒNG theo quý
// VÀ ĐÃ MANG SẴN DẤU đúng quy ước (dương = tăng Nợ/dòng vốn vào, âm = tăng Tài sản/dòng vốn ra —
// xem note ở _IMF_BOP_FIELD_MAP trong fetch_macro_data.py) — nên ở đây VẼ THẲNG từng dòng, KHÔNG tự
// suy diễn công thức trừ/cộng nào thêm (tránh bịa số), chỉ thêm 1 đường "TỔNG Cán cân tài chính"
// (bop_sbv_financial_account) để đối chiếu — đường Tổng này LẤY TỪ CHÍNH bảng NHNN, KHÔNG PHẢI
// tổng cộng dồn các đường bên dưới (NHNN còn nhiều dòng nhỏ khác như tiền & tiền gửi, công cụ tài
// chính... không lấy hết nên các đường chi tiết sẽ KHÔNG cộng khớp 100% với đường Tổng).
// SUA 2026-10-10 (user: "có cái VN đầu tư trực tiếp và gián tiếp ra ngoài bỏ đi vì bé quá, vẽ vào
// rối biểu đồ" — bỏ 2 đường "Tài sản" (bop_sbv_fdi_assets_bop/portfolio_assets_bop, VN đầu tư ra
// nước ngoài), giá trị quá nhỏ so với 3 đường còn lại nên nằm dẹt gần 0, chỉ gây rối mắt.
function renderFxFinancialAccountChart(indicators) {
    const canvas = document.getElementById('chart-fx-financial-account');
    const card = document.getElementById('fx-financial-account-chart-card');
    if (!canvas) return;
    const RAW_SERIES = [
        { key: 'bop_sbv_fdi_liabilities_bop', label: 'FDI vào VN (Nợ — dòng vốn chính)', color: '#10b981' },
        { key: 'bop_sbv_portfolio_liabilities_bop', label: 'Đầu tư gián tiếp — Nợ (khối ngoại mua/bán CP-TP VN)', color: '#60a5fa' },
        { key: 'bop_sbv_external_debt_net', label: 'Vay nợ nước ngoài (ròng)', color: '#f59e0b' },
        { key: 'bop_sbv_financial_account', label: 'TỔNG Cán cân tài chính (từ NHNN, không phải tổng 3 đường trên)', color: '#e5e7eb', bold: true },
    ];
    const validSeries = RAW_SERIES.filter(s => indicators[s.key] && indicators[s.key].series.length);
    if (!validSeries.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    const periods = Array.from(new Set(validSeries.flatMap(s =>
        indicators[s.key].series.map(p => p.period)))).filter(p => p >= '2020-Q1').sort();
    const datasets = validSeries.map(s => {
        const byPeriod = Object.fromEntries(indicators[s.key].series.map(p => [p.period, p.value]));
        const data = periods.map(p => { const v = byPeriod[p]; return (v === null || v === undefined) ? null : v; });
        return {
            label: s.label, data, yAxisID: 'y',
            borderColor: s.color, backgroundColor: s.color + '15', fill: false,
            borderWidth: s.bold ? 3 : 2, borderDash: s.dash || [], tension: 0.25,
            pointRadius: s.bold ? 2 : 3, pointBackgroundColor: s.color, spanGaps: true,
            datalabels: _endpointDatalabelsConfig(0),
        };
    });
    // USD/VND thực tế (tham chiếu, trục phải) — cùng cách quy về cuối quý như renderFxSupplyDemandChart.
    const usdvnd = indicators['usdvnd_monthly_avg'];
    if (usdvnd && usdvnd.series.length) {
        const QUARTER_END_MONTH = { '1': '03', '2': '06', '3': '09', '4': '12' };
        const usdvndByMonth = Object.fromEntries(usdvnd.series.map(p => [p.period, p.value]));
        const usdvndArr = periods.map(period => {
            const [year, q] = period.split('-Q');
            if (!q) return null;
            return usdvndByMonth[`${year}-${QUARTER_END_MONTH[q]}`] ?? null;
        });
        if (usdvndArr.some(v => v !== null)) {
            datasets.push({
                label: 'USD/VND thực tế (bình quân tháng, cuối quý)', data: usdvndArr, yAxisID: 'y1',
                borderColor: '#ef4444', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0,
                fill: false, tension: 0.2, spanGaps: true,
                datalabels: _endpointDatalabelsConfig(0),
            });
        }
    }
    const chart = new Chart(canvas, {
        type: 'line',
        data: { labels: periods, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     title: { display: true, text: 'Triệu USD (dương = dòng vào ròng, âm = dòng ra ròng)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM 2026-10-10 (user: "thêm TRƯỚC biểu đồ cán cân vãng lai cho tôi biểu đồ tỷ giá và đường
// cán cân tài chính tổng và cán cân vãng lai tổng để tôi xem tỷ giá biến động do đường nào chính,
// và cột thể hiện cán cân vãng lai + cán cân tài chính để xem tổng vào ra của 2 cái này như nào"
// — 2 đường TỔNG (current_account/financial_account LẤY THẲNG từ bảng NHNN, đã là số ròng/quý,
// không tự tính lại từ cấu phần con) để so trực quan đường nào biến động mạnh/đồng pha hơn với
// tỷ giá, + 1 cột = current_account + financial_account (cộng thẳng, không cần lo quy ước dấu vì
// cả 2 đã cùng quy ước dương=vào/âm=ra). CỐ Ý KHÔNG gọi cột này là "Cán cân tổng thể" (dù theo
// BPM6 CA+FA+E&O ≈ Cán cân tổng thể) vì CÒN THIẾU Lỗi & sai sót (bop_sbv_errors_omissions, user
// đã được giải thích là rất lớn với VN) — ghi rõ cảnh báo trong subtitle HTML để không gây hiểu
// lầm đây là số tổng thể chính thức.
function renderFxBalanceOverviewChart(indicators) {
    const canvas = document.getElementById('chart-fx-balance-overview');
    const card = document.getElementById('fx-balance-overview-chart-card');
    if (!canvas) return;
    const ca = indicators['bop_sbv_current_account'];
    const fa = indicators['bop_sbv_financial_account'];
    if (!ca || !fa || !ca.series.length || !fa.series.length) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    const periods = Array.from(new Set([...ca.series.map(p => p.period), ...fa.series.map(p => p.period)]))
        .filter(p => p >= '2020-Q1').sort();
    const caByPeriod = Object.fromEntries(ca.series.map(p => [p.period, p.value]));
    const faByPeriod = Object.fromEntries(fa.series.map(p => [p.period, p.value]));
    const caArr = periods.map(p => caByPeriod[p] ?? null);
    const faArr = periods.map(p => faByPeriod[p] ?? null);
    const sumArr = periods.map((p, i) => (caArr[i] === null || faArr[i] === null) ? null : caArr[i] + faArr[i]);

    const datasets = [
        { label: 'Vãng lai + Tài chính (cộng, KHÔNG phải Cán cân tổng thể — xem ghi chú)', data: sumArr,
          type: 'bar', yAxisID: 'y', order: 3,
          backgroundColor: sumArr.map(v => v === null ? '#999' : (v >= 0 ? 'rgba(16,185,129,0.55)' : 'rgba(239,68,68,0.55)')) },
        { label: 'Cán cân vãng lai (tổng)', data: caArr, type: 'line', yAxisID: 'y', order: 1,
          borderColor: '#60a5fa', borderWidth: 2.5, pointRadius: 3, pointBackgroundColor: '#60a5fa', tension: 0.2, spanGaps: true,
          fill: { target: 'origin', above: '#60a5fa22', below: '#60a5fa22' },
          datalabels: _endpointDatalabelsConfig(0) },
        { label: 'Cán cân tài chính (tổng)', data: faArr, type: 'line', yAxisID: 'y', order: 1,
          borderColor: '#f59e0b', borderWidth: 2.5, pointRadius: 3, pointBackgroundColor: '#f59e0b', tension: 0.2, spanGaps: true,
          fill: { target: 'origin', above: '#f59e0b22', below: '#f59e0b22' },
          datalabels: _endpointDatalabelsConfig(0) },
    ];
    const usdvnd = indicators['usdvnd_monthly_avg'];
    if (usdvnd && usdvnd.series.length) {
        const QUARTER_END_MONTH = { '1': '03', '2': '06', '3': '09', '4': '12' };
        const usdvndByMonth = Object.fromEntries(usdvnd.series.map(p => [p.period, p.value]));
        const usdvndArr = periods.map(period => {
            const [year, q] = period.split('-Q');
            if (!q) return null;
            return usdvndByMonth[`${year}-${QUARTER_END_MONTH[q]}`] ?? null;
        });
        if (usdvndArr.some(v => v !== null)) {
            datasets.push({
                label: 'USD/VND thực tế (bình quân tháng, cuối quý)', data: usdvndArr, type: 'line', yAxisID: 'y1', order: 0,
                borderColor: '#ef4444', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0,
                fill: false, tension: 0.2, spanGaps: true,
                datalabels: _endpointDatalabelsConfig(0),
            });
        }
    }
    const chart = new Chart(canvas, {
        type: 'bar',
        data: { labels: periods, datasets },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, offset: true, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, position: 'left',
                     // SUA 2026-10-10 (user: "chả biết khi nào cán cân âm lúc nào dương cả" — kẻ
                     // ĐẬM/SÁNG riêng đúng mốc 0 để thấy ngay ranh giới âm/dương, không cần dò
                     // ngược lên trục mới biết).
                     grid: { color: (ctx) => ctx.tick.value === 0 ? 'rgba(255,255,255,0.55)' : 'rgba(255,255,255,0.04)',
                             lineWidth: (ctx) => ctx.tick.value === 0 ? 1.5 : 1 },
                     title: { display: true, text: 'Triệu USD (ròng)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false },
                      title: { display: true, text: 'VND/USD', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    });
    chartInstances.push(chart);
}

// THEM (user 2026-10-01): "thêm biểu đồ tổng cung và tổng cầu ngoại tệ thì sẽ rõ hơn về áp
// lực" — CỘNG CÁC CẤU PHẦN CÙNG CHIỀU (4 thành phần Cung với nhau, 4 thành phần Cầu với nhau —
// phép cộng CÙNG ĐƠN VỊ, KHÔNG rủi ro quy ước dấu như trừ Cung-Cầu thành 1 số) thành 2 ĐƯỜNG
// riêng (Tổng Cung/Tổng Cầu), KHÔNG gộp tiếp thành 1 hiệu số/điểm áp lực duy nhất — vẫn giữ
// đúng nguyên tắc "không tính FX Pressure = Demand − Supply" mà user đã chốt trước đó, chỉ dễ
// so trực quan hơn biểu đồ 8 đường chi tiết ở renderFxSupplyDemandChart.
function renderFxSupplyDemandTotalChart(indicators) {
    const canvas = document.getElementById('chart-fx-supply-demand-total');
    const card = document.getElementById('fx-supply-demand-total-chart-card');
    if (!canvas) return;
    const SUPPLY_KEYS = ['bop_sbv_goods_export', 'bop_sbv_services_export', 'bop_sbv_investment_income_received', 'bop_sbv_secondary_income_received'];
    const DEMAND_KEYS = ['bop_sbv_goods_import', 'bop_sbv_services_import', 'bop_sbv_investment_income_paid', 'bop_sbv_secondary_income_paid'];
    const allKeys = [...SUPPLY_KEYS, ...DEMAND_KEYS];
    if (!allKeys.some(k => indicators[k] && indicators[k].series.length)) { if (card) card.style.display = 'none'; return; }
    if (card) card.style.display = '';

    // Gioi han hien thi tu 2020 (nhu renderFxPressureSignalsChart) - BOP da duoc backfill tu IMF
    // ve 1996-Q1 nen neu khong gioi han, nhan truc X se qua nhieu (>100 ky) gay chong chat/roi mat.
    const periods = Array.from(new Set(allKeys.flatMap(k => (indicators[k]?.series || []).map(p => p.period)))).filter(p => p >= '2020-Q1').sort();
    const sumByPeriod = (keys) => periods.map(p => {
        let sum = null;
        keys.forEach(k => {
            const pt = indicators[k]?.series.find(x => x.period === p);
            if (pt && pt.value !== null && pt.value !== undefined) {
                sum = (sum ?? 0) + pt.value;
            }
        });
        return sum;
    });
    const totalSupply = sumByPeriod(SUPPLY_KEYS);
    const totalDemand = sumByPeriod(DEMAND_KEYS);

    const chart = new Chart(canvas, {
        type: 'line',
        data: {
            labels: periods,
            datasets: [
                { label: 'Tổng Cung (Xuất khẩu + Dịch vụ XK + Thu nhập ĐT thu + Chuyển giao vãng lai thu)',
                  data: totalSupply, borderColor: '#10b981', backgroundColor: '#10b98120', fill: true,
                  tension: 0.25, pointRadius: 3, pointBackgroundColor: '#10b981', borderWidth: 2.5, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(0) },
                { label: 'Tổng Cầu (Nhập khẩu + Dịch vụ NK + Thu nhập ĐT chi + Chuyển giao vãng lai chi)',
                  data: totalDemand, borderColor: '#ef4444', backgroundColor: '#ef444420', fill: true,
                  tension: 0.25, pointRadius: 3, pointBackgroundColor: '#ef4444', borderWidth: 2.5, spanGaps: true,
                  datalabels: _endpointDatalabelsConfig(0) },
            ],
        },
        options: {
            ...CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 10 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false, maxTicksLimit: periods.length } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Triệu USD', color: '#9aa5bd', font: { size: 9 } } },
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

// Bảng GDP theo ngành / CPI theo nhóm hàng (heatmap YoY, user 2026-10-01) — dựng sẵn trong
// _build_level_yoy_heatmap() (template_vimo.py): %YoY TỰ TÍNH từ MỨC (GDP giá so sánh / chỉ số
// CPI, dulieukinhte.com) vì MỨC khác quy mô giữa các ngành/nhóm không so màu trực tiếp được. Cùng
// quy ước màu với renderMonitoringTable() (màu theo thang RIÊNG từng hàng, không so hàng khác) —
// viết hàm RIÊNG (không sửa renderMonitoringTable) vì 2 bảng mới không có cột "key"/"unit"/
// altSourceIdx, period có thể là quý ("YYYY-Qn") hoặc tháng ("YYYY-MM").
function _periodToShortLabelQOrM(period) {
    const mQ = /^(\d{4})-Q([1-4])$/.exec(period);
    if (mQ) return `Q${mQ[2]}-${mQ[1].slice(2)}`;
    return _periodToShortLabel(period);
}

function _renderHeatmapTableGeneric(cardId, elId, table, firstColLabel) {
    const card = document.getElementById(cardId);
    const el = document.getElementById(elId);
    if (!card || !el) return;
    if (!table || !table.rows || !table.rows.length) { card.style.display = 'none'; return; }
    card.style.display = '';

    const periods = table.periods;
    const thead = `<thead><tr><th>${firstColLabel}</th>${periods.map(p => `<th>${_periodToShortLabelQOrM(p)}</th>`).join('')}</tr></thead>`;
    const tbody = table.rows.map(row => {
        const lo = row.colorMin, hi = row.colorMax;
        const cells = row.values.map(v => {
            if (v === null || v === undefined) return `<td class="na">—</td>`;
            let g = hi === lo ? 0.5 : (v - lo) / (hi - lo);
            if (row.goodDirection === 'lower') g = 1 - g;
            const bg = _heatmapColor(g);
            return `<td style="background:${bg}">${formatNumber(v)}%</td>`;
        }).join('');
        return `<tr><th>${row.label}</th>${cells}</tr>`;
    }).join('');
    el.innerHTML = thead + `<tbody>${tbody}</tbody>`;
}

function renderGdpSectorTable(table) {
    _renderHeatmapTableGeneric('gdp-sector-table-card', 'gdp-sector-table', table, 'Ngành (GDP, YoY)');
}

function renderCpiGroupTable(table) {
    _renderHeatmapTableGeneric('cpi-group-table-card', 'cpi-group-table', table, 'Nhóm hàng (CPI, YoY)');
}

function renderExportCommodityTable(table) {
    _renderHeatmapTableGeneric('export-commodity-table-card', 'export-commodity-table', table, 'Mặt hàng XK (YoY)');
}

function renderImportCommodityTable(table) {
    _renderHeatmapTableGeneric('import-commodity-table-card', 'import-commodity-table', table, 'Mặt hàng NK (YoY)');
}

// Bảng giá XK/NK bình quân theo mặt hàng (user 2026-10-01: "cái nào giá tăng thì màu đỏ còn thấp
// thì trắng rồi xanh lá") — màu PHÂN KỲ quanh mốc 0% (KHÁC _heatmapColor min-max theo lịch sử
// riêng từng hàng ở trên): đỏ = giá tăng (YoY dương), trắng = quanh 0%, xanh = giá giảm (YoY âm).
// Thang màu NGƯỠNG CỐ ĐỊNH (user 2026-10-01: "cứ giá tăng trên 30% là đỏ, dưới 30% thì nhạt dần")
// — KHÔNG dùng maxAbs của cả bảng nữa (1 mặt hàng nhảy đột biến, vd "Quặng và khoáng sản khác"
// +575%, sẽ kéo thang giãn ra làm mọi ô khác nhạt màu hẳn dù bản thân tăng/giảm 20-40% vẫn đáng
// chú ý) — DIVERGING_SCALE_PCT = 30 là mốc bão hoà màu (đỏ/xanh đậm nhất), giá trị VƯỢT 30% vẫn
// giữ màu đậm nhất (clamp), không đậm hơn nữa.
const DIVERGING_SCALE_PCT = 30;

function _heatmapColorDiverging(g) {
    // g từ -1 (giảm mạnh, xanh #10b981) tới 0 (trắng) tới +1 (tăng mạnh, đỏ #ef4444).
    const white = [255, 255, 255];
    const red = [239, 68, 68], green = [16, 185, 129];
    const target = g >= 0 ? red : green;
    const t = Math.min(Math.abs(g), 1);
    const r = Math.round(white[0] + (target[0] - white[0]) * t);
    const gr = Math.round(white[1] + (target[1] - white[1]) * t);
    const b = Math.round(white[2] + (target[2] - white[2]) * t);
    return `rgba(${r},${gr},${b},1)`;
}

function _renderHeatmapTableDiverging(cardId, elId, table, firstColLabel) {
    const card = document.getElementById(cardId);
    const el = document.getElementById(elId);
    if (!card || !el) return;
    if (!table || !table.rows || !table.rows.length) { card.style.display = 'none'; return; }
    card.style.display = '';

    const periods = table.periods;
    const thead = `<thead><tr><th>${firstColLabel}</th>${periods.map(p => `<th>${_periodToShortLabelQOrM(p)}</th>`).join('')}</tr></thead>`;
    const tbody = table.rows.map(row => {
        const cells = row.values.map(v => {
            if (v === null || v === undefined) return `<td class="na">—</td>`;
            const bg = _heatmapColorDiverging(v / DIVERGING_SCALE_PCT);
            return `<td style="background:${bg};color:#0b1220">${formatNumber(v)}%</td>`;
        }).join('');
        return `<tr><th>${row.label}</th>${cells}</tr>`;
    }).join('');
    el.innerHTML = thead + `<tbody>${tbody}</tbody>`;
}

function renderExportPriceTable(table) {
    _renderHeatmapTableDiverging('export-price-table-card', 'export-price-table', table, 'Mặt hàng (Giá XK, YoY)');
}

function renderImportPriceTable(table) {
    _renderHeatmapTableDiverging('import-price-table-card', 'import-price-table', table, 'Mặt hàng (Giá NK, YoY)');
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
                legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 9 } } },
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
                    // SUA 2026-10-09 (user: "tôi tắt hết legend rồi mà vẫn còn nét đứt" — đường
                    // "(bình quân lũy kế)" là 1 DATASET RIÊNG, không có mục chú giải (bị filter ở
                    // trên) nên click tắt đường nét liền KHÔNG tự ẩn đường nét đứt cùng cặp theo
                    // mặc định của Chart.js — ghi đè onClick để ẩn/hiện CẢ 2 dataset cùng lúc.
                    onClick: (e, legendItem, legend) => {
                        const chart = legend.chart;
                        const index = legendItem.datasetIndex;
                        const pairLabel = `${chart.data.datasets[index].label} (bình quân lũy kế)`;
                        const pairIndex = chart.data.datasets.findIndex(d => d.label === pairLabel);
                        if (chart.isDatasetVisible(index)) {
                            chart.hide(index);
                            legendItem.hidden = true;
                            if (pairIndex >= 0) chart.hide(pairIndex);
                        } else {
                            chart.show(index);
                            legendItem.hidden = false;
                            if (pairIndex >= 0) chart.show(pairIndex);
                        }
                    },
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

// ══════════════════════════════════════════════════════════════════════════
// THEM 2026-10-09: tab "Báo cáo" — clone layout/nội dung báo cáo tháng kiểu DNL Capital (user gửi
// 4 ảnh mẫu). Dữ liệu + toàn bộ đoạn phân tích đến từ vn_report_tab.build_vn_report() (RULE-BASED,
// không AI viết văn) — hàm renderVnReport() ở đây CHỈ dựng HTML/chart, không tự suy ra số liệu mới.
// ══════════════════════════════════════════════════════════════════════════
function vnMonthlyTable(t) {
    if (!t || !t.rows || !t.rows.length) return '';
    const fmtv = v => (v === null || v === undefined) ? '—' : v;
    const head = t.periods.map(p => `<th>${p}</th>`).join('');
    // SUA 2026-10-09 (user: "bảng dữ liệu này không điền số à" — số liệu THẬT SỰ có trong DOM, chỉ
    // là VÔ HÌNH: ".monitoring-table tbody td { color: #0b1220 }" (vimo.html) là quy ước CHUNG cho
    // các bảng heatmap khác trong dự án, nơi JS tô NỀN SÁNG cho từng ô nên chữ gần đen mới đọc
    // được — bảng này KHÔNG tô nền ô nào nên chữ gần đen chìm hẳn vào nền tối chung của trang) —
    // ghi đè màu chữ TRẮNG trực tiếp trên từng ô, không đụng vào rule CSS dùng chung.
    const rows = t.rows.map(r => `<tr>
        <th style="text-align:left;white-space:nowrap">${r.label} <span class="ind-source-note">(${r.unit})</span></th>
        ${r.values.map(v => `<td style="color:#e5e7eb">${fmtv(v)}</td>`).join('')}
    </tr>`).join('');
    return `<div class="monitoring-table-scroll" style="overflow-x:auto"><table class="monitoring-table">
        <thead><tr><th style="text-align:left">Chỉ báo</th>${head}</tr></thead>
        <tbody>${rows}</tbody></table></div>`;
}

// SUA 2026-10-09 (user: "có chữ rồi nhưng chất lượng điểm ảnh quá thấp, nhòe lắm" — pixelRatio:2
// của html-to-image chỉ nhân độ phân giải phần TEXT/SVG vẽ bởi trình duyệt, còn các <canvas> của
// Chart.js bên trong vẫn chỉ có đúng số PIXEL THẬT đã render ở màn hình màn hình 1x (vd canvas
// rộng 649px) — html-to-image chỉ PHÓNG TO ảnh đã có, không vẽ lại canvas ở độ phân giải cao hơn,
// nên riêng phần biểu đồ bị nhòe khi phóng to. Ép TẤT CẢ chart trong tab Báo cáo tự vẽ ở
// devicePixelRatio cao hơn màn hình thật (không phụ thuộc màn hình người dùng là 1x hay 2x) để
// canvas gốc đã đủ nét ngay từ đầu, không chỉ riêng lúc xuất ảnh.
const VN_CHART_DEFAULTS = { ...CHART_DEFAULTS, devicePixelRatio: 2 };

function vnRetailChart(c) {
    const canvas = document.getElementById('chart-vn-retail');
    if (!canvas || !c || !c.periods.length) return;
    chartInstances.push(new Chart(canvas, {
        type: 'bar',
        data: { labels: c.periods, datasets: [
            { label: 'Bán lẻ hàng hóa', data: c.goods, backgroundColor: '#60a5fa', stack: 'retail', order: 2 },
            { label: 'Dịch vụ lưu trú, ăn uống', data: c.hospitality, backgroundColor: '#f59e0b', stack: 'retail', order: 2 },
            { label: 'Du lịch lữ hành', data: c.travel, backgroundColor: '#a78bfa', stack: 'retail', order: 2 },
            { label: 'Dịch vụ khác', data: c.other, backgroundColor: '#94a3b8', stack: 'retail', order: 2 },
            // SUA 2026-10-09 (user: "line tăng trưởng bị ẩn đằng sau khó nhìn quá" — màu xanh lá
            // mảnh (2px, không điểm) lẫn vào các cột xám/cam phía trên) — đổi màu vàng tương phản
            // mạnh, dày hơn, có điểm + viền đen để nổi hẳn lên trên nền cột, order thấp hơn để
            // Chart.js vẽ ĐÈ lên sau cùng (giống quy ước order ở vnTradeChart bên dưới).
            { type: 'line', label: 'Tăng trưởng THỰC (YoY, trục phải)', data: c.real_yoy, yAxisID: 'y1', order: 1,
              borderColor: '#facc15', backgroundColor: '#facc15', borderWidth: 3, pointRadius: 3,
              pointBackgroundColor: '#facc15', pointBorderColor: '#1a1f2e', pointBorderWidth: 1.5,
              tension: 0.2, spanGaps: true,
              datalabels: { ..._endpointAboveLabelConfig(1, '%'), color: '#facc15' } },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            layout: { padding: { right: 24 } },
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } }, datalabels: { display: false } },
            scales: {
                // SUA 2026-10-09 (user: "cột tháng gần nhất hẹp quá, kéo dịch sang chút" — đo thực
                // tế bằng Chart.getChart().getDatasetMeta() thấy TÂM cột cuối NẰM ĐÚNG TẠI mép phải
                // vùng vẽ (el.x === chartArea.right), nghĩa là NỬA cột cuối bị khuất ra ngoài —
                // category scale của biểu đồ TRỘN bar+line không tự nhận offset:true như bar
                // thường, phải khai báo TƯỜNG MINH mới có khoảng đệm nửa-cột ở 2 đầu trục X).
                x: { ...CHART_DEFAULTS.scales.x, offset: true, stacked: true, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Nghìn tỷ đồng', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false }, title: { display: true, text: '% YoY (thực)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    }));
}

// SUA 2026-10-09 lần 2 (user: "cái CPI này vẫn xấu, trả lại như cũ trước kia đi... CPI cũng đúng
// chứ không thấp bò như này đâu" — đã thử nguồn tươi hơn (11 nhóm %YoY THÔ chưa nhân trọng số) 2
// lần, cả multi-line và stacked-bar đều lộ cùng 1 vấn đề: %YoY thô 5-11 nhóm CỘNG LẠI cao gấp
// nhiều lần CPI thật (không có trọng số để giảm tỷ trọng trước khi cộng), cột luôn vọt cao hẳn so
// với đường CPI — REVERT HẲN về nguồn VBMA gốc (ĐÃ nhân trọng số, cộng đúng ra CPI, cân đối như
// bản user khen đẹp). Xem _build_consumption (vn_report_tab.py) — ưu tiên ĐÚNG TỶ LỆ hơn tươi hơn.
// SUA 2026-10-09 lần 3 (user: "dữ liệu thô kia không tính ra được à, dựa theo cả dữ liệu thô và tỷ
// trọng CPI các kỳ trong quá khứ để tính ra trọng số rồi nhân vào ra tạm số của tháng gần nhất đi
// chứ, để tôi còn tham khảo" — xem _add_provisional_contrib trong vn_report_tab.py: suy ngược
// trọng số từ các kỳ VBMA đã có + %YoY thô cùng kỳ, áp vào %YoY thô mới nhất để ước tính TẠM cột
// còn thiếu. c.provisionalPeriod đánh dấu kỳ đó — vẽ viền NÉT ĐỨT cho CẢ 5 cột của kỳ này để biết
// là số TẠM TÍNH, không phải số VBMA chính thức.
function vnCpiGroupChart(c) {
    const canvas = document.getElementById('chart-vn-cpi-group');
    if (!canvas || !c || !c.periods.length) return;
    const meta = [['food', 'Thực phẩm', '#f59e0b'], ['housing_utilities', 'Nhà, điện, nước', '#60a5fa'],
                  ['healthcare', 'Y tế', '#10b981'], ['transport', 'Vận tải', '#ef4444'], ['other', 'Khác', '#94a3b8']];
    const provIdx = c.provisionalPeriod ? c.periods.indexOf(c.provisionalPeriod) : -1;
    chartInstances.push(new Chart(canvas, {
        type: 'bar',
        data: { labels: c.periods, datasets: [
            ...meta.map(([k, label, color]) => ({
                label, data: c[k], backgroundColor: color, stack: 'cpi', order: 2,
                borderColor: '#fff',
                borderWidth: (ctx) => ctx.dataIndex === provIdx ? 2 : 0,
                borderDash: (ctx) => ctx.dataIndex === provIdx ? [4, 2] : [],
            })),
            { type: 'line', label: 'CPI YoY (tổng, %)', data: c.cpi_yoy, order: 1,
              borderColor: '#22d3ee', backgroundColor: '#22d3ee', borderWidth: 3, pointRadius: 3,
              pointBackgroundColor: '#22d3ee', pointBorderColor: '#1a1f2e', pointBorderWidth: 1.5,
              tension: 0.2, spanGaps: true,
              datalabels: { ..._endpointAboveLabelConfig(2, '%'), color: '#22d3ee' } },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            layout: { padding: { right: 24 } },
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } }, datalabels: { display: false } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, offset: true, stacked: true, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Điểm % đóng góp / CPI YoY', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
        plugins: [ChartDataLabels],
    }));
}

function vnTradeChart(c) {
    const canvas = document.getElementById('chart-vn-trade');
    if (!canvas || !c || !c.periods.length) return;
    chartInstances.push(new Chart(canvas, {
        type: 'bar',
        data: { labels: c.periods, datasets: [
            // SUA 2026-10-09 (user: "biểu đồ cột thấp quá nhìn bé khó nhìn" — cột cán cân TM (biên
            // độ nhỏ, ±5 tỷ USD) dùng CHUNG trục với 2 đường xuất/nhập khẩu (biên độ 0-60 tỷ USD)
            // nên bị "đè" cho lùn hẳn xuống đáy) — tách cột sang trục phải RIÊNG (y1), Chart.js tự
            // giãn trục đó theo đúng biên độ ±5-10 tỷ USD của riêng cán cân, cột sẽ cao hẳn lên.
            { label: 'Cán cân thương mại (tỷ USD, trục phải)', data: c.balance, yAxisID: 'y1',
              backgroundColor: c.balance.map(v => v === null ? '#999' : (v >= 0 ? 'rgba(16,185,129,0.65)' : 'rgba(239,68,68,0.65)')), order: 3 },
            { type: 'line', label: 'Xuất khẩu (tỷ USD)', data: c.export, borderColor: '#60a5fa', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true, order: 1 },
            { type: 'line', label: 'Nhập khẩu (tỷ USD)', data: c.import, borderColor: '#f59e0b', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true, order: 1 },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                // SUA 2026-10-09 (xem ghi chú offset:true ở vnRetailChart).
                x: { ...CHART_DEFAULTS.scales.x, offset: true, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: 'Tỷ USD (xuất/nhập khẩu)', color: '#9aa5bd', font: { size: 9 } } },
                y1: { ...CHART_DEFAULTS.scales.y, position: 'right', grid: { display: false }, title: { display: true, text: 'Tỷ USD (cán cân)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    }));
}

function vnIipChart(iipOverall, sector) {
    const canvas = document.getElementById('chart-vn-iip');
    if (!canvas || !iipOverall || !sector) return;
    const periods = [...new Set([...iipOverall.periods, ...sector.periods])].sort();
    const lookup = (per, val) => { const m = {}; per.forEach((p, i) => { m[p] = val[i]; }); return periods.map(p => (p in m ? m[p] : null)); };
    chartInstances.push(new Chart(canvas, {
        type: 'line',
        data: { labels: periods, datasets: [
            { label: 'IIP tổng (YoY, lũy kế)', data: lookup(iipOverall.periods, iipOverall.values), borderColor: '#e5e7eb', borderWidth: 2, pointRadius: 2, tension: 0.2, spanGaps: true },
            { label: 'Chế biến, chế tạo', data: lookup(sector.periods, sector.manufacturing), borderColor: '#60a5fa', borderWidth: 2, pointRadius: 3, spanGaps: true },
            { label: 'SX & phân phối điện', data: lookup(sector.periods, sector.electricity), borderColor: '#f59e0b', borderWidth: 2, pointRadius: 3, spanGaps: true },
            { label: 'Cấp nước, xử lý rác thải', data: lookup(sector.periods, sector.water_waste), borderColor: '#10b981', borderWidth: 2, pointRadius: 3, spanGaps: true },
            { label: 'Khai khoáng', data: lookup(sector.periods, sector.mining), borderColor: '#ef4444', borderWidth: 2, pointRadius: 3, spanGaps: true },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: '% YoY', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    }));
}

function vnPmiChart(canvasId, pmi) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !pmi || !pmi.periods.length) return;
    chartInstances.push(new Chart(canvas, {
        type: 'line',
        data: { labels: pmi.periods, datasets: [
            { label: 'PMI sản xuất', data: pmi.values, borderColor: '#60a5fa', backgroundColor: '#60a5fa', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true },
            { label: 'Ngưỡng 50 (mở rộng/thu hẹp)', data: pmi.periods.map(() => 50), borderColor: 'rgba(255,255,255,0.35)', borderWidth: 1, borderDash: [4, 4], pointRadius: 0 },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
                y: CHART_DEFAULTS.scales.y,
            },
        },
    }));
}

// So sánh quỹ đạo theo THÁNG (1-12) giữa các năm cho các chỉ báo lũy kế YTD (tín dụng/huy động/
// FDI đăng ký/giải ngân đầu tư công) — xem _year_compare_chart() trong vn_report_tab.py.
function vnYearCompareChart(canvasId, chart, unitLabel) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !chart) return;
    const years = Object.keys(chart.series);
    if (!years.length) return;
    const palette = ['#60a5fa', '#f59e0b', '#a78bfa'];
    // THEM 2026-10-09 (user: "TTHD/TTTD nếu thiếu thì lấy tổng toàn hệ thống bank áp tạm sang, vẽ
    // nét khác để biết là số tạm" — chart.provisionalFrom (xem _with_bank_fallback trong vn_report_
    // tab.py) đánh dấu {year, month} là kỳ ĐẦU TIÊN dùng số dự phòng (tổng 26 NH niêm yết/UPCoM
    // theo BCTC, KHÁC phạm vi "toàn nền kinh tế" của nguồn chính thức) — vẽ NÉT ĐỨT từ đó trở đi
    // bằng segment.borderDash, để không lẫn với số chính thức (nét liền).
    const pf = chart.provisionalFrom;
    chartInstances.push(new Chart(canvas, {
        type: 'line',
        data: { labels: chart.months.map(m => `T${m}`), datasets: years.map((y, i) => {
            const ds = { label: y, data: chart.series[y], borderColor: palette[i % palette.length], borderWidth: 2, pointRadius: 2, spanGaps: true };
            if (pf && pf.year === y) {
                const pIdx = pf.month - 1;
                ds.segment = { borderDash: (ctx) => ctx.p1DataIndex >= pIdx ? [6, 4] : undefined };
                ds.pointStyle = chart.months.map((_, idx) => idx >= pIdx ? 'star' : 'circle');
                ds.pointRadius = chart.months.map((_, idx) => idx >= pIdx ? 5 : 2);
            }
            return ds;
        }) },
        options: {
            ...VN_CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: false } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: unitLabel, color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    }));
}

// THEM 2026-10-09 (user: "sự phụ thuộc nhập khẩu vào nhóm FDI nó ở đây nhé" — tỷ trọng khu vực FDI
// trong tổng KNXK/KNNK, tính từ export_monthly_fdi/domestic + import_monthly_fdi/domestic đã có
// sẵn trong indicators, xem _build_production trong vn_report_tab.py).
function vnFdiDependencyChart(c) {
    const canvas = document.getElementById('chart-vn-fdi-dependency');
    if (!canvas || !c || !c.periods.length) return;
    chartInstances.push(new Chart(canvas, {
        type: 'line',
        data: { labels: c.periods, datasets: [
            { label: 'Tỷ trọng FDI trong Xuất khẩu (%)', data: c.export_fdi_share, borderColor: '#60a5fa', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true },
            { label: 'Tỷ trọng FDI trong Nhập khẩu (%)', data: c.import_fdi_share, borderColor: '#f59e0b', borderWidth: 2, pointRadius: 0, tension: 0.2, spanGaps: true },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, title: { display: true, text: '% tỷ trọng FDI', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    }));
}

// THEM 2026-10-09 (user gửi nguồn dulieukinhte.com/du-lieu/von-fdi-dang-ky-cap-moi-405): FDI đăng
// ký tách cấp mới/điều chỉnh, vẽ stacked-bar CẢ CHUỖI lịch sử (khác các chart "theo năm" khác ở
// mục Đầu tư) để thấy xu hướng CƠ CẤU, không phải so sánh tốc độ 2 năm.
function vnFdiBreakdownChart(c) {
    const canvas = document.getElementById('chart-vn-fdi-breakdown');
    if (!canvas || !c || !c.periods.length) return;
    chartInstances.push(new Chart(canvas, {
        type: 'bar',
        data: { labels: c.periods, datasets: [
            { label: 'Cấp mới (dự án mới)', data: c.new, backgroundColor: '#60a5fa', stack: 'fdi' },
            { label: 'Điều chỉnh (dự án cũ tăng vốn)', data: c.adjusted, backgroundColor: '#f59e0b', stack: 'fdi' },
        ] },
        options: {
            ...VN_CHART_DEFAULTS,
            plugins: { legend: { display: true, labels: { boxWidth: 10, font: { size: 9 } } } },
            scales: {
                x: { ...CHART_DEFAULTS.scales.x, offset: true, stacked: true, ticks: { ...CHART_DEFAULTS.scales.x.ticks, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
                y: { ...CHART_DEFAULTS.scales.y, stacked: true, title: { display: true, text: 'Tỷ USD (lũy kế)', color: '#9aa5bd', font: { size: 9 } } },
            },
        },
    }));
}

function renderVnReport(report) {
    const box = document.getElementById('vn-report-container');
    if (!box || !report) return;
    box.style.display = '';
    // SUA 2026-10-09 (user: "vẫn mờ lắm, chả nhìn thấy gì luôn, chia thành các cụm đi cho đỡ vỡ
    // điểm ảnh" — 1 ảnh DUY NHẤT gộp cả 5 mục cao ~9500px (tỷ lệ cạnh quá dị dạng) bị các nơi hiển
    // thị/nén ảnh (chat, mạng xã hội...) tự thu nhỏ rất mạnh để fit khung xem, nhìn mờ dù dữ liệu
    // gốc đã nét — đổi sang MỖI THẺ (card) có nút chụp RIÊNG, ảnh nhỏ gọn đúng tỷ lệ bình thường,
    // không bị nén khi hiển thị. Nút chụp gắn class "vn-report-no-capture" để saveVnReportSection()
    // LOẠI nó ra khỏi ảnh chụp (xem filter trong htmlToImage.toBlob).
    let sectionSeq = 0;
    const card = (title, body) => {
        sectionSeq += 1;
        const id = `vn-report-section-${sectionSeq}`;
        return `<div class="card margin-top-20" id="${id}">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px">
                <h3 class="border-blue" style="margin:0">${title}</h3>
                <button class="vn-report-no-capture" onclick="saveVnReportSection('${id}', this)"
                    title="Lưu/Copy riêng mục này thành ảnh"
                    style="flex-shrink:0;background:#0ea5e9;border:none;color:#fff;font-weight:600;font-size:0.78em;padding:6px 10px;border-radius:7px;cursor:pointer">
                    📷 Copy ảnh mục này
                </button>
            </div>
            ${body}
        </div>`;
    };
    const para = txt => `<p style="margin:6px 0;line-height:1.6">${txt}</p>`;

    const cons = report.consumption, prod = report.production, inv = report.investment;
    box.innerHTML = `
      ${card('📋 Tóm tắt & Bảng chỉ số kinh tế tháng ' + report.asOf, para(report.summaryText)
          + '<p class="ind-source-note">Toàn bộ số liệu lấy từ các nguồn tự động đã dùng trong các tab khác (GSO/NSO, Hải quan qua dulieukinhte.com, VBMA, vietnambiz) — xem từng chỉ báo ở tab "Giám sát chỉ số" để tra nguồn gốc chi tiết.</p>'
          + '<h4 style="margin:14px 0 6px">1. Bảng chỉ số kinh tế tháng (13 tháng gần nhất)</h4>'
          + vnMonthlyTable(report.monthlyTable)
          + '<p class="ind-source-note">Giải ngân đầu tư công chỉ công bố ở báo cáo DẠNG THÁNG — các kỳ báo cáo quý (3 lần/năm) không có số này, ô sẽ để trống. Khách quốc tế/CPI MoM/IIP theo ngành là chỉ báo MỚI, chuỗi sẽ dài dần theo mỗi lần cập nhật.</p>')}
      ${card('2. Tiêu dùng & Dịch vụ', `
          <div class="ind-chart" style="height:320px"><canvas id="chart-vn-retail"></canvas></div>
          ${para(cons.paragraphs[0])}
          <div class="ind-chart" style="height:320px;margin-top:14px"><canvas id="chart-vn-cpi-group"></canvas></div>
          ${para(cons.paragraphs[1])}
          ${para(cons.paragraphs[2])}
          <p class="ind-source-note">4 phân khúc bán lẻ + CPI theo nhóm: Hải quan/GSO qua dulieukinhte.com và VBMA (đóng góp điểm % đã có trọng số thật). Tăng trưởng THỰC = tăng trưởng danh nghĩa trừ CPI YoY (xấp xỉ). VBMA (cột đóng góp theo nhóm) có thể công bố TRỄ hơn 1 tháng so với CPI tổng (đường line) — khi đó, cột viền NÉT ĐỨT là số TẠM TÍNH (suy ngược trọng số từ các kỳ VBMA đã có + %YoY thô mới nhất của NSO qua dulieukinhte.com), chỉ để tham khảo, sẽ tự thay bằng số chính thức khi VBMA cập nhật.</p>`)}
      ${card('3. Sản xuất & Thương mại', `
          <div class="ind-chart" style="height:320px"><canvas id="chart-vn-trade"></canvas></div>
          ${para(prod.paragraphs[0])}
          <div class="ind-chart" style="height:300px;margin-top:14px"><canvas id="chart-vn-iip"></canvas></div>
          ${para(prod.paragraphs[1])}
          <div class="ind-chart" style="height:260px;margin-top:14px"><canvas id="chart-vn-pmi"></canvas></div>
          ${para(prod.paragraphs[2])}
          <p class="ind-source-note">IIP theo 4 ngành là chỉ báo MỚI (trích từ báo cáo tháng NSO) — chuỗi còn ngắn, sẽ dài dần theo thời gian, KHÔNG lùi được lịch sử xa hơn.</p>
          <div class="ind-chart" style="height:260px;margin-top:14px"><canvas id="chart-vn-fdi-dependency"></canvas></div>
          ${para(prod.paragraphs[3])}
          <p class="ind-source-note">% = Khu vực FDI / (Khu vực FDI + Khu vực trong nước), tính từ kim ngạch THÁNG ĐƠN LẺ (Hải quan qua dulieukinhte.com) — đo mức độ ngoại thương VN phụ thuộc khối FDI (Samsung, Foxconn...) so với DN nội địa.</p>`)}
      ${card('4. Đầu tư, Tín dụng & FDI', `
          <div class="ind-chart" style="height:260px"><canvas id="chart-vn-pmi-long"></canvas></div>
          ${para(inv.paragraphs[0])}
          <div class="bank-chart-row-title">Dòng vốn lũy kế — đầu tư công &amp; FDI (% kế hoạch năm / tỷ USD)</div>
          <div class="bank-chart-grid-3" style="margin-top:6px">
              <div><div class="ind-chart" style="height:260px"><canvas id="chart-vn-public-inv"></canvas></div><p class="ind-source-note" style="text-align:center">Giải ngân đầu tư công (lũy kế, % kế hoạch năm)</p></div>
              <div><div class="ind-chart" style="height:260px"><canvas id="chart-vn-fdi-reg"></canvas></div><p class="ind-source-note" style="text-align:center">FDI đăng ký — vốn CAM KẾT (lũy kế, tỷ USD)</p></div>
              <div><div class="ind-chart" style="height:260px"><canvas id="chart-vn-fdi-disb"></canvas></div><p class="ind-source-note" style="text-align:center">FDI giải ngân — vốn THỰC TẾ đã rót vào nền kinh tế (lũy kế, tỷ USD)</p></div>
          </div>
          <div class="bank-chart-row-title">Tăng trưởng tín dụng &amp; huy động (lũy kế YTD, %)</div>
          <div class="bank-chart-grid-2" style="margin-top:6px">
              <div><div class="ind-chart" style="height:260px"><canvas id="chart-vn-credit"></canvas></div><p class="ind-source-note" style="text-align:center">Tăng trưởng tín dụng (lũy kế YTD, %)</p></div>
              <div><div class="ind-chart" style="height:260px"><canvas id="chart-vn-deposit"></canvas></div><p class="ind-source-note" style="text-align:center">Tăng trưởng huy động vốn (lũy kế YTD, %)</p></div>
          </div>
          ${para(inv.paragraphs[1])}
          <p class="ind-source-note">Các chuỗi lũy kế YTD (tín dụng/huy động/FDI đăng ký/giải ngân đầu tư công) RESET mỗi tháng 1 — so sánh giữa các năm TẠI CÙNG mốc tháng để biết năm nay đang nhanh/chậm hơn năm trước, không so 2 giá trị cuối kỳ khác tháng. ⭐ Điểm nét đứt/hình sao (nếu có) là số TẠM TÍNH từ tổng 26 ngân hàng niêm yết/UPCoM (theo BCTC) khi số chính thức toàn nền kinh tế (NHNN/VBMA) chưa kịp cập nhật — sẽ tự thay bằng số chính thức (nét liền) ngay khi có.</p>
          <div class="ind-chart" style="height:280px;margin-top:14px"><canvas id="chart-vn-fdi-breakdown"></canvas></div>
          ${para(inv.paragraphs[2])}
          <p class="ind-source-note">FDI đăng ký tách theo loại hình (cấp mới/điều chỉnh) — nguồn dulieukinhte.com (Bộ KH&amp;ĐT/Hải quan), lũy kế từ đầu năm. Tổng 2 cột XẤP XỈ fdi_registered_usd_bn (còn thiếu phần "góp vốn, mua cổ phần" không có ở nguồn này).</p>`)}
      <div class="card margin-top-20" id="vn-report-narrative-section">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap">
              <h3 class="border-blue" style="margin:0">📝 Trích Nhận Định — Copy sang bài viết</h3>
              <button class="vn-report-no-capture" onclick="copyVnReportNarrative(this)"
                  title="Copy toàn bộ lời nhận định (văn bản thuần, không kèm ảnh/bảng) để dán sang bài viết khác"
                  style="flex-shrink:0;background:#22c55e;border:none;color:#fff;font-weight:600;font-size:0.78em;padding:6px 10px;border-radius:7px;cursor:pointer">
                  📋 Copy toàn bộ nhận định
              </button>
          </div>
          <p class="ind-source-note" style="margin-top:10px">Gộp lại toàn bộ lời nhận định rule-based (ghép câu từ số liệu thật, không phải văn AI tự viết tự do) ở các mục trên vào 1 chỗ để copy nhanh — bấm nút rồi dán (Ctrl+V) sang bài viết/báo cáo khác.</p>
          <div id="vn-report-narrative-text" style="margin-top:6px">
              <h4 style="margin:14px 0 6px">📋 Tóm tắt tháng ${report.asOf}</h4>
              ${para(report.summaryText)}
              <h4 style="margin:14px 0 6px">🛍️ Tiêu dùng &amp; Dịch vụ</h4>
              ${cons.paragraphs.map(para).join('')}
              <h4 style="margin:14px 0 6px">🏭 Sản xuất &amp; Thương mại</h4>
              ${prod.paragraphs.map(para).join('')}
              <h4 style="margin:14px 0 6px">💰 Đầu tư, Tín dụng &amp; FDI</h4>
              ${inv.paragraphs.map(para).join('')}
          </div>
      </div>
    `;

    vnRetailChart(cons.retailChart);
    vnCpiGroupChart(cons.cpiGroupChart);
    vnTradeChart(prod.tradeChart);
    vnIipChart(prod.iipChart, prod.iipSectorChart);
    vnPmiChart('chart-vn-pmi', prod.pmiChart);
    vnFdiDependencyChart(prod.fdiDependencyChart);
    vnPmiChart('chart-vn-pmi-long', inv.pmiChart);
    vnYearCompareChart('chart-vn-public-inv', inv.publicInvestmentChart, '% kế hoạch năm');
    vnYearCompareChart('chart-vn-fdi-disb', inv.fdiDisbursedChart, 'Tỷ USD');
    vnYearCompareChart('chart-vn-fdi-reg', inv.fdiChart, 'Tỷ USD');
    vnYearCompareChart('chart-vn-credit', inv.creditChart, '%');
    vnYearCompareChart('chart-vn-deposit', inv.depositChart, '%');
    vnFdiBreakdownChart(inv.fdiBreakdownChart);

    // SUA 2026-10-09 (xem ghi chú ở cuối renderUsMacro — cùng lý do, cuộn bảng dài kỳ về tháng mới
    // nhất NGAY SAU KHI render, không chỉ đợi lúc bấm chuyển tab; gọi 2 lần để bắt kịp scrollWidth
    // tăng thêm sau khi Chart.js hoàn tất layout).
    const scrollTablesToLatestVn = () => {
        if (!box.classList.contains('vimo-tab-hidden')) {
            box.querySelectorAll('.monitoring-table-scroll').forEach(el => { el.scrollLeft = el.scrollWidth; });
        }
    };
    setTimeout(scrollTablesToLatestVn, 0);
    setTimeout(scrollTablesToLatestVn, 300);
}

// THEM 2026-10-10 (user báo trên điện thoại "xuất ảnh ra thì ... thiếu biểu đồ, bảng biểu" — điều
// tra bằng Playwright thật với engine WebKit (Safari iPhone/Samsung Internet) mất cả buổi mới ra
// đúng nguyên nhân: html-to-image dựng 1 SVG có <foreignObject> chứa toàn bộ DOM, rồi nạp SVG đó
// làm nguồn cho 1 <img>/<canvas> để rasterize ra ảnh cuối — nhưng WebKit có lỗi lâu năm KHÔNG vẽ
// được ảnh raster (canvas/img) nằm LỒNG BÊN TRONG foreignObject khi SVG đó lại đang được dùng làm
// nguồn cho 1 ảnh khác (coi là "external resource", bị chặn vẽ) — dù html-to-image vẫn tạo ra SVG
// với data:image/png đầy đủ (đã log ra kiểm tra, base64 dài hàng trăm KB, không rỗng), WebKit vẫn
// không chịu vẽ nó ra. Chromium (desktop + máy ảo Android Chrome) không bị lỗi này, nên trước đó
// test trên Chromium tưởng đã ổn. Thử thay canvas bằng <img> rồi mới chụp cũng KHÔNG ăn thua — vì
// vẫn là "ảnh lồng trong foreignObject", y nguyên tình trạng lỗi.
// Khắc phục THẬT: không để canvas/img nào lọt vào trong foreignObject nữa — ẩn tạm các canvas
// (visibility:hidden, vẫn giữ chỗ layout) rồi mới cho html-to-image chụp (lúc này khu vực chụp chỉ
// còn chữ/bảng/nền, phần html-to-image làm tốt), sau đó dán đè chính canvas đó lên vị trí cũ bằng
// ctx.drawImage() ngay trên <canvas> kết quả — drawImage giữa 2 canvas là lệnh 2D thuần, không qua
// SVG/foreignObject nên mọi engine (gồm WebKit) vẽ đúng (đây cũng là cách saveAstroCombinedImage ở
// app_chiemtinh.js đã ghép 2 ảnh chụp riêng thành công — cùng nguyên lý).
async function _captureSectionToCanvas(area, opts) {
    const canvases = Array.from(area.querySelectorAll('canvas'));
    const areaRect = area.getBoundingClientRect();
    const positions = canvases.map((canvas) => {
        const r = canvas.getBoundingClientRect();
        return {
            canvas, prevVisibility: canvas.style.visibility,
            x: r.left - areaRect.left, y: r.top - areaRect.top, w: r.width, h: r.height,
        };
    });
    positions.forEach((p) => { p.canvas.style.visibility = 'hidden'; });
    let outCanvas;
    try {
        outCanvas = await htmlToImage.toCanvas(area, opts);
    } finally {
        positions.forEach((p) => { p.canvas.style.visibility = p.prevVisibility; });
    }
    const scale = outCanvas.width / areaRect.width;
    const ctx = outCanvas.getContext('2d');
    positions.forEach((p) => {
        if (p.w > 0 && p.h > 0) ctx.drawImage(p.canvas, p.x * scale, p.y * scale, p.w * scale, p.h * scale);
    });
    return outCanvas;
}

// THEM 2026-10-09 (user: "toàn bộ phần tab báo cáo này tôi muốn lưu được dưới dạng ảnh, bấm vào là
// copy được nó như 1 tấm ảnh" — chụp DOM thành ảnh, thử COPY vào clipboard trước; nếu trình duyệt
// không hỗ trợ/bị chặn quyền thì TỰ ĐỘNG TẢI FILE xuống thay thế.
// SUA 2026-10-09 (user: "ảnh copy lại chất lượng thấp, không có chữ, mờ lắm" — html2canvas (thử
// trước) tự VIẼ LẠI layout bằng JS nên xử lý SAI với các kiểu CSS hiện đại dự án đang dùng nhiều:
// màu chữ qua CSS custom property và nền card bán trong suốt (rgba) — 2 lỗi html2canvas nổi
// tiếng. Đổi sang html-to-image: lấy toàn bộ DOM làm <foreignObject> trong SVG rồi để CHÍNH TRÌNH
// DUYỆT vẽ, mọi CSS đều ĐÚNG như hiển thị thật.
// SUA 2026-10-09 lần 2 (user: "vẫn mờ lắm, chả nhìn thấy gì luôn, chia thành các cụm đi cho đỡ vỡ
// điểm ảnh" — 1 ảnh DUY NHẤT gộp toàn bộ tab cao ~9500px bị các nơi hiển thị/nén ảnh thu nhỏ rất
// mạnh để fit khung xem, nhìn mờ dù dữ liệu gốc đã nét) — đổi từ 1 nút "chụp toàn bộ" sang MỖI THẺ
// (card) có nút chụp RIÊNG (xem card() trong renderVnReport) — ảnh nhỏ gọn, tỷ lệ cạnh hợp lý,
// không bị nén khi hiển thị trên chat/mạng xã hội. filter loại nút bấm (.vn-report-no-capture) ra
// khỏi ảnh chụp, vì nó là UI điều khiển, không phải nội dung báo cáo.
async function saveVnReportSection(sectionId, btnEl) {
    const area = document.getElementById(sectionId);
    if (!area || typeof htmlToImage === 'undefined') {
        alert('Không tải được công cụ xuất ảnh (html-to-image) — kiểm tra kết nối mạng rồi thử lại.');
        return;
    }
    const prevText = btnEl ? btnEl.innerHTML : '';
    if (btnEl) { btnEl.innerHTML = '⏳ Đang tạo...'; btnEl.disabled = true; }
    try {
        if (document.fonts && document.fonts.ready) await document.fonts.ready;
        // skipFonts:true — tránh html-to-image cố INLINE font Google Fonts (bị CORS chặn đọc
        // cssRules của stylesheet cross-origin, chỉ log lỗi vô hại nhưng gây chậm/ồn console) —
        // không cần nhúng font vào ảnh, trình duyệt đã render chữ đúng font trước khi chụp rồi.
        const outCanvas = await _captureSectionToCanvas(area, {
            backgroundColor: '#0b1220', pixelRatio: 2, skipFonts: true,
            filter: (node) => !(node.classList && node.classList.contains('vn-report-no-capture')),
        });
        const blob = await new Promise((resolve) => outCanvas.toBlob(resolve, 'image/png'));
        if (!blob) {
            if (btnEl) { btnEl.innerHTML = prevText; btnEl.disabled = false; }
            return;
        }
        const filename = `bao-cao-vi-mo-${sectionId}-${new Date().toISOString().slice(0, 10)}.png`;
        const resultLabel = await _deliverImageBlob(blob, filename);
        if (resultLabel === null) {
            // người dùng tự bấm Hủy trên hộp chia sẻ — không phải lỗi, không làm gì thêm.
            if (btnEl) { btnEl.innerHTML = prevText; btnEl.disabled = false; }
            return;
        }
        if (btnEl) {
            btnEl.innerHTML = resultLabel;
            setTimeout(() => { btnEl.innerHTML = prevText; btnEl.disabled = false; }, 2200);
        }
    } catch (e) {
        console.error('Lỗi xuất ảnh báo cáo:', e);
        if (btnEl) { btnEl.innerHTML = prevText; btnEl.disabled = false; }
        alert('Không tạo được ảnh, thử lại sau.');
    }
}

// THEM 2026-10-10 (user: "trên mobile khi bấm lưu ảnh thì lúc xem ok nhưng không thấy tải về" —
// <a download> với blob: URL không đáng tin cậy trên Safari/WebKit di động (thường chỉ MỞ ảnh ra
// xem thay vì lưu file), và navigator.clipboard.write(image) cũng hay bị chặn trên mobile (yêu cầu
// gesture người dùng còn "nóng" ngay lúc gọi, nhưng code đã qua nhiều await trước đó nên mất).
// Khắc phục: ưu tiên Web Share API (navigator.share với file) khi máy hỗ trợ — mở đúng màn hình
// chia sẻ/lưu ảnh gốc của hệ điều hành (iOS: "Lưu ảnh" vào Photos; Android: chia sẻ tới app khác),
// đây là cách CHUẨN và chắc chắn hoạt động để lưu ảnh trên di động. Copy clipboard vẫn thử TRƯỚC
// (tiện cho desktop dán thẳng vào chat), rồi mới tới Share, cuối cùng mới rơi về <a download> (vẫn
// ăn tốt trên desktop Chrome/Firefox và Android Chrome). Trả về null nếu người dùng tự hủy hộp
// chia sẻ (không phải lỗi, không cần tải thay).
async function _deliverImageBlob(blob, filename) {
    try {
        if (navigator.clipboard && window.ClipboardItem) {
            await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
            return '✅ Đã copy ảnh!';
        }
    } catch (e) { /* trình duyệt không hỗ trợ/không cấp quyền clipboard ảnh — thử cách khác */ }
    try {
        const file = new File([blob], filename, { type: 'image/png' });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
            await navigator.share({ files: [file] });
            return '✅ Đã chia sẻ!';
        }
    } catch (e) {
        if (e && e.name === 'AbortError') return null;
        /* không hỗ trợ Web Share API hoặc lỗi khác — rơi xuống tải file */
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    return '✅ Đã tải ảnh!';
}

// THEM 2026-10-10 (user: "tạo thêm cho tôi 1 vùng dưới cuối cùng để copy phần lời đánh giá nhé...
// tôi sẽ copy phần lời này và đẩy sang bài viết cho đẹp" — khác với saveVnReportSection() ở trên
// (chụp ẢNH để dán vào chat/mạng xã hội), đây copy VĂN BẢN THUẦN (không bảng/biểu đồ) để dán vào
// trình soạn bài viết/Word — dùng innerText của #vn-report-narrative-text để giữ đúng xuống dòng
// giữa các tiêu đề phụ/đoạn văn mà không phải tự ghép lại chuỗi text 1 lần nữa.
async function copyVnReportNarrative(btnEl) {
    const el = document.getElementById('vn-report-narrative-text');
    if (!el) return;
    const text = el.innerText.trim();
    const prevText = btnEl ? btnEl.innerHTML : '';
    try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(text);
            if (btnEl) {
                btnEl.innerHTML = '✅ Đã copy!';
                setTimeout(() => { btnEl.innerHTML = prevText; }, 2200);
            }
            return;
        }
        throw new Error('Clipboard API không khả dụng');
    } catch (e) {
        // Trình duyệt không hỗ trợ/không cấp quyền clipboard văn bản (hay gặp trên mobile qua
        // HTTP không an toàn) — hiện hộp thoại có sẵn nội dung để người dùng tự bôi đen & copy.
        if (btnEl) btnEl.innerHTML = prevText;
        window.prompt('Không copy tự động được — bôi đen (Ctrl+A) rồi copy (Ctrl+C) đoạn dưới đây:', text);
    }
}
