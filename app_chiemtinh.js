/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app_chiemtinh.js (Chiêm tinh Tài chính)
   Tải data/astro.json (tính toán thiên văn thật + diễn giải rule-based, xem
   fetch_astro_data.py/template_astro.py). Trọng tâm trang (user 2026-08-04): "Đánh giá Hiện tại"
   (tác động tới tài chính/tâm lý thị trường) + "Dự báo Giai đoạn Tới" (dòng thời gian sự kiện sắp
   tới) — 5 buổi tham khảo (izumi.edu.vn) giữ lại phía dưới làm kiến thức nền. Xem lưu ý quan
   trọng ở đầu fetch_astro_data.py về bản chất khung lý thuyết này.
   ════════════════════════════════════════════════════════ */

'use strict';

const HARD_ASPECT_NAMES = ['Hợp', 'Vuông', 'Xung'];

function _isHardAspect(aspectLabel) {
    return HARD_ASPECT_NAMES.some(n => aspectLabel.startsWith(n));
}

function _fmtDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    if (!m) return iso;
    return `${m[3]}/${m[2]}/${m[1]}`;
}

document.addEventListener('DOMContentLoaded', async () => {
    const data = await fetch('data/astro.json').then(r => r.ok ? r.json() : null).catch(() => null);
    if (!data) {
        document.querySelector('main.view').innerHTML =
            '<div class="loading-state card">Chưa có dữ liệu chiêm tinh. Hãy chạy template_astro.py hoặc GitHub Action "Cập nhật Chiêm tinh Tài chính".</div>';
        return;
    }

    const genEl = document.getElementById('astro-generated-at');
    if (genEl) genEl.textContent = `Dữ liệu thiên văn tính lúc: ${data.generatedAt}`;

    renderSunMoonToday(data.positions);
    renderPositionsTable(data.positions);
    renderCurrentAspectsTable(data.currentAspects);
    renderAssessment(data.currentAssessment);
    renderForecastTimeline(data.forecastTimeline);
    initCycleTool();
});

function renderAssessment(assessment) {
    const card = document.getElementById('assessment-card');
    if (!assessment || !card) return;
    card.style.display = 'block';

    const overallEl = document.getElementById('assessment-overall');
    if (overallEl) overallEl.textContent = assessment.overallText;

    const retroEl = document.getElementById('assessment-retrograde');
    if (retroEl) {
        retroEl.innerHTML = (assessment.retrogradeLines || [])
            .map(l => `<div class="astro-line">℞ ${l}</div>`).join('');
    }

    const aspEl = document.getElementById('assessment-aspects');
    if (aspEl) {
        aspEl.innerHTML = (assessment.aspectLines || [])
            .map(l => `<div class="astro-line">🪐 ${l}</div>`).join('');
    }
}

function renderForecastTimeline(events) {
    const card = document.getElementById('forecast-card');
    const el = document.getElementById('astro-forecast-table');
    if (!events || !card || !el) return;
    card.style.display = 'block';
    if (!events.length) {
        el.innerHTML = '<tbody><tr><td>Không có sự kiện nào đáng chú ý trong giai đoạn tới.</td></tr></tbody>';
        return;
    }
    const EVENT_BADGE = {
        aspect: 'hard', solar: 'solar', lunar: 'lunar',
        station_retrograde: 'hard', station_direct: 'soft',
    };
    const EVENT_ICON = {
        aspect: '🪐', solar: '☀️', lunar: '🌕', station_retrograde: '℞', station_direct: '➡️',
    };
    el.innerHTML = `
        <thead><tr><th>Ngày</th><th>Sự kiện</th><th>Diễn giải (lý thuyết chiêm tinh tài chính)</th></tr></thead>
        <tbody>${events.map(e => `
            <tr>
                <td style="white-space:nowrap">${_fmtDate(e.date)}</td>
                <td style="white-space:nowrap"><span class="astro-badge ${EVENT_BADGE[e.type] || 'soft'}">${EVENT_ICON[e.type] || ''} ${e.title}</span></td>
                <td style="font-size:0.92em">${e.interpretation}</td>
            </tr>
        `).join('')}</tbody>`;
}

function renderSunMoonToday(positions) {
    const el = document.getElementById('astro-sun-moon-today');
    if (!el || !positions) return;
    const sun = positions.find(p => p.name === 'Mặt Trời');
    const moon = positions.find(p => p.name === 'Mặt Trăng');
    if (!sun || !moon) return;
    el.innerHTML = `Hôm nay: Mặt Trời đang ở cung <b>${sun.sign}</b> (${sun.degInSign.toFixed(1)}°), Mặt Trăng đang ở cung <b>${moon.sign}</b> (${moon.degInSign.toFixed(1)}°).`;
}

function renderPositionsTable(positions) {
    const el = document.getElementById('astro-positions-table');
    if (!el || !positions) return;
    el.innerHTML = `
        <thead><tr><th>Hành tinh</th><th>Kinh độ hoàng đạo</th><th>Cung hoàng đạo</th><th>Độ trong cung</th><th>Trạng thái</th></tr></thead>
        <tbody>${positions.map(p => `
            <tr><td>${p.name}</td><td class="num">${p.lon.toFixed(2)}°</td><td>${p.sign}</td><td class="num">${p.degInSign.toFixed(2)}°</td>
                <td>${p.retrograde ? '<span class="astro-badge hard">℞ Nghịch hành</span>' : ''}</td></tr>
        `).join('')}</tbody>`;
}

function renderCurrentAspectsTable(aspects) {
    const el = document.getElementById('astro-current-aspects-table');
    if (!el) return;
    if (!aspects || !aspects.length) {
        el.innerHTML = '<tbody><tr><td>Không có góc chiếu nào trong orb ±3° hiện tại.</td></tr></tbody>';
        return;
    }
    el.innerHTML = `
        <thead><tr><th>Hành tinh 1</th><th>Hành tinh 2</th><th>Góc chiếu</th><th>Orb</th><th>Trạng thái</th></tr></thead>
        <tbody>${aspects.map(a => `
            <tr><td>${a.a}</td><td>${a.b}</td>
                <td><span class="astro-badge ${_isHardAspect(a.aspect) ? 'hard' : 'soft'}">${a.aspect}</span></td>
                <td class="num">${a.orb.toFixed(2)}°</td>
                <td>${a.applying === true ? 'Đang tới (applying)' : a.applying === false ? 'Đang qua (separating)' : ''}</td></tr>
        `).join('')}</tbody>`;
}

// Buổi 2 — công cụ tính mốc chu kỳ, THUẦN CLIENT-SIDE (không cần dữ liệu server) vì mốc neo là
// lựa chọn chủ quan của người phân tích, không có 1 ngày "đúng" duy nhất.
function initCycleTool() {
    const lengthSel = document.getElementById('cycle-length');
    const customWrap = document.getElementById('cycle-custom-wrap');
    const btn = document.getElementById('cycle-calc-btn');
    const resultEl = document.getElementById('cycle-result');
    if (!lengthSel || !btn) return;

    lengthSel.addEventListener('change', () => {
        customWrap.style.display = lengthSel.value === 'custom' ? 'block' : 'none';
    });

    btn.addEventListener('click', () => {
        const anchorStr = document.getElementById('cycle-anchor').value;
        if (!anchorStr) {
            resultEl.innerHTML = '<p class="ind-note">Vui lòng chọn ngày mốc.</p>';
            return;
        }
        const anchor = new Date(anchorStr + 'T00:00:00Z');
        let weeksPerCycle;
        if (lengthSel.value === '50w') weeksPerCycle = 50;
        else if (lengthSel.value === '4y') weeksPerCycle = 208.43; // 4 năm ~ 208.43 tuần (tính theo 365.25 ngày/năm)
        else weeksPerCycle = Math.max(1, parseInt(document.getElementById('cycle-custom-weeks').value, 10) || 13);

        const msPerCycle = weeksPerCycle * 7 * 86400000;
        const now = Date.now();
        const anchorMs = anchor.getTime();
        const cyclesSoFar = Math.floor((now - anchorMs) / msPerCycle);
        const rows = [];
        for (let i = Math.max(0, cyclesSoFar - 1); i < cyclesSoFar + 9; i++) {
            const d = new Date(anchorMs + i * msPerCycle);
            const iso = d.toISOString().slice(0, 10);
            const isPast = d.getTime() < now;
            rows.push(`<tr><td>Chu kỳ #${i}</td><td>${_fmtDate(iso)}</td><td>${isPast ? '(đã qua)' : ''}</td></tr>`);
        }
        resultEl.innerHTML = `
            <table class="astro-table">
                <thead><tr><th>Chu kỳ</th><th>Ngày dự kiến</th><th></th></tr></thead>
                <tbody>${rows.join('')}</tbody>
            </table>`;
    });
}
