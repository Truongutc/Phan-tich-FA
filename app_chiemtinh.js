/* ════════════════════════════════════════════════════════
   AIC FA SYSTEM — app_chiemtinh.js (Chiêm tinh Tài chính)
   Tải data/astro.json (tính toán thiên văn thật + diễn giải rule-based, xem
   fetch_astro_data.py/template_astro.py). Trang tập trung vào DỮ LIỆU SỐNG: vị trí hành tinh +
   góc chiếu hiện tại (kèm diễn giải ngay dưới mỗi bảng), biểu đồ vòng hoàng đạo, và dự báo giai
   đoạn tới — KHÔNG còn nội dung khung "5 buổi học" tham khảo (user 2026-08-04: bỏ nội dung khóa
   học, chỉ giữ dữ liệu + diễn giải). Xem lưu ý quan trọng ở đầu fetch_astro_data.py về bản chất
   khung lý thuyết này.
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
    // ?t=... để tránh trình duyệt/CDN cache lại astro.json cũ — file này cập nhật hàng tuần, cache
    // cứng khiến user thấy dữ liệu backtest lệch ngày dù server đã có bản mới (2026-08-04).
    const data = await fetch('data/astro.json?t=' + Date.now()).then(r => r.ok ? r.json() : null).catch(() => null);
    if (!data) {
        document.querySelector('main.view').innerHTML =
            '<div class="loading-state card">Chưa có dữ liệu chiêm tinh. Hãy chạy template_astro.py hoặc GitHub Action "Cập nhật Chiêm tinh Tài chính".</div>';
        return;
    }

    const genEl = document.getElementById('astro-generated-at');
    if (genEl) genEl.textContent = `Dữ liệu thiên văn tính lúc: ${data.generatedAt}`;

    renderSunMoonToday(data.positions);
    renderChartWheel(data.positions, data.currentAspects, data.houses);
    renderPositionsTable(data.positions);
    renderCurrentAspectsTable(data.currentAspects);
    renderHousesTable(data.positions, data.houses);
    renderVerdictBanner(data.marketVerdict);
    renderPressureChart(data.dailyPressure);
    renderBacktestChart(data.backtest);
    renderAssessment(data.currentAssessment);
    renderForecastTimeline(data.forecastTimeline);
    initCycleTool();
});

function renderVerdictBanner(verdict) {
    const banner = document.getElementById('verdict-banner');
    if (!banner || !verdict) return;
    banner.className = `astro-verdict-banner ${verdict.colorKey}`;
    const labelEl = document.getElementById('verdict-label');
    const detailEl = document.getElementById('verdict-detail');
    if (labelEl) labelEl.textContent = verdict.label;
    if (detailEl) detailEl.textContent = verdict.detail;
}

function _findPressureTurningPoints(scores, threshold) {
    // Thuật toán zigzag chuẩn: xác nhận 1 đỉnh/đáy khi đảo chiều đủ lớn (>= threshold)
    const points = [];
    let trend = 0; // 0=chưa rõ, 1=đang tăng, -1=đang giảm
    let extremeIdx = 0;
    let extremeVal = scores[0];
    for (let i = 1; i < scores.length; i++) {
        const v = scores[i];
        if (trend >= 0 && v >= extremeVal) {
            extremeVal = v; extremeIdx = i; trend = 1;
        } else if (trend <= 0 && v <= extremeVal) {
            extremeVal = v; extremeIdx = i; trend = -1;
        } else if (trend === 1 && (extremeVal - v) >= threshold) {
            points.push({ idx: extremeIdx, type: 'peak' });
            trend = -1; extremeVal = v; extremeIdx = i;
        } else if (trend === -1 && (v - extremeVal) >= threshold) {
            points.push({ idx: extremeIdx, type: 'trough' });
            trend = 1; extremeVal = v; extremeIdx = i;
        }
    }
    if (extremeIdx !== (points.length ? points[points.length - 1].idx : -1)) {
        points.push({ idx: extremeIdx, type: trend === -1 ? 'trough' : 'peak' });
    }
    return points;
}

function renderPressureChart(series) {
    const container = document.getElementById('astro-pressure-chart-container');
    if (!container || !series || !series.length) return;

    const w = 900, h = 250, padL = 40, padR = 10, padT = 26, padB = 24;
    const scores = series.map(p => p.score);
    const minS = Math.min(...scores, 0), maxS = Math.max(...scores, 0);
    const range = (maxS - minS) || 1;
    const plotW = w - padL - padR, plotH = h - padT - padB;

    const xAt = i => padL + (i / (series.length - 1)) * plotW;
    const yAt = s => padT + plotH - ((s - minS) / range) * plotH;
    const zeroY = yAt(0);
    const UP = '#10b981', DOWN = '#ef4444';

    const linePoints = series.map((p, i) => `${xAt(i).toFixed(1)},${yAt(p.score).toFixed(1)}`).join(' ');
    const areaPoints = `${padL.toFixed(1)},${zeroY.toFixed(1)} ` + linePoints + ` ${xAt(series.length - 1).toFixed(1)},${zeroY.toFixed(1)}`;

    let svg = `<svg viewBox="0 0 ${w} ${h}" style="width:100%;height:auto;display:block" preserveAspectRatio="none">`;
    svg += `<line x1="${padL}" y1="${zeroY.toFixed(1)}" x2="${w - padR}" y2="${zeroY.toFixed(1)}" stroke="#6b7280" stroke-width="1" stroke-dasharray="4,4"/>`;
    svg += `<polygon points="${areaPoints}" fill="#8b5cf615"/>`;

    // Vẽ từng đoạn line, tô màu theo chiều tăng/giảm so với ngày trước
    for (let i = 0; i < series.length - 1; i++) {
        const rising = scores[i + 1] >= scores[i];
        svg += `<line x1="${xAt(i).toFixed(1)}" y1="${yAt(scores[i]).toFixed(1)}" x2="${xAt(i + 1).toFixed(1)}" y2="${yAt(scores[i + 1]).toFixed(1)}" stroke="${rising ? UP : DOWN}" stroke-width="3.5"/>`;
    }

    // Lưới ngày trục hoành (mỗi ~10 ngày)
    series.forEach((p, i) => {
        if (i % 10 === 0 || i === series.length - 1) {
            const x = xAt(i);
            svg += `<line x1="${x.toFixed(1)}" y1="${padT}" x2="${x.toFixed(1)}" y2="${h - padB}" stroke="#374151" stroke-width="0.5"/>`;
            svg += `<text x="${x.toFixed(1)}" y="${h - 6}" fill="#9ca3af" font-size="10" text-anchor="middle">${_fmtDate(p.date).slice(0, 5)}</text>`;
        }
    });

    // Đánh dấu ngày tại các đỉnh/đáy đảo chiều đáng kể (lọc nhiễu bằng ngưỡng biên độ)
    const threshold = Math.max(range * 0.15, 0.5);
    const turningPoints = _findPressureTurningPoints(scores, threshold);
    turningPoints.forEach(({ idx, type }) => {
        const x = xAt(idx), y = yAt(scores[idx]);
        const isPeak = type === 'peak';
        const color = isPeak ? UP : DOWN;
        const label = _fmtDate(series[idx].date);
        let labelY = isPeak ? y - 12 : y + 20;
        labelY = Math.min(Math.max(labelY, padT + 8), h - padB - 4);
        const boxW = label.length * 5.6 + 8;
        const labelX = Math.min(Math.max(x, padL + boxW / 2), w - padR - boxW / 2);
        svg += `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3.5" fill="${color}" stroke="#111827" stroke-width="1"/>`;
        svg += `<rect x="${(labelX - boxW / 2).toFixed(1)}" y="${(labelY - 10).toFixed(1)}" width="${boxW.toFixed(1)}" height="13" rx="3" fill="#111827cc"/>`;
        svg += `<text x="${labelX.toFixed(1)}" y="${labelY.toFixed(1)}" fill="${color}" font-size="10" font-weight="600" text-anchor="middle">${label}</text>`;
    });

    svg += `<text x="${padL - 6}" y="${padT + 4}" fill="#9ca3af" font-size="10" text-anchor="end">${maxS.toFixed(1)}</text>`;
    svg += `<text x="${padL - 6}" y="${(h - padB).toFixed(1)}" fill="#9ca3af" font-size="10" text-anchor="end">${minS.toFixed(1)}</text>`;
    svg += `</svg>`;
    container.innerHTML = svg;
}

function _d2ts(dateStr) {
    return Math.floor(new Date(dateStr + 'T00:00:00Z').getTime() / 1000);
}

// Cấu hình theme dùng chung cho biểu đồ Lightweight Charts (khớp phong cách site tham chiếu
// aic-proweb.vercel.app — "🌐 Cuộn để zoom", "🖐️ Kéo để pan").
const LWC_THEME = {
    layout: { background: { color: 'transparent' }, textColor: '#9ca3af' },
    grid: { vertLines: { color: 'rgba(255,255,255,0.05)' }, horzLines: { color: 'rgba(255,255,255,0.05)' } },
    crosshair: { mode: 1 },
    timeScale: {
        borderColor: 'rgba(255,255,255,0.12)', timeVisible: false, secondsVisible: false,
        fixLeftEdge: false, fixRightEdge: false, rightOffset: 5, minBarSpacing: 1,
    },
    handleScale: { mouseWheel: true, pinch: true, axisPressedMouseMove: { time: true, price: true } },
    handleScroll: { mouseWheel: true, pressedMouseMove: true, horzTouchDrag: true, vertTouchDrag: true },
};

function renderBacktestChart(backtest) {
    const card = document.getElementById('backtest-card');
    const container = document.getElementById('astro-backtest-chart-container');
    if (!card || !container || !backtest || !backtest.astro || !backtest.astro.length || !backtest.vnindex || !backtest.vnindex.length) return;
    if (typeof LightweightCharts === 'undefined') {
        container.innerHTML = '<div class="astro-generated-note">Không tải được thư viện biểu đồ (lightweight-charts) — kiểm tra kết nối mạng.</div>';
        card.style.display = 'block';
        return;
    }
    card.style.display = 'block';

    const astroSeriesAll = backtest.astro; // [{date, score}] đủ mỗi ngày lịch
    const vnSeries = backtest.vnindex;  // [{date, close}] chỉ các phiên giao dịch thật

    // Tạo Set các ngày giao dịch của VN-Index để lọc astroSeries khớp 1:1 theo phiên giao dịch
    const vnDateSet = new Set(vnSeries.map(p => p.date));
    const astroSeries = astroSeriesAll.filter(p => vnDateSet.has(p.date));

    const chart = LightweightCharts.createChart(container, {
        ...LWC_THEME,
        width: container.clientWidth,
        height: 420,
        rightPriceScale: { borderColor: 'rgba(255,255,255,0.12)', autoScale: true, scaleMargins: { top: 0.1, bottom: 0.1 } },
        leftPriceScale: { visible: true, borderColor: 'rgba(255,255,255,0.12)', autoScale: true, scaleMargins: { top: 0.1, bottom: 0.1 } },
    });

    // VN-Index thật (trục phải, trắng)
    const vnLine = chart.addLineSeries({
        priceScaleId: 'right', color: '#e5e7eb', lineWidth: 2, title: 'VN-Index',
        priceFormat: { type: 'price', precision: 0, minMove: 1 },
    });
    vnLine.setData(vnSeries.map(p => ({ time: _d2ts(p.date), value: p.close })));

    // Chỉ số áp lực/thuận lợi chiêm tinh (trục trái) — tô XANH khi tăng so với hôm trước, ĐỎ khi
    // giảm (user yêu cầu, để đối chiếu trực quan với chiều VN-Index), vẽ dạng ĐƯỜNG liền (không
    // phải cột) để đọc dễ như biểu đồ dự báo phía trên.
    const astroAnchor = chart.addLineSeries({
        // series "neo" vô hình — chỉ để tra priceToCoordinate đúng theo trục trái, không hiện line thật
        priceScaleId: 'left', color: 'transparent', lineWidth: 1, lineVisible: false,
        lastValueVisible: false, priceLineVisible: false, crosshairMarkerVisible: false, title: '',
    });
    astroAnchor.setData(astroSeries.map(p => ({ time: _d2ts(p.date), value: p.score })));

    const astroTimes = astroSeries.map(p => _d2ts(p.date));
    const astroByDate = {};
    astroSeriesAll.forEach(p => { astroByDate[p.date] = p.score; });
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'position:absolute;top:0;left:0;pointer-events:none;z-index:2;';
    
    // Đặt canvas overlay vào ĐÚNG khung vẽ (pane container) của Lightweight Charts thay vì container tổng ngoài cùng.
    // Trục giá trái (leftPriceScale) chiếm khoảng ~50-60px bên trái container tổng.
    // Hàm timeToCoordinate() trả về tọa độ pixel tính từ mép trái của vùng vẽ (Plot Area).
    // Nếu gắn canvas ở container tổng ngoài cùng (left:0), đường vẽ sẽ bị lệch sang trái đúng bằng chiều rộng trục trái.
    // Khi zoom out (mỗi bar chỉ 2px), lệch 50px = lệch 25 ngày! Khi zoom in (mỗi bar 10px), lệch 50px = 5 ngày.
    // Gắn canvas vào parentElement của chart canvas sẽ giúp gốc (0,0) của canvas trùng 100% với Plot Area.
    const chartCanvas = container.querySelector('canvas');
    const paneContainer = chartCanvas ? chartCanvas.parentElement : container;
    if (paneContainer !== container) {
        paneContainer.style.position = 'relative';
    }
    paneContainer.appendChild(canvas);

    function drawAstroOverlayNow() {
        const rect = paneContainer.getBoundingClientRect();
        const dpr = window.devicePixelRatio || 1;
        canvas.width = rect.width * dpr;
        canvas.height = rect.height * dpr;
        canvas.style.width = rect.width + 'px';
        canvas.style.height = rect.height + 'px';
        const ctx = canvas.getContext('2d');
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.clearRect(0, 0, rect.width, rect.height);

        const range = chart.timeScale().getVisibleRange();
        if (!range) return;
        let startIdx = astroTimes.findIndex(t => t >= range.from);
        if (startIdx === -1) startIdx = astroTimes.length - 1;
        startIdx = Math.max(0, startIdx - 1);
        let endIdx = astroTimes.length - 1;
        for (let i = startIdx; i < astroTimes.length; i++) {
            if (astroTimes[i] > range.to) { endIdx = i; break; }
        }
        endIdx = Math.min(astroTimes.length - 1, endIdx + 1);

        ctx.lineWidth = 1.6;
        for (let i = startIdx; i < endIdx; i++) {
            const x1 = chart.timeScale().timeToCoordinate(astroTimes[i]);
            const y1 = astroAnchor.priceToCoordinate(astroSeries[i].score);
            const x2 = chart.timeScale().timeToCoordinate(astroTimes[i + 1]);
            const y2 = astroAnchor.priceToCoordinate(astroSeries[i + 1].score);
            if (x1 === null || y1 === null || x2 === null || y2 === null) continue;
            ctx.strokeStyle = astroSeries[i + 1].score >= astroSeries[i].score ? '#10b981' : '#ef4444';
            ctx.beginPath();
            ctx.moveTo(x1, y1);
            ctx.lineTo(x2, y2);
            ctx.stroke();
        }
    }

    // Gộp nhiều lần gọi vẽ lại trong cùng 1 khung hình (rAF) — khi zoom/pan bằng chuột/pinch, chart
    // bắn RẤT NHIỀU sự kiện subscribeVisibleTimeRangeChange liên tiếp trong lúc đang cử động; nếu
    // vẽ đồng bộ (synchronous) mỗi lần thì canvas có thể vẽ chậm hơn tốc độ sự kiện dồn tới, gây
    // cảm giác đường chiêm tinh "thụt lại"/trễ so với chart gốc trong lúc đang zoom/kéo (user phản
    // ánh 2026-08-04). rAF đảm bảo mỗi khung hình chỉ vẽ ĐÚNG 1 LẦN, luôn dùng trạng thái MỚI NHẤT.
    let rafScheduled = false;
    function drawAstroOverlay() {
        if (rafScheduled) return;
        rafScheduled = true;
        requestAnimationFrame(() => { rafScheduled = false; drawAstroOverlayNow(); });
    }
    chart.timeScale().subscribeVisibleTimeRangeChange(drawAstroOverlay);

    // Chú thích theo con trỏ chuột — hiện ĐÚNG giá trị 2 đường TẠI CÙNG 1 NGÀY khi di chuột, để tự
    // đối chiếu chính xác thay vì áng chừng bằng mắt (biểu đồ nhiều điểm dao động rất dễ nhìn nhầm
    // sang đỉnh/đáy lân cận — user 2026-08-04 nghi ngờ lệch ngày; đã verify bằng cách so khớp từng
    // pixel màu thật trên canvas, tọa độ tính đúng 100%, không có lỗi — nhưng vẫn thêm chú thích này
    // để không ai phải đoán bằng mắt nữa).
    const legendEl = document.getElementById('astro-backtest-legend');
    if (legendEl) {
        chart.subscribeCrosshairMove(param => {
            if (!param || !param.time) return;
            const dateStr = new Date(param.time * 1000).toISOString().slice(0, 10);
            const vnData = param.seriesData.get(vnLine);
            const astroScore = astroByDate[dateStr];
            const vnPart = vnData ? `<b>${vnData.value.toFixed(1)}</b>` : '<span style="color:#6b7280">không có phiên</span>';
            const astroPart = typeof astroScore === 'number' ? `<b>${astroScore >= 0 ? '+' : ''}${astroScore.toFixed(2)}</b>` : '—';
            legendEl.innerHTML = `Ngày <b>${_fmtDate(dateStr)}</b>: VN-Index ${vnPart} · Chỉ số chiêm tinh ${astroPart}`;
        });
    }

    // Mặc định hiển thị ~1 năm gần nhất; cuộn chuột/pinch để zoom ra xa xem hết 10 năm, kéo để pan
    const fromIdx = Math.max(0, astroSeries.length - 366);
    chart.timeScale().setVisibleRange({
        from: _d2ts(astroSeries[fromIdx].date),
        to: _d2ts(astroSeries[astroSeries.length - 1].date),
    });

    const ro = new ResizeObserver(entries => {
        const rect = entries[0].contentRect;
        if (rect.width > 0) chart.resize(rect.width, 420);
        drawAstroOverlay();
    });
    ro.observe(container);
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
        <thead><tr><th>Hành tinh</th><th>Kinh độ hoàng đạo</th><th>Cung hoàng đạo</th><th>Độ trong cung</th><th>Nhà</th><th>Trạng thái</th></tr></thead>
        <tbody>${positions.map(p => `
            <tr><td>${p.name}</td><td class="num">${p.lon.toFixed(2)}°</td><td>${p.sign}</td><td class="num">${p.degInSign.toFixed(2)}°</td>
                <td class="num">${p.house}</td>
                <td>${p.retrograde ? '<span class="astro-badge hard">℞ Nghịch hành</span>' : ''}</td></tr>
        `).join('')}</tbody>`;
}

function renderHousesTable(positions, houses) {
    const introEl = document.getElementById('astro-house-intro');
    const el = document.getElementById('astro-houses-table');
    if (!el || !houses) return;

    if (introEl) {
        introEl.innerHTML = `"Nhà" (House) khác vị trí hành tinh/cung hoàng đạo ở trên — Nhà phụ thuộc 1 ĐỊA ĐIỂM cụ thể trên Trái Đất (dựa trên đường chân trời tại nơi đó lúc quan sát). Địa điểm tham chiếu: <b>${houses.locationName}</b>. Điểm Mọc (Ascendant, khởi đầu Nhà 1): <b>${houses.ascendant.toFixed(1)}°</b> · Thiên Đỉnh (Midheaven): <b>${houses.midheaven.toFixed(1)}°</b>. Hệ Nhà Đều (Equal House — mỗi Nhà đúng 30° từ Ascendant).`;
    }

    const byHouse = {};
    positions.forEach(p => {
        (byHouse[p.house] = byHouse[p.house] || []).push(p.name);
    });

    const rows = [];
    for (let h = 1; h <= 12; h++) {
        const planetsHere = byHouse[h] || [];
        const highlight = [2, 5, 8].includes(h);
        rows.push(`
            <tr${highlight ? ' style="background:rgba(139,92,246,0.06)"' : ''}>
                <td class="num">${h}${highlight ? ' 💰' : ''}</td>
                <td>${houses.meanings[h] || houses.meanings[String(h)] || ''}</td>
                <td>${planetsHere.length ? planetsHere.join(', ') : '—'}</td>
            </tr>`);
    }
    el.innerHTML = `
        <thead><tr><th>Nhà</th><th>Ý nghĩa (lăng kính tài chính)</th><th>Hành tinh đang trú</th></tr></thead>
        <tbody>${rows.join('')}</tbody>`;
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

// Diễn giải rule-based (từ template_astro.py) — phân bổ vào ĐÚNG bảng liên quan: kết luận tổng
// thể lên đầu trang, dòng nghịch hành xuống dưới bảng vị trí, dòng góc chiếu xuống dưới bảng aspect.
function renderAssessment(assessment) {
    if (!assessment) return;

    const overallCard = document.getElementById('assessment-overall-card');
    const overallEl = document.getElementById('assessment-overall');
    if (overallCard && overallEl) {
        overallCard.style.display = 'block';
        overallEl.textContent = assessment.overallText;
    }

    const retroEl = document.getElementById('assessment-retrograde');
    if (retroEl) {
        retroEl.innerHTML = (assessment.retrogradeLines || []).length
            ? assessment.retrogradeLines.map(l => `<div class="astro-line">℞ ${l}</div>`).join('')
            : '<div class="astro-line">Không có hành tinh chính nào đang nghịch hành.</div>';
    }

    const aspEl = document.getElementById('assessment-aspects');
    if (aspEl) {
        aspEl.innerHTML = (assessment.aspectLines || []).length
            ? `<p class="astro-section-title" style="font-size:0.85em">Diễn biến này nói lên điều gì? (góc chiếu cứng giữa các hành tinh chu kỳ chậm — quan trọng nhất cho xu hướng dài hơi)</p>`
              + assessment.aspectLines.map(l => `<div class="astro-line">🪐 ${l}</div>`).join('')
            : '<div class="astro-line">Không có góc chiếu cứng lớn nào giữa các hành tinh chu kỳ chậm hiện tại.</div>';
    }

    const houseEl = document.getElementById('assessment-houses');
    if (houseEl) {
        houseEl.innerHTML = (assessment.houseLines || []).length
            ? `<p class="astro-section-title" style="font-size:0.85em">💰 Hành tinh đang trú tại Nhà tài chính (2, 5, 8) — nói lên điều gì?</p>`
              + assessment.houseLines.map(l => `<div class="astro-line">🏠 ${l}</div>`).join('')
            : '<div class="astro-line">Không có hành tinh nào đang ở Nhà 2 (Tiền bạc)/5 (Đầu cơ)/8 (Nợ-Khủng hoảng) hiện tại.</div>';
    }
}

const SENTIMENT_CLASS = { positive: 'astro-sentiment-positive', negative: 'astro-sentiment-negative', tension: 'astro-sentiment-tension', neutral: 'astro-sentiment-neutral' };
const SENTIMENT_BADGE = { positive: 'soft', negative: 'hard', tension: 'tension', neutral: 'solar' };
const EVENT_ICON = { aspect: '🪐', solar: '☀️', lunar: '🌕', station_retrograde: '℞', station_direct: '➡️' };

function renderForecastTimeline(events) {
    const card = document.getElementById('forecast-card');
    const el = document.getElementById('astro-forecast-table');
    if (!events || !card || !el) return;
    card.style.display = 'block';
    if (!events.length) {
        el.innerHTML = '<tbody><tr><td>Không có sự kiện nào đáng chú ý trong giai đoạn tới.</td></tr></tbody>';
        return;
    }
    el.innerHTML = `
        <thead><tr><th>Ngày</th><th>Sự kiện</th><th>Diễn giải (lý thuyết chiêm tinh tài chính)</th><th>Điểm ròng hôm đó</th></tr></thead>
        <tbody>${events.map(e => {
            const sentimentClass = SENTIMENT_CLASS[e.sentiment] || 'astro-sentiment-neutral';
            const badgeClass = SENTIMENT_BADGE[e.sentiment] || 'solar';
            let netCell = '<span style="color:#6b7280">—</span>';
            if (typeof e.netScore === 'number') {
                const trendUp = e.netTrend === 'up';
                const trendDown = e.netTrend === 'down';
                const color = trendUp ? '#10b981' : (trendDown ? '#ef4444' : '#9ca3af');
                const arrow = trendUp ? '↑' : (trendDown ? '↓' : '');
                netCell = `<span style="color:${color};font-weight:600;white-space:nowrap">${e.netScore >= 0 ? '+' : ''}${e.netScore.toFixed(2)} ${arrow}</span>`;
            }
            return `
            <tr>
                <td style="white-space:nowrap">${_fmtDate(e.date)}</td>
                <td style="white-space:nowrap"><span class="astro-badge ${badgeClass}">${EVENT_ICON[e.type] || ''} ${e.title}</span></td>
                <td style="font-size:0.92em" class="${sentimentClass}">${e.interpretation}</td>
                <td class="num">${netCell}</td>
            </tr>`;
        }).join('')}</tbody>`;
}

// Biểu đồ vòng hoàng đạo đơn giản hóa — SVG thuần client-side từ positions/currentAspects đã có
// sẵn trong data/astro.json, không cần tính toán gì thêm. Quy ước: kinh độ 0° (Bạch Dương) đặt ở
// đỉnh (12h), tăng dần THEO CHIỀU KIM ĐỒNG HỒ — đây là cách trình bày đơn giản để dễ đọc, KHÔNG
// nhất thiết khớp quy ước phần mềm chiêm tinh chuyên nghiệp (thường đặt 0° Bạch Dương bên trái,
// tăng ngược chiều kim đồng hồ theo góc nhìn "từ Trái Đất nhìn lên bầu trời").
const PLANET_SYMBOL = {
    'Mặt Trời': '☉', 'Mặt Trăng': '☽', 'Sao Thủy': '☿', 'Sao Kim': '♀', 'Sao Hỏa': '♂',
    'Sao Mộc': '♃', 'Sao Thổ': '♄', 'Sao Thiên Vương': '♅', 'Sao Hải Vương': '♆', 'Sao Diêm Vương': '♇',
};
const PLANET_COLOR = {
    'Mặt Trời': '#f59e0b', 'Mặt Trăng': '#cbd5e1', 'Sao Thủy': '#22d3ee', 'Sao Kim': '#f472b6',
    'Sao Hỏa': '#ef4444', 'Sao Mộc': '#a78bfa', 'Sao Thổ': '#94a3b8', 'Sao Thiên Vương': '#34d399',
    'Sao Hải Vương': '#60a5fa', 'Sao Diêm Vương': '#c084fc',
};
const ZODIAC_SYMBOLS = ['♈', '♉', '♊', '♋', '♌', '♍', '♎', '♏', '♐', '♑', '♒', '♓'];

function renderChartWheel(positions, aspects, houses) {
    const container = document.getElementById('astro-wheel-container');
    const legendEl = document.getElementById('astro-wheel-legend');
    if (!container || !positions || !positions.length) return;

    const size = 500, cx = size / 2, cy = size / 2, outerR = 230, ringR = 205, planetR = 175, houseR = 140;
    const toXY = (lon, r) => {
        const rad = (lon - 90) * Math.PI / 180;
        return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
    };

    let svg = `<svg viewBox="0 0 ${size} ${size}" style="width:100%;max-width:480px;display:block;margin:0 auto">`;
    svg += `<circle cx="${cx}" cy="${cy}" r="${outerR}" fill="none" stroke="#374151" stroke-width="1.5"/>`;
    svg += `<circle cx="${cx}" cy="${cy}" r="${ringR}" fill="none" stroke="#374151" stroke-width="1"/>`;
    for (let i = 0; i < 12; i++) {
        const [x1, y1] = toXY(i * 30, ringR);
        const [x2, y2] = toXY(i * 30, outerR);
        svg += `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="#374151" stroke-width="1"/>`;
        const [lx, ly] = toXY(i * 30 + 15, (outerR + ringR) / 2);
        svg += `<text x="${lx.toFixed(1)}" y="${ly.toFixed(1)}" fill="#9ca3af" font-size="15" text-anchor="middle" dominant-baseline="central">${ZODIAC_SYMBOLS[i]}</text>`;
    }
    // 12 vạch Nhà (House) — mờ hơn, xoay theo Ascendant (khác vòng cung hoàng đạo cố định ở trên).
    if (houses && houses.cusps) {
        houses.cusps.forEach((cusp, i) => {
            const [hx1, hy1] = toXY(cusp, 30);
            const [hx2, hy2] = toXY(cusp, houseR);
            const financial = [2, 5, 8].includes(i + 1);
            svg += `<line x1="${hx1.toFixed(1)}" y1="${hy1.toFixed(1)}" x2="${hx2.toFixed(1)}" y2="${hy2.toFixed(1)}" stroke="${financial ? '#8b5cf699' : '#4b556388'}" stroke-width="${financial ? 1.4 : 0.8}" stroke-dasharray="3,3"/>`;
            const nextCusp = houses.cusps[(i + 1) % 12];
            const midAngle = cusp + (((nextCusp - cusp + 360) % 360) / 2);
            const [nx, ny] = toXY(midAngle, houseR + 15);
            svg += `<text x="${nx.toFixed(1)}" y="${ny.toFixed(1)}" fill="${financial ? '#c4b5fd' : '#6b7280'}" font-size="11" text-anchor="middle" dominant-baseline="central">${i + 1}</text>`;
        });
    }
    (aspects || []).forEach(a => {
        const pa = positions.find(p => p.name === a.a);
        const pb = positions.find(p => p.name === a.b);
        if (!pa || !pb) return;
        const [x1, y1] = toXY(pa.lon, planetR);
        const [x2, y2] = toXY(pb.lon, planetR);
        const color = _isHardAspect(a.aspect) ? '#ef444499' : '#10b98199';
        svg += `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="${color}" stroke-width="1.3"/>`;
    });
    positions.forEach(p => {
        const [x, y] = toXY(p.lon, planetR);
        const color = PLANET_COLOR[p.name] || '#e5e7eb';
        svg += `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="12" fill="#0b1220" stroke="${color}" stroke-width="1.5"/>`;
        svg += `<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" fill="${color}" font-size="14" text-anchor="middle" dominant-baseline="central">${PLANET_SYMBOL[p.name] || '?'}</text>`;
        if (p.retrograde) {
            svg += `<text x="${(x + 13).toFixed(1)}" y="${(y - 10).toFixed(1)}" fill="#ef4444" font-size="10" text-anchor="middle">℞</text>`;
        }
    });
    svg += `</svg>`;
    container.innerHTML = svg;

    if (legendEl) {
        legendEl.innerHTML = Object.entries(PLANET_SYMBOL)
            .map(([name, sym]) => `<span style="color:${PLANET_COLOR[name]}">${sym} ${name}</span>`)
            .join(' &nbsp;·&nbsp; ');
    }
}

// Công cụ tính mốc chu kỳ, THUẦN CLIENT-SIDE (không cần dữ liệu server) vì mốc neo là lựa chọn
// chủ quan của người phân tích, không có 1 ngày "đúng" duy nhất.
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
