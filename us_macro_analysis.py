"""Phân tích lạm phát & tiêu dùng Mỹ theo mô-tuýp cắt lớp (headline → composition → breadth →
persistence → pipeline). Đọc các chuỗi usm_* đã tải từ FRED vào data/vimo_raw.json; không tự nội suy.
"""
import statistics


# Trọng số tương đối (relative importance, %) của BLS — Table 2 trong news release CPI, cột
# "Relative importance Jul. 2026" (công bố cùng số liệu tháng 8/2026). KHÔNG lấy được qua
# api.bls.gov (bảng báo cáo, không phải time series) và www.bls.gov/download.bls.gov đều chặn
# fetch tự động (403) — lấy TAY qua WebFetch ngày 2026-10-07, 1 lần, KHÔNG tự cập nhật theo lịch.
# Trọng số thật thay đổi chậm theo năm nên dùng xấp xỉ cho cả lịch sử là chấp nhận được, nhưng
# cần làm lại thủ công (nhờ fetch lại cpi.t02.htm) khi muốn cập nhật vintage mới.
#
# SUA 2026-10-07 (user: "CPI tăng 3% thì food +1 điểm, nhiên liệu +2 điểm, y tế -1 điểm... kiểu
# kiểu vậy" — cần đủ 9 nhóm như bảng nhiệt 2b, không phải rút gọn 4 phần) — quay lại 9 nhóm
# major-group, nhưng SỬA ĐÚNG lỗi đếm trùng đã phát hiện: "Giao thông" (CPITRNSL) chứa Motor fuel
# (xăng dầu), mà Motor fuel GỐC đã nằm trong "Energy" (CPIENGSL). Thay vì bỏ hẳn 9 nhóm (bản
# trước), giữ Transportation NGUYÊN (đủ xăng dầu, khớp CPITRNSL + weight 17,060) và đổi "Năng
# lượng" trong PHẦN ĐÓNG GÓP này sang "Energy services" (CUSR0000SEHF = điện+gas, KHÔNG xăng dầu,
# weight 3,303, xem usm_cpi_energy_services) — xăng dầu chỉ còn tính 1 lần, trong Transportation.
# LƯU Ý: bảng nhiệt 2b vẫn hiện YoY của "Năng lượng" ĐẦY ĐỦ (usm_cpi_energy/CPIENGSL, gồm xăng
# dầu) vì đó chỉ là %, không cộng — không có vấn đề đếm trùng ở biểu đồ đó.
US_GROUP_WEIGHTS_VINTAGE = "2026-07 (BLS CPI news release Table 2, công bố 2026-08)"
US_GROUP_WEIGHTS_SOURCE = "https://www.bls.gov/news.release/cpi.t02.htm"
US_GROUP_WEIGHTS_FOOD_ENERGY_CORE = {
    "food": 13.540, "energy": 7.347, "shelter": 35.343, "core": 79.114,
}

US_GROUPS = [
    ("usm_cpi_food", "Thực phẩm"), ("usm_cpi_energy", "Năng lượng"), ("usm_cpi_shelter", "Nhà ở"),
    ("usm_cpi_transport", "Giao thông"), ("usm_cpi_medical", "Y tế"), ("usm_cpi_apparel", "May mặc"),
    ("usm_cpi_recreation", "Giải trí"), ("usm_cpi_education_comm", "Giáo dục & Truyền thông"),
    ("usm_cpi_other", "Hàng hóa & dịch vụ khác"),
]

# Trọng số (%) + series DÙNG RIÊNG cho phần đóng góp (contribution) — key khớp US_GROUPS, nhưng
# "usm_cpi_energy" ở đây trỏ sang series_key="usm_cpi_energy_services" (KHÔNG xăng dầu) để tránh
# đếm trùng với Transportation (đã có xăng dầu). Tổng trọng số ~94,45% (phần dư ~5,55%, chủ yếu
# là "Fuel oil" 0,162 không nằm trong nhóm nào + các mục nhỏ chưa gán — xem residual khi tính).
US_GROUP9_WEIGHTS = {
    "usm_cpi_food": 13.540, "usm_cpi_energy": 3.303, "usm_cpi_shelter": 35.343,
    "usm_cpi_transport": 17.060, "usm_cpi_medical": 8.267, "usm_cpi_apparel": 2.406,
    "usm_cpi_recreation": 5.081, "usm_cpi_education_comm": 5.716, "usm_cpi_other": 3.732,
}
US_GROUP9_SERIES_OVERRIDE = {"usm_cpi_energy": "usm_cpi_energy_services"}


def _series(raw, key):
    return {p["period"]: p["value"] for p in raw.get(key, {}).get("series", [])}


def _shift(period, months):
    y, m = int(period[:4]), int(period[5:7])
    idx = y * 12 + (m - 1) + months
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def _yoy(s, period):
    cur, prev = s.get(period), s.get(_shift(period, -12))
    return round((cur / prev - 1) * 100, 2) if cur is not None and prev else None


def _mom(s, period):
    cur, prev = s.get(period), s.get(_shift(period, -1))
    return (cur / prev - 1) * 100 if cur is not None and prev else None


def _annualized(s, period, months):
    ratio = 1.0
    for k in range(months):
        m = _mom(s, _shift(period, -k))
        if m is None:
            return None
        ratio *= 1 + m / 100
    return round((ratio ** (12 / months) - 1) * 100, 2)


def _corr(a, b):
    if len(a) < 24:
        return None
    try:
        return round(statistics.correlation(a, b), 3)
    except statistics.StatisticsError:
        return None


def build_us_macro(raw):
    """Trả dict 'usMacro' cho vimo.json, hoặc None nếu chưa có dữ liệu FRED."""
    cpi = _series(raw, "usm_cpi")
    if not cpi:
        return None
    core = _series(raw, "usm_core_cpi")
    goods = _series(raw, "usm_cpi_goods")
    services = _series(raw, "usm_cpi_services")
    ppi = _series(raw, "usm_ppi_final_demand")
    imp = _series(raw, "usm_import_price")
    wti = _series(raw, "usm_oil_wti")
    retail = _series(raw, "usm_retail_sales")
    pce_n = _series(raw, "usm_pce_nominal")
    pce_r = _series(raw, "usm_pce_real")
    latest = max(cpi)
    periods = sorted(cpi)

    headline = {
        "latest": latest,
        "cpi_yoy": _yoy(cpi, latest), "cpi_mom": round(_mom(cpi, latest) or 0, 2),
        "cpi_3m_ann": _annualized(cpi, latest, 3), "cpi_6m_ann": _annualized(cpi, latest, 6),
        "core_yoy": _yoy(core, latest), "core_3m_ann": _annualized(core, latest, 3),
        "core_6m_ann": _annualized(core, latest, 6),
        "goods_yoy": _yoy(goods, latest), "services_yoy": _yoy(services, latest),
    }

    groups = []
    for key, label in US_GROUPS:
        s = _series(raw, key)
        if s:
            groups.append({"key": key, "label": label, "yoy": _yoy(s, latest)})
    vals = [g["yoy"] for g in groups if g["yoy"] is not None]
    breadth = {
        "n_groups": len(vals),
        "pct_gt_3": round(sum(v > 3 for v in vals) / len(vals) * 100, 1) if vals else None,
        "pct_gt_5": round(sum(v > 5 for v in vals) / len(vals) * 100, 1) if vals else None,
        "pct_rising_mom": round(sum((_mom(_series(raw, g["key"]), latest) or 0) > 0 for g in groups) / len(groups) * 100, 1) if groups else None,
    }

    group_series = {key: _series(raw, key) for key, _ in US_GROUPS}
    breadth_hist = []
    for p in periods:
        vs = [_yoy(group_series[k], p) for k, _ in US_GROUPS if group_series[k]]
        vs = [v for v in vs if v is not None]
        breadth_hist.append(round(sum(v > 3 for v in vs) / len(vs) * 100, 1) if vs else None)

    # Headline vs core YoY history (2015+)
    hist = {
        "periods": periods,
        "breadth_gt3_pct": breadth_hist,
        "cpi_yoy": [_yoy(cpi, p) for p in periods],
        "core_yoy": [_yoy(core, p) for p in periods],
        "goods_yoy": [_yoy(goods, p) for p in periods],
        "services_yoy": [_yoy(services, p) for p in periods],
    }

    # Pipeline: PPI MoM (t-L) → CPI goods MoM (t)
    lead = []
    common = [p for p in periods if p in goods and p in ppi]
    for lag in range(0, 7):
        xs, ys = [], []
        for p in common:
            prev = _shift(p, -lag)
            a, b = _mom(ppi, prev), _mom(goods, p)
            if a is not None and b is not None:
                xs.append(a)
                ys.append(b)
        lead.append({"lag_months": lag, "corr": _corr(xs, ys), "n": len(xs)})

    # Lạm phát nhập khẩu: MoM giá nhập khẩu (t-L) → MoM CPI hàng hóa (t). Dùng MoM (không YoY) để
    # tránh tương quan giả do xu hướng chung kéo dài.
    imp_lead = []
    common_i = [p for p in periods if p in goods and p in imp]
    for lag in range(0, 7):
        xs, ys = [], []
        for p in common_i:
            a, b = _mom(imp, _shift(p, -lag)), _mom(goods, p)
            if a is not None and b is not None:
                xs.append(a)
                ys.append(b)
        imp_lead.append({"lag_months": lag, "corr": _corr(xs, ys), "n": len(xs)})

    pipeline = {
        "ppi_final_yoy": _yoy(ppi, latest), "import_price_yoy": _yoy(imp, latest),
        "oil_yoy": _yoy(wti, latest),
        "ppi_to_cpi_goods_corr": lead, "import_to_cpi_goods_corr": imp_lead,
        "ppi_hist": {"periods": periods, "ppi_yoy": [_yoy(ppi, p) for p in periods],
                     "import_yoy": [_yoy(imp, p) for p in periods],
                     "oil_yoy": [_yoy(wti, p) for p in periods]},
    }

    # Real consumption: nominal retail / PCE vs real PCE (CPI-deflated)
    rc_periods = sorted(p for p in retail if p >= "2015-01")
    real = {
        "latest": latest,
        "retail_nominal_yoy": _yoy(retail, latest),
        "cpi_yoy": _yoy(cpi, latest),
        "pce_nominal_yoy": _yoy(pce_n, latest),
        "pce_real_yoy": _yoy(pce_r, latest),
        "hist": {"periods": rc_periods,
                 "retail_nominal_yoy": [_yoy(retail, p) for p in rc_periods],
                 "cpi_yoy": [_yoy(cpi, p) for p in rc_periods]},
    }

    rates = {}
    for key, label in [("usm_fed_funds", "Lãi suất quỹ liên bang (%)"),
                       ("usm_yield_10y", "Lợi suất 10 năm (%)"),
                       ("usm_unemployment", "Thất nghiệp (%)"),
                       ("usm_spread_10y_2y", "10Y-2Y (%)")]:
        s = _series(raw, key)
        if s:
            last = max(s)
            rates[key] = {"label": label, "latest": s[last], "period": last}

    heatmap = _monthly_heatmap(raw, latest)
    cracks = _crack_spreads(raw)
    contributions = _cpi_contributions(raw, periods)

    return {"headline": headline, "groups": groups, "breadth": breadth, "history": hist, "heatmap": heatmap, "cracks": cracks, "contributions": contributions,
            "pipeline": pipeline, "real_consumption": real, "rates": rates,
            "note": ("Đóng góp (contribution) từng nhóm vào CPI chưa tính: FRED không cung cấp trọng số "
                     "tương đối (relative importance) ổn định theo kỳ, không tự ghép trọng số đoán mò.")}


def _monthly_heatmap(raw, latest):
    """YoY theo THÁNG của từng nhóm CPI (+ CPI tổng để đối chiếu) — SUA 2026-10-07 (user: "Mỹ công
    bố CPI theo tháng, chỉnh lại được không") — bỏ gộp quý (bản cũ _quarterly_heatmap), dùng đúng
    tần suất gốc BLS công bố. KHÔNG cần ghi chú "kỳ cuối thiếu dữ liệu" như bản quý vì mỗi tháng
    đã là 1 kỳ công bố đầy đủ."""
    rows = [("CPI tổng", "usm_cpi")] + [(label, key) for key, label in US_GROUPS]
    series = {key: _series(raw, key) for _, key in rows}
    months = sorted({m for s in series.values() for m in s if m <= latest})
    out_rows = []
    for label, key in rows:
        s = series[key]
        vals = [round(_yoy(s, m), 2) if _yoy(s, m) is not None else None for m in months]
        out_rows.append({"label": label, "values": vals})
    # SUA 2026-10-07 (user: "xóa bớt đi, tôi cần hiện từ T1-2016 thôi, đằng trước có dữ liệu đâu")
    # — 2015 chưa có YoY (chuỗi gốc chỉ có từ 2015-01, cần 12 tháng trước nên YoY sớm nhất là
    # 2016-01) — cắt bỏ các tháng đầu mà CPI tổng (và do đó hầu hết nhóm) toàn None.
    cpi_row = out_rows[0]["values"]
    start_idx = next((i for i, v in enumerate(cpi_row) if v is not None), 0)
    months = months[start_idx:]
    for r in out_rows:
        r["values"] = r["values"][start_idx:]
    return {"months": months, "rows": out_rows}


def _crack_spreads(raw):
    """Crack spread theo tháng ($/thùng) từ EIA qua FRED: diesel ULSD & xăng Vịnh Mexico (USD/gallon × 42)
    trừ WTI. Crack 3-2-1 = (2 × xăng + 1 × diesel)/3 − WTI — chuẩn lọc dầu Mỹ (3 thùng dầu → 2 xăng + 1 diesel)."""
    diesel = _series(raw, "usm_diesel_gulf")
    gas = _series(raw, "usm_gasoline_gulf")
    wti = _series(raw, "usm_oil_wti")
    periods = sorted(p for p in diesel if p in gas and p in wti)
    d_crack, g_crack, c321 = [], [], []
    for p in periods:
        dsl, gsl, w = diesel[p] * 42, gas[p] * 42, wti[p]
        d_crack.append(round(dsl - w, 2))
        g_crack.append(round(gsl - w, 2))
        c321.append(round((2 * gsl + dsl) / 3 - w, 2))
    return {"periods": periods, "diesel_crack": d_crack, "gasoline_crack": g_crack, "crack_321": c321,
            "latest": {"period": periods[-1] if periods else None,
                       "diesel_crack": d_crack[-1] if periods else None,
                       "crack_321": c321[-1] if periods else None}}


# Nhãn RIÊNG cho phần đóng góp (khác nhãn dùng ở bảng nhiệt 2b) — vì "Năng lượng" ở ĐÂY là
# điện+gas (ex xăng dầu, tránh đếm trùng với "Giao thông"), nếu dùng chung nhãn "Năng lượng" với
# bảng nhiệt (ở đó là NĂNG LƯỢNG ĐỦ, gồm xăng dầu) thì 2 số % khác hẳn nhau dưới CÙNG 1 tên —
# đúng điều user phản ánh "nhìn chả hiểu gì" (Năng lượng heatmap +16% nhưng bảng đóng góp +4%).
# SUA 2026-10-07 (user vẫn hỏi lại "năng lượng tăng có 3% thôi á, có sai không" — nhãn dài bị cắt
# trên biểu đồ cột, không thấy được chú thích) — RÚT NGẮN nhãn, đưa lời giải thích đầy đủ vào
# ind-source-note (văn bản dưới chart, không bị cắt) thay vì nhét hết vào label trục X.
CONTRIB_LABEL_OVERRIDE = {
    "usm_cpi_energy": "Năng lượng (chỉ điện/gas)",
    "usm_cpi_transport": "Giao thông (gồm xăng dầu)",
}


def _cpi_contributions(raw, periods):
    """Đóng góp (contribution, điểm %) vào CPI YoY — ĐỦ 9 NHÓM giống bảng nhiệt 2b (user 2026-10-07:
    "CPI tăng 3% thì food +1 điểm, nhiên liệu +2 điểm, y tế -1 điểm... kiểu kiểu vậy"). SỬA lỗi
    đếm trùng xăng dầu: "Năng lượng" ở đây dùng series usm_cpi_energy_services (CUSR0000SEHF —
    điện+gas, KHÔNG xăng dầu) + trọng số 3,303, vì xăng dầu đã tính trong "Giao thông" (CPITRNSL,
    ĐỦ xăng dầu, trọng số 17,060) — xem US_GROUP9_WEIGHTS/US_GROUP9_SERIES_OVERRIDE và
    CONTRIB_LABEL_OVERRIDE (đổi nhãn hiển thị để không nhầm với "Năng lượng" ĐỦ ở bảng nhiệt).
    9 nhóm cộng ~94,45% (residual ~5,55% là "Fuel oil" + vài mục nhỏ chưa gán)."""
    cpi = _series(raw, "usm_cpi")
    group_series = {key: _series(raw, US_GROUP9_SERIES_OVERRIDE.get(key, key)) for key, _ in US_GROUPS}

    rows = []
    for key, label in US_GROUPS:
        w = US_GROUP9_WEIGHTS[key]
        s = group_series[key]
        rows.append({"key": key, "label": CONTRIB_LABEL_OVERRIDE.get(key, label), "weight_pct": w,
                     "values": [round(w / 100 * _yoy(s, p), 3) if _yoy(s, p) is not None else None for p in periods]})
    residual = []
    for i, p in enumerate(periods):
        hy = _yoy(cpi, p)
        if hy is None:
            residual.append(None)
            continue
        known = sum(r["values"][i] for r in rows if r["values"][i] is not None)
        residual.append(round(hy - known, 3))
    rows.append({"key": "residual", "label": "Phần dư (Fuel oil + mục nhỏ chưa gán)", "weight_pct": None, "values": residual})

    # Snapshot kỳ gần nhất — CHỈ Weight/YoY/Contribution (user 2026-10-07: "nhìn chả hiểu gì" với
    # bảng cũ có thêm 3M/6M/Pressure — bỏ bớt, giữ đúng 3 cái user cần: "food làm CPI tăng bao
    # nhiêu điểm, bản thân food tăng bao nhiêu %"). Sắp theo |đóng góp| giảm dần để thấy ngay cái
    # gì đang "gánh" CPI nhiều nhất lên đầu, khớp ví dụ user: "học phí tăng rất cao nhưng đóng góp
    # ít vì tiêu dùng ít, năng lượng tăng nhẹ nhưng đóng góp nhiều vì tiêu dùng nhiều".
    latest = periods[-1]
    snapshot = []
    for key, label in US_GROUPS:
        w = US_GROUP9_WEIGHTS[key]
        s = group_series[key]
        yoy = _yoy(s, latest)
        snapshot.append({"key": key, "label": CONTRIB_LABEL_OVERRIDE.get(key, label), "weight_pct": w, "yoy": yoy,
                          "contribution": round(w / 100 * yoy, 3) if yoy is not None else None})
    snapshot.sort(key=lambda r: abs(r["contribution"]) if r["contribution"] is not None else -1, reverse=True)

    return {"periods": periods, "rows": rows, "weights_vintage": US_GROUP_WEIGHTS_VINTAGE,
            "weights_source": US_GROUP_WEIGHTS_SOURCE, "snapshot": snapshot, "snapshot_period": latest}
