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
# CHỈ dùng đúng 3 cấp BLS tự công bố CỘNG CHÍNH XÁC 100% (Food + Energy + "All items less food
# and energy"/Core = 100.001, không chồng chéo) — KHÔNG tự ghép lại 9 nhóm major-group kiểu cũ vì
# phát hiện bug: "Giao thông" (CPITRNSL) chứa Motor fuel, mà Motor fuel ĐÃ nằm trong "Energy" —
# cộng cả 2 vào contribution sẽ đếm trùng giá xăng, làm residual bị thổi phồng rất lớn (thử tính
# ra residual ~-0,9 điểm %, quá to để là "sai số xấp xỉ"). Tách tiếp Core thành Shelter (35,343,
# có series riêng usm_cpi_shelter) + "Core trừ Shelter" (suy ra bằng trừ trọng số, xem
# _cpi_contributions) — vẫn cộng ĐÚNG 100%, không chồng chéo.
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

    heatmap = _quarterly_heatmap(raw, latest)
    cracks = _crack_spreads(raw)
    contributions = _cpi_contributions(raw, periods)

    return {"headline": headline, "groups": groups, "breadth": breadth, "history": hist, "heatmap": heatmap, "cracks": cracks, "contributions": contributions,
            "pipeline": pipeline, "real_consumption": real, "rates": rates,
            "note": ("Đóng góp (contribution) từng nhóm vào CPI chưa tính: FRED không cung cấp trọng số "
                     "tương đối (relative importance) ổn định theo kỳ, không tự ghép trọng số đoán mò.")}


def _quarterly_heatmap(raw, latest):
    """YoY trung bình theo quý của từng nhóm CPI (+ CPI tổng để đối chiếu). Quý cuối có thể chưa đủ
    3 tháng — ghi rõ số tháng thực có để không đọc nhầm."""
    rows = [("CPI tổng", "usm_cpi")] + [(label, key) for key, label in US_GROUPS]
    series = {key: _series(raw, key) for _, key in rows}
    months = sorted({m for s in series.values() for m in s})
    quarters = sorted({f"{m[:4]}-Q{(int(m[5:7]) - 1) // 3 + 1}" for m in months})
    out_rows = []
    for label, key in rows:
        s = series[key]
        vals = []
        for q in quarters:
            y, qq = int(q[:4]), int(q[-1])
            ms = [f"{y}-{mm:02d}" for mm in range(3 * qq - 2, 3 * qq + 1)]
            yoys = [_yoy(s, m) for m in ms if m <= latest]
            yoys = [v for v in yoys if v is not None]
            vals.append(round(sum(yoys) / len(yoys), 2) if yoys else None)
        out_rows.append({"label": label, "values": vals})
    last_q = quarters[-1]
    last_months = sum(1 for m in months if f"{m[:4]}-Q{(int(m[5:7]) - 1) // 3 + 1}" == last_q)
    return {"quarters": quarters, "rows": out_rows, "last_quarter_months": last_months}


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


def _cpi_contributions(raw, periods):
    """Đóng góp (contribution, điểm %) vào CPI YoY — CÁCH SẠCH, không chồng chéo: BLS tự công bố
    CPI = Food (trọng số 13,540) + Energy (7,347) + "All items less food and energy"/Core (79,114)
    — 3 phần CỘNG ĐÚNG 100% theo đúng cấu trúc chính thức (US_GROUP_WEIGHTS_FOOD_ENERGY_CORE).
    Tách tiếp Core thành Shelter (35,343) + "Core trừ nhà ở" (43,771, SUY RA bằng phép trừ trọng
    số — KHÔNG có series riêng, vì "Core trừ Shelter" không phải 1 chỉ số BLS công bố thẳng) để
    thấy nhà ở tách khỏi phần lõi còn lại, như 2 bài phân tích tham khảo. "residual" = headline
    YoY thật trừ tổng 4 phần — CHỈ còn sai số xấp xỉ do compounding (thường < 0,2 điểm %), không
    còn lỗi chồng trọng số như bản cũ (Motor fuel vừa trong Energy vừa trong Transportation)."""
    cpi, food, energy, shelter, core = (_series(raw, k) for k in
        ("usm_cpi", "usm_cpi_food", "usm_cpi_energy", "usm_cpi_shelter", "usm_core_cpi"))
    wF = US_GROUP_WEIGHTS_FOOD_ENERGY_CORE["food"]
    wE = US_GROUP_WEIGHTS_FOOD_ENERGY_CORE["energy"]
    wS = US_GROUP_WEIGHTS_FOOD_ENERGY_CORE["shelter"]
    wCore = US_GROUP_WEIGHTS_FOOD_ENERGY_CORE["core"]

    food_c, energy_c, shelter_c, core_ex_c, residual = [], [], [], [], []
    for p in periods:
        fy, ey, sy, cy, hy = (_yoy(food, p), _yoy(energy, p), _yoy(shelter, p), _yoy(core, p), _yoy(cpi, p))
        fc = round(wF / 100 * fy, 3) if fy is not None else None
        ec = round(wE / 100 * ey, 3) if ey is not None else None
        sc = round(wS / 100 * sy, 3) if sy is not None else None
        cec = round((wCore * cy - wS * sy) / 100, 3) if (cy is not None and sy is not None) else None
        food_c.append(fc); energy_c.append(ec); shelter_c.append(sc); core_ex_c.append(cec)
        known = sum(v for v in (fc, ec, sc, cec) if v is not None)
        residual.append(round(hy - known, 3) if hy is not None else None)

    rows = [
        {"key": "usm_cpi_food", "label": "Thực phẩm", "weight_pct": wF, "values": food_c},
        {"key": "usm_cpi_energy", "label": "Năng lượng", "weight_pct": wE, "values": energy_c},
        {"key": "usm_cpi_shelter", "label": "Nhà ở (Shelter)", "weight_pct": wS, "values": shelter_c},
        {"key": "core_ex_shelter", "label": "Lõi, trừ nhà ở (suy ra)", "weight_pct": round(wCore - wS, 2), "values": core_ex_c},
        {"key": "residual", "label": "Phần dư (sai số xấp xỉ, do compounding)", "weight_pct": None, "values": residual},
    ]
    return {"periods": periods, "rows": rows, "weights_vintage": US_GROUP_WEIGHTS_VINTAGE,
            "weights_source": US_GROUP_WEIGHTS_SOURCE}
