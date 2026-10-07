"""Phân tích lạm phát & tiêu dùng Mỹ theo mô-tuýp cắt lớp (headline → composition → breadth →
persistence → pipeline). Đọc các chuỗi usm_* đã tải từ FRED vào data/vimo_raw.json; không tự nội suy.
"""
import statistics

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

    return {"headline": headline, "groups": groups, "breadth": breadth, "history": hist,
            "pipeline": pipeline, "real_consumption": real, "rates": rates,
            "note": ("Đóng góp (contribution) từng nhóm vào CPI chưa tính: FRED không cung cấp trọng số "
                     "tương đối (relative importance) ổn định theo kỳ, không tự ghép trọng số đoán mò.")}
