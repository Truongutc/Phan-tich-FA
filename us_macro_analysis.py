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

# SUA 2026-10-08 (user: "tách cho đồng nhất với 2b, hoặc 2c đưa về đúng như 2b" — tức MUỐN
# "Năng lượng" ở 2c khớp Y HỆT "Năng lượng" ở bảng nhiệt 2b, ĐỦ xăng dầu) — đảo hướng xử lý: giữ
# "Năng lượng" = CPIENGSL ĐỦ (7,347%, khớp 100% với 2b, KHÔNG override series nữa), và trừ xăng
# dầu ra khỏi "Giao thông" thay vì ra khỏi "Năng lượng" như bản trước. Trọng số Giao thông giảm
# từ 17,060 xuống 13,178 (= 17,060 − Motor fuel 3,882). "Giao thông" ở ĐÂY dùng series SUY RA
# (xem _transport_ex_fuel_yoy) vì KHÔNG có sẵn 1 chỉ số FRED bó gọn "Transportation trừ motor
# fuel" — suy ra bằng trừ trọng số, dùng Gasoline (usm_gasoline, CUSR0000SETB01, 3,770/3,882 =
# 97% trọng số Motor fuel) làm proxy cho Motor fuel (thiếu phần nhỏ "Other motor fuels" 0,112%,
# ảnh hưởng không đáng kể). Tổng trọng số ~94,45% không đổi (phần dư vẫn ~5,55%).
US_GROUP9_WEIGHTS = {
    "usm_cpi_food": 13.540, "usm_cpi_energy": 7.347, "usm_cpi_shelter": 35.343,
    "usm_cpi_transport": 13.178, "usm_cpi_medical": 8.267, "usm_cpi_apparel": 2.406,
    "usm_cpi_recreation": 5.081, "usm_cpi_education_comm": 5.716, "usm_cpi_other": 3.732,
}
_TRANSPORT_FULL_WEIGHT = 17.060
_MOTOR_FUEL_WEIGHT = 3.882


def _transport_ex_fuel_yoy(raw, p):
    """YoY SUY RA của "Giao thông trừ xăng dầu" = (TransportĐủ×YoY − MotorFuel×YoY) / TrọngSốCòn
    lại — xem ghi chú US_GROUP9_WEIGHTS."""
    trn_y = _yoy(_series(raw, "usm_cpi_transport"), p)
    gas_y = _yoy(_series(raw, "usm_gasoline"), p)
    if trn_y is None or gas_y is None:
        return None
    return (_TRANSPORT_FULL_WEIGHT * trn_y - _MOTOR_FUEL_WEIGHT * gas_y) / US_GROUP9_WEIGHTS["usm_cpi_transport"]


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


def _chg_abs(s, latest, months):
    cur, prev = s.get(latest), s.get(_shift(latest, -months))
    return round(cur - prev, 4) if cur is not None and prev is not None else None


def _chg_pct(s, latest, months):
    cur, prev = s.get(latest), s.get(_shift(latest, -months))
    return round((cur / prev - 1) * 100, 2) if cur is not None and prev else None


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
    # SUA 2026-10-08 (user: "vẽ từ T1-2016 tới nay" — chuỗi gốc có từ 2015-01 nhưng YoY cần 12
    # tháng trước nên 2015 toàn None, vẽ ra khoảng trống vô nghĩa đầu mọi biểu đồ theo tháng) —
    # cắt bỏ các tháng đầu chưa có YoY, khớp cách đã sửa cho bảng nhiệt _monthly_heatmap.
    periods = sorted(cpi)
    periods = [p for p in periods if _yoy(cpi, p) is not None]

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
    liquidity = _fed_liquidity(raw)
    growth = _growth(raw)
    labor = _labor(raw)
    treasury_credit = _treasury_credit(raw)
    usd = _usd(raw)
    capital_flows = _capital_flows(raw)

    return {"headline": headline, "groups": groups, "breadth": breadth, "history": hist, "heatmap": heatmap, "cracks": cracks, "contributions": contributions,
            "pipeline": pipeline, "real_consumption": real, "rates": rates, "liquidity": liquidity,
            "growth": growth, "labor": labor, "treasury_credit": treasury_credit, "usd": usd, "capital_flows": capital_flows,
            "note": ("Đóng góp (contribution) từng nhóm vào CPI chưa tính: FRED không cung cấp trọng số "
                     "tương đối (relative importance) ổn định theo kỳ, không tự ghép trọng số đoán mò.")}


def _fed_liquidity(raw):
    """Bảng cân đối Fed (H.4.1, FRED mirror trực tiếp) — user 2026-10-08: "QE/QT, thu hẹp/mở rộng
    bảng cân đối". KHÔNG tự chấm điểm "nới/thắt" — chỉ đưa số liệu thật + % thay đổi 6 tháng/12
    tháng để người đọc tự kết luận (đúng nguyên tắc không gộp nhiều tín hiệu thành 1 điểm số đã
    chốt ở phần vĩ mô VN).
    Net Liquidity (cách giới phân tích thị trường hay dùng, KHÔNG phải định nghĩa chính thức của
    Fed) = Tổng tài sản − RRP − TGA — phần tiền "sẵn sàng" lưu thông trong hệ thống ngân hàng/thị
    trường, sau khi trừ 2 "bể chứa" hút tiền ra khỏi hệ thống (RRP: tiền gửi vào Fed qua repo
    ngược; TGA: tiền Chính phủ gửi tại Fed, chưa chi ra)."""
    keys = ["usm_fed_assets", "usm_fed_treasury", "usm_fed_mbs", "usm_fed_reserves", "usm_fed_rrp", "usm_fed_tga"]
    series = {k: _series(raw, k) for k in keys}
    if not series["usm_fed_assets"]:
        return None
    periods = sorted(series["usm_fed_assets"])
    net_liq = {}
    for p in periods:
        a, r, t = series["usm_fed_assets"].get(p), series["usm_fed_rrp"].get(p), series["usm_fed_tga"].get(p)
        if a is not None and r is not None and t is not None:
            net_liq[p] = a - r - t
    latest = periods[-1]

    def _chg(s, months):
        cur, prev = s.get(latest), s.get(_shift(latest, -months))
        return round(cur - prev, 0) if cur is not None and prev is not None else None

    def _chg_pct(s, months):
        cur, prev = s.get(latest), s.get(_shift(latest, -months))
        return round((cur / prev - 1) * 100, 2) if cur is not None and prev else None

    latest_vals = {k: series[k].get(latest) for k in keys}
    latest_vals["usm_fed_net_liquidity"] = net_liq.get(latest)
    chg = {k: {"chg_6m": _chg(series[k], 6), "chg_12m": _chg(series[k], 12),
               "chg_pct_6m": _chg_pct(series[k], 6), "chg_pct_12m": _chg_pct(series[k], 12)} for k in keys}
    chg["usm_fed_net_liquidity"] = {"chg_6m": _chg(net_liq, 6), "chg_12m": _chg(net_liq, 12),
                                     "chg_pct_6m": _chg_pct(net_liq, 6), "chg_pct_12m": _chg_pct(net_liq, 12)}

    ecb = _series(raw, "usm_ecb_assets")
    ecb_latest = max(ecb) if ecb else None

    # Xu hướng QE/QT — SỰ KIỆN (tăng/giảm liên tục), KHÔNG phải điểm số: so tổng tài sản hiện tại
    # với đỉnh gần nhất (all-time high trong dữ liệu có) để biết đang ở pha thu hẹp (QT) bao lâu.
    peak_period = max(series["usm_fed_assets"], key=lambda p: series["usm_fed_assets"][p])
    assets_vs_peak_pct = round((latest_vals["usm_fed_assets"] / series["usm_fed_assets"][peak_period] - 1) * 100, 2)

    return {
        "periods": periods, "latest": latest, "latest_values": latest_vals, "changes": chg,
        "peak_period": peak_period, "peak_value": series["usm_fed_assets"][peak_period],
        "assets_vs_peak_pct": assets_vs_peak_pct,
        "ecb_latest": {"period": ecb_latest, "value": ecb.get(ecb_latest)} if ecb_latest else None,
        "history": {
            "periods": periods,
            "assets": [series["usm_fed_assets"].get(p) for p in periods],
            "treasury": [series["usm_fed_treasury"].get(p) for p in periods],
            "mbs": [series["usm_fed_mbs"].get(p) for p in periods],
            "reserves": [series["usm_fed_reserves"].get(p) for p in periods],
            "rrp": [series["usm_fed_rrp"].get(p) for p in periods],
            "tga": [series["usm_fed_tga"].get(p) for p in periods],
            "net_liquidity": [net_liq.get(p) for p in periods],
        },
    }


# THEM 2026-10-08 (user: "tiếp tục triển khai theo ma trận đã bàn" — "US Macro Liquidity Matrix"
# user gửi: Growth/Labor/Treasury&Credit/USD. KHÔNG chấm điểm gộp Macro Score/Liquidity Score như
# bản gốc đề xuất — giữ đúng nguyên tắc đã chốt: chỉ đưa số liệu thô độc lập, người đọc tự kết
# luận. "Capital Flows" (TIC — nước ngoài nắm giữ Treasury) KHÔNG có trên FRED, cần nguồn riêng từ
# treasury.gov/tic (giống tình trạng EIA dầu mỏ/OPEC+ trước đây) — CHƯA làm, bỏ qua tạm.
def _growth(raw):
    """GDP thực (quý) + sản xuất công nghiệp (tháng). Không có Leading Index (USSLIND đã bị FRB
    Philadelphia ngừng công bố từ 2020-02, xem ghi chú ở US_MACRO_SERIES)."""
    gdp = _series(raw, "usm_gdp_real")
    indpro = _series(raw, "usm_indpro")
    if not gdp and not indpro:
        return None

    def _qoq_ann(s, period):
        cur, prev = s.get(period), s.get(_shift(period, -3))
        return round(((cur / prev) ** 4 - 1) * 100, 2) if cur is not None and prev else None

    gdp_periods = sorted(gdp) if gdp else []
    ip_periods = sorted(p for p in indpro if _yoy(indpro, p) is not None) if indpro else []
    gdp_latest = gdp_periods[-1] if gdp_periods else None
    ip_latest = ip_periods[-1] if ip_periods else None

    return {
        "gdp": {"latest": gdp_latest, "yoy": _yoy(gdp, gdp_latest), "qoq_annualized": _qoq_ann(gdp, gdp_latest)} if gdp_latest else None,
        "indpro": {"latest": ip_latest, "yoy": _yoy(indpro, ip_latest), "mom_3m_ann": _annualized(indpro, ip_latest, 3)} if ip_latest else None,
        "history": {
            "gdp": {"periods": gdp_periods, "yoy": [_yoy(gdp, p) for p in gdp_periods], "qoq_ann": [_qoq_ann(gdp, p) for p in gdp_periods]},
            "indpro": {"periods": ip_periods, "yoy": [_yoy(indpro, p) for p in ip_periods]},
        },
    }


def _labor(raw):
    """Việc làm phi NN (thay đổi MoM, nghìn người — số được theo dõi nhiều nhất mỗi báo cáo BLS),
    trợ cấp thất nghiệp lần đầu, JOLTS job openings, tỷ lệ tham gia LLLĐ, thu nhập bình quân giờ
    (so YoY với CPI YoY ra "lương thực" — cùng kiểu real-vs-nominal đã làm ở mục tiêu dùng)."""
    payrolls = _series(raw, "usm_payrolls")
    if not payrolls:
        return None
    claims = _series(raw, "usm_claims")
    openings = _series(raw, "usm_job_openings")
    participation = _series(raw, "usm_participation")
    earnings = _series(raw, "usm_avg_earnings")
    unemployment = _series(raw, "usm_unemployment")
    cpi = _series(raw, "usm_cpi")

    periods = sorted(p for p in payrolls if _shift(p, -1) in payrolls)
    mom = {p: round(payrolls[p] - payrolls[_shift(p, -1)], 1) for p in periods}

    def _avg3(p):
        i = periods.index(p)
        return round(sum(mom[x] for x in periods[i - 2:i + 1]) / 3, 1) if i >= 2 else None

    latest = periods[-1]
    openings_latest = max(openings) if openings else None
    participation_latest = max(participation) if participation else None
    earnings_latest = max(earnings) if earnings else None
    unemployment_latest = max(unemployment) if unemployment else None

    def _real_wage(p):
        ey, cy = _yoy(earnings, p), _yoy(cpi, p)
        return round(ey - cy, 2) if ey is not None and cy is not None else None

    return {
        "latest_period": latest, "payrolls_mom": mom[latest], "payrolls_mom_3m_avg": _avg3(latest),
        "claims": {"period": max(claims), "latest": claims[max(claims)]} if claims else None,
        "openings": {"period": openings_latest, "latest": openings.get(openings_latest), "yoy": _yoy(openings, openings_latest)} if openings_latest else None,
        "participation": {"period": participation_latest, "latest": participation.get(participation_latest), "chg_12m_pp": _chg_abs(participation, participation_latest, 12)} if participation_latest else None,
        "earnings_yoy": _yoy(earnings, earnings_latest) if earnings_latest else None,
        "real_wage_yoy": _real_wage(earnings_latest) if earnings_latest else None,
        "unemployment_latest": unemployment.get(unemployment_latest) if unemployment_latest else None,
        "history": {
            "periods": periods,
            "payrolls_mom": [mom[p] for p in periods],
            "payrolls_mom_3m_avg": [_avg3(p) for p in periods],
            "unemployment": [unemployment.get(p) for p in periods] if unemployment else None,
            "claims": [claims.get(p) for p in periods] if claims else None,
            "real_wage_yoy": [_real_wage(p) for p in periods],
        },
    }


def _treasury_credit(raw):
    """Đường cong lợi suất (2Y/10Y), lợi suất thực TIPS 10Y, lạm phát kỳ vọng hòa vốn (breakeven),
    chênh lệch tín dụng High Yield & Investment Grade (OAS) — tất cả từ FRED, daily, lấy TB tháng."""
    y2, y10 = _series(raw, "usm_yield_2y"), _series(raw, "usm_yield_10y")
    spread, real10 = _series(raw, "usm_spread_10y_2y"), _series(raw, "usm_real_yield_10y")
    breakeven, hy, ig = _series(raw, "usm_breakeven_10y"), _series(raw, "usm_hy_oas"), _series(raw, "usm_ig_oas")
    if not y10:
        return None
    series_map = {"usm_yield_2y": y2, "usm_yield_10y": y10, "usm_spread_10y_2y": spread,
                  "usm_real_yield_10y": real10, "usm_breakeven_10y": breakeven, "usm_hy_oas": hy, "usm_ig_oas": ig}
    latest = max(y10)
    periods = sorted(y10)
    latest_values = {k: s.get(latest) for k, s in series_map.items()}
    changes = {k: {"chg_6m": _chg_abs(s, latest, 6), "chg_12m": _chg_abs(s, latest, 12)} for k, s in series_map.items()}
    return {
        "latest": latest, "latest_values": latest_values, "changes": changes,
        "history": {"periods": periods, **{k.replace("usm_", ""): [s.get(p) for p in periods] for k, s in series_map.items()}},
    }


def _capital_flows(raw):
    """Nước ngoài nắm giữ Treasury Mỹ (TIC — Treasury International Capital System, Bộ Tài chính
    Mỹ, ticdata.treasury.gov) — KHÔNG có trên FRED, lấy riêng qua fetch_tic_major_foreign_holders().
    Tách khối "chính thức" (NHTW/chính phủ nước ngoài — phản ánh hành vi dự trữ ngoại hối quốc gia)
    và khối "tư nhân" (= Tổng − Chính thức, quỹ đầu tư/doanh nghiệp/cá nhân, đầu cơ/lợi suất nhiều
    hơn). KHÔNG chấm điểm — chỉ số liệu + % thay đổi."""
    total = _series(raw, "usm_tic_total")
    if not total:
        return None
    official = _series(raw, "usm_tic_official")
    latest = max(total)
    periods = sorted(total)
    private = {p: round(total[p] - official[p], 1) for p in periods if p in official}
    return {
        "latest": latest, "total_latest": total[latest],
        "official_latest": official.get(latest), "private_latest": private.get(latest),
        "total_chg_pct_6m": _chg_pct(total, latest, 6), "total_chg_pct_12m": _chg_pct(total, latest, 12),
        "official_chg_pct_6m": _chg_pct(official, latest, 6) if official else None,
        "official_chg_pct_12m": _chg_pct(official, latest, 12) if official else None,
        "history": {"periods": periods, "total": [total.get(p) for p in periods],
                    "official": [official.get(p) for p in periods], "private": [private.get(p) for p in periods]},
    }


def _usd(raw):
    """Chỉ số USD trọng số thương mại rộng (DTWEXBGS, FRED) — USD mạnh/yếu ảnh hưởng ngược chiều
    tới hàng hóa định giá bằng USD (dầu, vàng...) và dòng vốn vào thị trường mới nổi (VN)."""
    dxy = _series(raw, "usm_dxy_broad")
    if not dxy:
        return None
    latest = max(dxy)
    periods = sorted(dxy)
    return {
        "latest": latest, "latest_value": dxy[latest],
        "chg_pct_6m": _chg_pct(dxy, latest, 6), "chg_pct_12m": _chg_pct(dxy, latest, 12),
        "history": {"periods": periods, "values": [dxy.get(p) for p in periods]},
    }


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


# SUA 2026-10-08 (user: "tách cho đồng nhất với 2b, hoặc 2c đưa về đúng như 2b" — Năng lượng ở
# đây giờ khớp Y HỆT bảng nhiệt 2b nên KHÔNG cần đổi nhãn nữa) — chỉ còn Giao thông cần ghi chú vì
# đã trừ xăng dầu ra (khác CPITRNSL thô dùng ở bảng nhiệt 2b).
CONTRIB_LABEL_OVERRIDE = {"usm_cpi_transport": "Giao thông (trừ xăng dầu — xem Năng lượng)"}


def _cpi_contributions(raw, periods):
    """Đóng góp (contribution, điểm %) vào CPI YoY — ĐỦ 9 NHÓM, "Năng lượng" khớp Y HỆT bảng nhiệt
    2b (CPIENGSL đủ xăng dầu). SỬA lỗi đếm trùng xăng dầu bằng cách trừ xăng dầu khỏi "Giao
    thông" (xem _transport_ex_fuel_yoy) THAY VÌ trừ khỏi "Năng lượng" như bản trước — user
    2026-10-08: "muốn Năng lượng ở 2c khớp 2b". 9 nhóm cộng ~94,45% (residual ~5,55% là "Fuel
    oil" + vài mục nhỏ chưa gán)."""
    cpi = _series(raw, "usm_cpi")
    group_series = {key: _series(raw, key) for key, _ in US_GROUPS}

    def _group_yoy(key, p):
        return _transport_ex_fuel_yoy(raw, p) if key == "usm_cpi_transport" else _yoy(group_series[key], p)

    rows = []
    for key, label in US_GROUPS:
        w = US_GROUP9_WEIGHTS[key]
        rows.append({"key": key, "label": CONTRIB_LABEL_OVERRIDE.get(key, label), "weight_pct": w,
                     "values": [round(w / 100 * _group_yoy(key, p), 3) if _group_yoy(key, p) is not None else None for p in periods]})
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
        yoy = _group_yoy(key, latest)
        snapshot.append({"key": key, "label": CONTRIB_LABEL_OVERRIDE.get(key, label), "weight_pct": w, "yoy": yoy,
                          "contribution": round(w / 100 * yoy, 3) if yoy is not None else None})
    snapshot.sort(key=lambda r: abs(r["contribution"]) if r["contribution"] is not None else -1, reverse=True)

    return {"periods": periods, "rows": rows, "weights_vintage": US_GROUP_WEIGHTS_VINTAGE,
            "weights_source": US_GROUP_WEIGHTS_SOURCE, "snapshot": snapshot, "snapshot_period": latest}
