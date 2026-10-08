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

    # THEM 2026-10-08 (user gửi tài liệu "Inflation Transmission" — "CPI tăng từ đâu và shock đó
    # có đang truyền sang các nhóm khác không?") — thêm 2 tầng truyền dẫn SỚM/MUỘN hơn tầng PPI→CPI
    # hàng hóa đã có: (1) Dầu → PPI (tầng ĐẦU chuỗi, trước khi vào PPI), (2) Lương → CPI dịch vụ
    # (tầng dịch vụ — "wage-price spiral" nếu có). Dùng MoM + lag correlation giống cách đã làm,
    # KHÔNG chấm điểm — chỉ cho thấy độ trễ/độ mạnh tương quan ở TỪNG tầng để người đọc tự đánh giá
    # cú sốc đang "cô lập" hay "lan truyền".
    earnings = _series(raw, "usm_avg_earnings")
    oil_lead = []
    common_o = [p for p in periods if p in ppi and p in wti]
    for lag in range(0, 7):
        xs, ys = [], []
        for p in common_o:
            a, b = _mom(wti, _shift(p, -lag)), _mom(ppi, p)
            if a is not None and b is not None:
                xs.append(a)
                ys.append(b)
        oil_lead.append({"lag_months": lag, "corr": _corr(xs, ys), "n": len(xs)})

    wage_lead = []
    common_w = [p for p in periods if p in services and p in earnings] if earnings else []
    for lag in range(0, 7):
        xs, ys = [], []
        for p in common_w:
            a, b = _mom(earnings, _shift(p, -lag)), _mom(services, p)
            if a is not None and b is not None:
                xs.append(a)
                ys.append(b)
        wage_lead.append({"lag_months": lag, "corr": _corr(xs, ys), "n": len(xs)})

    pipeline = {
        "ppi_final_yoy": _yoy(ppi, latest), "import_price_yoy": _yoy(imp, latest),
        "oil_yoy": _yoy(wti, latest),
        "ppi_to_cpi_goods_corr": lead, "import_to_cpi_goods_corr": imp_lead,
        "oil_to_ppi_corr": oil_lead, "wage_to_services_corr": wage_lead,
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
            rp = sorted(p for p in s if p >= "2021-01")
            rates[key] = {"label": label, "latest": s[last], "period": last,
                          "history": {"periods": rp, "values": [s.get(p) for p in rp]}}

    heatmap = _monthly_heatmap(raw, latest)
    cracks = _crack_spreads(raw)
    contributions = _cpi_contributions(raw, periods)
    liquidity = _fed_liquidity(raw)
    growth = _growth(raw)
    labor = _labor(raw)
    treasury_credit = _treasury_credit(raw)
    usd = _usd(raw)
    capital_flows = _capital_flows(raw)
    states = _assess_states(headline, rates, liquidity, growth, labor, treasury_credit, usd, capital_flows, breadth, contributions)
    synthesis = _build_synthesis(states)

    return {"headline": headline, "groups": groups, "breadth": breadth, "history": hist, "heatmap": heatmap, "cracks": cracks, "contributions": contributions,
            "pipeline": pipeline, "real_consumption": real, "rates": rates, "liquidity": liquidity,
            "growth": growth, "labor": labor, "treasury_credit": treasury_credit, "usd": usd, "capital_flows": capital_flows, "states": states, "synthesis": synthesis,
            "note": ("Đóng góp (contribution) từng nhóm vào CPI chưa tính: FRED không cung cấp trọng số "
                     "tương đối (relative importance) ổn định theo kỳ, không tự ghép trọng số đoán mò.")}


# THEM 2026-10-08 (user: "không nên làm chấm điểm gộp, vì cùng 1 mức điểm nhưng nói lên nhiều
# trạng thái, thay vì đó thì bạn có thể làm đánh giá từng chỉ tiêu để đánh giá trạng thái") — mỗi
# NHÓM trong ma trận có 1 đánh giá trạng thái RULE-BASED riêng (ngưỡng số thực tế, không AI, giống
# cách build_synthesis_vimo() đã làm cho vĩ mô VN) — label + màu (good/neutral/warn/bad) + 1 câu
# giải thích TỪ SỐ LIỆU THẬT. KHÔNG cộng dồn các label này thành điểm tổng — mỗi nhóm đứng độc lập,
# người đọc tự tổng hợp bức tranh chung từ nhiều trạng thái khác nhau.
def _assess_states(headline, rates, liquidity, growth, labor, treasury_credit, usd, capital_flows, breadth, contributions):
    states = {}

    if headline.get("cpi_3m_ann") is not None:
        m = headline["cpi_3m_ann"]
        if m < 2:
            label, color = "Đang về mục tiêu (~2%)", "good"
        elif m < 3.5:
            label, color = "Hạ nhiệt nhưng còn trên mục tiêu", "neutral"
        elif m < 5:
            label, color = "Dai dẳng, chưa hạ nhiệt rõ", "warn"
        else:
            label, color = "Nóng trở lại", "bad"
        top2 = (contributions or {}).get("snapshot", [])[:2]
        source_note = ""
        if top2:
            source_note = "; nhóm đóng góp nhiều nhất: " + ", ".join(
                f"{r['label']} ({r['contribution']:+.2f}pp)" for r in top2 if r.get("contribution") is not None)
        states["inflation"] = {"label": label, "color": color,
            "detail": f"CPI YoY {headline.get('cpi_yoy')}%, lõi {headline.get('core_yoy')}%, momentum 3 tháng năm hóa {m}% (so mục tiêu Fed ~2%){source_note}."}

    # THEM 2026-10-08 (user gửi tài liệu "Inflation Persistence" — "CPI tăng từ đâu và có đang lan
    # truyền sang nhóm khác không, hay chỉ là 1 cú sốc đơn lẻ (energy) Fed có thể 'nhìn xuyên
    # qua'?") — tách RIÊNG khỏi "mức độ lạm phát" ở trên. Rule dùng ĐÚNG số liệu đã có, không thêm
    # nguồn mới: độ lan tỏa (breadth, % trong 9 nhóm có YoY>3%) + CPI lõi có cao hay không (lõi cao
    # = áp lực không chỉ tới từ năng lượng/thực phẩm dễ biến động).
    pct_gt_3 = (breadth or {}).get("pct_gt_3")
    core_yoy = headline.get("core_yoy")
    if pct_gt_3 is not None and core_yoy is not None:
        if pct_gt_3 > 60 and core_yoy > 3:
            label, color = "Cao — lan rộng, không chỉ 1 cú sốc đơn lẻ", "bad"
        elif pct_gt_3 > 40 or core_yoy > 2.5:
            label, color = "Trung bình — có dấu hiệu lan truyền một phần", "warn"
        else:
            label, color = "Thấp — tương đối cô lập, Fed có thể 'nhìn xuyên qua'", "good"
        states["inflation_persistence"] = {"label": label, "color": color,
            "detail": f"{pct_gt_3}% trong 9 nhóm CPI có YoY &gt; 3% (độ lan tỏa); CPI lõi {core_yoy}% so CPI toàn phần {headline.get('cpi_yoy')}% (lõi cao = không chỉ do năng lượng/thực phẩm)."}

    if growth and growth.get("gdp"):
        g = growth["gdp"]["qoq_annualized"]
        if g is None:
            pass
        elif g < 0:
            label, color = "Suy giảm", "bad"
        elif g < 1.5:
            label, color = "Tăng trưởng yếu", "warn"
        elif g < 3:
            label, color = "Tăng trưởng vừa phải", "neutral"
        else:
            label, color = "Tăng trưởng mạnh", "good"
        if g is not None:
            ip_note = f", sản xuất công nghiệp YoY {growth['indpro']['yoy']}%" if growth.get("indpro") else ""
            states["growth"] = {"label": label, "color": color,
                "detail": f"GDP thực quý {growth['gdp']['latest']} tăng {g}%/năm (QoQ năm hóa), YoY {growth['gdp']['yoy']}%{ip_note}."}

    if labor:
        avg3 = labor.get("payrolls_mom_3m_avg")
        if avg3 is not None:
            if avg3 < 0:
                label, color = "Suy yếu — mất việc làm ròng", "bad"
            elif avg3 < 100:
                label, color = "Hạ nhiệt rõ rệt", "warn"
            elif avg3 < 200:
                label, color = "Tăng trưởng vừa phải", "neutral"
            else:
                label, color = "Tăng trưởng mạnh", "good"
            rw = labor.get("real_wage_yoy")
            rw_note = f", lương thực YoY {rw}%" if rw is not None else ""
            states["labor"] = {"label": label, "color": color,
                "detail": f"Việc làm phi NN thêm TB 3 tháng {avg3}k/tháng (tháng gần nhất {labor.get('payrolls_mom')}k), thất nghiệp {labor.get('unemployment_latest')}%{rw_note}."}

    # THEM 2026-10-08 (user gửi tài liệu "Fed reaction function" — "không nên viết CPI tăng → Fed
    # hawkish [trực tiếp]; phải nhìn Inflation → Persistence → Labor → Demand/GDP → Fed reaction")
    # — ghép chuỗi NHÂN QUẢ từ các trạng thái ĐÃ TÍNH Ở TRÊN (inflation_persistence/labor/growth)
    # vào câu giải thích TRƯỚC KHI nêu sự kiện/số liệu thực (lãi suất thực) — không tạo điểm số Fed
    # Hawkish Pressure nào, chỉ liệt kê nhãn qua dấu "+".
    fed_funds = rates.get("usm_fed_funds", {}).get("latest")
    cpi_yoy = headline.get("cpi_yoy")
    if fed_funds is not None and cpi_yoy is not None:
        real_rate = round(fed_funds - cpi_yoy, 2)
        if real_rate > 1.5:
            policy_label, policy_color = "Thắt chặt rõ rệt", "warn"
        elif real_rate > 0:
            policy_label, policy_color = "Thắt chặt nhẹ", "neutral"
        else:
            policy_label, policy_color = "Nới lỏng thực (lãi suất thực âm)", "good"
        pressure_bits = []
        if states.get("inflation_persistence"):
            pressure_bits.append(f"độ dai dẳng lạm phát '{states['inflation_persistence']['label']}'")
        if states.get("labor"):
            pressure_bits.append(f"lao động '{states['labor']['label']}'")
        if states.get("growth"):
            pressure_bits.append(f"tăng trưởng '{states['growth']['label']}'")
        pressure_text = (" + ".join(pressure_bits) + " → ") if pressure_bits else ""
        states["fed_policy"] = {"label": policy_label, "color": policy_color,
            "detail": f"{pressure_text}lãi suất thực hiện tại {real_rate}% (Fed funds {fed_funds}% − CPI YoY {cpi_yoy}%)."}

    # Tách RIÊNG "bảng cân đối" khỏi "lãi suất" (user: "không thể chỉ có Fed Funds Rate... phải
    # phân biệt QT thật với reserve management") — dùng Liquidity Impulse (ΔReserves−ΔRRP−ΔTGA,
    # xem _fed_liquidity) thay vì chỉ % thay đổi tổng tài sản, vì tổng tài sản có thể đi ngang do
    # cơ cấu kỳ hạn trong khi thanh khoản ròng NGÂN HÀNG vẫn đang thay đổi.
    if liquidity:
        chg6 = liquidity["changes"]["usm_fed_assets"]["chg_pct_6m"]
        impulse = liquidity.get("liquidity_impulse_latest")
        if chg6 is not None:
            if chg6 > 0.5:
                bs_label, bs_color = "Mở rộng trở lại (QE)", "good"
            elif chg6 < -0.5:
                bs_label, bs_color = "Thu hẹp (QT)", "warn"
            else:
                bs_label, bs_color = "Đi ngang", "neutral"
            impulse_note = ""
            if impulse is not None:
                impulse_bn = round(impulse / 1000, 1)
                impulse_note = (f"; Liquidity Impulse (ΔReserves−ΔRRP−ΔTGA) tháng gần nhất {'+' if impulse_bn >= 0 else ''}{impulse_bn} tỷ$ "
                                 + ("(bơm ròng vào hệ thống ngân hàng)" if impulse_bn >= 0 else "(rút ròng khỏi hệ thống ngân hàng)"))
            states["fed_balance_sheet"] = {"label": bs_label, "color": bs_color,
                "detail": f"Tổng tài sản Fed thay đổi {chg6}%/6 tháng{impulse_note}."}

    if treasury_credit:
        v, c = treasury_credit["latest_values"], treasury_credit["changes"]
        spread = v.get("usm_spread_10y_2y")
        curve_label = "Đường cong ĐẢO NGƯỢC — tín hiệu cảnh báo suy thoái kinh điển" if (spread is not None and spread < 0) else "Đường cong bình thường (dốc lên)"
        curve_color = "bad" if (spread is not None and spread < 0) else "good"
        hy_chg6 = c.get("usm_hy_oas", {}).get("chg_6m")
        credit_note = ""
        if hy_chg6 is not None:
            if hy_chg6 > 0.3:
                credit_note = "; chênh lệch tín dụng High Yield NỚI RỘNG rõ — khẩu vị rủi ro đang giảm"
            elif hy_chg6 < -0.3:
                credit_note = "; chênh lệch tín dụng High Yield THU HẸP — khẩu vị rủi ro đang cao (cẩn trọng nếu quá chủ quan)"
        states["treasury_credit"] = {"label": curve_label, "color": curve_color,
            "detail": f"10Y-2Y = {spread}%, HY OAS {v.get('usm_hy_oas')}%{credit_note}."}

    if usd:
        c12 = usd.get("chg_pct_12m")
        if c12 is not None:
            if c12 > 3:
                label, color = "USD mạnh lên rõ rệt", "neutral"
            elif c12 < -3:
                label, color = "USD yếu đi rõ rệt", "neutral"
            else:
                label, color = "USD dao động trong biên hẹp", "neutral"
            states["usd"] = {"label": label, "color": color,
                "detail": f"Chỉ số USD Broad {usd.get('latest_value')} ({usd.get('latest')}), thay đổi {c12}%/12 tháng — ảnh hưởng ngược chiều giá hàng hóa USD và dòng vốn vào thị trường mới nổi."}

    if capital_flows:
        oc12 = capital_flows.get("official_chg_pct_12m")
        if oc12 is not None:
            if oc12 < -3:
                label, color = "Khối chính thức/NHTW đang RÚT rõ rệt khỏi Treasury Mỹ", "warn"
            elif oc12 > 3:
                label, color = "Khối chính thức/NHTW đang TĂNG mua Treasury Mỹ", "good"
            else:
                label, color = "Khối chính thức/NHTW ổn định", "neutral"
            states["capital_flows"] = {"label": label, "color": color,
                "detail": f"Tổng nước ngoài nắm giữ thay đổi {capital_flows.get('total_chg_pct_12m')}%/12 tháng; khối chính thức {oc12}%/12 tháng."}

    return states


# THEM 2026-10-08 (user gửi tài liệu "3 câu hỏi độc lập": "A. Nền kinh tế đang khỏe hay yếu? B.
# Chính sách tiền tệ đang nới hay thắt? C. Môi trường có thuận lợi cho tài sản rủi ro không? Ba cái
# này có thể cho ba kết quả hoàn toàn khác nhau" — vd Economy=GOOD, Monetary=BAD, Investment=BAD
# khi GDP khỏe nhưng Fed phải tăng lãi suất vì lạm phát) — gộp các "states" ĐỘC LẬP đã có thành 3
# ĐOẠN VĂN BẢN ngắn riêng biệt theo 3 câu hỏi trên. Đây CHỈ LÀ GHÉP CÂU rule-based (if/else trên
# color đã có, giống build_synthesis_vimo() cho vĩ mô VN) — KHÔNG tính thêm bất kỳ con số/điểm nào,
# và KHÔNG gộp 3 đoạn này lại thành 1 kết luận chung — để riêng vì 3 câu hỏi có thể trả lời khác
# nhau (đúng tinh thần "Economy tốt nhưng Monetary xấu" mà user nêu).
def _build_synthesis(states):
    def _combine(parts, colors, good_text, mixed_text, bad_text):
        if not colors:
            return ""
        if all(c in ("good", "neutral") for c in colors):
            verdict = good_text
        elif any(c == "bad" for c in colors):
            verdict = bad_text
        else:
            verdict = mixed_text
        return (", ".join(parts) + ". " + verdict).strip()

    g, l = states.get("growth"), states.get("labor")
    economic = _combine(
        [f"Tăng trưởng: {g['label'].lower()}" for g in [g] if g] + [f"lao động: {l['label'].lower()}" for l in [l] if l],
        [s["color"] for s in (g, l) if s],
        "Nhìn chung nền kinh tế vẫn trụ vững, chưa có dấu hiệu suy thoái rõ rệt.",
        "Nền kinh tế đang hạ nhiệt nhưng chưa tới mức báo động.",
        "Nền kinh tế đang có dấu hiệu suy yếu rõ rệt ở ít nhất 1 trụ cột (tăng trưởng hoặc lao động).")

    ip, fp, fb = states.get("inflation_persistence"), states.get("fed_policy"), states.get("fed_balance_sheet")
    monetary = _combine(
        [f"Độ dai dẳng lạm phát: {s['label'].lower()}" for s in [ip] if s]
        + [f"chính sách lãi suất: {s['label'].lower()}" for s in [fp] if s]
        + [f"bảng cân đối: {s['label'].lower()}" for s in [fb] if s],
        [s["color"] for s in (ip, fp, fb) if s],
        "Điều kiện tiền tệ nhìn chung đang khá thuận lợi/nới lỏng.",
        "Điều kiện tiền tệ ở trạng thái trung tính, chưa nghiêng rõ về nới hay thắt.",
        "Điều kiện tiền tệ nhìn chung đang THẮT CHẶT hơn là nới lỏng — không nên mặc định Fed sắp dovish chỉ vì lạm phát hạ nhiệt một phần.")

    tc, u, cf = states.get("treasury_credit"), states.get("usd"), states.get("capital_flows")
    investment = _combine(
        [f"Lợi suất/tín dụng: {s['label'].lower()}" for s in [tc] if s]
        + [f"{s['label']}" for s in [u] if s]
        + [f"dòng vốn nước ngoài vào Treasury: {s['label'].lower()}" for s in [cf] if s],
        [s["color"] for s in (tc, u, cf) if s],
        "Điều kiện tài chính (financial conditions) nhìn chung chưa gây cản trở lớn cho tài sản rủi ro.",
        "Điều kiện tài chính đang pha trộn — vừa có yếu tố thuận lợi vừa có yếu tố bất lợi.",
        "Điều kiện tài chính đang có yếu tố bất lợi rõ cho tài sản rủi ro — nên chọn lọc/phòng thủ hơn là risk-on toàn diện.")

    return {"economic": economic, "monetary": monetary, "investment": investment}


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

    # THEM 2026-10-08 (user gửi tài liệu đề xuất "Liquidity Impulse = ΔReserves − ΔTGA − ΔRRP +
    # ΔFed Lending") — dùng DỰ TRỮ NGÂN HÀNG (Reserves) thay vì Tổng tài sản làm gốc (khác "Net
    # Liquidity" ở trên dùng Assets) vì Reserves là tiền THỰC SỰ nằm trong hệ thống ngân hàng, không
    # bị nhiễu bởi thay đổi kỳ hạn/cơ cấu danh mục SOMA. Bỏ "ΔFed Lending" (không có chuỗi discount
    # window riêng). Impulse = thay đổi MoM của (Reserves − RRP − TGA) — dương = bơm thanh khoản
    # ròng vào hệ thống tháng đó, âm = rút ròng.
    reserves_liq = {}
    for p in periods:
        r, rrp, t = series["usm_fed_reserves"].get(p), series["usm_fed_rrp"].get(p), series["usm_fed_tga"].get(p)
        if r is not None and rrp is not None and t is not None:
            reserves_liq[p] = r - rrp - t
    impulse = {periods[i]: round(reserves_liq[periods[i]] - reserves_liq[periods[i - 1]], 0)
               for i in range(1, len(periods)) if periods[i] in reserves_liq and periods[i - 1] in reserves_liq}

    return {
        "periods": periods, "latest": latest, "latest_values": latest_vals, "changes": chg,
        "peak_period": peak_period, "peak_value": series["usm_fed_assets"][peak_period],
        "assets_vs_peak_pct": assets_vs_peak_pct,
        "ecb_latest": {"period": ecb_latest, "value": ecb.get(ecb_latest)} if ecb_latest else None,
        "liquidity_impulse_latest": impulse.get(latest),
        "history": {
            "periods": periods,
            "assets": [series["usm_fed_assets"].get(p) for p in periods],
            "treasury": [series["usm_fed_treasury"].get(p) for p in periods],
            "mbs": [series["usm_fed_mbs"].get(p) for p in periods],
            "reserves": [series["usm_fed_reserves"].get(p) for p in periods],
            "rrp": [series["usm_fed_rrp"].get(p) for p in periods],
            "tga": [series["usm_fed_tga"].get(p) for p in periods],
            "net_liquidity": [net_liq.get(p) for p in periods],
            "liquidity_impulse": [impulse.get(p) for p in periods],
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
    # THEM 2026-10-08 (user gửi tài liệu đề xuất "Contribution change": "Energy contribution +0.3pp
    # → +0.8pp => đỏ; Shelter +1.2pp → +1.0pp => xanh" — biết nhóm nào đang TĂNG áp lực lên CPI,
    # không chỉ mức đóng góp hiện tại) — so đóng góp kỳ này với đúng 3 tháng trước, CÙNG công thức.
    latest = periods[-1]
    latest_3m_ago = _shift(latest, -3)
    snapshot = []
    for key, label in US_GROUPS:
        w = US_GROUP9_WEIGHTS[key]
        yoy = _group_yoy(key, latest)
        contribution = round(w / 100 * yoy, 3) if yoy is not None else None
        yoy_3m = _group_yoy(key, latest_3m_ago) if latest_3m_ago in periods else None
        contribution_3m_ago = round(w / 100 * yoy_3m, 3) if yoy_3m is not None else None
        contribution_chg_3m = round(contribution - contribution_3m_ago, 3) if contribution is not None and contribution_3m_ago is not None else None
        snapshot.append({"key": key, "label": CONTRIB_LABEL_OVERRIDE.get(key, label), "weight_pct": w, "yoy": yoy,
                          "contribution": contribution, "contribution_3m_ago": contribution_3m_ago,
                          "contribution_chg_3m": contribution_chg_3m})
    snapshot.sort(key=lambda r: abs(r["contribution"]) if r["contribution"] is not None else -1, reverse=True)

    return {"periods": periods, "rows": rows, "weights_vintage": US_GROUP_WEIGHTS_VINTAGE,
            "weights_source": US_GROUP_WEIGHTS_SOURCE, "snapshot": snapshot, "snapshot_period": latest}
