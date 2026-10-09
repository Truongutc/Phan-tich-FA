"""Tab "Báo cáo" — tái tạo lại layout/nội dung 1 báo cáo kinh tế THÁNG kiểu DNL Capital (user gửi
4 ảnh "Tình hình kinh tế tháng 9/2026" làm mẫu, 2026-10-09), dựng HOÀN TOÀN từ dữ liệu THẬT đã có
trong data/vimo_raw.json — KHÔNG dùng AI viết văn, mọi đoạn phân tích đều RULE-BASED ghép câu từ số
liệu (cùng quy ước với build_synthesis_vimo trong template_vimo.py và _build_synthesis trong
us_macro_analysis.py). Chỗ nào thiếu dữ liệu thì bỏ qua câu đó hoặc ghi chú rõ, KHÔNG suy diễn số.
"""


def _series(raw, key):
    return {p["period"]: p["value"] for p in raw.get(key, {}).get("series", [])
            if p.get("value") is not None}


def _monthly_only(s):
    """Chỉ giữ period dạng thuần 'YYYY-MM' — loại bỏ các mốc lũy kế dạng 'YYYY-Qn/Hn/9M/FY' trộn
    lẫn trong vài chỉ báo (fdi_disbursed, fdi_registered_usd_bn, public_investment_disbursement_*)."""
    return {p: v for p, v in s.items() if len(p) == 7 and p[4] == "-" and p[5:7].isdigit()}


def _shift(period, months):
    y, m = int(period[:4]), int(period[5:7])
    idx = y * 12 + (m - 1) + months
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def _yoy(s, period):
    cur, prev = s.get(period), s.get(_shift(period, -12))
    return round((cur / prev - 1) * 100, 2) if cur is not None and prev else None


def _mom(s, period):
    cur, prev = s.get(period), s.get(_shift(period, -1))
    return round((cur / prev - 1) * 100, 2) if cur is not None and prev else None


def _last_n_periods(latest, n):
    return [_shift(latest, -(n - 1 - i)) for i in range(n)]


def _latest_period(*dicts):
    periods = [max(d) for d in dicts if d]
    return max(periods) if periods else None


def _by_year_month(s):
    out = {}
    for p, v in s.items():
        if len(p) == 7 and p[4] == "-":
            y, m = int(p[:4]), int(p[5:7])
            out.setdefault(y, {})[m] = v
    return out


def _year_compare_chart(s, n_years=2):
    """So sánh quỹ đạo theo THÁNG (1-12) giữa n_years năm gần nhất có dữ liệu — dùng cho các chỉ
    báo lũy kế YTD (tín dụng/huy động/FDI đăng ký/giải ngân đầu tư công) để thấy năm nay đang nhanh
    hay chậm hơn năm trước TẠI CÙNG mốc tháng, thay vì so 2 giá trị cuối kỳ không cùng tháng."""
    bym = _by_year_month(s)
    years = sorted(bym)[-n_years:]
    months = list(range(1, 13))
    return {
        "months": months,
        "series": {str(y): [bym.get(y, {}).get(m) for m in months] for y in years},
    }


def _build_overview_text(latest, export_m, import_m, pmi, cpi_yoy, iip):
    sentences = []
    exp_v, imp_v = export_m.get(latest), import_m.get(latest)
    if exp_v is not None and imp_v is not None:
        bal = exp_v - imp_v
        sentences.append(
            f"Tháng {latest}: xuất khẩu đạt {exp_v / 1000:.2f} tỷ USD, nhập khẩu đạt {imp_v / 1000:.2f} tỷ USD, "
            f"{'xuất siêu' if bal >= 0 else 'nhập siêu'} {abs(bal) / 1000:.2f} tỷ USD.")
    cpi_v = cpi_yoy.get(latest)
    if cpi_v is not None:
        sentences.append(f"CPI tăng {cpi_v:.2f}% so với cùng kỳ năm trước.")
    pmi_v = pmi.get(latest)
    if pmi_v is not None:
        sentences.append(f"PMI sản xuất ở mức {pmi_v:.1f} điểm ({'vùng mở rộng' if pmi_v >= 50 else 'vùng thu hẹp'}).")
    iip_v = iip.get(latest)
    if iip_v is not None:
        sentences.append(f"Sản xuất công nghiệp (IIP) tăng {iip_v:.1f}% so với cùng kỳ năm trước (lũy kế).")
    return " ".join(sentences) if sentences else "Chưa đủ dữ liệu tổng quan cho kỳ gần nhất."


def _build_monthly_table(latest, export_m, import_m, fdi_disb, pub_inv_rate, pmi, iip,
                           retail_total, visitors_m, credit_ytd, cpi_yoy, cpi_mom):
    periods = _last_n_periods(latest, 13)
    trade_balance = {p: round((export_m[p] - import_m[p]) / 1000, 2)
                      for p in periods if p in export_m and p in import_m}
    retail_growth = {p: _yoy(retail_total, p) for p in periods}

    def row(key, label, unit, series, scale=1, nd=2):
        return {"key": key, "label": label, "unit": unit,
                "values": [round(series[p] * scale, nd) if series.get(p) is not None else None for p in periods]}

    rows = [
        row("export", "Xuất khẩu", "tỷ USD", export_m, scale=1 / 1000),
        row("import", "Nhập khẩu", "tỷ USD", import_m, scale=1 / 1000),
        row("trade_balance", "Cán cân thương mại", "tỷ USD", trade_balance),
        row("fdi_disbursed", "FDI giải ngân (lũy kế từ đầu năm)", "tỷ USD", fdi_disb),
        row("public_investment_rate", "Giải ngân đầu tư công (lũy kế, %KH năm)", "%", pub_inv_rate, nd=1),
        row("pmi", "PMI sản xuất", "điểm", pmi, nd=1),
        row("iip", "Sản xuất công nghiệp — IIP (YoY, lũy kế)", "%", iip, nd=1),
        row("retail_total", "Tổng mức bán lẻ HH & DV tiêu dùng", "nghìn tỷ đồng", retail_total, scale=1 / 1000),
        row("retail_growth", "— trong đó tăng trưởng (YoY)", "%", retail_growth),
        row("visitors", "Khách quốc tế đến Việt Nam", "triệu lượt", visitors_m),
        row("credit_ytd", "Tăng trưởng tín dụng (YTD từ đầu năm)", "%", credit_ytd),
        row("cpi_yoy", "CPI (YoY)", "%", cpi_yoy),
        row("cpi_mom", "CPI (MoM)", "%", cpi_mom),
    ]
    return {"periods": periods, "rows": rows}


def _build_consumption(latest, retail_total, retail_goods, retail_hosp, retail_travel, retail_other,
                         cpi_yoy, cpi_mom, cpi_food, cpi_housing, cpi_health, cpi_transport, cpi_other):
    periods = sorted(retail_total)

    def scaled(s):
        return [round(s[p] / 1000, 2) if s.get(p) is not None else None for p in periods]

    nominal_yoy = [_yoy(retail_total, p) for p in periods]
    real_yoy = []
    for p, ny in zip(periods, nominal_yoy):
        cy = cpi_yoy.get(p)
        real_yoy.append(round((1 + ny / 100) / (1 + cy / 100) * 100 - 100, 2)
                          if ny is not None and cy is not None else None)

    retail_chart = {
        "periods": periods,
        "goods": scaled(retail_goods), "hospitality": scaled(retail_hosp),
        "travel": scaled(retail_travel), "other": scaled(retail_other),
        "total": scaled(retail_total),
        "nominal_yoy": nominal_yoy, "real_yoy": real_yoy,
    }

    cpi_periods = sorted(set(cpi_food) | set(cpi_housing) | set(cpi_health)
                          | set(cpi_transport) | set(cpi_other))[-24:]
    cpi_chart = {
        "periods": cpi_periods,
        "food": [cpi_food.get(p) for p in cpi_periods],
        "housing_utilities": [cpi_housing.get(p) for p in cpi_periods],
        "healthcare": [cpi_health.get(p) for p in cpi_periods],
        "transport": [cpi_transport.get(p) for p in cpi_periods],
        "other": [cpi_other.get(p) for p in cpi_periods],
        "cpi_yoy": [cpi_yoy.get(p) for p in cpi_periods],
    }

    rt_latest, rt_mom, rt_yoy = retail_total.get(latest), _mom(retail_total, latest), _yoy(retail_total, latest)
    s1 = []
    if rt_latest is not None and rt_mom is not None and rt_yoy is not None:
        s1.append(
            f"Tổng mức bán lẻ hàng hóa & doanh thu dịch vụ tiêu dùng tháng {latest} đạt {rt_latest / 1000:.1f} "
            f"nghìn tỷ đồng, {'tăng' if rt_mom >= 0 else 'giảm'} {abs(rt_mom):.1f}% so với tháng trước và "
            f"{'tăng' if rt_yoy >= 0 else 'giảm'} {abs(rt_yoy):.1f}% so với cùng kỳ năm trước.")
        seg_yoys = {"bán lẻ hàng hóa": _yoy(retail_goods, latest), "dịch vụ lưu trú, ăn uống": _yoy(retail_hosp, latest),
                     "du lịch lữ hành": _yoy(retail_travel, latest), "dịch vụ khác": _yoy(retail_other, latest)}
        valid = {k: v for k, v in seg_yoys.items() if v is not None}
        if valid:
            leader = max(valid, key=valid.get)
            s1.append(f"Trong 4 phân khúc, {leader} dẫn đầu mức tăng trưởng với {valid[leader]:+.1f}% YoY.")
    txt1 = " ".join(s1) if s1 else "Chưa đủ dữ liệu bán lẻ tháng gần nhất."

    cpi_y, cpi_m = cpi_yoy.get(latest), cpi_mom.get(latest)
    groups_latest = {"Thực phẩm": cpi_food.get(latest), "Nhà, điện, nước": cpi_housing.get(latest),
                       "Y tế": cpi_health.get(latest), "Vận tải": cpi_transport.get(latest),
                       "Khác": cpi_other.get(latest)}
    valid_groups = {k: v for k, v in groups_latest.items() if v is not None}
    s2 = []
    if cpi_y is not None:
        tail = f" và {'tăng' if cpi_m >= 0 else 'giảm'} {abs(cpi_m):.2f}% so với tháng trước." if cpi_m is not None else "."
        s2.append(f"CPI tháng {latest} tăng {cpi_y:.2f}% so với cùng kỳ năm trước{tail}")
    if valid_groups:
        leader = max(valid_groups, key=lambda k: abs(valid_groups[k]))
        s2.append(f"Nhóm {leader} đóng góp nhiều nhất vào mức tăng CPI chung, với {valid_groups[leader]:+.2f} điểm %.")
    txt2 = " ".join(s2) if s2 else "Chưa đủ dữ liệu CPI tháng gần nhất."

    s3 = []
    if rt_yoy is not None and cpi_y is not None:
        real = round((1 + rt_yoy / 100) / (1 + cpi_y / 100) * 100 - 100, 2)
        s3.append(f"Sau khi trừ lạm phát (CPI {cpi_y:+.2f}% YoY), tăng trưởng THỰC của tiêu dùng bán lẻ ước đạt "
                   f"{real:+.2f}% — sức mua thực {'vẫn đang tăng' if real >= 0 else 'đang bị lạm phát ăn mòn, giảm thực tế'}.")
    txt3 = " ".join(s3) if s3 else "Chưa đủ dữ liệu để tính tăng trưởng thực của tiêu dùng."

    return {"retailChart": retail_chart, "cpiGroupChart": cpi_chart, "paragraphs": [txt1, txt2, txt3]}


def _build_production(latest, export_m, import_m, iip, iip_manuf, iip_elec, iip_water, iip_mining, pmi):
    periods = sorted(set(export_m) | set(import_m))
    trade_chart = {
        "periods": periods,
        "export": [round(export_m[p] / 1000, 2) if p in export_m else None for p in periods],
        "import": [round(import_m[p] / 1000, 2) if p in import_m else None for p in periods],
        "balance": [round((export_m[p] - import_m[p]) / 1000, 2) if p in export_m and p in import_m else None
                     for p in periods],
    }
    iip_periods = sorted(iip)
    iip_chart = {"periods": iip_periods, "values": [iip[p] for p in iip_periods]}
    iip_sector_periods = sorted(set(iip_manuf) | set(iip_elec) | set(iip_water) | set(iip_mining))
    iip_sector_chart = {
        "periods": iip_sector_periods,
        "manufacturing": [iip_manuf.get(p) for p in iip_sector_periods],
        "electricity": [iip_elec.get(p) for p in iip_sector_periods],
        "water_waste": [iip_water.get(p) for p in iip_sector_periods],
        "mining": [iip_mining.get(p) for p in iip_sector_periods],
    }
    pmi_periods = sorted(pmi)
    pmi_chart = {"periods": pmi_periods, "values": [pmi[p] for p in pmi_periods]}

    paragraphs = []
    exp_v, imp_v = export_m.get(latest), import_m.get(latest)
    exp_yoy, imp_yoy = _yoy(export_m, latest), _yoy(import_m, latest)
    if exp_v is not None and imp_v is not None:
        bal = exp_v - imp_v
        s1 = (f"Xuất khẩu tháng {latest} đạt {exp_v / 1000:.2f} tỷ USD"
               + (f" ({exp_yoy:+.1f}% YoY)" if exp_yoy is not None else "")
               + f", nhập khẩu đạt {imp_v / 1000:.2f} tỷ USD"
               + (f" ({imp_yoy:+.1f}% YoY)" if imp_yoy is not None else "")
               + f" — {'xuất siêu' if bal >= 0 else 'nhập siêu'} {abs(bal) / 1000:.2f} tỷ USD trong tháng.")
        paragraphs.append(s1)
    else:
        paragraphs.append("Chưa đủ dữ liệu xuất nhập khẩu tháng gần nhất.")

    iip_v = iip.get(latest)
    sector_vals = {"chế biến, chế tạo": iip_manuf.get(latest), "sản xuất & phân phối điện": iip_elec.get(latest),
                    "cung cấp nước, xử lý rác thải": iip_water.get(latest), "khai khoáng": iip_mining.get(latest)}
    valid_sectors = {k: v for k, v in sector_vals.items() if v is not None}
    s2 = []
    if iip_v is not None:
        s2.append(f"Chỉ số sản xuất công nghiệp (IIP) lũy kế đến tháng {latest} tăng {iip_v:.1f}% so với cùng kỳ năm trước.")
    if valid_sectors:
        leader = max(valid_sectors, key=valid_sectors.get)
        laggard = min(valid_sectors, key=valid_sectors.get)
        s2.append(f"Trong đó ngành {leader} dẫn đầu mức tăng lũy kế ({valid_sectors[leader]:+.1f}%), ngành "
                   f"{laggard} tăng chậm nhất ({valid_sectors[laggard]:+.1f}%).")
    paragraphs.append(" ".join(s2) if s2 else
                        "Chưa đủ dữ liệu IIP theo ngành (chuỗi mới bắt đầu ghi nhận từ kỳ gần đây, sẽ dài dần theo mỗi lần cập nhật).")

    pmi_v, pmi_prev = pmi.get(latest), pmi.get(_shift(latest, -1))
    s3 = []
    if pmi_v is not None:
        s3.append(f"PMI sản xuất tháng {latest} đạt {pmi_v:.1f} điểm, {'trên' if pmi_v >= 50 else 'dưới'} ngưỡng 50 — "
                   f"ngành sản xuất đang {'mở rộng' if pmi_v >= 50 else 'thu hẹp'}.")
        if pmi_prev is not None:
            diff = pmi_v - pmi_prev
            s3.append(f"So với tháng trước ({pmi_prev:.1f} điểm), PMI {'tăng' if diff >= 0 else 'giảm'} {abs(diff):.1f} điểm.")
    paragraphs.append(" ".join(s3) if s3 else "Chưa có dữ liệu PMI tháng gần nhất.")

    return {"tradeChart": trade_chart, "iipChart": iip_chart, "iipSectorChart": iip_sector_chart,
            "pmiChart": pmi_chart, "paragraphs": paragraphs}


def _build_investment(latest, pmi, pub_inv_val, pub_inv_rate, credit_ytd, deposit_ytd, fdi_reg, fdi_disb):
    pmi_periods = sorted(pmi)
    pmi_chart = {"periods": pmi_periods, "values": [pmi[p] for p in pmi_periods]}
    pub_inv_chart = _year_compare_chart(pub_inv_rate)
    credit_chart = _year_compare_chart(credit_ytd)
    deposit_chart = _year_compare_chart(deposit_ytd)
    fdi_chart = _year_compare_chart(fdi_reg)
    # THEM 2026-10-09 (user: "khu vực này thêm biểu đồ giá trị FDI giải ngân đầu tư nhé, kia là FDI
    # đăng ký thôi" — fdi_chart ở trên là FDI ĐĂNG KÝ (vốn cam kết), KHÁC fdi_disbursed (vốn THỰC TẾ
    # đã giải ngân, cùng nguồn NSO đã dùng ở monthlyTable) — thêm biểu đồ so sánh theo năm riêng.
    fdi_disbursed_chart = _year_compare_chart(fdi_disb)

    # Các chỉ báo YTD (VBMA/NSO) thường TRỄ 1-3 tháng so với xuất/nhập khẩu (latest toàn báo cáo)
    # — dùng kỳ MỚI NHẤT CỦA RIÊNG từng chỉ báo (không ép theo latest chung) để không bị rơi vào
    # "chưa đủ dữ liệu" một cách giả tạo chỉ vì lệch nhịp công bố, đồng thời luôn NÊU RÕ kỳ thực tế
    # đang nói tới trong câu để không gây hiểu nhầm là số của tháng mới nhất.
    paragraphs = []
    pub_inv_period = max(pub_inv_rate) if pub_inv_rate else None
    rate_latest = pub_inv_rate.get(pub_inv_period) if pub_inv_period else None
    val_latest = pub_inv_val.get(pub_inv_period) if pub_inv_period else None
    s1 = []
    if rate_latest is not None:
        s1.append(f"Giải ngân vốn đầu tư công lũy kế đến tháng {pub_inv_period} đạt {rate_latest:.1f}% kế hoạch năm"
                    + (f" ({val_latest:,.1f} nghìn tỷ đồng)." if val_latest is not None else "."))
        prev_year_period = f"{int(pub_inv_period[:4]) - 1}-{pub_inv_period[5:7]}"
        prev_rate = pub_inv_rate.get(prev_year_period)
        if prev_rate is not None:
            diff = rate_latest - prev_rate
            s1.append(f"So với cùng kỳ năm trước ({prev_rate:.1f}% KH), tiến độ giải ngân đang "
                        f"{'nhanh' if diff >= 0 else 'chậm'} hơn {abs(diff):.1f} điểm %.")
    paragraphs.append(" ".join(s1) if s1 else "Chưa đủ dữ liệu giải ngân đầu tư công.")

    credit_period = max(credit_ytd) if credit_ytd else None
    deposit_period = max(deposit_ytd) if deposit_ytd else None
    fdi_period = max(fdi_reg) if fdi_reg else None
    credit_latest = credit_ytd.get(credit_period) if credit_period else None
    deposit_latest = deposit_ytd.get(deposit_period) if deposit_period else None
    fdi_latest = fdi_reg.get(fdi_period) if fdi_period else None
    credit_part = f"tín dụng lũy kế từ đầu năm đến tháng {credit_period} đạt {credit_latest:.2f}%" if credit_latest is not None else None
    deposit_part = f"huy động vốn (đến tháng {deposit_period}) đạt {deposit_latest:.2f}%" if deposit_latest is not None else None
    s2 = [p for p in (credit_part, deposit_part) if p]
    if s2:
        txt2 = "Tăng trưởng " + ", ".join(s2) + " (YTD)."
        if credit_latest is not None and deposit_latest is not None and credit_period == deposit_period:
            gap = credit_latest - deposit_latest
            txt2 += (f" Tín dụng {'tăng nhanh hơn' if gap > 0 else 'tăng chậm hơn'} huy động {abs(gap):.2f} điểm % — "
                      + ("có thể tạo áp lực lên lãi suất liên ngân hàng." if gap > 0 else "thanh khoản hệ thống nhìn chung dư dả hơn."))
    else:
        txt2 = "Chưa đủ dữ liệu tín dụng/huy động."
    if fdi_latest is not None:
        txt2 += f" FDI đăng ký lũy kế từ đầu năm đến tháng {fdi_period} đạt {fdi_latest:.1f} tỷ USD."
    paragraphs.append(txt2)

    return {"pmiChart": pmi_chart, "publicInvestmentChart": pub_inv_chart, "creditChart": credit_chart,
            "depositChart": deposit_chart, "fdiChart": fdi_chart, "fdiDisbursedChart": fdi_disbursed_chart,
            "paragraphs": paragraphs}


def build_vn_report(raw):
    """Trả dict 'vnReport' cho vimo.json, hoặc None nếu chưa đủ dữ liệu lõi (xuất/nhập khẩu)."""
    export_m = _monthly_only(_series(raw, "export_monthly_total"))
    import_m = _monthly_only(_series(raw, "import_monthly_total"))
    if not export_m or not import_m:
        return None

    pmi = _series(raw, "pmi_manufacturing")
    cpi_yoy = _series(raw, "cpi_yoy")
    cpi_mom = _series(raw, "cpi_mom")
    iip = _series(raw, "iip_growth")
    fdi_disb = _monthly_only(_series(raw, "fdi_disbursed"))
    fdi_reg = _monthly_only(_series(raw, "fdi_registered_usd_bn"))
    pub_inv_rate = _monthly_only(_series(raw, "public_investment_disbursement_rate"))
    pub_inv_val = _monthly_only(_series(raw, "public_investment_disbursement_value"))
    credit_ytd = _series(raw, "credit_growth_ytd_monthly")
    deposit_ytd = _series(raw, "deposit_growth_ytd_monthly")
    retail_total = _series(raw, "retail_sales_total_monthly")
    retail_goods = _series(raw, "retail_sales_goods_monthly")
    retail_hosp = _series(raw, "retail_sales_hospitality_monthly")
    retail_travel = _series(raw, "retail_sales_travel_monthly")
    retail_other = _series(raw, "retail_sales_other_monthly")
    visitors_m = _series(raw, "international_visitors_monthly")
    cpi_food = _series(raw, "cpi_contrib_food")
    cpi_housing = _series(raw, "cpi_contrib_housing_utilities")
    cpi_health = _series(raw, "cpi_contrib_healthcare")
    cpi_transport = _series(raw, "cpi_contrib_transport")
    cpi_other = _series(raw, "cpi_contrib_other")
    iip_manuf = _series(raw, "iip_manufacturing_ytd_yoy")
    iip_elec = _series(raw, "iip_electricity_ytd_yoy")
    iip_water = _series(raw, "iip_water_waste_ytd_yoy")
    iip_mining = _series(raw, "iip_mining_ytd_yoy")

    latest = _latest_period(export_m, import_m, cpi_yoy, pmi)
    if not latest:
        return None

    return {
        "asOf": latest,
        "summaryText": _build_overview_text(latest, export_m, import_m, pmi, cpi_yoy, iip),
        "monthlyTable": _build_monthly_table(latest, export_m, import_m, fdi_disb, pub_inv_rate, pmi, iip,
                                                retail_total, visitors_m, credit_ytd, cpi_yoy, cpi_mom),
        "consumption": _build_consumption(latest, retail_total, retail_goods, retail_hosp, retail_travel,
                                             retail_other, cpi_yoy, cpi_mom, cpi_food, cpi_housing, cpi_health,
                                             cpi_transport, cpi_other),
        "production": _build_production(latest, export_m, import_m, iip, iip_manuf, iip_elec, iip_water,
                                           iip_mining, pmi),
        "investment": _build_investment(latest, pmi, pub_inv_val, pub_inv_rate, credit_ytd, deposit_ytd, fdi_reg, fdi_disb),
    }
