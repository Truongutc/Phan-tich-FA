"""
template_astro.py — Đóng gói dữ liệu chiêm tinh tài chính (data/astro.json) cho chiemtinh.html.

Xem fetch_astro_data.py để biết chi tiết cách tính + lưu ý quan trọng về bản chất khung lý thuyết
này (chiêm tinh tài chính/W.D. Gann — dữ liệu thiên văn chính xác, mối liên hệ với thị trường
KHÔNG được khoa học chính thống kiểm chứng).

Module này THÊM lớp DIỄN GIẢI rule-based (KHÔNG dùng AI/LLM — ráp câu chữ từ template có điều
kiện, cùng triết lý với build_synthesis_vimo() trong template_vimo.py) — biến dữ liệu thiên văn
thô thành đánh giá tác động hiện tại + dự báo giai đoạn tới, theo đúng yêu cầu user (2026-08-04):
"đưa ra đánh giá kết luận hiện tại tác động tới tài chính chứng khoán và tâm lí con người như nào,
và sự vận động đó trong tương lai gần có những ngày hoặc giai đoạn nào quan trọng".
"""
import datetime
import json
import os

import requests

import fetch_astro_data as astro

BACKTEST_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
               "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

FORECAST_DAYS_AHEAD = 90

# Ý nghĩa hành tinh trong chiêm tinh tài chính (hiển thị trong chiemtinh.html).
PLANET_MEANING = {
    "Mặt Trời": "xu hướng chính, niềm tin thị trường",
    "Mặt Trăng": "tâm lý đám đông, biến động ngắn hạn",
    "Sao Thủy": "dòng thông tin, thanh khoản, hoạt động giao dịch",
    "Sao Kim": "định giá tài sản, nhóm ngành tiêu dùng/tài chính",
    "Sao Hỏa": "biến động mạnh, hoạt động đầu cơ",
    "Sao Mộc": "mở rộng, tăng trưởng, tâm lý lạc quan",
    "Sao Thổ": "thắt chặt, kỷ luật, rủi ro suy thoái/khủng hoảng",
    "Sao Thiên Vương": "biến động đột ngột/cú sốc, nhóm công nghệ",
    "Sao Hải Vương": "bong bóng định giá, sự mơ hồ/ảo tưởng thị trường",
    "Sao Diêm Vương": "biến đổi cấu trúc dài hạn, khủng hoảng/tái cấu trúc",
}

# Ý nghĩa TRUYỀN THỐNG của từng hành tinh khi nghịch hành — lấy tinh thần từ cách W.D. Gann diễn
# giải trong "The Tunnel Thru the Air" (ứng dụng vào lĩnh vực tài chính mà hành tinh đó cai quản,
# xem PLANET_MEANING) — luôn dùng ngôn ngữ "theo lý thuyết truyền thống"/"thường được cho là",
# KHÔNG khẳng định chắc chắn (nhất quán với cảnh báo ở đầu trang).
RETROGRADE_MEANING = {
    "Sao Thủy": "thông tin/giao dịch dễ nhiễu loạn, quyết định vội vàng dễ sai sót, hợp đồng/thỏa thuận dễ trục trặc — thị trường thường biến động thất thường, xu hướng không rõ ràng, nên cẩn trọng với các quyết định đầu tư quan trọng trong giai đoạn này",
    "Sao Kim": "định giá tài sản/nhóm ngành tiêu dùng-tài chính dễ được thị trường nhìn lại và điều chỉnh, tâm lý định giá lại những gì đã thiết lập trước đó",
    "Sao Hỏa": "năng lượng/tốc độ hành động của thị trường bị dồn nén, các xu hướng đầu cơ ngắn hạn dễ chững lại hoặc đảo chiều bất ngờ",
    "Sao Mộc": "đà mở rộng/tăng trưởng chậm lại, thời điểm thị trường có xu hướng nhìn lại các quyết định đầu tư/mở rộng đã đưa ra thay vì mở rộng mới",
    "Sao Thổ": "nhìn lại kỷ luật/cấu trúc rủi ro đã thiết lập, dễ bộc lộ điểm yếu cấu trúc tài chính tiềm ẩn",
    "Sao Thiên Vương": "biến động bị dồn nén, thường giải phóng đột ngột quanh thời điểm hành tinh quay lại thuận hành (station direct)",
    "Sao Hải Vương": "ảo tưởng định giá/kỳ vọng thị trường dễ bị nhìn lại và điều chỉnh về gần thực tế hơn",
    "Sao Diêm Vương": "thay đổi cấu trúc dài hạn diễn ra âm thầm, ít tạo biến động tức thời rõ rệt",
}

# Ý nghĩa 12 Nhà (House) — theo đúng lý thuyết chiêm tinh học kinh điển (nhất quán giữa mọi trường
# phái phương Tây), diễn giải qua lăng kính tài chính. Nhà PHỤ THUỘC ĐỊA ĐIỂM (xem HOUSE_LOCATION_
# NAME/HOUSE_LAT/HOUSE_LON trong fetch_astro_data.py — user 2026-08-04 chọn Hà Nội làm chuẩn).
HOUSE_MEANING = {
    1: "Bản ngã/Diện mạo — hình ảnh, động lượng ban đầu của thị trường",
    2: "Tài sản/Tiền bạc — tài sản hữu hình, giá trị, dòng tiền cá nhân/doanh nghiệp",
    3: "Giao tiếp/Thông tin — tin tức, dòng thông tin, giao dịch ngắn hạn",
    4: "Gia đình/Nền tảng — bất động sản, nền tảng cơ bản của nền kinh tế",
    5: "Đầu cơ/Rủi ro — ĐẦU CƠ, cờ bạc tài chính, khẩu vị rủi ro",
    6: "Công việc/Dịch vụ — lao động, năng suất, chi phí vận hành",
    7: "Đối tác/Hợp đồng — hợp tác, M&A, đối thủ cạnh tranh, hợp đồng",
    8: "Tài chính chung/Nợ — NỢ, thuế, tiền người khác, khủng hoảng/chuyển đổi",
    9: "Mở rộng/Quốc tế — thương mại quốc tế, luật pháp/quy định, triết lý đầu tư",
    10: "Sự nghiệp/Vị thế — chính sách, uy tín, vai trò lãnh đạo thị trường",
    11: "Cộng đồng/Kỳ vọng — tâm lý đám đông, dòng vốn tổ chức, kỳ vọng tập thể",
    12: "Tiềm ẩn/Bí mật — rủi ro ẩn, gian lận, điểm mù hệ thống",
}
# 3 Nhà được coi là liên quan TRỰC TIẾP NHẤT tới tài chính/đầu cơ trong chiêm tinh tài chính truyền
# thống — dùng để làm nổi bật trong đánh giá hiện tại (Nhà 2 = tiền bạc, Nhà 5 = đầu cơ, Nhà 8 = nợ/
# khủng hoảng).
FINANCIAL_HOUSES = {2, 5, 8}

# (nhãn ngắn, mô tả) cho từng loại góc chiếu — dùng cả cho đánh giá hiện tại lẫn dự báo.
ASPECT_TONE = {
    "Hợp (Conjunction)": ("khởi đầu", "thường đánh dấu điểm KHỞI ĐẦU của 1 chu kỳ/xu hướng mới liên quan tới 2 lĩnh vực này"),
    "Vuông (Square)": ("căng thẳng", "thường tạo CĂNG THẲNG/xung đột, liên quan tới biến động mạnh hoặc thời điểm đảo chiều đột ngột"),
    "Xung (Opposition)": ("đối lập", "thường tạo thế GIẰNG CO/đối lập, liên quan tới các điểm đỉnh-đáy cục bộ, thị trường lưỡng lự"),
    "Tam hợp (Trine)": ("hài hòa", "thường tạo sự HÀI HÒA/thuận lợi, liên quan tới giai đoạn ổn định hoặc tăng trưởng êm"),
    "Lục phân (Sextile)": ("cơ hội nhẹ", "thường tạo CƠ HỘI nhẹ nhàng, ít biến động mạnh"),
}

# Theo trường phái Bill Meridian (Planetary Stock Trading, xem tài liệu user gửi 2026-08-03) — góc
# chiếu giữa các hành tinh CHẬM (Mộc trở ra) quyết định chu kỳ kinh tế LỚN, quan trọng hơn góc
# chiếu chỉ liên quan Mặt Trời/Mặt Trăng/hành tinh nhanh (vốn đổi liên tục, chỉ tạo biến động ngày).
SLOW_PLANETS = {"Sao Mộc", "Sao Thổ", "Sao Thiên Vương", "Sao Hải Vương", "Sao Diêm Vương"}

# Chiều hướng (sentiment) RULE-BASED cho bảng dự báo. SỬA LẠI (user 2026-08-04 chỉ ra bảng dự báo
# "đỏ lòe hết cả" suốt 2 tháng, hỏi thẳng "giữa 2 mốc thì thị trường lên hay xuống") — lỗi thiết kế
# TRƯỚC ĐÓ: gán Vuông/Xung = "negative" (đỏ) ngầm khẳng định "sẽ giảm", trong khi đúng lý thuyết
# chiêm tinh tài chính, góc chiếu CỨNG chỉ có nghĩa CĂNG THẲNG/BIẾN ĐỘNG MẠNH — CÓ THỂ đi lên hoặc
# xuống, không có chiều rõ ràng. Sửa: tách riêng "tension" (màu hổ phách/cam — biến động mạnh,
# KHÔNG khẳng định chiều) khỏi "negative" thật (chỉ dành cho tín hiệu tương đối MỘT CHIỀU theo lý
# thuyết truyền thống: hành tinh BẮT ĐẦU nghịch hành). Tam hợp/lục phân (hài hòa) = positive.
# Hợp (conjunction, khởi đầu) = neutral. Hành tinh THUẬN HÀNH TRỞ LẠI = positive (giải phóng/rõ
# ràng). Nhật/nguyệt thực = tension (biến động mạnh, không rõ chiều — giống vuông/xung).
ASPECT_SENTIMENT = {
    "Hợp (Conjunction)": "neutral",
    "Vuông (Square)": "tension",
    "Xung (Opposition)": "tension",
    "Tam hợp (Trine)": "positive",
    "Lục phân (Sextile)": "positive",
}


def build_market_verdict(daily_pressure_series):
    """Kết luận TỔNG QUÁT ở đầu trang (user 2026-08-04: "hiện tại đang thuận hay đang rủi ro xấu,
    hay đang chuẩn bị đảo chiều tăng, đảo chiều giảm") — dựa trên build_daily_pressure_series()
    (fetch_astro_data.py): điểm HÔM NAY (current_score) so với điểm TRUNG BÌNH 10 ngày tới
    (near_term_avg) để phát hiện xu hướng đang cải thiện/xấu đi, không chỉ nhìn 1 thời điểm. Trả
    {label, colorKey, currentScore, nearTermAvg, detail}."""
    if not daily_pressure_series:
        return None
    current_score = daily_pressure_series[0]["score"]
    near_term = daily_pressure_series[1:11]
    near_term_avg = round(sum(p["score"] for p in near_term) / len(near_term), 3) if near_term else current_score
    trend = round(near_term_avg - current_score, 3)

    FAV_THRESHOLD, TREND_THRESHOLD = 0.8, 0.5
    is_favorable_now = current_score > FAV_THRESHOLD
    is_risky_now = current_score < -FAV_THRESHOLD
    improving = trend > TREND_THRESHOLD
    worsening = trend < -TREND_THRESHOLD

    if is_risky_now and improving:
        label, color_key = "🔄 Đang chịu áp lực nhưng có dấu hiệu CHUẨN BỊ ĐẢO CHIỀU TĂNG", "reversal_up"
    elif is_favorable_now and worsening:
        label, color_key = "🔄 Đang thuận lợi nhưng có dấu hiệu CHUẨN BỊ ĐẢO CHIỀU GIẢM", "reversal_down"
    elif is_risky_now:
        label, color_key = "⚠️ Đang chịu ÁP LỰC/RỦI RO", "risk"
    elif is_favorable_now:
        label, color_key = "✅ Đang THUẬN LỢI", "favorable"
    else:
        label, color_key = "⚪ TRUNG TÍNH — chưa có tín hiệu rõ ràng", "neutral"

    detail = (f"Điểm áp lực/thuận lợi hôm nay: {current_score:+.2f} · Trung bình 10 ngày tới: "
              f"{near_term_avg:+.2f} ({'đang cải thiện' if trend > 0 else 'đang xấu đi' if trend < 0 else 'ổn định'} "
              f"{abs(trend):+.2f} điểm). Điểm dương = nghiêng thuận lợi (nhiều góc chiếu hài hòa/ít nghịch hành), "
              f"điểm âm = nghiêng áp lực (nhiều góc chiếu căng thẳng/nhiều hành tinh nghịch hành cùng lúc). "
              f"Đây là chỉ số TỔNG HỢP nhiều tín hiệu cùng lúc, không phải chỉ 1 sự kiện đơn lẻ — có thể trái "
              f"chiều với 1 sự kiện cụ thể trong bảng dự báo bên dưới nếu các tín hiệu khác đang mạnh hơn.")

    return {"label": label, "colorKey": color_key, "currentScore": current_score,
            "nearTermAvg": near_term_avg, "trend": trend, "detail": detail}


def build_current_assessment(positions, current_aspects):
    """Sinh nhận định RULE-BASED (không AI) về tác động HIỆN TẠI của vị trí/góc chiếu hành tinh
    tới tài chính & tâm lý thị trường, theo lý thuyết chiêm tinh tài chính truyền thống. Trả
    {overallText, retrogradeLines: [...], aspectLines: [...]}."""
    retro_planets = [p["name"] for p in positions if p["retrograde"]]

    retrograde_lines = [
        f"{name} đang nghịch hành: theo lý thuyết chiêm tinh tài chính truyền thống, đây là giai đoạn {RETROGRADE_MEANING[name]}."
        for name in retro_planets if name in RETROGRADE_MEANING
    ]

    major_aspects = [a for a in current_aspects
                     if a["a"] in SLOW_PLANETS and a["b"] in SLOW_PLANETS and a["angle"] in (0, 90, 180)]
    aspect_lines = []
    for a in major_aspects:
        _, desc = ASPECT_TONE.get(a["aspect"], ("", ""))
        status = "đang TỚI gần chính xác (ảnh hưởng mạnh dần)" if a["applying"] else "đang QUA đỉnh điểm (ảnh hưởng nhạt dần)"
        aspect_lines.append(
            f"{a['a']} {a['aspect']} {a['b']} (orb {a['orb']}°, {status}): liên quan \"{PLANET_MEANING.get(a['a'], '')}\" "
            f"và \"{PLANET_MEANING.get(a['b'], '')}\" — {desc}."
        )

    if retro_planets and major_aspects:
        overall = (f"Hiện có {len(retro_planets)} hành tinh nghịch hành VÀ {len(major_aspects)} góc chiếu cứng lớn "
                    f"(giữa các hành tinh chu kỳ chậm) đang hoạt động — theo lý thuyết chiêm tinh tài chính, đây là "
                    f"giai đoạn có khả năng biến động/bất định cao hơn bình thường, nên thận trọng với quyết định lớn.")
    elif retro_planets:
        overall = (f"Hiện có {len(retro_planets)} hành tinh nghịch hành, KHÔNG có góc chiếu cứng lớn nào giữa các "
                    f"hành tinh chu kỳ chậm — biến động (nếu có) nhiều khả năng mang tính ngắn hạn/cục bộ hơn là "
                    f"thay đổi cấu trúc lớn.")
    elif major_aspects:
        overall = (f"Có {len(major_aspects)} góc chiếu cứng lớn giữa các hành tinh chu kỳ chậm đang hoạt động, "
                    f"không có hành tinh nào nghịch hành — đáng chú ý cho biến động mang tính chu kỳ dài hơi hơn.")
    else:
        overall = ("Không có hành tinh nghịch hành hay góc chiếu cứng lớn nào đáng chú ý giữa các hành tinh chu kỳ "
                    "chậm — theo lý thuyết chiêm tinh tài chính, đây được coi là giai đoạn tương đối ổn định/trung tính.")

    # Nhà tài chính (2/5/8) đang có hành tinh nào trú ngụ — góc nhìn bổ sung theo địa điểm tham
    # chiếu (xem HOUSE_LOCATION_NAME trong fetch_astro_data.py).
    house_lines = []
    for p in positions:
        h = p.get("house")
        if h in FINANCIAL_HOUSES:
            house_lines.append(
                f"{p['name']} đang ở Nhà {h} ({HOUSE_MEANING[h]}) — chủ đề của hành tinh này "
                f"(\"{PLANET_MEANING.get(p['name'], '')}\") đang được nhấn mạnh trong lĩnh vực Nhà {h}."
            )

    return {"overallText": overall, "retrogradeLines": retrograde_lines, "aspectLines": aspect_lines,
            "houseLines": house_lines}


def build_forecast_timeline(upcoming_aspects, upcoming_eclipses, retro_stations, days_ahead=FORECAST_DAYS_AHEAD):
    """Hợp nhất góc chiếu sắp tới + nhật/nguyệt thực + ngày hành tinh đổi chiều (station) trong
    `days_ahead` ngày tới thành 1 dòng thời gian duy nhất, mỗi sự kiện kèm diễn giải rule-based.
    Trả list sắp theo ngày tăng dần."""
    cutoff = (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=days_ahead)).strftime("%Y-%m-%d")
    events = []

    for a in upcoming_aspects:
        if a["date"] > cutoff:
            continue
        _, desc = ASPECT_TONE.get(a["aspect"], ("", ""))
        events.append({
            "date": a["date"], "type": "aspect",
            "title": f"{a['a']} {a['aspect']} {a['b']}",
            "interpretation": f"Liên quan \"{PLANET_MEANING.get(a['a'], '')}\" và \"{PLANET_MEANING.get(a['b'], '')}\" — {desc}.",
            "sentiment": ASPECT_SENTIMENT.get(a["aspect"], "neutral"),
        })

    for e in upcoming_eclipses:
        if e["date"] > cutoff:
            continue
        is_solar = e["type"] == "solar"
        events.append({
            "date": e["date"], "type": e["type"],
            "title": "Nhật thực" if is_solar else "Nguyệt thực",
            "interpretation": ("Nhiều nhà phân tích chiêm tinh tài chính coi đây là \"mốc thời gian động\" lớn — thời "
                                "điểm thị trường có xác suất biến động/đảo chiều cao hơn bình thường trong vài tuần "
                                "quanh ngày này." if is_solar else
                                "Tương tự nhật thực, được coi là mốc thời gian động — thường liên quan tới biến động "
                                "tâm lý đám đông/thanh khoản ngắn hạn."),
            "sentiment": "tension",
        })

    for s in retro_stations:
        if s["date"] > cutoff:
            continue
        is_retro_start = s["type"] == "station_retrograde"
        events.append({
            "date": s["date"], "type": s["type"],
            "title": f"{s['planet']} {'bắt đầu nghịch hành' if is_retro_start else 'thuận hành trở lại'}",
            "interpretation": (f"Theo lý thuyết truyền thống: {RETROGRADE_MEANING.get(s['planet'], '')}."
                                if is_retro_start else
                                f"Đánh dấu thời điểm năng lượng bị dồn nén trong giai đoạn {s['planet']} nghịch hành "
                                f"trước đó được GIẢI PHÓNG — thị trường dễ có biến động rõ rệt quanh mốc này."),
            "sentiment": "negative" if is_retro_start else "positive",
        })

    events.sort(key=lambda e: e["date"])
    return events


def fetch_vnindex_price_history(days_back=380):
    """Lấy lịch sử giá đóng cửa VN-Index theo ngày từ Vietcap IQ (API nội bộ, xem
    .agents/skills/giaodichvietcap/SKILL.md) — dùng để đối chiếu (backtest) trực quan với chỉ số
    áp lực/thuận lợi chiêm tinh, theo yêu cầu user (2026-08-04): "vẽ backtest lại trong 1 năm trở
    lại đây để test". Trả list [{date, close}] tăng dần theo thời gian, hoặc [] nếu lỗi mạng
    (fail-safe — không được làm hỏng phần dữ liệu chiêm tinh chính nếu API ngoài lỗi/đổi cấu trúc)."""
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": BACKTEST_UA,
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://trading.vietcap.com.vn/",
        })
        s.get("https://trading.vietcap.com.vn/iq/market?tab=information", timeout=15)
        to_date = datetime.date.today()
        from_date = to_date - datetime.timedelta(days=days_back)
        url = "https://trading.vietcap.com.vn/api/iq-insight-service/v1/market-indices/history"
        points = {}
        page = 0
        page_size = 200
        while page <= 40:  # 40*200=8000 phiên, dư sức phủ 10 năm (~2500 phiên)
            r = s.get(url, params={
                "index": "VNINDEX", "page": page, "size": page_size,
                "fromDate": from_date.isoformat(), "toDate": to_date.isoformat(),
            }, timeout=20)
            r.raise_for_status()
            payload = r.json()
            if not payload.get("successful"):
                break
            content = payload.get("data", {}).get("content", [])
            if not content:
                break
            for row in content:
                d, c = row.get("tradingDate"), row.get("closeIndex")
                if d and c is not None:
                    points[d] = c
            if len(content) < page_size:
                break
            page += 1
        return [{"date": d, "close": points[d]} for d in sorted(points)]
    except Exception as e:
        print(f"  [WARN] VNIndex price history (Vietcap) thất bại: {e}")
        return []


def build_backtest_data(days_back=3650):
    """Ghép chỉ số áp lực/thuận lợi chiêm tinh QUÁ KHỨ với giá đóng cửa VN-Index THẬT cùng khoảng
    thời gian — user (2026-08-04, mở rộng lên 10 năm) muốn tự quan sát trực quan có tương quan hay
    không. Đây là công cụ KIỂM THỬ/tham khảo cá nhân, KHÔNG phải bằng chứng khoa học đã kiểm chứng.
    days_back=3650 (~10 năm) — Vietcap IQ có dữ liệu VNINDEX từ 2016 trở đi cho endpoint này."""
    return {
        "astro": astro.build_historical_pressure_series(days_back=days_back),
        "vnindex": fetch_vnindex_price_history(days_back=days_back + 5),
    }


def build_astro_data():
    positions_raw = astro.get_planet_positions()
    signs = {name: astro.get_zodiac_sign(lon) for name, lon in positions_raw.items()}
    retro = {name: astro.is_retrograde(name) for name in positions_raw}
    house_cusps = astro.get_house_cusps()
    ascendant, midheaven = astro.get_ascendant_mc()

    positions = [
        {"name": name, "lon": lon, "sign": signs[name][0], "degInSign": signs[name][1],
         "retrograde": retro[name], "house": astro.get_house_of(lon, house_cusps)}
        for name, lon in positions_raw.items()
    ]
    current_aspects = astro.find_current_aspects(positions_raw, orb=3.0)
    # only_hard=False (SỬA user 2026-08-04): trước đây chỉ lấy góc chiếu cứng khiến bảng dự báo
    # toàn màu căng thẳng suốt 2 tháng, không có đối trọng — giờ lấy CẢ góc chiếu mềm (tam hợp/lục
    # phân) để bức tranh cân bằng, đúng thực tế hơn (không phải lúc nào cũng chỉ có căng thẳng).
    upcoming_aspects = astro.find_upcoming_exact_aspects(days_ahead=FORECAST_DAYS_AHEAD, step_hours=6, only_hard=False)
    upcoming_eclipses = astro.find_upcoming_eclipses(months_ahead=12)
    retro_stations = astro.find_retrograde_stations(days_ahead=FORECAST_DAYS_AHEAD, step_hours=12)
    daily_pressure = astro.build_daily_pressure_series(days_ahead=FORECAST_DAYS_AHEAD)

    out = {
        "generatedAt": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "positions": positions,
        "currentAspects": current_aspects,
        "upcomingAspects": upcoming_aspects,
        "upcomingEclipses": upcoming_eclipses,
        "retroStations": retro_stations,
        "dailyPressure": daily_pressure,
        "marketVerdict": build_market_verdict(daily_pressure),
        "houses": {
            "locationName": astro.HOUSE_LOCATION_NAME, "ascendant": round(ascendant, 2),
            "midheaven": round(midheaven, 2), "cusps": [round(c, 2) for c in house_cusps],
            "meanings": HOUSE_MEANING,
        },
        "currentAssessment": build_current_assessment(positions, current_aspects),
        "forecastTimeline": build_forecast_timeline(upcoming_aspects, upcoming_eclipses, retro_stations),
        "backtest": build_backtest_data(days_back=3650),
    }

    json_path = os.path.join(PROJECT_ROOT, "data", "astro.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON: {json_path}")
    print(f"  -> {len(out['positions'])} hành tinh, {len(out['currentAspects'])} aspect hiện tại, "
          f"{len(out['upcomingAspects'])} aspect sắp tới, {len(out['upcomingEclipses'])} nhật/nguyệt thực, "
          f"{len(out['retroStations'])} station, {len(out['forecastTimeline'])} sự kiện trong dòng thời gian dự báo, "
          f"backtest: {len(out['backtest']['astro'])} điểm chiêm tinh / {len(out['backtest']['vnindex'])} điểm VN-Index")
    return json_path


if __name__ == "__main__":
    build_astro_data()
