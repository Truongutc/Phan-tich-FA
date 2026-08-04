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

import fetch_astro_data as astro

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

FORECAST_DAYS_AHEAD = 90

# Ý nghĩa hành tinh trong chiêm tinh tài chính — khớp nội dung đã viết ở Buổi 1 (chiemtinh.html).
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

    return {"overallText": overall, "retrogradeLines": retrograde_lines, "aspectLines": aspect_lines}


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
        })

    events.sort(key=lambda e: e["date"])
    return events


def build_astro_data():
    positions_raw = astro.get_planet_positions()
    signs = {name: astro.get_zodiac_sign(lon) for name, lon in positions_raw.items()}
    retro = {name: astro.is_retrograde(name) for name in positions_raw}

    positions = [
        {"name": name, "lon": lon, "sign": signs[name][0], "degInSign": signs[name][1],
         "retrograde": retro[name]}
        for name, lon in positions_raw.items()
    ]
    current_aspects = astro.find_current_aspects(positions_raw, orb=3.0)
    upcoming_aspects = astro.find_upcoming_exact_aspects(days_ahead=FORECAST_DAYS_AHEAD, step_hours=6)
    upcoming_eclipses = astro.find_upcoming_eclipses(months_ahead=12)
    retro_stations = astro.find_retrograde_stations(days_ahead=FORECAST_DAYS_AHEAD, step_hours=12)

    out = {
        "generatedAt": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "positions": positions,
        "currentAspects": current_aspects,
        "upcomingAspects": upcoming_aspects,
        "upcomingEclipses": upcoming_eclipses,
        "retroStations": retro_stations,
        "currentAssessment": build_current_assessment(positions, current_aspects),
        "forecastTimeline": build_forecast_timeline(upcoming_aspects, upcoming_eclipses, retro_stations),
    }

    json_path = os.path.join(PROJECT_ROOT, "data", "astro.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON: {json_path}")
    print(f"  -> {len(out['positions'])} hành tinh, {len(out['currentAspects'])} aspect hiện tại, "
          f"{len(out['upcomingAspects'])} aspect sắp tới, {len(out['upcomingEclipses'])} nhật/nguyệt thực, "
          f"{len(out['retroStations'])} station, {len(out['forecastTimeline'])} sự kiện trong dòng thời gian dự báo")
    return json_path


if __name__ == "__main__":
    build_astro_data()
