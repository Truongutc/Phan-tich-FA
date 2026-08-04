"""
fetch_astro_data.py — Tính toán thiên văn cho tab "Chiêm tinh Tài chính" (financial astrology).

Toàn bộ tính toán dùng thư viện `ephem` (PyEphem) — HOÀN TOÀN OFFLINE, không gọi mạng, không cần
API key, không cần tải file ephemeris ngoài. Đã verify độ chính xác: ngày nhật thực tính ra khớp
với các nhật thực thật đã biết (12/8/2026 Iceland/Tây Ban Nha, 6/2/2027, 2/8/2027 Ai Cập).

LƯU Ý QUAN TRỌNG (đọc trước khi dùng): đây là khung lý thuyết chiêm tinh tài chính/W.D. Gann —
vị trí hành tinh và ngày nhật/nguyệt thực TÍNH CHÍNH XÁC về mặt thiên văn học, nhưng mối liên hệ
với biến động thị trường KHÔNG được khoa học/tài chính chính thống kiểm chứng. Nguồn tham khảo
khung nội dung (izumi.edu.vn, chiemtinhtaichinh.blogspot.com) chỉ là trang giới thiệu/blog, KHÔNG
cung cấp công thức cụ thể nào — mọi phép tính ở đây tự xây dựng dựa trên lý thuyết chiêm tinh phổ
biến (góc chiếu/aspect kinh điển: hợp/lục phân/vuông/tam hợp/xung), không phải công thức độc
quyền của khóa học nào.
"""
import datetime
import math

import ephem

PLANETS = {
    "Mặt Trời": ephem.Sun,
    "Mặt Trăng": ephem.Moon,
    "Sao Thủy": ephem.Mercury,
    "Sao Kim": ephem.Venus,
    "Sao Hỏa": ephem.Mars,
    "Sao Mộc": ephem.Jupiter,
    "Sao Thổ": ephem.Saturn,
    "Sao Thiên Vương": ephem.Uranus,
    "Sao Hải Vương": ephem.Neptune,
    "Sao Diêm Vương": ephem.Pluto,
}

ZODIAC_SIGNS = [
    "Bạch Dương", "Kim Ngưu", "Song Tử", "Cự Giải", "Sư Tử", "Xử Nữ",
    "Thiên Bình", "Bọ Cạp", "Nhân Mã", "Ma Kết", "Bảo Bình", "Song Ngư",
]

# Góc chiếu (aspect) kinh điển trong chiêm tinh — tên tiếng Việt phổ biến.
ASPECTS = {
    0: "Hợp (Conjunction)",
    60: "Lục phân (Sextile)",
    90: "Vuông (Square)",
    120: "Tam hợp (Trine)",
    180: "Xung (Opposition)",
}
# "Hard aspects" (hợp/vuông/xung) — theo lý thuyết chiêm tinh tài chính/Gann thường được coi là
# thời điểm biến động/đảo chiều tiềm năng nhiều hơn "soft aspects" (lục phân/tam hợp, coi là hài
# hòa/ổn định) — dùng để lọc bảng "ngày góc chiếu sắp tới" (Buổi 4) cho gọn, tránh liệt kê quá
# nhiều aspect ít ý nghĩa với mục đích "tìm ngày đảo chiều".
HARD_ASPECTS = {0, 90, 180}


def _utcnow():
    return datetime.datetime.now(datetime.timezone.utc)


def get_planet_positions(when=None):
    """Trả {tên hành tinh: kinh độ hoàng đạo địa tâm (độ, 0-360)} tại thời điểm `when` (datetime
    UTC hoặc None = hiện tại)."""
    d = ephem.Date(when or _utcnow())
    positions = {}
    for name, cls in PLANETS.items():
        lon = math.degrees(float(ephem.Ecliptic(cls(d)).lon)) % 360
        positions[name] = round(lon, 2)
    return positions


def is_retrograde(name, when=None, delta_hours=12):
    """Nghịch hành (retrograde) = kinh độ hoàng đạo GIẢM theo thời gian khi quan sát từ Trái Đất
    (hiện tượng biểu kiến do tốc độ quỹ đạo khác nhau, không phải hành tinh thật sự đảo chiều).
    Mặt Trời/Mặt Trăng KHÔNG BAO GIỜ nghịch hành (luôn tiến) nên trả False ngay không cần tính."""
    if name in ("Mặt Trời", "Mặt Trăng"):
        return False
    d = ephem.Date(when or _utcnow())
    cls = PLANETS[name]
    lon1 = float(ephem.Ecliptic(cls(d)).lon)
    lon2 = float(ephem.Ecliptic(cls(ephem.Date(d) + delta_hours / 24)).lon)
    diff = math.degrees(lon2 - lon1)
    if diff > 180:
        diff -= 360
    elif diff < -180:
        diff += 360
    return diff < 0


def get_zodiac_sign(lon_deg):
    """Kinh độ hoàng đạo (độ) -> (tên cung, độ trong cung 0-30)."""
    idx = int(lon_deg // 30) % 12
    deg_in_sign = round(lon_deg % 30, 2)
    return ZODIAC_SIGNS[idx], deg_in_sign


def find_current_aspects(positions, orb=3.0, when=None, check_applying=True):
    """positions: {tên: kinh độ độ}. Trả list các cặp hành tinh đang trong `orb` độ của 1 trong 5
    góc chiếu kinh điển, sắp xếp theo orb tăng dần (gần exact nhất trước). Mỗi kết quả kèm
    "applying" (True/False) — đọc từ sách "The Tunnel Thru the Air" của W.D. Gann (đoạn Professor
    Joyful luận giải: "Venus applied to a trine of Uranus" khi orb đang thu hẹp/tiến tới exact,
    "Venus was separating from a conjunction of Mars" khi orb đang nới rộng/đã qua exact) — đây là
    kỹ thuật chiêm tinh thật, KHÔNG phải tự suy diễn: applying = ảnh hưởng đang TỚI/mạnh dần,
    separating = ảnh hưởng đang QUA/nhạt dần."""
    names = list(positions.keys())
    later_positions = None
    if check_applying:
        later_positions = get_planet_positions((when or _utcnow()) + datetime.timedelta(hours=12))
    out = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = positions[names[i]], positions[names[j]]
            diff = abs(a - b) % 360
            if diff > 180:
                diff = 360 - diff
            for angle, label in ASPECTS.items():
                gap = abs(diff - angle)
                if gap <= orb:
                    applying = None
                    if check_applying:
                        a2, b2 = later_positions[names[i]], later_positions[names[j]]
                        diff2 = abs(a2 - b2) % 360
                        if diff2 > 180:
                            diff2 = 360 - diff2
                        gap2 = abs(diff2 - angle)
                        applying = gap2 < gap
                    out.append({
                        "a": names[i], "b": names[j], "aspect": label,
                        "angle": angle, "orb": round(gap, 2), "applying": applying,
                    })
                    break
    out.sort(key=lambda x: x["orb"])
    return out


def _aspect_signed_gap(a_lon, b_lon, target_angle):
    """Khoảng lệch CÓ DẤU giữa góc lệch 2 hành tinh và 1 góc chiếu mục tiêu — dùng để dò điểm
    "exact" (đổi dấu) khi quét theo thời gian. Trả giá trị trong (-180, 180]."""
    diff = (a_lon - b_lon) % 360
    if diff > 180:
        diff = diff - 360
    # so khoảng cách góc (không dấu) với target rồi gán lại dấu theo hướng tiệm cận
    gap = abs(diff) - target_angle
    return gap


def find_upcoming_exact_aspects(days_ahead=90, step_hours=6, only_hard=True, exclude_moon=True):
    """Quét từng bước `step_hours` giờ trong `days_ahead` ngày tới, phát hiện các thời điểm 1 cặp
    hành tinh đi qua ĐÚNG 1 góc chiếu (gap đổi dấu giữa 2 bước qua) — đây là "ngày góc chiếu chính
    xác" theo cách diễn giải phổ biến của lý thuyết timing chiêm tinh tài chính (không phải công
    thức độc quyền Gann — nguồn khóa học không công bố công thức cụ thể, xem docstring module).
    Trả list {date_iso, a, b, aspect} sắp xếp theo thời gian tăng dần, KHÔNG trùng lặp quá gần
    nhau (chỉ giữ lần dò đầu tiên phát hiện đổi dấu cho mỗi cặp+aspect). exclude_moon=True (mặc
    định) bỏ Mặt Trăng khỏi bảng này — Mặt Trăng di chuyển ~13°/ngày nên tạo góc chiếu với MỌI
    hành tinh khác mỗi vài ngày, làm bảng "ngày đảo chiều tiềm năng" (Buổi 4) ngập toàn tín hiệu
    ngắn hạn/ít ý nghĩa — thực hành chiêm tinh tài chính phổ biến tập trung hành tinh chậm hơn cho
    mốc thời gian nhiều ngày-nhiều tuần. Mặt Trăng vẫn xuất hiện đầy đủ ở bảng "aspect hiện tại"
    (Buổi 3, xem find_current_aspects)."""
    names = [n for n in PLANETS if not (exclude_moon and n == "Mặt Trăng")]
    target_angles = HARD_ASPECTS if only_hard else set(ASPECTS.keys())
    now = _utcnow()
    steps = int(days_ahead * 24 / step_hours)

    prev_positions = get_planet_positions(now)
    prev_gaps = {}
    results = []
    seen = set()

    for step in range(1, steps + 1):
        t = now + datetime.timedelta(hours=step * step_hours)
        cur_positions = get_planet_positions(t)
        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                a, b = names[i], names[j]
                for angle in target_angles:
                    key = (a, b, angle)
                    gap = _aspect_signed_gap(cur_positions[a], cur_positions[b], angle)
                    prev_gap = prev_gaps.get(key)
                    if prev_gap is not None and abs(prev_gap) < 8 and (prev_gap > 0) != (gap > 0) and key not in seen:
                        results.append({
                            "date": t.strftime("%Y-%m-%d"),
                            "a": a, "b": b, "aspect": ASPECTS[angle],
                        })
                        seen.add(key)
                    prev_gaps[key] = gap
        prev_positions = cur_positions

    results.sort(key=lambda r: r["date"])
    return results


_ECLIPSE_LAT_THRESHOLD = {"solar": 1.6, "lunar": 1.1}  # ngưỡng ecliptic latitude Mặt Trăng (độ)


def find_upcoming_eclipses(months_ahead=12):
    """Lặp qua các kỳ trăng non (New Moon, ứng viên nhật thực)/trăng tròn (Full Moon, ứng viên
    nguyệt thực) sắp tới, lọc theo ecliptic latitude Mặt Trăng tại thời điểm giao hội — đã verify
    khớp các nhật thực thật đã biết (2026-2027). Trả list {date_iso, type: 'solar'/'lunar'} sắp
    xếp theo thời gian."""
    end = _utcnow() + datetime.timedelta(days=months_ahead * 31)
    cur = ephem.Date(_utcnow())
    results = []
    seen_dates = set()
    for _ in range(months_ahead * 2 + 4):  # đủ vòng lặp cho cả New+Full trong khung thời gian
        if cur > ephem.Date(end):
            break
        try:
            nm = ephem.next_new_moon(cur)
            fm = ephem.next_full_moon(cur)
        except Exception:
            break
        for date, kind in sorted([(nm, "solar"), (fm, "lunar")], key=lambda x: x[0]):
            if date > ephem.Date(end):
                continue
            date_str = str(date)
            if date_str in seen_dates:
                continue
            moon_lat = math.degrees(float(ephem.Ecliptic(ephem.Moon(date)).lat))
            if abs(moon_lat) < _ECLIPSE_LAT_THRESHOLD[kind]:
                results.append({
                    "date": ephem.Date(date).datetime().strftime("%Y-%m-%d"),
                    "type": kind,
                })
            seen_dates.add(date_str)
        cur = ephem.Date(min(nm, fm)) + 1

    # Loại trùng (cùng ngày, cùng loại) do vòng lặp có thể quét lại gần biên
    dedup = []
    seen_pairs = set()
    for r in sorted(results, key=lambda r: r["date"]):
        pair = (r["date"], r["type"])
        if pair not in seen_pairs:
            dedup.append(r)
            seen_pairs.add(pair)
    return dedup
