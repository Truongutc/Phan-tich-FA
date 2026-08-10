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

# Nhà (House) — KHÁC vị trí hành tinh/cung hoàng đạo ở trên (không phụ thuộc địa điểm): Nhà cần 1
# ĐỊA ĐIỂM cụ thể trên Trái Đất (dựa trên đường chân trời/kinh tuyến tại nơi đó). User (2026-08-04)
# ban đầu chọn Hà Nội; đổi sang TP.HCM (2026-08-09) — nơi đặt trụ sở HOSE (Sở Giao dịch Chứng
# khoán TP.HCM, đơn vị vận hành VN-Index), phù hợp hơn vì các tính toán "Nhà" trên trang dùng để
# đối chiếu với biến động VN-Index. Toạ độ trung tâm Q.1, TP.HCM (khu vực trụ sở HOSE).
HOUSE_LOCATION_NAME = "TP. Hồ Chí Minh"
HOUSE_LAT = 10.7769
HOUSE_LON = 106.7009
# Độ nghiêng trục Trái Đất (obliquity) — xấp xỉ hiện tại, đủ chính xác cho việc tính Nhà (không
# cần độ chính xác cấp giây cung như hiệu chỉnh thiên văn chuyên nghiệp).
OBLIQUITY_DEG = 23.4367


def get_ascendant_mc(when=None):
    """Tính Điểm Mọc (Ascendant) và Thiên Đỉnh (Midheaven/MC) tại HOUSE_LAT/HOUSE_LON — công thức
    chuẩn chiêm tinh học (dùng RAMC = giờ sao địa phương). Đã verify: Ascendant tính ra khớp kinh
    độ hoàng đạo Mặt Trời lúc mặt trời mọc tại Hà Nội (sai lệch <0.6°, do bán kính đĩa Mặt Trời/
    khúc xạ khí quyển — nằm trong dung sai chấp nhận được). Trả (asc_deg, mc_deg)."""
    obs = ephem.Observer()
    obs.lat = str(HOUSE_LAT)
    obs.lon = str(HOUSE_LON)
    obs.elevation = 0
    obs.pressure = 0
    obs.date = ephem.Date(when or _utcnow())
    ramc = math.degrees(float(obs.sidereal_time()))
    eps = math.radians(OBLIQUITY_DEG)
    phi = math.radians(HOUSE_LAT)
    r = math.radians(ramc)
    asc = math.degrees(math.atan2(math.cos(r), -(math.sin(r) * math.cos(eps) + math.tan(phi) * math.sin(eps)))) % 360
    mc = math.degrees(math.atan2(math.tan(r), math.cos(eps))) % 360
    if 90 < ramc < 270:
        mc = (mc + 180) % 360
    return asc, mc


def get_house_cusps(when=None):
    """Trả list 12 độ khởi đầu (cusp) của Nhà 1-12, hệ NHÀ ĐỀU (Equal House — mỗi nhà đúng 30°
    tính từ Ascendant) — chọn hệ này vì tính toán đơn giản, minh bạch, không cần thư viện chuyên
    dụng (khác hệ Placidus phổ biến hơn nhưng đòi hỏi phép lặp số phức tạp, dễ sai nếu tự cài đặt
    lại từ đầu không qua thư viện đã kiểm chứng kỹ). cusps[0] = Nhà 1 = Ascendant."""
    asc, _ = get_ascendant_mc(when)
    return [(asc + i * 30) % 360 for i in range(12)]


def get_house_of(lon_deg, cusps):
    """Hành tinh ở kinh độ hoàng đạo `lon_deg` đang nằm ở NHÀ số mấy (1-12), theo `cusps` (list 12
    độ khởi đầu, xem get_house_cusps)."""
    asc = cusps[0]
    offset = (lon_deg - asc) % 360
    return int(offset // 30) + 1

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
# hòa/ổn định).
HARD_ASPECTS = {0, 90, 180}

# Hành tinh CHẬM — dùng cho điểm áp lực liên tục (build_daily_pressure_series) và đánh giá hiện tại
# (template_astro.py:SLOW_PLANETS) vì quyết định chu kỳ kinh tế LỚN hơn hành tinh nhanh/Mặt Trời-
# Mặt Trăng (đổi vị trí liên tục, chỉ tạo biến động ngày).
SLOW_PLANET_NAMES = ("Sao Mộc", "Sao Thổ", "Sao Thiên Vương", "Sao Hải Vương", "Sao Diêm Vương")


def _utcnow():
    return datetime.datetime.now(datetime.timezone.utc)


def _anchor_noon_utc(base=None):
    """Chuẩn hóa 1 thời điểm về đúng 12:00:00 UTC cùng ngày — dùng làm mốc CỐ ĐỊNH cho chuỗi điểm
    áp lực/thuận lợi theo ngày trong build_daily_pressure_series/build_historical_pressure_series,
    để KHÔNG phụ thuộc giờ chạy thực tế của cron (xem lý do sửa 2026-08-06 tại nơi gọi hàm này)."""
    b = base or _utcnow()
    return datetime.datetime(b.year, b.month, b.day, 12, 0, 0, tzinfo=datetime.timezone.utc)


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


def find_retrograde_stations(days_ahead=90, step_hours=12):
    """Quét phát hiện ngày hành tinh ĐỔI CHIỀU chuyển động biểu kiến — "station retrograde" (bắt
    đầu nghịch hành) hoặc "station direct" (kết thúc nghịch hành, thuận hành trở lại). Đây là các
    mốc thời gian được chiêm tinh tài chính coi là QUAN TRỌNG hơn cả giai đoạn nghịch hành đang
    diễn ra — thời điểm hành tinh "đứng yên" biểu kiến trên bầu trời trong vài ngày. Mặt Trời/Mặt
    Trăng không có nghịch hành nên không quét. Trả list {date, planet, type} sắp theo thời gian."""
    names = [n for n in PLANETS if n not in ("Mặt Trời", "Mặt Trăng")]
    now = _utcnow()
    steps = int(days_ahead * 24 / step_hours)
    results = []
    prev_state = {name: is_retrograde(name, now) for name in names}
    for step in range(1, steps + 1):
        t = now + datetime.timedelta(hours=step * step_hours)
        for name in names:
            cur_state = is_retrograde(name, t)
            if cur_state != prev_state[name]:
                results.append({
                    "date": t.strftime("%Y-%m-%d"),
                    "planet": name,
                    "type": "station_retrograde" if cur_state else "station_direct",
                })
            prev_state[name] = cur_state
    results.sort(key=lambda r: r["date"])
    return results


def build_daily_pressure_series(days_ahead=90):
    """Điểm ÁP LỰC/THUẬN LỢI thị trường LIÊN TỤC theo từng ngày — trả lời trực tiếp câu hỏi user
    (2026-08-04): "giữa 2 mốc sự kiện [rời rạc trong bảng dự báo] thì thị trường ra sao?" — bảng
    sự kiện chỉ đánh dấu ngày CHÍNH XÁC (exact) của góc chiếu/station, nhưng ảnh hưởng thực tế trải
    dài NHIỀU NGÀY quanh mốc đó (orb) — hàm này tính 1 đường LIÊN TỤC thay vì chỉ các điểm rời rạc.

    Mỗi ngày: (a) mỗi hành tinh (trừ Mặt Trời/Mặt Trăng) đang nghịch hành trừ 0.4 điểm (ảnh hưởng
    nhẹ nhưng kéo dài nhiều tuần); (b) mỗi cặp hành tinh (trừ Mặt Trăng — di chuyển quá nhanh, xem
    lý do ở find_upcoming_exact_aspects) đang trong orb ±6° của 1 góc chiếu cộng/trừ điểm có TRỌNG
    SỐ theo orb (gần exact = ảnh hưởng mạnh hơn, giảm tuyến tính về 0 ở biên orb) — hài hòa (tam
    hợp/lục phân) CỘNG điểm, căng thẳng (vuông/xung) TRỪ điểm. Điểm dương = nghiêng thuận lợi, điểm
    âm = nghiêng áp lực/căng thẳng — đây là góc nhìn TỔNG HỢP ĐỊNH LƯỢNG mang tính quyết đoán hơn
    nhãn "tension" trung lập ở bảng sự kiện rời rạc (2 góc nhìn bổ sung cho nhau: bảng sự kiện =
    từng góc chiếu riêng lẻ trung lập, đường điểm này = tổng hợp NHIỀU góc chiếu + nghịch hành cùng
    lúc GIỮA CÙNG 1 TẬP HÀNH TINH với bảng sự kiện — user 2026-08-04 chỉ ra ví dụ Sao Hỏa vuông Sao
    Thổ phải thể hiện được trong đường điểm này, nên DÙNG CHUNG tập hành tinh với
    find_upcoming_exact_aspects (mọi hành tinh trừ Mặt Trăng), KHÔNG giới hạn ở 5 hành tinh chậm
    như bản đầu (lỗi khiến ảnh hưởng Sao Hỏa/Mặt Trời/Sao Thủy/Sao Kim bị bỏ sót hoàn toàn). Trả
    list [{date, score}].

    LƯU Ý (sửa lỗi 2026-08-06): mốc "now" TRƯỚC ĐÂY giữ nguyên giờ/phút/giây lúc script chạy, nên
    "ngày X" ở 2 lần chạy cron khác giờ nhau (VD ~10h UTC tuần trước vs ~1h UTC tuần này) thực chất
    là 2 THỜI ĐIỂM khác nhau trong ngày X — với góc chiếu đang gần đúng orb (sắp exact), chênh lệch
    vài giờ đủ làm điểm số xê dịch nhẹ, đôi khi khiến ĐỈNH của đường dao động "nhảy" sang ngày liền
    kề giữa 2 lần cập nhật (user 2026-08-06 phát hiện: đỉnh báo 12/08 tuần trước (4.347 vs 4.195),
    13/08 tuần này (4.198 vs 4.327) — 2 ngày gần như hòa điểm, chỉ cần lệch giờ tính là lật đỉnh).
    Dùng _anchor_noon_utc() để mốc mỗi ngày LUÔN là đúng 12:00 UTC bất kể cron chạy giờ nào — kết
    quả sẽ ỔN ĐỊNH, tái lập được giữa các lần chạy khác nhau."""
    now = _anchor_noon_utc()
    return [_pressure_score_at(now + datetime.timedelta(days=d)) for d in range(days_ahead + 1)]


def _pressure_score_at(t):
    """Điểm áp lực/thuận lợi tại 1 thời điểm t bất kỳ (quá khứ hay tương lai đều tính được vì ephem
    không giới hạn khoảng thời gian) — logic dùng chung cho build_daily_pressure_series (tương lai)
    và build_historical_pressure_series (quá khứ, để đối chiếu ngược với diễn biến VN-Index thật)."""
    ORB_WINDOW = 6.0
    score_planets = [n for n in PLANETS if n != "Mặt Trăng"]
    positions = get_planet_positions(t)
    retro_penalty = sum(0.4 for name in score_planets if name != "Mặt Trời" and is_retrograde(name, t))

    aspect_score = 0.0
    for i in range(len(score_planets)):
        for j in range(i + 1, len(score_planets)):
            a, b = positions[score_planets[i]], positions[score_planets[j]]
            diff = abs(a - b) % 360
            if diff > 180:
                diff = 360 - diff
            for angle in (0, 60, 90, 120, 180):
                gap = abs(diff - angle)
                if gap <= ORB_WINDOW:
                    weight = (ORB_WINDOW - gap) / ORB_WINDOW
                    if angle in (60, 120):
                        aspect_score += weight
                    elif angle in (90, 180):
                        aspect_score -= weight
                    break

    return {"date": t.strftime("%Y-%m-%d"), "score": round(aspect_score - retro_penalty, 3)}


def build_historical_pressure_series(days_back=365):
    """Đường điểm áp lực/thuận lợi trong QUÁ KHỨ (days_back ngày trở lại đây) — dùng để đối chiếu
    ngược (backtest) với diễn biến VN-Index thật, kiểm tra trực quan mối tương quan (nếu có) giữa
    chỉ số chiêm tinh và biến động thị trường thực tế. KHÔNG phải bằng chứng khoa học, chỉ là công
    cụ quan sát trực quan theo yêu cầu user (2026-08-04). Dùng mốc 12:00 UTC cố định mỗi ngày (xem
    _anchor_noon_utc, lý do tại build_daily_pressure_series) để kết quả ổn định giữa các lần chạy."""
    now = _anchor_noon_utc()
    return [_pressure_score_at(now - datetime.timedelta(days=d)) for d in range(days_back, -1, -1)]
