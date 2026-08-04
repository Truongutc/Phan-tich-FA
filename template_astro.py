"""
template_astro.py — Đóng gói dữ liệu chiêm tinh tài chính (data/astro.json) cho chiemtinh.html.

Xem fetch_astro_data.py để biết chi tiết cách tính + lưu ý quan trọng về bản chất khung lý thuyết
này (chiêm tinh tài chính/W.D. Gann — dữ liệu thiên văn chính xác, mối liên hệ với thị trường
KHÔNG được khoa học chính thống kiểm chứng).
"""
import datetime
import json
import os

import fetch_astro_data as astro

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))


def build_astro_data():
    positions = astro.get_planet_positions()
    signs = {name: astro.get_zodiac_sign(lon) for name, lon in positions.items()}
    retro = {name: astro.is_retrograde(name) for name in positions}

    out = {
        "generatedAt": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "positions": [
            {"name": name, "lon": lon, "sign": signs[name][0], "degInSign": signs[name][1],
             "retrograde": retro[name]}
            for name, lon in positions.items()
        ],
        "currentAspects": astro.find_current_aspects(positions, orb=3.0),
        "upcomingAspects": astro.find_upcoming_exact_aspects(days_ahead=90, step_hours=6),
        "upcomingEclipses": astro.find_upcoming_eclipses(months_ahead=12),
    }

    json_path = os.path.join(PROJECT_ROOT, "data", "astro.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON: {json_path}")
    print(f"  -> {len(out['positions'])} hành tinh, {len(out['currentAspects'])} aspect hiện tại, "
          f"{len(out['upcomingAspects'])} aspect sắp tới, {len(out['upcomingEclipses'])} nhật/nguyệt thực")
    return json_path


if __name__ == "__main__":
    build_astro_data()
