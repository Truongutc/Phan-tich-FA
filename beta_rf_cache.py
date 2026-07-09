#!/usr/bin/env python3
"""
beta_rf_cache.py — Cơ chế fallback Beta/Rf dùng chung cho template_banking/securities/kcn.py.

Khi fetch Beta (Vietstock/Vietcap) hoặc Rf (worldgovernmentbonds.com) trực tiếp thất bại hoàn toàn
(vd. IP datacenter của GitHub Actions bị các nguồn dữ liệu VN chặn/hạn chế khác với máy chạy local —
cùng nguyên nhân đã gặp với tin sản lượng thép HPG), dùng lại giá trị đã fetch thành công ở lần chạy
gần nhất (lưu trong data/<ticker>.json, field "betaRfCache") thay vì rơi thẳng về hằng số cứng
(Beta=1.0, Rf=4.5%) làm sai lệch COE và kéo theo toàn bộ định giá RI.

RECALCULATE_FRESH=true (input "Tính toán lại mới" trên GitHub Actions) bỏ qua cache này hoàn toàn —
tính lại từ đầu, chấp nhận rơi về hằng số cứng nếu fetch thất bại, đúng ý nghĩa "phân tích mới".
"""
import os
import json
from datetime import datetime


def apply_beta_rf_cache_fallback(ticker, project_root, rf_val, rf_src, beta_val, beta_src,
                                  is_enough_sessions, beta_web):
    """Nhận kết quả fetch Beta/Rf trực tiếp (đã tính xong ở caller), vá lại bằng cache lần chạy trước
    nếu fetch thất bại (rơi về nguồn "Fallback (manual)"/beta_web mặc định) và RECALCULATE_FRESH khác
    "true". Trả về (rf_val, rf_src, beta_val, beta_src, cache_entry) — caller tự lưu cache_entry vào
    key "betaRfCache" của data/<ticker>.json khi ghi file JSON tổng kết."""
    force_fresh = os.environ.get("RECALCULATE_FRESH", "false").strip().lower() == "true"

    cache_path = os.path.join(project_root, "data", f"{ticker}.json")
    prev = None
    if not force_fresh:
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                prev = json.load(f).get("betaRfCache")
        except Exception:
            pass

    rf_is_fresh = rf_src != "Fallback (manual)"
    beta_is_fresh = is_enough_sessions or beta_web != 1.0

    if not rf_is_fresh and prev and prev.get("rf") is not None:
        print(f"  [DIAG] Rf {ticker}: fetch trực tiếp thất bại (rơi về fallback cứng) - dùng lại Rf lần "
              f"chạy trước (lúc {prev.get('fetchedAt', '?')}): {prev['rf']*100:.2f}%")
        rf_val = prev["rf"]
        rf_src = f"Cache lần chạy trước ({prev.get('fetchedAt', '?')})"
    if not beta_is_fresh and prev and prev.get("beta") is not None:
        print(f"  [DIAG] Beta {ticker}: fetch/scrape trực tiếp đều thất bại - dùng lại Beta lần chạy "
              f"trước (lúc {prev.get('fetchedAt', '?')}): {prev['beta']}")
        beta_val = prev["beta"]
        beta_src = f"Cache lần chạy trước ({prev.get('fetchedAt', '?')})"

    cache_entry = {
        "beta": beta_val if beta_is_fresh else (prev.get("beta") if prev else beta_val),
        "rf": rf_val if rf_is_fresh else (prev.get("rf") if prev else rf_val),
        "fetchedAt": datetime.now().strftime("%Y-%m-%d %H:%M") if (rf_is_fresh or beta_is_fresh)
                     else (prev.get("fetchedAt") if prev else None),
    }
    return rf_val, rf_src, beta_val, beta_src, cache_entry
