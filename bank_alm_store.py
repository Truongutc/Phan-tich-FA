#!/usr/bin/env python3
"""
bank_alm_store.py — Kho JSON lưu dữ liệu RỦI RO LÃI SUẤT/THANH KHOẢN theo TỪNG NGÂN HÀNG, TỪNG KỲ
(data/bank_alm/<TICKER>.json). Đây là nguồn dữ liệu THÔ (raw inputs) duy nhất — chỉ lưu gap buckets
+ balance sheet snapshot đã trích được, KHÔNG lưu tỷ lệ/chỉ số đã tính (sensitivity_level, stress
scenarios...) — các chỉ số đó LUÔN được tính lại từ đầu bằng compute_interest_rate_risk_metrics/
compute_liquidity_risk_metrics mỗi khi cần dùng, để không có 2 nơi lưu cùng 1 con số có thể lệch
nhau nếu công thức tính sau này thay đổi.

Dùng bởi:
- bank_risk_notes.fetch_bank_risk_gaps_cached() — cache xuyên suốt giữa pipeline phân tích 1 mã lẻ
  (template_banking.py) và pipeline tổng hợp toàn hệ thống (bank_system_risk.py).
- bank_system_risk.py — vòng lặp kiểm tra độ mới (refresh_bank_alm_data) và tổng hợp hệ thống
  (recompute_system_aggregate_all_periods).

Quy ước khóa kỳ (period_key): DÙNG CHUNG định dạng có gạch ngang của vimo ("YYYY-H1"/"YYYY-FY" cho
2 bảng gap chỉ có ở BCTC đã kiểm toán/soát xét, "YYYY-Qn" cho dữ liệu bảng cân đối theo quý không
cần OCR) — KHÔNG dùng định dạng "YYYYQn" không gạch ngang của segments_kcn, vì
_PERIOD_SHEET_PATTERNS trong template_vimo.py đã nhận diện sẵn đúng định dạng có gạch ngang này để
tự xếp đúng sheet Excel (Luy_Ke/Theo_Quy) khi dữ liệu chảy vào cơ chế indicator có sẵn của vimo,
không cần code chuyển đổi gì thêm.
"""
import os
import json
import datetime

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
STORE_DIR = os.path.join(PROJECT_ROOT, "data", "bank_alm")

_VN_TZ = datetime.timezone(datetime.timedelta(hours=7))


def _now_iso():
    return datetime.datetime.now(_VN_TZ).isoformat(timespec="seconds")


def _store_path(ticker):
    return os.path.join(STORE_DIR, f"{ticker.upper()}.json")


def load_bank_store(ticker):
    """Trả về dict store của 1 ticker, hoặc {"ticker": TICKER, "gap_periods": {}, "quarterly_balance_sheet":
    {}, "last_cheap_check": None} nếu chưa có file/lỗi đọc (KHÔNG BAO GIỜ raise)."""
    ticker = ticker.upper()
    path = _store_path(ticker)
    if not os.path.exists(path):
        return {"ticker": ticker, "gap_periods": {}, "quarterly_balance_sheet": {}, "last_cheap_check": None}
    try:
        with open(path, "r", encoding="utf-8") as f:
            store = json.load(f)
        store.setdefault("gap_periods", {})
        store.setdefault("quarterly_balance_sheet", {})
        store.setdefault("last_cheap_check", None)
        return store
    except Exception:
        return {"ticker": ticker, "gap_periods": {}, "quarterly_balance_sheet": {}, "last_cheap_check": None}


def save_bank_store(ticker, store):
    os.makedirs(STORE_DIR, exist_ok=True)
    with open(_store_path(ticker.upper()), "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)


def upsert_reported_period(ticker, period_key, gaps_dict, source):
    """Ghi 1 kỳ ĐÃ CÓ SỐ LIỆU THẬT (status="reported") — gaps_dict là dict trả về từ
    fetch_bank_risk_gaps()/fetch_bank_risk_gaps_for_period() (đơn vị TRIỆU đồng, giữ nguyên đơn vị
    gốc, KHÔNG quy đổi ở đây — nơi dùng dữ liệu tự quy đổi tỷ đồng khi cần, giống cách
    template_banking.py đang làm).

    KHÔNG lưu kèm snapshot bảng cân đối ở đây (cố tình) — số liệu tổng tài sản/VCSH/NII cùng kỳ lấy
    từ `quarterly_balance_sheet` (xem gap_period_to_quarter()) vì đó là nguồn ĐỘC LẬP, luôn cập
    nhật sẵn từ Vietcap (không cần OCR) bất kể lệnh ghi gap đến từ pipeline nào — tránh phải xác
    định "ai chịu trách nhiệm ghi balance_sheet_snapshot" giữa 2 nơi gọi khác nhau
    (template_banking.py cho 1 mã lẻ, bank_system_risk.py cho toàn hệ thống)."""
    store = load_bank_store(ticker)
    store["gap_periods"][period_key] = {
        "status": "reported",
        "patched_from": None,
        "interest_rate_gap": gaps_dict.get("interest_rate_gap"),
        "liquidity_gap": gaps_dict.get("liquidity_gap"),
        "liabilities_by_bucket": gaps_dict.get("liabilities_by_bucket"),
        "interest_rate_sensitivity_disclosed": gaps_dict.get("interest_rate_sensitivity_disclosed"),
        "source": source,
        "fetched_at": _now_iso(),
    }
    save_bank_store(ticker, store)
    return store["gap_periods"][period_key]


def gap_period_to_quarter(period_key):
    """Ánh xạ kỳ gap ("YYYY-H1"/"YYYY-FY") sang khóa quý tương ứng trong quarterly_balance_sheet
    ("YYYY-Qn") để lấy tổng tài sản/VCSH/NII/tiền gửi CÙNG THỜI ĐIỂM — bán niên = hết quý 2, cả năm
    = hết quý 4."""
    year_str, suffix = period_key.split("-", 1)
    q = {"H1": 2, "FY": 4}.get(suffix)
    return f"{year_str}-Q{q}" if q else None


def upsert_patched_period(ticker, target_period_key, source_period_key):
    """Ghi 1 kỳ CHƯA CÓ SỐ LIỆU THẬT (ngân hàng chưa công bố) bằng cách "vá" (carry-forward) dữ liệu
    từ 1 kỳ CŨ HƠN đã có trong store — đánh dấu RÕ status="patched" + patched_from=nguồn, KHÔNG lặng
    lẽ trộn vào coi như số liệu thật. Trả về False (không ghi gì) nếu source_period_key chưa tồn tại
    trong store (không có gì để vá từ đó)."""
    store = load_bank_store(ticker)
    source_entry = store["gap_periods"].get(source_period_key)
    if not source_entry or source_entry.get("status") == "missing":
        return False
    patched = dict(source_entry)
    patched["status"] = "patched"
    patched["patched_from"] = source_period_key
    patched["fetched_at"] = _now_iso()
    store["gap_periods"][target_period_key] = patched
    save_bank_store(ticker, store)
    return True


def mark_missing_period(ticker, period_key):
    """Ghi nhận RÕ RÀNG rằng kỳ này không có số liệu (chưa niêm yết lúc đó, hoặc OCR thất bại liên
    tục) — status="missing" thay vì để trống im lặng, để phần tổng hợp hệ thống biết chính xác ngân
    hàng nào đang thiếu ở kỳ nào (yêu cầu kiểm soát dữ liệu của user)."""
    store = load_bank_store(ticker)
    store["gap_periods"][period_key] = {
        "status": "missing", "patched_from": None,
        "interest_rate_gap": None, "liquidity_gap": None, "liabilities_by_bucket": None,
        "interest_rate_sensitivity_disclosed": None, "balance_sheet_snapshot": None,
        "source": None, "fetched_at": _now_iso(),
    }
    save_bank_store(ticker, store)


def upsert_quarterly_balance_sheet(ticker, quarter_key, snapshot_dict):
    """Ghi đè snapshot bảng cân đối 1 quý (không cần OCR, luôn ghi đè mới nhất từ Vietcap mỗi lần
    gọi — không cần theo dõi độ mới riêng, .cache/{TICKER}_bctc.json của fetch_data.py đã lo việc
    đó)."""
    store = load_bank_store(ticker)
    store["quarterly_balance_sheet"][quarter_key] = snapshot_dict
    save_bank_store(ticker, store)


def record_cheap_check(ticker, period_found):
    store = load_bank_store(ticker)
    store["last_cheap_check"] = {"period_found": period_found, "checked_at": _now_iso()}
    save_bank_store(ticker, store)


_PERIOD_SORT_SUFFIX = {"H1": 1, "FY": 2}


def _period_sort_key(period_key):
    year_str, suffix = period_key.split("-", 1)
    return (int(year_str), _PERIOD_SORT_SUFFIX.get(suffix, 0))


def latest_reported_period(ticker):
    """Kỳ gap MỚI NHẤT trong store có status="reported" (bỏ qua "patched"/"missing") — dùng làm
    nguồn "vá" khi 1 ngân hàng chưa công bố kỳ mới nhất của hệ thống. Trả về None nếu store chưa có
    kỳ "reported" nào."""
    store = load_bank_store(ticker)
    reported = [pk for pk, entry in store["gap_periods"].items() if entry.get("status") == "reported"]
    if not reported:
        return None
    return max(reported, key=_period_sort_key)


def get_period_entry(ticker, period_key):
    """Trả về dict entry (hoặc None) của 1 kỳ cụ thể trong store — tiện dùng trực tiếp không cần
    load cả store."""
    store = load_bank_store(ticker)
    return store["gap_periods"].get(period_key)
