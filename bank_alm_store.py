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

Quy ước khóa kỳ (period_key) trong gap_periods (2026-08-31, sau khi xác nhận qua ảnh chụp BCTC quý
thường thật MBB/TCB rằng 2 bảng gap CŨNG có ở báo cáo quý thường, không chỉ ở bản kiểm toán/soát
xét): "YYYY-Qn" (n=1-4) cho MỌI báo cáo quý thường, "YYYY-FY" CHỈ cho báo cáo năm đã kiểm toán —
KHÔNG BAO GIỜ dùng "YYYY-H1" nữa. Một ngân hàng có thể có ĐỒNG THỜI cả "YYYY-Q4" (báo cáo quý
thường, công bố sớm) VÀ "YYYY-FY" (báo cáo năm kiểm toán, công bố sau, đáng tin hơn) cho CÙNG 1
thời điểm cuối năm — đây là 2 LẦN CÔNG BỐ khác nhau nên KHÔNG ghi đè lên nhau trong dict này; việc
gộp chúng lại thành 1 kỳ hệ thống duy nhất khi tổng hợp toàn ngành nằm ở
bank_system_risk._ticker_gap_entry(), không phải ở đây. quarterly_balance_sheet LUÔN dùng "YYYY-Qn"
(không bao giờ "-FY", vì đây là dữ liệu Vietcap theo quý, không phân biệt kiểm toán/soát xét).

Định dạng có gạch ngang này khớp SẴN với _PERIOD_SHEET_PATTERNS trong template_vimo.py (không phải
"YYYYQn" không gạch ngang của segments_kcn) để tự xếp đúng sheet Excel khi dữ liệu chảy vào cơ chế
indicator có sẵn của vimo — không cần code chuyển đổi gì thêm.
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


def _empty_store(ticker, load_error=None):
    st = {"ticker": ticker, "gap_periods": {}, "quarterly_balance_sheet": {}, "last_cheap_check": None}
    if load_error:
        st["_load_error"] = str(load_error)
    return st


def load_bank_store(ticker):
    """Tra ve dict store cua 1 ticker. File CHUA TON TAI -> store rong hop le. Loi DOC that su (file
    bi khoa, dang duoc ghi do, JSON hong...) -> thu lai vai lan; van loi thi tra ve store rong co co
    "_load_error" - save_bank_store() TU CHOI ghi store nay (xem duoi). KHONG BAO GIO raise.

    SUA 2026-09-26 (mat du lieu that: chay fetch_macro_data.py lam 3 file LPB/KLB/MBB mat het
    gap_periods): ban cu nuot MOI loi doc roi tra ve store rong nhu "chua co file" - buoc cap nhat khac
    (record_cheap_check/upsert_quarterly_balance_sheet) sau do load->sua->save GHI DE file that bang
    store rong, xoa sach du lieu da doi chieu thu cong."""
    import time
    ticker = ticker.upper()
    path = _store_path(ticker)
    if not os.path.exists(path):
        return _empty_store(ticker)
    last_err = None
    for attempt in range(4):
        try:
            with open(path, "r", encoding="utf-8") as f:
                store = json.load(f)
            store.setdefault("gap_periods", {})
            store.setdefault("quarterly_balance_sheet", {})
            store.setdefault("last_cheap_check", None)
            return store
        except Exception as e:
            last_err = e
            time.sleep(0.3 * (attempt + 1))
    print(f"  [ERROR] bank_alm_store: khong doc duoc {path} ({last_err}) - KHONG ghi de file nay")
    return _empty_store(ticker, load_error=last_err)


def save_bank_store(ticker, store):
    """Ghi store (atomic: ghi file tam roi thay the, tranh file nua chung). TU CHOI ghi khi (a) store
    den tu 1 lan doc loi ("_load_error"), (b) ghi lam MAT >50% gap_periods so voi file dang co tren dia
    (dau hieu ghi de bang store rong/hong) - de ghi co chu dich thi truyen allow_shrink qua
    store["_allow_shrink"] = True truoc khi luu."""
    ticker = ticker.upper()
    if store.get("_load_error"):
        raise RuntimeError(f"tu choi ghi {ticker}: store duoc tao tu lan doc loi ({store['_load_error']})")
    allow_shrink = store.pop("_allow_shrink", False)
    path = _store_path(ticker)
    new_n = len((store.get("gap_periods") or {}))
    if os.path.exists(path) and not allow_shrink:
        try:
            with open(path, "r", encoding="utf-8") as f:
                old_n = len((json.load(f).get("gap_periods") or {}))
        except Exception:
            old_n = 0
        if old_n >= 4 and new_n < old_n / 2:
            raise RuntimeError(f"tu choi ghi {ticker}: gap_periods giam {old_n} -> {new_n} (nghi ghi de bang store rong)")
    os.makedirs(STORE_DIR, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def upsert_reported_period(ticker, period_key, gaps_dict, source):
    """Ghi 1 kỳ ĐÃ CÓ SỐ LIỆU THẬT (status="reported") — gaps_dict là dict trả về từ
    fetch_bank_risk_gaps()/fetch_bank_risk_gaps_for_period() (đơn vị TRIỆU đồng, giữ nguyên đơn vị
    gốc, KHÔNG quy đổi ở đây — nơi dùng dữ liệu tự quy đổi tỷ đồng khi cần, giống cách
    template_banking.py đang làm).

    KHÔNG lưu kèm snapshot bảng cân đối ở đây (cố tình) — số liệu tổng tài sản/VCSH/NII cùng kỳ lấy
    từ `quarterly_balance_sheet` (xem gap_period_to_quarter()) vì đó là nguồn ĐỘC LẬP, luôn cập
    nhật sẵn từ Vietcap (không cần OCR) bất kể lệnh ghi gap đến từ pipeline nào — tránh phải xác
    định "ai chịu trách nhiệm ghi balance_sheet_snapshot" giữa 2 nơi gọi khác nhau
    (template_banking.py cho 1 mã lẻ, bank_system_risk.py cho toàn hệ thống).

    SỬA (user 2026-09-18): MỌI field OCR đều ưu tiên giá trị MỚI, nhưng GIỮ LẠI giá trị CŨ nếu lần
    này không trích được (thay vì luôn ghi đè bằng None) — bug thật tự phát hiện 2 lần liên tiếp:
    (1) fx_position/fx_source (ghi RIÊNG qua upsert_fx_position(), hàm này trước đây không biết tới
    2 field đó, OCR lại sẽ âm thầm xóa mất FX đã nhập); (2) VIB/VAB — nhiều kỳ có status="reported"
    (báo "đã xong") nhưng CHỈ interest_rate_gap thành công, liquidity_gap ÂM THẦM là None — nếu sau
    này sửa fix rồi backfill_period() bỏ qua các kỳ này (coi status="reported" là XONG, không thử
    lại), sẽ KHÔNG BAO GIỜ tự lấy được phần thiếu (xem is_fully_reported() ở dưới, dùng để quyết định
    có bỏ qua hay thử lại). Giờ MỌI lần OCR lại chỉ có thể THÊM/CẢI THIỆN dữ liệu, không bao giờ làm
    MẤT field đã trích được trước đó chỉ vì lần này không trích lại được field đó."""
    store = load_bank_store(ticker)
    existing = store["gap_periods"].get(period_key) or {}

    def _merge(key):
        new_val = gaps_dict.get(key)
        return new_val if new_val is not None else existing.get(key)

    # Chot kiem tra hop ly 2026-09 (xem bank_alm_validate.py): KHONG chan ghi (van luu de con so de xem lai)
    # nhung gan `validation` + dua ky loi vao hang doi xem lai, de loi doc nham khong am tham nam trong so
    # lieu he thong nhu truoc day.
    validation = None
    try:
        import bank_alm_validate as _v
        _cand = {
            "interest_rate_gap": _merge("interest_rate_gap"),
            "interest_rate_liabilities_by_bucket": _merge("interest_rate_liabilities_by_bucket"),
            "liquidity_gap": _merge("liquidity_gap"),
            "liabilities_by_bucket": _merge("liabilities_by_bucket"),
        }
        _snap = (store.get("quarterly_balance_sheet") or {}).get(gap_period_to_quarter(period_key)) or {}
        _res = _v.validate_entry(_cand, _snap.get("total_assets"))
        validation = {"ok": _res["ok"], "level": _res["level"], "issues": [i["code"] for i in _res["issues"]]}
        _v.queue_review(ticker, period_key, _res, source)
        if not _res["ok"]:
            print(f"  [VALIDATE] {ticker} {period_key}: NGHI DOC SAI - " +
                  "; ".join(i["msg"] for i in _res["issues"] if i["level"] == "fail"))
            # Ky da duoc DOI CHIEU THU CONG voi PDF (verified) thi OCR moi doc sai KHONG duoc ghi de
            if existing.get("verified"):
                print(f"  [KEEP] {ticker} {period_key}: giu so lieu da doi chieu thu cong "
                      f"(verified={existing.get('verified')!r}), bo ket qua OCR moi vi khong qua kiem tra")
                return existing
    except Exception as _e:
        print(f"  [WARN] validate {ticker} {period_key}: {_e}")

    store["gap_periods"][period_key] = {
        "status": "reported",
        "patched_from": None,
        "validation": validation,
        "verified": None,
        "interest_rate_gap": _merge("interest_rate_gap"),
        "liquidity_gap": _merge("liquidity_gap"),
        "liabilities_by_bucket": _merge("liabilities_by_bucket"),
        "interest_rate_liabilities_by_bucket": _merge("interest_rate_liabilities_by_bucket"),
        "interest_rate_sensitivity_disclosed": _merge("interest_rate_sensitivity_disclosed"),
        # 3 dong chi tiet bo sung tu bang thanh khoan (user 2026-09-18, xem bank_risk_notes.py
        # _ROW_LABELS_CASH/_SBV_DEP/_CUST_DEP) — CHUA CHAC luon trich duoc (chi khop tuyen tinh,
        # khong co tang du phong), None neu khong trich duoc thay vi doan.
        "cash_by_bucket": _merge("cash_by_bucket"),
        "sbv_dep_by_bucket": _merge("sbv_dep_by_bucket"),
        "customer_deposits_by_bucket": _merge("customer_deposits_by_bucket"),
        "fx_position": _merge("fx_position"),
        "fx_source": _merge("fx_source"),
        "source": source,
        "fetched_at": _now_iso(),
    }
    save_bank_store(ticker, store)
    return store["gap_periods"][period_key]


def is_fully_reported(entry):
    """Kiểm tra entry có THẬT SỰ đầy đủ CẢ 2 bảng gap (lãi suất + thanh khoản) hay chỉ 1 trong 2 —
    status="reported" KHÔNG đảm bảo cả 2 đều có (bug thật phát hiện 2026-09-18 qua VIB/VAB: bảng lãi
    suất trích thành công trong khi bảng thanh khoản ÂM THẦM thất bại, entry vẫn được đánh dấu
    "reported" như thể đã xong hoàn toàn). Dùng để quyết định "bỏ qua vì đã xong" hay "vẫn nên thử
    lại" khi backfill — CHỈ coi là xong khi CẢ 2 đều có, nếu chỉ có 1 thì vẫn còn cơ hội lấy nốt phần
    thiếu (đặc biệt hữu ích sau khi sửa 1 fix OCR — tự động thử lại đúng những kỳ còn dở dang thay vì
    mãi mãi bị coi là "đã xong")."""
    if not entry or entry.get("status") != "reported":
        return False
    return bool(entry.get("interest_rate_gap")) and bool(entry.get("liquidity_gap"))


def gap_period_to_quarter(period_key):
    """Ánh xạ kỳ gap ("YYYY-Qn"/"YYYY-FY") sang khóa quý tương ứng trong quarterly_balance_sheet
    ("YYYY-Qn") để lấy tổng tài sản/VCSH/NII/tiền gửi CÙNG THỜI ĐIỂM. Từ 2026-08 (sau khi xác nhận
    BCTC quý thường CŨNG có đủ 2 bảng gap — xem bank_risk_notes.py _period_key_for_candidate), khóa
    kỳ gap dùng CHUNG định dạng "YYYY-Qn" với bảng cân đối theo quý nên hầu hết đã KHỚP SẴN, chỉ
    "YYYY-FY" (báo cáo năm đã kiểm toán, không có "Qn" tương ứng) mới cần ánh xạ riêng sang quý 4
    (hết năm = hết quý 4)."""
    year_str, suffix = period_key.split("-", 1)
    if suffix == "FY":
        return f"{year_str}-Q4"
    if suffix.startswith("Q") and suffix[1:].isdigit():
        return period_key
    return None


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


def _period_sort_key(period_key):
    """Sắp xếp "YYYY-Qn" theo đúng thứ tự thời gian (Q1<Q2<Q3<Q4) — "YYYY-FY" (báo cáo năm đã kiểm
    toán) xếp SAU Q4 CÙNG NĂM vì luôn công bố sau, dù cùng phản ánh thời điểm cuối năm."""
    year_str, suffix = period_key.split("-", 1)
    year = int(year_str)
    if suffix == "FY":
        return (year, 5)
    if suffix.startswith("Q") and suffix[1:].isdigit():
        return (year, int(suffix[1:]))
    return (year, 0)


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


def upsert_fx_position(ticker, period_key, fx_position, source=None):
    """Ghi/cập nhật trường fx_position (trạng thái ngoại tệ theo từng đồng tiền — xem
    "Danh gia rui ro tien te.docx", user 2026-09-18) cho 1 entry gap_periods ĐÃ CÓ SẴN — KHÔNG tạo
    entry mới nếu chưa có (dữ liệu FX luôn đi kèm cùng 1 báo cáo với bảng lãi suất/thanh khoản, nên kỳ
    đó phải đã tồn tại trước). fx_position là dict {ma_tien: {"assets", "liabilities", "onbalance",
    "offbalance", "net"}} theo đơn vị TRIỆU đồng (giữ nguyên đơn vị gốc, giống quy ước liquidity_gap/
    interest_rate_gap — nơi dùng tự quy đổi tỷ đồng khi cần). Trả về False nếu entry chưa tồn tại
    (gọi nơi dùng nên backfill/nhập bảng lãi suất-thanh khoản trước)."""
    store = load_bank_store(ticker)
    entry = store["gap_periods"].get(period_key)
    if not entry:
        return False
    entry["fx_position"] = fx_position
    if source:
        entry["fx_source"] = source
    save_bank_store(ticker, store)
    return True
