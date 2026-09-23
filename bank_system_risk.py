#!/usr/bin/env python3
"""
bank_system_risk.py — Tổng hợp rủi ro lãi suất/thanh khoản TOÀN HỆ THỐNG ngân hàng niêm yết/UPCoM
(26 mã, xem bank_universe.py), dùng cho phần "Rủi ro hệ thống ngân hàng" trong báo cáo Vĩ mô.

Kiến trúc (xem plan đã duyệt 2026-08-30, sửa lại 2026-08-31 sau khi xác nhận qua ảnh chụp thật
MBB/TCB rằng 2 bảng gap CŨNG có ở BCTC quý thường, không chỉ ở bản kiểm toán/soát xét):
- Dữ liệu THÔ (gap buckets từ 2 bảng thuyết minh BCTC + snapshot bảng cân đối) lưu RIÊNG TỪNG NGÂN
  HÀNG, TỪNG KỲ trong data/bank_alm/<TICKER>.json (bank_alm_store.py) — KHÔNG lưu số liệu tổng hợp
  hệ thống dạng lịch sử riêng, để tránh 2 nơi lưu cùng 1 con số có thể lệch nhau. Số liệu hệ thống
  LUÔN được tính lại (recompute_system_aggregate_all_periods) từ các file per-bank này.
- 2 bảng gap (lãi suất, thanh khoản) có ở CẢ báo cáo quý thường (Q1-Q4) LẪN báo cáo đã kiểm
  toán/soát xét — nên dữ liệu giờ là THEO QUÝ THẬT (không còn chỉ ~2 điểm/năm như thiết kế ban đầu).
  Một ngân hàng có thể có ĐỒNG THỜI 2 bản ghi cho cùng 1 thời điểm cuối năm: "YYYY-Q4" (báo cáo quý
  thường, công bố sớm) và "YYYY-FY" (báo cáo năm kiểm toán, công bố sau, đáng tin hơn) — 2 khóa
  KHÁC NHAU trong gap_periods vì đây là 2 LẦN CÔNG BỐ khác nhau, không ghi đè lên nhau. Khi tổng hợp
  hệ thống, MỌI so khớp kỳ giữa các ngân hàng PHẢI qua kỳ QUÝ CHUẨN HÓA (gap_period_to_quarter) —
  xem _ticker_gap_entry() — KHÔNG BAO GIỜ so khớp period_key nguyên văn, nếu không sẽ tính nhầm 1
  ngân hàng đã có "YYYY-Q4" là "thiếu dữ liệu" chỉ vì ngân hàng khác đang dùng để so là "YYYY-FY".
- Kiểm tra độ mới: RẺ (chỉ gọi API liệt kê BCTC qua bank_risk_notes.latest_reviewed_period, không
  OCR) cho toàn bộ 26 ngân hàng mỗi lần chạy; CHỈ ngân hàng nào thực sự có kỳ mới hơn store hiện có
  mới bị OCR lại (refresh_bank_alm_data). Ngân hàng chưa công bố kỳ mới nhất của hệ thống được "vá"
  (patch) bằng kỳ gần nhất đã có, đánh dấu RÕ status="patched" — không lặng lẽ trộn vào coi như thật.
"""
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import re

import bank_alm_store

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))


def _normalize_period_input(raw):
    """Chuẩn hóa nhiều cách nhập kỳ backfill khác nhau về ĐÚNG 1 định dạng chuẩn "YYYY-Qn" hoặc
    "YYYY-FY" mà toàn bộ hệ thống (bank_alm_store, workflow_dispatch) đang dùng — user (2026-08-31)
    muốn ô nhập trong workflow chấp nhận "Q1-2025", "2025-Q1", "1-2025", "2025-1", "FY-2025",
    "2025-FY", "2025FY"... đều hiểu đúng ý, không cần nhớ đúng thứ tự năm/quý. Không phân biệt hoa
    thường, khoảng trắng/dấu gạch ngang tùy ý. Trả None nếu không nhận diện được năm HOẶC không
    nhận diện được quý/FY (không đoán bừa khi mơ hồ, vd chỉ nhập "2025")."""
    if not raw:
        return None
    s = raw.strip().upper()
    year_m = re.search(r"(20\d{2})", s)
    if not year_m:
        return None
    year = year_m.group(1)
    rest = (s[:year_m.start()] + s[year_m.end():]).strip(" -_/")
    if "FY" in rest:
        return f"{year}-FY"
    q_m = re.search(r"Q?\s*([1-4])", rest)
    if q_m:
        return f"{year}-Q{q_m.group(1)}"
    return None

# Field code Vietcap (xem template_banking.py get_yr — CHIA /1e9 để ra tỷ đồng, khớp đơn vị dùng
# xuyên suốt template_banking.py/bank_risk_notes.py). Nguồn BALANCE_SHEET, trừ khi ghi chú khác.
# inv_sec_bs/von_cg/tctd_dep: chưa được dùng ở đâu trong file này — giữ lại vì rẻ (đã lấy sẵn cùng
# 1 lượt fetch_data.fetch_all) và dành cho phần mở rộng LDR/CASA/NIM theo quý hệ thống (user đã
# đồng ý phạm vi "từ Q1-2025 tới nay" 2026-08-30) — PHẦN ĐÓ CHƯA XÂY, đây chỉ là input để dành.
_BS_FIELD_MAP = {
    "total_assets": "bsa53", "equity": "bsa78", "customer_deposits": "bsb113",
    "cash": "bsa2", "sbv_dep": "bsb97", "bank_dep": "bsb98", "interbank_liab": "bsb112",
    "loans": "bsb103", "bonds_issued": "bsb116", "inv_sec_bs": "bsb106",
    "von_cg": "bsb115", "tctd_dep": "bsb270",
}

_IR_HORIZON_1Y_KEYS = ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang"]
_IR_WEIGHTS = {"den_1_thang": 11.5 / 12, "tu_1_3_thang": 10 / 12, "tu_3_6_thang": 7.5 / 12, "tu_6_12_thang": 3 / 12}
_LIQ_ST_KEYS = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang"]

# Thu tu ky han dung de dung DUONG CONG gap luy ke (user 2026-08-31) — "khong anh huong lai suat"
# va "qua han" (bang lai suat) bi LOAI khoi duong cong IR vi khong phai muc dinh gia lai theo thoi
# gian (tai san/no khong nhay cam lai suat), khop voi _IR_HORIZON_1Y_KEYS da dung tu truoc. Bang
# thanh khoan giu ca 2 bucket "qua han" (coi nhu ky han ngan nhat, can thanh khoan ngay).
_LIQ_CUMULATIVE_ORDER = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang",
                         "tu_3_12_thang", "tu_1_5_nam", "tren_5_nam"]
_LIQ_ST_1Y_KEYS = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang", "tu_3_12_thang"]
_LIQ_LT_KEYS = ["tu_1_5_nam", "tren_5_nam"]
_IR_CUMULATIVE_ORDER = ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang", "tu_1_5_nam", "tren_5_nam"]
# Ban SONG SONG bao gom ca bucket "qua han" trong duong cong luy ke (theo file huong dan "Danh gia
# rui ro lai suat.docx", user 2026-09-18: "Sau khi bo dong 'khong chiu lai', ta co [duong cong luy
# ke]" — CHI loai "khong chiu lai", GIU LAI "qua han", khac voi _IR_CUMULATIVE_ORDER hien co (da loai
# ca 2 tu truoc, dung lam co so cho stress_nii_100bp — mot chi so DA SHIP, KHONG doi de tranh anh
# huong nguoc). Verify khop CHINH XAC voi vi du that trong file (VAB Quy 1/2026: cumulative ≤1 thang
# = -2.462 ty, ≤3 thang = +12.489 ty, ≤6 thang = +17.258 ty, ≤12 thang = +5.148 ty, ≤5 nam = -13.143
# ty, >5 nam = +3.313 ty — dung 100% so trong file).
_IR_CUMULATIVE_ORDER_INCL_OVERDUE = ["qua_han"] + _IR_CUMULATIVE_ORDER
# 2 bucket "qua han" cua bang thanh khoan (tai san da qua han hop dong) — theo huong dan phan tich
# rui ro thanh khoan (user 2026-09-18, file "Danh gia rui ro thanh khoan.docx"): tai san DA QUA HAN
# KHONG NEN mac nhien coi la nguon thanh khoan kha dung (mot khoan vay qua han 500 ty khong co nghia
# ngan hang co 500 ty tien mat de tra nguoi gui) — dung de tinh them 1 phien ban THAN TRONG cua gap
# luy ke (loai tai san qua han khoi nguon bu dap), song song ban "hop dong" (contractual) hien co.
_LIQ_OVERDUE_KEYS = ["qua_han_tren_3t", "qua_han_den_3t"]


def _cumulative_curve(gap_ty, order, ta):
    """Duong cong gap luy ke theo tung moc ky han trong `order` — moi diem la TONG DON (khong phai
    trung binh) cac bucket TU DAU DEN diem do, chia cho tong tai san. Dung chung cho ca lai suat va
    thanh khoan de tranh viet trung logic."""
    curve_abs, curve_ratio = {}, {}
    running = 0.0
    for k in order:
        running += gap_ty.get(k, 0.0)
        curve_abs[k] = running
        curve_ratio[k] = (running / ta) if ta else None
    return curve_abs, curve_ratio


def _quarter_key_from_record(rec):
    y = rec.get("yearReport")
    q = rec.get("lengthReport")
    if y is None or q not in (1, 2, 3, 4):
        return None
    return f"{y}-Q{q}"


# ── Cập nhật bảng cân đối theo quý (KHÔNG cần OCR — Vietcap, đã cache sẵn) ─────────────────────

def refresh_quarterly_balance_sheet_all_banks():
    """Cập nhật quarterly_balance_sheet cho TẤT CẢ 26 ngân hàng — rẻ (chỉ gọi fetch_data.fetch_all,
    Vietcap, tự cache ở .cache/{TICKER}_bctc.json), chạy thoải mái mỗi tuần. Ghi đè MỌI quý tìm
    được (không chỉ quý mới) vì chi phí gần 0 và Vietcap đôi khi sửa lại số liệu quý cũ."""
    from bank_universe import BANKING_TICKERS
    import fetch_data
    n_ok, n_fail = 0, 0
    for ticker in sorted(BANKING_TICKERS):
        try:
            raw = fetch_data.fetch_all(ticker, use_cache=True)
            bs_q = raw["sections"]["BALANCE_SHEET"].get("quarters", [])
            is_q = raw["sections"]["INCOME_STATEMENT"].get("quarters", [])
            is_by_qkey = {}
            for rec in is_q:
                qk = _quarter_key_from_record(rec)
                if qk:
                    is_by_qkey[qk] = rec
            for rec in bs_q:
                qk = _quarter_key_from_record(rec)
                if not qk:
                    continue
                snap = {name: (rec.get(code) or 0) / 1e9 for name, code in _BS_FIELD_MAP.items()}
                nii_rec = is_by_qkey.get(qk)
                snap["nii"] = (nii_rec.get("isb27") or 0) / 1e9 if nii_rec else None
                bank_alm_store.upsert_quarterly_balance_sheet(ticker, qk, snap)
            n_ok += 1
        except Exception as e:
            print(f"  [WARN] He thong ALM ({ticker}): loi cap nhat bang can doi quy ({e})")
            n_fail += 1
    print(f"  [INFO] He thong ALM: da cap nhat bang can doi quy cho {n_ok}/{n_ok+n_fail} ngan hang")


# ── Kiểm tra độ mới + vá dữ liệu thiếu (chạy hàng tuần, OCR CHỈ khi thực sự cần) ────────────────

def refresh_bank_alm_data():
    """Vòng lặp chính gọi từ fetch_macro_data.py mỗi tuần. Bước 1 (RẺ): kiểm tra kỳ đã kiểm
    toán/soát xét mới nhất hiện có của cả 26 ngân hàng (chỉ gọi API liệt kê, không OCR) để xác định
    "kỳ mục tiêu" = kỳ mới nhất mà BẤT KỲ ngân hàng nào đã thực sự công bố (tự điều chỉnh theo thực
    tế, không cố định lịch). Bước 2: với TỪNG ngân hàng, chỉ OCR (ĐẮT) nếu ngân hàng đó THỰC SỰ đã
    công bố đúng kỳ mục tiêu mà store chưa có; ngược lại vá bằng kỳ gần nhất đã có (đánh dấu rõ) hoặc
    bỏ qua nếu đã vá đúng nguồn rồi. Không bao giờ raise — lỗi 1 ngân hàng không chặn 25 ngân hàng
    còn lại. Trả về {"changed": bool, "target_period": str|None, "actions": {ticker: mo_ta}}."""
    from bank_universe import BANKING_TICKERS
    from bank_risk_notes import latest_reviewed_period, fetch_bank_risk_gaps_for_period

    cheap_periods = {}
    for ticker in sorted(BANKING_TICKERS):
        try:
            cheap_periods[ticker] = latest_reviewed_period(ticker)
        except Exception as e:
            print(f"  [WARN] He thong ALM ({ticker}): loi kiem tra ky moi nhat ({e})")
            cheap_periods[ticker] = None
        if cheap_periods[ticker]:
            try:
                bank_alm_store.record_cheap_check(ticker, cheap_periods[ticker])
            except Exception:
                pass

    known_periods = [p for p in cheap_periods.values() if p]
    if not known_periods:
        print("  [SKIP] He thong ALM: khong kiem tra duoc ky nao cho ngan hang nao (loi mang toan bo?)")
        return {"changed": False, "target_period": None, "actions": {}}
    target_period = max(known_periods, key=bank_alm_store._period_sort_key)
    target_canon = bank_alm_store.gap_period_to_quarter(target_period)
    print(f"  [INFO] He thong ALM: ky muc tieu hien tai = {target_period} (quy chuan hoa {target_canon})")

    actions = {}
    changed = False
    for ticker in sorted(BANKING_TICKERS):
        try:
            store = bank_alm_store.load_bank_store(ticker)
            # So khop qua QUY CHUAN HOA, KHONG so period_key nguyen van - 1 ngan hang co the da co
            # du lieu dung thoi diem duoi dang bao cao quy thuong ("...-Q4") trong khi target_period
            # he thong dang la ban kiem toan ("...-FY") cua ngan hang KHAC, hoac nguoc lai.
            existing_entry, existing_pk = _ticker_gap_entry(store, target_canon)
            if existing_entry and existing_entry.get("status") == "reported":
                actions[ticker] = f"skip_da_co ({existing_pk})"
                continue

            cp = cheap_periods.get(ticker)
            if cp and bank_alm_store.gap_period_to_quarter(cp) == target_canon:
                gaps = fetch_bank_risk_gaps_for_period(ticker, cp)
                if gaps and (gaps.get("interest_rate_gap") or gaps.get("liquidity_gap")):
                    source = {"title": gaps.get("source_title"), "url": gaps.get("source_url"),
                              "fetched_year": gaps.get("fetched_year")}
                    bank_alm_store.upsert_reported_period(ticker, cp, gaps, source)
                    actions[ticker] = f"da_OCR_thanh_cong ({cp})"
                    changed = True
                else:
                    actions[ticker] = "OCR_that_bai"
                continue

            latest_reported = bank_alm_store.latest_reported_period(ticker)
            if latest_reported:
                if (existing_entry and existing_entry.get("status") == "patched"
                        and existing_entry.get("patched_from") == latest_reported):
                    actions[ticker] = "skip_da_va_dung_nguon"
                else:
                    bank_alm_store.upsert_patched_period(ticker, target_canon, latest_reported)
                    actions[ticker] = f"da_va_tu_{latest_reported}"
                    changed = True
            else:
                if not existing_entry:
                    bank_alm_store.mark_missing_period(ticker, target_canon)
                    actions[ticker] = "danh_dau_thieu"
                    changed = True
                else:
                    actions[ticker] = "skip_da_danh_dau_thieu"
        except Exception as e:
            print(f"  [WARN] He thong ALM ({ticker}): loi xu ly ky {target_period} ({e})")
            actions[ticker] = f"loi: {e}"

    return {"changed": changed, "target_period": target_period, "actions": actions}


# ── Backfill lịch sử (workflow_dispatch riêng, KHÔNG chạy trong cron hàng tuần) ─────────────────

def backfill_period(period_key, force=False):
    """Backfill 1 KỲ LỊCH SỬ cụ thể (vd "2024-FY") cho TẤT CẢ 26 ngân hàng — dùng cho
    .github/workflows/backfill_bank_alm.yml (workflow_dispatch thủ công, KHÔNG chạy tự động).
    Bỏ qua ngân hàng đã có status="reported" đúng kỳ này (không OCR lại vô ích). Ghi "missing" cho
    ngân hàng không có báo cáo đúng kỳ (không phải lỗi — có thể chưa niêm yết lúc đó, hoặc
    CafeF/24hmoney không còn lưu bản cũ, xem fetch_bank_risk_gaps_for_period).

    2 SỬA (user 2026-08-31):
    1. period_key được CHUẨN HÓA qua _normalize_period_input() trước — chấp nhận "Q1-2025",
       "2025-Q1", "1-2025", "2025-1", "FY-2025", "2025-FY"... đều hiểu đúng, không cần đúng thứ tự.
    2. Nếu kỳ chuẩn hóa ra là quý 4 ("YYYY-Q4"), CHỦ ĐỘNG thử bản báo cáo NĂM ("YYYY-FY", đã kiểm
       toán, đáng tin hơn) TRƯỚC, chỉ fallback về đúng "YYYY-Q4" (báo cáo quý thường) nếu ngân hàng
       đó chưa có bản năm — nhất quán với thứ tự ưu tiên FY>Q4 đã dùng khi TỔNG HỢP hệ thống (xem
       _ticker_gap_entry), giờ áp dụng luôn từ bước BACKFILL/FETCH thay vì chỉ ở bước tổng hợp.

    `force=True` (user 2026-09-23, sau khi phat hien qua kiem tra cheo LS/TK: RAT NHIEU ky da
    "reported" (co du du lieu) nhung THUC RA sai - lech cot, tron nham dong, sai vi tri... - cac loi
    nay khong duoc backfill_period() thuong phat hien/OCR lai vi da "reported" thi bo qua): OCR LAI
    KHONG DIEU KIEN CA 26 ngan hang cho ky nay du da "reported" tu truoc, tan dung cac fix da sua
    trong code OCR (khop nham tieu de, mat dau am...) - CO THE tu sua duoc 1 phan, nhung KHONG chac
    sua het moi loi (vd loi lech cot No phai tra theo bucket rieng, tron nham dong noi/ngoai bang -
    2 loi nay CHUA duoc sua tan goc trong code, van can anh chup thuc te de sua tay)."""
    from bank_universe import BANKING_TICKERS

    normalized = _normalize_period_input(period_key)
    if normalized is None:
        raise ValueError(
            f"Khong nhan dien duoc ky '{period_key}' - nhap dang YYYY-Qn (vd 2025-Q1, Q1-2025, "
            f"1-2025 deu duoc) hoac YYYY-FY (vd 2025-FY, FY-2025)."
        )
    if normalized != period_key:
        print(f"  [INFO] Chuan hoa ky nhap '{period_key}' -> '{normalized}'")
    period_key = normalized

    year = period_key.split("-")[0]
    # Uu tien ban NAM khi ky la quy 4 (xem docstring) — thu tung candidate theo dung thu tu, dung
    # ban DAU TIEN thanh cong, chi ghi "thieu" khi CA 2 deu khong co.
    candidates = [f"{year}-FY", f"{year}-Q4"] if period_key.endswith("-Q4") else [period_key]

    results = {ticker: _backfill_ticker_period(ticker, period_key, candidates, force=force)
               for ticker in sorted(BANKING_TICKERS)}
    print(f"[DONE] Backfill {period_key} (force={force}): {results}")
    return results


def backfill_all_missing_periods(force=False):
    """Quét TẤT CẢ kỳ đã từng được backfill/nhập ít nhất 1 lần (bất kỳ trạng thái gì, ở BẤT KỲ ngân
    hàng nào — tức UNION của mọi period_key xuất hiện trong data/bank_alm/*.json) rồi chạy
    backfill_period() cho TỪNG kỳ đó — mỗi kỳ TỰ ĐỘNG bỏ qua (ticker, kỳ) đã có dữ liệu thật
    ("reported"), CHỈ tốn OCR cho đúng những tổ hợp đang "missing". Dùng cho workflow_dispatch
    backfill_bank_alm.yml khi để trống ô "period" — user (2026-09-18) muốn 1 lần chạy tự kiểm tra +
    lấy đúng phần dữ liệu còn thiếu trên TOÀN BỘ lịch sử đã backfill, thay vì phải tự tay chạy lại
    riêng từng kỳ (Q1, Q2, Q3...) mỗi khi có 1 fix OCR mới.

    KHÔNG tự "phát minh" thêm kỳ chưa ai từng backfill (vd sẽ không tự chạy kỳ tương lai chưa có báo
    cáo) — chỉ quét lại đúng các kỳ ĐÃ CÓ ít nhất 1 bản ghi (kể cả "missing") trong store, tức đúng
    những kỳ user đã từng chủ động backfill trước đây.

    `force=True`: xem docstring backfill_period() - truyen thang xuong, OCR LAI KHONG DIEU KIEN toan
    bo (ticker, ky) da tung backfill (26 ngan hang x N ky), khong chi cac to hop "missing". Ton
    NHIEU thoi gian hon han (moi to hop deu OCR lai, khong chi to hop thieu) - workflow can nguong
    timeout du rong (xem backfill_bank_alm.yml)."""
    from bank_universe import BANKING_TICKERS
    all_periods = set()
    for ticker in BANKING_TICKERS:
        store = bank_alm_store.load_bank_store(ticker)
        all_periods.update(store.get("gap_periods", {}).keys())
    if not all_periods:
        print("[SKIP] Chua co ky nao tung duoc backfill - khong co gi de quet lai.")
        return {}
    sorted_periods = sorted(all_periods, key=bank_alm_store._period_sort_key)
    print(f"[INFO] Quet lai {len(sorted_periods)} ky da tung backfill (force={force}): {sorted_periods}")
    results = {}
    for period_key in sorted_periods:
        results[period_key] = backfill_period(period_key, force=force)
    return results


def _backfill_ticker_period(ticker, period_key, candidates, force=False):
    """Phần thân DÙNG CHUNG cho cả backfill_period() (lặp qua 26 ngân hàng) và backfill_single_ticker()
    (1 ngân hàng lẻ) — thử từng candidate trong `candidates` (đã tính sẵn thứ tự ưu tiên FY>Q4 nếu
    cần) cho ĐÚNG 1 ticker, trả về chuỗi mô tả kết quả. force=True bỏ qua kiểm tra "đã có dữ liệu
    thật" (luôn OCR lại) — dùng khi biết dữ liệu cũ sai/thiếu (vd sau khi sửa 1 fix OCR, muốn lấy lại
    đúng 1 mã cụ thể mà không phải chờ/đợi toàn bộ 26 mã chạy lại)."""
    from bank_risk_notes import fetch_bank_risk_gaps_for_period
    try:
        for try_period in candidates:
            if not force:
                existing = bank_alm_store.get_period_entry(ticker, try_period)
                # SUA (user 2026-09-18, phat hien qua VIB/VAB): status="reported" KHONG dam bao ca 2
                # bang (lai suat + thanh khoan) deu co - dung is_fully_reported() thay vi chi kiem
                # tra status, de tu dong THU LAI dung nhung ky con do dang (1 trong 2 bang bi thieu)
                # thay vi mai mai bi coi la "da xong". upsert_reported_period() gio da MERGE (uu tien
                # gia tri moi, giu gia tri cu neu lan nay khong trich lai duoc) nen thu lai an toan,
                # khong lam mat field da co truoc do.
                if bank_alm_store.is_fully_reported(existing):
                    print(f"  [SKIP] {ticker} {try_period}: da co du lieu that DAY DU (ca lai suat + "
                          f"thanh khoan), bo qua")
                    return f"da_co ({try_period})"
                if existing and existing.get("status") == "reported":
                    print(f"  [INFO] {ticker} {try_period}: da 'reported' nhung con thieu 1 trong 2 "
                          f"bang gap - thu lai de lay not phan thieu")
            gaps = fetch_bank_risk_gaps_for_period(ticker, try_period)
            if gaps and (gaps.get("interest_rate_gap") or gaps.get("liquidity_gap")):
                source = {"title": gaps.get("source_title"), "url": gaps.get("source_url"),
                          "fetched_year": gaps.get("fetched_year")}
                bank_alm_store.upsert_reported_period(ticker, try_period, gaps, source)
                return f"da_co_du_lieu ({try_period})"
        # KHONG candidate nao thanh cong — TRUOC KHI ghi "missing", kiem tra xem da co du lieu
        # "reported" TOT tu truoc chua (bug that phat hien 2026-09-17 khi test cuc bo
        # backfill_single_ticker(force=True) tren BID 2025-Q1: OCR lai that bai vi may test khong co
        # tesseract, roi mark_missing_period() GHI DE THANG len du lieu that/nhap tay tot da co san,
        # xoa mat du lieu dung — mark_missing_period() KHONG tu kiem tra, ghi de VO DIEU KIEN). CHI
        # ghi "missing" khi truoc do THAT SU chua co gi (hoac da la "missing"/"patched") — force=True
        # + OCR lai that bai thi GIU NGUYEN du lieu tot cu, khong lam mat du lieu vi 1 lan thu lai
        # khong thanh cong (vd tam thoi mat mang, hoac OCR khong on dinh giua cac lan chay).
        existing = bank_alm_store.get_period_entry(ticker, period_key)
        # SUA (user 2026-09-23, phat hien qua SHB/NVB 2024-Q1): truoc day CHI giu nguyen du lieu cu
        # khi status="reported" - status="patched" (vd upsert_patched_period() vua ghi tay theo yeu
        # cau user vi ky nay THAT SU khong co bao cao rieng, "va" tam tu ky gan nhat) bi coi nhu
        # "chua co gi", bi GHI DE THANH "missing" ngay khi lan OCR thu lai nay khong tim thay gi (dung
        # nhu du doan cua user - bao cao khong ton tai) - XOA MAT quyet dinh "va" da chu dong lam,
        # lap lai VO HAN moi lan workflow chay lai (backfill_all_missing_periods() quet lai TAT CA ky
        # da tung co, bao gom ca ky "patched"). Gio giu nguyen CA "reported" VA "patched".
        if existing and existing.get("status") in ("reported", "patched"):
            print(f"  [WARN] {ticker} {period_key}: OCR lai that bai, GIU NGUYEN du lieu cu da co "
                  f"(status={existing.get('status')!r} - khong ghi de thanh missing)")
            return f"giu_nguyen_du_lieu_cu ({period_key})"
        bank_alm_store.mark_missing_period(ticker, period_key)
        return "thieu"
    except Exception as e:
        print(f"  [WARN] Backfill {ticker} {period_key}: loi ({e})")
        return f"loi: {e}"


def backfill_single_ticker(ticker, period_key, force=False):
    """Backfill/cập nhật lại ĐÚNG 1 ngân hàng cho 1 kỳ cụ thể — dùng cho workflow riêng
    backfill_bank_alm_single.yml (workflow_dispatch nhập tay 1 mã + 1 kỳ, KHÔNG cần chờ chạy qua cả
    26 mã như backfill_period()/backfill_bank_alm.yml — user 2026-09-17 muốn 1 action riêng để sửa
    nhanh 1 mã cụ thể đang thiếu/sai, đặc biệt sau khi đã sửa 1 fix OCR và muốn thử lại ngay 1 mã mà
    không tốn thời gian chờ + OCR lại 25 mã khác đã đúng."""
    from bank_universe import BANKING_TICKERS
    ticker = ticker.upper().strip()
    if ticker not in BANKING_TICKERS:
        raise ValueError(
            f"'{ticker}' khong nam trong danh sach 26 ngan hang niem yet/UPCoM dang theo doi "
            f"(xem bank_universe.BANKING_TICKERS) - kiem tra lai ma vua nhap."
        )
    normalized = _normalize_period_input(period_key)
    if normalized is None:
        raise ValueError(
            f"Khong nhan dien duoc ky '{period_key}' - nhap dang YYYY-Qn (vd 2025-Q1, Q1-2025, "
            f"1-2025 deu duoc) hoac YYYY-FY (vd 2025-FY, FY-2025)."
        )
    if normalized != period_key:
        print(f"  [INFO] Chuan hoa ky nhap '{period_key}' -> '{normalized}'")
    period_key = normalized

    year = period_key.split("-")[0]
    candidates = [f"{year}-FY", f"{year}-Q4"] if period_key.endswith("-Q4") else [period_key]
    result = _backfill_ticker_period(ticker, period_key, candidates, force=force)
    print(f"[DONE] Backfill {ticker} {period_key}: {result}")
    return {ticker: result}


# ── Tổng hợp hệ thống (arithmetic thuần, luôn tính lại từ đầu, không cache riêng) ────────────────

def _ticker_gap_entry(store, canonical_quarter):
    """Trả (entry, period_key_thuc_te) của 1 ngân hàng tại kỳ QUÝ CHUẨN HÓA (vd "2025-Q4") — nếu
    canonical_quarter là quý 4, ƯU TIÊN bản "YYYY-FY" (kiểm toán, công bố sau nhưng đáng tin hơn)
    nếu ngân hàng đó đã có, chỉ dùng "YYYY-Q4" (báo cáo quý thường) khi chưa có bản FY. BẮT BUỘC
    dùng hàm này thay vì so khớp period_key nguyên văn khi tổng hợp toàn hệ thống — 2 ngân hàng
    cùng phản ánh 1 thời điểm cuối năm có thể đang ở 2 "giai đoạn công bố" khác nhau (1 bên đã có
    FY, bên kia mới có Q4 thường), so khớp y hệt sẽ tính nhầm bên có Q4 là "thiếu dữ liệu"."""
    gp = store.get("gap_periods", {})
    if canonical_quarter.endswith("-Q4"):
        fy_key = f"{canonical_quarter.split('-')[0]}-FY"
        fy_entry = gp.get(fy_key)
        if fy_entry and fy_entry.get("status") != "missing":
            return fy_entry, fy_key
    entry = gp.get(canonical_quarter)
    return entry, (canonical_quarter if entry else None)


def _bank_period_metrics(entry, bs_snap):
    """Tính các chỉ số THÔ của 1 ngân hàng tại 1 kỳ từ (entry gap_periods, snapshot bảng cân đối
    cùng kỳ) — dùng CHUNG cho cả _aggregate_for_period (cộng dồn hệ thống) và
    update_bank_alm_excel_sheet (ghi hàng chi tiết từng ngân hàng), để 2 nơi này KHÔNG BAO GIỜ lệch
    công thức nhau. Trả về {} nếu entry rỗng/missing hoặc thiếu snapshot bảng cân đối cùng kỳ."""
    if not entry or entry.get("status") == "missing" or not bs_snap or not bs_snap.get("total_assets"):
        return {}
    ta = bs_snap["total_assets"]
    ir_gap_ty = {k: (v or 0) / 1000 for k, v in (entry.get("interest_rate_gap") or {}).items()}
    cum_1y = sum(ir_gap_ty.get(k, 0.0) for k in _IR_HORIZON_1Y_KEYS)
    # "weighted" = gap đã nhân trọng số thời gian còn lại tới khi định giá lại (KHÔNG PHẢI đã ở 1
    # mức sốc cụ thể) — nhân với bps/10000 mới ra đúng mức NII thay đổi tại 1 mức sốc — GIỮ Ở DẠNG
    # THÔ (chưa nhân bp, chưa làm tròn) để _aggregate_for_period cộng dồn TRƯỚC rồi mới nhân/làm
    # tròn 1 LẦN ở cấp hệ thống, tránh sai số cộng dồn của làm tròn từng ngân hàng riêng lẻ.
    weighted = sum(ir_gap_ty.get(k, 0.0) * _IR_WEIGHTS[k] for k in _IR_HORIZON_1Y_KEYS)
    liq_gap_ty = {k: (v or 0) / 1000 for k, v in (entry.get("liquidity_gap") or {}).items()}
    cum_1m = sum(liq_gap_ty.get(k, 0.0) for k in _LIQ_ST_KEYS)
    liquid_assets = (bs_snap.get("cash") or 0) + (bs_snap.get("sbv_dep") or 0) + (bs_snap.get("bank_dep") or 0)
    cust_dep = bs_snap.get("customer_deposits") or 0.0
    equity = bs_snap.get("equity") or 0.0
    nii = bs_snap.get("nii") or 0.0
    stress_nii_100 = round(weighted * 0.01)

    # ── Duong cong gap luy ke + cac moc ≤1 thang/≤3 thang/≤1 nam (user 2026-08-31, khung phan tich
    # day du rui ro thanh khoan + rui ro lai suat) — dung _cumulative_curve() chung cho ca 2 bang.
    liq_curve_abs, liq_curve_ratio = _cumulative_curve(liq_gap_ty, _LIQ_CUMULATIVE_ORDER, ta)
    ir_curve_abs, ir_curve_ratio = _cumulative_curve(ir_gap_ty, _IR_CUMULATIVE_ORDER, ta)
    liq_cum_3m, liq_cum_3m_ratio = liq_curve_abs["tu_1_3_thang"], liq_curve_ratio["tu_1_3_thang"]
    liq_cum_1y, liq_cum_1y_ratio = liq_curve_abs["tu_3_12_thang"], liq_curve_ratio["tu_3_12_thang"]
    ir_cum_1m, ir_cum_1m_ratio = ir_curve_abs["den_1_thang"], ir_curve_ratio["den_1_thang"]
    ir_cum_3m, ir_cum_3m_ratio = ir_curve_abs["tu_1_3_thang"], ir_curve_ratio["tu_1_3_thang"]
    # ── Them moc ≤6 thang/≤5 nam (theo "Danh gia rui ro lai suat.docx", user 2026-09-18: khuyen dung
    # ca day du 1M/3M/6M/12M/5Y/>5Y, khong chi dung lai o 3M nhu truoc) — dung LAI duong cong da tinh
    # o tren (khong tao them phep tinh moi), quy uoc LOAI "qua han" giu nguyen (xem ir_curve_abs).
    ir_cum_6m, ir_cum_6m_ratio = ir_curve_abs["tu_3_6_thang"], ir_curve_ratio["tu_3_6_thang"]
    ir_cum_5y, ir_cum_5y_ratio = ir_curve_abs["tu_1_5_nam"], ir_curve_ratio["tu_1_5_nam"]

    # ── Duong cong luy ke BAO GOM "qua han" (ban song song theo dung quy uoc file huong dan, xem
    # _IR_CUMULATIVE_ORDER_INCL_OVERDUE) — KHONG thay the ban tren (van la co so cho stress_nii_100bp,
    # 1 chi so DA SHIP), chi la GOC NHIN THAY THE de doi chieu voi cach doc phan tich trong file.
    ir_curve_incl_abs, ir_curve_incl_ratio = _cumulative_curve(ir_gap_ty, _IR_CUMULATIVE_ORDER_INCL_OVERDUE, ta)

    # ── Tai dung TAI SAN theo bucket = gap + no phai tra cung bucket (gap = TS - No), cho CA 2
    # bang, de tinh Short-term Funding/LT Assets, NSFR proxy, LMI, RSA/RSL — CHI tinh khi da trich
    # duoc dong "Tong no phai tra" tuong ung (interest_rate_liabilities_by_bucket moi them
    # 2026-08-31), neu khong co thi de None thay vi coi No=0 (se lam sai lech nghiem trong).
    #
    # KIEM TRA DU LIEU HONG (bug that phat hien 2026-08-31 qua ABB/STB 2025-Q1): doi luc buoc trich
    # dong "Tong no phai tra" theo toa do bi LAY NHAM chinh dong gap (7 gia tri "no phai tra" trung
    # KHOP TUYET DOI voi 7 gia tri gap) — coi nhu chua trich duoc gi, KHONG dung de tinh (se ra
    # Short-term Funding am/NSFR sai lech nghiem trong, xem thao luan voi user).
    liab_liq_raw = entry.get("liabilities_by_bucket") or {}
    if liab_liq_raw and liab_liq_raw == (entry.get("liquidity_gap") or {}):
        liab_liq_raw = {}
    liab_ir_raw = entry.get("interest_rate_liabilities_by_bucket") or {}
    if liab_ir_raw and liab_ir_raw == (entry.get("interest_rate_gap") or {}):
        liab_ir_raw = {}

    if liab_liq_raw:
        liab_liq_ty = {k: (v or 0) / 1000 for k, v in liab_liq_raw.items()}
        assets_liq_ty = {k: liq_gap_ty.get(k, 0.0) + liab_liq_ty.get(k, 0.0)
                          for k in set(liq_gap_ty) | set(liab_liq_ty)}
        short_term_funding = sum(liab_liq_ty.get(k, 0.0) for k in _LIQ_ST_1Y_KEYS)
        long_term_assets_liq = sum(assets_liq_ty.get(k, 0.0) for k in _LIQ_LT_KEYS)
        stable_funding = sum(liab_liq_ty.get(k, 0.0) for k in _LIQ_LT_KEYS) + equity
        st_funding_lt_assets_ratio = (short_term_funding / long_term_assets_liq) if long_term_assets_liq else None
        nsfr_proxy = (stable_funding / long_term_assets_liq) if long_term_assets_liq else None
        lmi = (long_term_assets_liq / stable_funding) if stable_funding else None

        # ── Gap luy ke THAN TRONG (Conservative Gap, xem "Danh gia rui ro thanh khoan.docx" muc 5) —
        # loai TAI SAN qua han khoi nguon bu dap (gan gia tri 0), GIU NGUYEN phia no phai tra (thuong
        # cung la 0, nhung khong gia dinh) — chi khac ban "hop dong" (contractual, cac bien cum_gap_1m/
        # liq_cum_gap_3m/1y phia tren) o 2 bucket qua han. Verify khop CHINH XAC vi du that trong file
        # huong dan (VAB Quy 1/2026: Conservative Gap 12 thang = -13.832 ty, tuong duong -9,62% tong
        # tai san — khop dung so lieu "-13,8 nghin ty"/"-9,6%" trong file).
        conservative_liq_gap_ty = dict(liq_gap_ty)
        for k in _LIQ_OVERDUE_KEYS:
            conservative_liq_gap_ty[k] = -liab_liq_ty.get(k, 0.0)
        cons_curve_abs, cons_curve_ratio = _cumulative_curve(conservative_liq_gap_ty, _LIQ_CUMULATIVE_ORDER, ta)
        liq_cum_1m_cons, liq_cum_1m_cons_ratio = cons_curve_abs["den_1_thang"], cons_curve_ratio["den_1_thang"]
        liq_cum_3m_cons, liq_cum_3m_cons_ratio = cons_curve_abs["tu_1_3_thang"], cons_curve_ratio["tu_1_3_thang"]
        liq_cum_1y_cons, liq_cum_1y_cons_ratio = cons_curve_abs["tu_3_12_thang"], cons_curve_ratio["tu_3_12_thang"]

        # ── Rollover Dependency (xem "Danh gia rui ro thanh khoan cau truc he thong.docx", user
        # 2026-09-21, dung vi du that MBB Quy 2/2026 lam mau) — KHAC HAN cac ty le luy ke o tren (chia
        # cho TONG TAI SAN): o day chia Conservative Gap (Forward Cumulative Gap, da tinh o tren) cho
        # TONG NO PHAI TRA DEN HAN trong dung khung thoi gian do — tra loi dung cau hoi file huong dan
        # dat ra: "bao nhieu % nghia vu den han KHONG duoc tai san cung ky han tu tai tro, buoc phai
        # rollover (huy dong moi/vay lien ngan hang/phat hanh GTCG...)". Ty le nay CANG CAO nghia la
        # ngan hang CANG PHAI CANH TRANH huy dong/rollover nguon von ky han do — day chinh la co che
        # noi gap thanh khoan cau truc VOI ap luc day lai suat huy dong ky han dai len (file huong dan
        # muc VIII/XI: "1M cao + 3M cao + 12M cao -> phu thuoc manh vao rollover"). Verify khop CHINH
        # XAC vi du that trong file (Nợ đến hạn ≤1th=437,97 ty -> rollover 1M=23,0%; ≤3th=659,96 ty ->
        # 20,1%; ≤12th=1.154,19 ty -> 7,1% — dung 100% 3/3 con so).
        liab_cum_abs, _ = _cumulative_curve(liab_liq_ty, _LIQ_CUMULATIVE_ORDER, ta)
        liab_due_1m = liab_cum_abs.get("den_1_thang")
        liab_due_3m = liab_cum_abs.get("tu_1_3_thang")
        liab_due_12m = liab_cum_abs.get("tu_3_12_thang")
        rollover_dep_1m = (abs(liq_cum_1m_cons) / liab_due_1m) if liab_due_1m else None
        rollover_dep_3m = (abs(liq_cum_3m_cons) / liab_due_3m) if liab_due_3m else None
        rollover_dep_12m = (abs(liq_cum_1y_cons) / liab_due_12m) if liab_due_12m else None

        # ── Ty le Tai san/No phai tra THEO TUNG BUCKET RIENG LE (khac han cac ty le luy ke o tren) —
        # xem muc 7 file huong dan: "<1 thang = 154,9%"/"1-3 thang = 60,6%"/"3-12 thang = 61,6%" — chỉ
        # 3 bucket nay duoc chon vi la 3 moc quan trong nhat de danh gia (< 1 thang = an toan tuc thoi,
        # 1-3 va 3-12 thang = vung ap luc chinh theo huong dan).
        liq_al_ratio_1m = (assets_liq_ty.get("den_1_thang", 0.0) / liab_liq_ty["den_1_thang"]) \
            if liab_liq_ty.get("den_1_thang") else None
        liq_al_ratio_3m = (assets_liq_ty.get("tu_1_3_thang", 0.0) / liab_liq_ty["tu_1_3_thang"]) \
            if liab_liq_ty.get("tu_1_3_thang") else None
        liq_al_ratio_12m = (assets_liq_ty.get("tu_3_12_thang", 0.0) / liab_liq_ty["tu_3_12_thang"]) \
            if liab_liq_ty.get("tu_3_12_thang") else None

        # ── Liquidity Buffer Coverage <1 thang (xem muc 9 file huong dan: "tai san gan tien" — tien
        # mat/tien gui NHNN/TCTD — so voi nghia vu <1 thang) — dung `liquid_assets` da co san (toan bo
        # bang can doi, tu bs_snap) thay vi chi rieng phan <1 thang trong bang gap (chua trich duoc o
        # muc do dong rieng le) — XAP XI rong hon dinh nghia trong file (dung toan bo thay vi chi phan
        # <1 thang), nhung van la 1 chi so huu ich, KHONG phai LCR that (file nhan manh ro diem nay).
        liquidity_buffer_coverage_1m = (liquid_assets / liab_liq_ty["den_1_thang"]) \
            if liab_liq_ty.get("den_1_thang") else None

        # ── Ban CHINH XAC HON cua "tai san gan tien <=1 thang" (muc 9 file huong dan) — dung DUNG
        # dong Tien mat + Tien gui NHNN trich truc tiep tu bang thanh khoan (user 2026-09-18), thay vi
        # xap xi bang liquid_assets toan bo bang can doi nhu tren. KHONG THAY THE
        # liquidity_buffer_coverage_1m (giu de tuong thich/luon co san du thieu du lieu chi tiet) —
        # them MOI, chinh xac hon, CHI co khi da trich duoc ca 2 dong chi tiet.
        cash_raw = entry.get("cash_by_bucket") or {}
        sbv_raw = entry.get("sbv_dep_by_bucket") or {}
        if cash_raw and sbv_raw and liab_liq_ty.get("den_1_thang"):
            near_cash_1m = ((cash_raw.get("den_1_thang") or 0) + (sbv_raw.get("den_1_thang") or 0)) / 1000
            near_cash_coverage_1m_precise = near_cash_1m / liab_liq_ty["den_1_thang"]
        else:
            near_cash_1m = None
            near_cash_coverage_1m_precise = None

        # ── Co cau ky han tien gui khach hang (muc 12 file huong dan) — % tien gui <=1 thang va
        # <=1 nam tren TONG tien gui khach hang (khong phai tren tong tai san) — CHI co khi da trich
        # duoc dong "Tien gui cua khach hang" theo bucket.
        cust_dep_raw = entry.get("customer_deposits_by_bucket") or {}
        if cust_dep_raw:
            cust_dep_bucket_ty = {k: (v or 0) / 1000 for k, v in cust_dep_raw.items()}
            cust_dep_total_bucketed = sum(cust_dep_bucket_ty.values())
            if cust_dep_total_bucketed:
                cust_dep_pct_1m = sum(cust_dep_bucket_ty.get(k, 0.0) for k in _LIQ_ST_KEYS) / cust_dep_total_bucketed
                cust_dep_pct_1y = sum(cust_dep_bucket_ty.get(k, 0.0) for k in _LIQ_ST_1Y_KEYS) / cust_dep_total_bucketed
            else:
                cust_dep_pct_1m = cust_dep_pct_1y = None
        else:
            cust_dep_pct_1m = cust_dep_pct_1y = None
    else:
        short_term_funding = long_term_assets_liq = stable_funding = None
        st_funding_lt_assets_ratio = nsfr_proxy = lmi = None
        liq_cum_1m_cons = liq_cum_1m_cons_ratio = None
        liq_cum_3m_cons = liq_cum_3m_cons_ratio = None
        liq_cum_1y_cons = liq_cum_1y_cons_ratio = None
        rollover_dep_1m = rollover_dep_3m = rollover_dep_12m = None
        liab_due_1m = liab_due_3m = liab_due_12m = None
        liq_al_ratio_1m = liq_al_ratio_3m = liq_al_ratio_12m = None
        liquidity_buffer_coverage_1m = None
        near_cash_1m = near_cash_coverage_1m_precise = None
        cust_dep_pct_1m = cust_dep_pct_1y = None

    if liab_ir_raw:
        liab_ir_ty = {k: (v or 0) / 1000 for k, v in liab_ir_raw.items()}
        assets_ir_ty = {k: ir_gap_ty.get(k, 0.0) + liab_ir_ty.get(k, 0.0)
                         for k in set(ir_gap_ty) | set(liab_ir_ty)}
        rsa = sum(assets_ir_ty.get(k, 0.0) for k in _IR_CUMULATIVE_ORDER)
        rsl = sum(liab_ir_ty.get(k, 0.0) for k in _IR_CUMULATIVE_ORDER)
        rsa_rsl_ratio = (rsa / rsl) if rsl else None

        # ── RSA/RSL THEO TUNG BUCKET RIENG LE (xem "Danh gia rui ro lai suat.docx" muc 17) — file
        # CANH BAO RO: mau so nho co the lam ty le bi phong dai (vd 3-6 thang co the len toi >1000%
        # neu RSL bucket do rat nho) — nen KHONG dua vao rieng ty le nay, luon xem cung GAP tuyet doi +
        # GAP/Tai san + Cumulative GAP (da co san o cac field khac). Chi tinh cho 6 bucket nhay cam lai
        # suat that su (loai "qua han" va "khong chiu lai", giong _IR_CUMULATIVE_ORDER).
        ir_al_ratio_by_bucket = {
            k: (assets_ir_ty.get(k, 0.0) / liab_ir_ty[k]) if liab_ir_ty.get(k) else None
            for k in _IR_CUMULATIVE_ORDER
        }
    else:
        rsa = rsl = rsa_rsl_ratio = None
        ir_al_ratio_by_bucket = {}

    loans = bs_snap.get("loans")
    ldr = (loans / cust_dep) if (loans and cust_dep) else None

    return {
        "total_assets": ta, "equity": bs_snap.get("equity"), "nii": bs_snap.get("nii"),
        "customer_deposits": cust_dep, "liquid_assets": liquid_assets, "loans": loans, "ldr": ldr,
        "cum_gap_1y": cum_1y, "cum_gap_1y_ratio": (cum_1y / ta) if ta else None,
        "cum_gap_1m": cum_1m, "cum_gap_1m_ratio": (cum_1m / ta) if ta else None,
        "liq_cum_gap_3m": liq_cum_3m, "liq_cum_gap_3m_ratio": liq_cum_3m_ratio,
        "liq_cum_gap_1y": liq_cum_1y, "liq_cum_gap_1y_ratio": liq_cum_1y_ratio,
        "liq_cum_gap_curve_ratio": liq_curve_ratio,
        "ir_cum_gap_1m": ir_cum_1m, "ir_cum_gap_1m_ratio": ir_cum_1m_ratio,
        "ir_cum_gap_3m": ir_cum_3m, "ir_cum_gap_3m_ratio": ir_cum_3m_ratio,
        "ir_cum_gap_6m": ir_cum_6m, "ir_cum_gap_6m_ratio": ir_cum_6m_ratio,
        "ir_cum_gap_5y": ir_cum_5y, "ir_cum_gap_5y_ratio": ir_cum_5y_ratio,
        "ir_cum_gap_curve_ratio": ir_curve_ratio,
        # ── Ban song song BAO GOM "qua han" (xem ghi chu _IR_CUMULATIVE_ORDER_INCL_OVERDUE) — theo
        # dung quy uoc "Danh gia rui ro lai suat.docx", KHONG thay the cac field tren.
        "ir_cum_gap_curve_incl_overdue_ratio": ir_curve_incl_ratio,
        "short_term_funding": short_term_funding, "long_term_assets": long_term_assets_liq,
        "st_funding_lt_assets_ratio": st_funding_lt_assets_ratio,
        "stable_funding": stable_funding, "nsfr_proxy": nsfr_proxy, "lmi": lmi,
        # ── Rollover Dependency + Long-term Structural Funding Gap (xem ghi chu chi tiet cong thuc
        # o khoi tinh toan phia tren) — "long_term_structural_gap" = 1 - nsfr_proxy (nsfr_proxy da la
        # "Long-term Funding Coverage" dung y file huong dan, khong can tinh lai tu dau). Cac field
        # liq_cum_gap_*_conservative da co san (xem duoi), KHONG lap lai o day.
        "rollover_dependency_1m": rollover_dep_1m, "rollover_dependency_3m": rollover_dep_3m,
        "rollover_dependency_12m": rollover_dep_12m,
        # Mau so (No den han luy ke, ty dong) cua 3 ty le rollover_dependency tren - de _aggregate_
        # for_period() cong dong TU/MAU rieng (tong hop CO TRONG SO, khong lay trung binh % tung NH).
        "liab_due_1m": liab_due_1m, "liab_due_3m": liab_due_3m, "liab_due_12m": liab_due_12m,
        "long_term_funding_coverage": nsfr_proxy,
        "long_term_structural_gap": (1 - nsfr_proxy) if nsfr_proxy is not None else None,
        "rsa": rsa, "rsl": rsl, "rsa_rsl_ratio": rsa_rsl_ratio,
        "ir_al_ratio_by_bucket": ir_al_ratio_by_bucket,
        "weighted_gap_raw": weighted,
        "stress_nii_100bp": stress_nii_100,
        "stress_nii_ratio_100bp": (stress_nii_100 / nii) if nii else None,
        "stress_nii_ratio_equity_100bp": (stress_nii_100 / equity) if equity else None,
        "stress_nii_200bp": stress_nii_100 * 2,
        "stress_nii_ratio_200bp": (stress_nii_100 * 2 / nii) if nii else None,
        "deposit_run_coverage_5pct": (liquid_assets / (cust_dep * 0.05)) if cust_dep else None,
        "deposit_run_coverage_10pct": (liquid_assets / (cust_dep * 0.10)) if cust_dep else None,
        "deposit_run_coverage_20pct": (liquid_assets / (cust_dep * 0.20)) if cust_dep else None,
        # ── Bo sung theo "Danh gia rui ro thanh khoan.docx" (user 2026-09-18) — xem cac ghi chu chi
        # tiet cong thuc o khoi tinh toan phia tren. Da verify khop CHINH XAC voi vi du that trong
        # file (VAB Quy 1/2026).
        "liq_cum_gap_1m_conservative": liq_cum_1m_cons, "liq_cum_gap_1m_conservative_ratio": liq_cum_1m_cons_ratio,
        "liq_cum_gap_3m_conservative": liq_cum_3m_cons, "liq_cum_gap_3m_conservative_ratio": liq_cum_3m_cons_ratio,
        "liq_cum_gap_1y_conservative": liq_cum_1y_cons, "liq_cum_gap_1y_conservative_ratio": liq_cum_1y_cons_ratio,
        "liq_al_ratio_1m": liq_al_ratio_1m, "liq_al_ratio_3m": liq_al_ratio_3m, "liq_al_ratio_12m": liq_al_ratio_12m,
        "liquidity_buffer_coverage_1m": liquidity_buffer_coverage_1m,
        "near_cash_1m": near_cash_1m, "near_cash_coverage_1m_precise": near_cash_coverage_1m_precise,
        "customer_deposits_pct_1m": cust_dep_pct_1m, "customer_deposits_pct_1y": cust_dep_pct_1y,
    }


def _aggregate_for_period(period_key):
    """Tổng hợp CÓ TRỌNG SỐ THEO QUY MÔ (cộng dồn số tuyệt đối tỷ VND rồi mới chia ra tỷ lệ) — KHÔNG
    lấy trung bình cộng % của từng ngân hàng, vì cách đó coi 1 ngân hàng nhỏ ngang 1 ngân hàng lớn,
    sai lệch nghiêm trọng ý nghĩa "rủi ro của TOÀN HỆ THỐNG". Chỉ tính trên ngân hàng có
    status in ("reported","patched") VÀ có snapshot bảng cân đối cùng kỳ — thiếu 1 trong 2 thì coi
    như thiếu dữ liệu cho kỳ này, không đoán.

    period_key ở đây LUÔN là KỲ QUÝ CHUẨN HÓA ("YYYY-Qn", không bao giờ "YYYY-FY") — xem
    recompute_system_aggregate_all_periods(). Việc từng ngân hàng thực tế đã công bố dưới dạng báo
    cáo quý thường hay báo cáo năm kiểm toán được giải quyết TRONG _ticker_gap_entry(), không ảnh
    hưởng đến khóa period dùng để nhóm toàn hệ thống — nhờ vậy 1 bảng chỉ số không bị tách rời giữa
    2 sheet Excel (Theo_Quy/Luy_Ke) và coverage không bị đếm thiếu ngân hàng do khác kiểu báo cáo."""
    from bank_universe import BANKING_TICKERS

    reported, patched, missing = [], [], []
    sum_ta = sum_ir_net = sum_ir_abs = sum_weighted = sum_nii = 0.0
    sum_liq_1m = sum_liquid_assets = sum_cust_dep = 0.0
    # Cong dong TU/MAU rieng cho Rollover Dependency + Long-term Funding Coverage he thong (xem
    # "Danh gia rui ro thanh khoan cau truc he thong.docx", user 2026-09-21) - GIONG nguyen tac cac
    # ty le khac trong ham nay: cong tong tuyet doi truoc, chia ty le SAU, khong lay trung binh %.
    sum_cons_gap_1m = sum_cons_gap_3m = sum_cons_gap_12m = 0.0
    sum_liab_due_1m = sum_liab_due_3m = sum_liab_due_12m = 0.0
    sum_long_term_assets = sum_stable_funding = 0.0
    sum_ta_structural = 0.0  # tong tai san CHI cua cac NH co du lieu hop le cho cau truc ky han
    n_banks_structural = 0
    # Do phu RIENG cho ty le rui ro lai suat/thanh khoan chinh (user 2026-09-23, sau khi them guard
    # kiem tra cheo tong lai suat vs tong thanh khoan o duoi - guard nay co the loai RAT NHIEU ngan
    # hang cung 1 luc (vd 19/26 tai 1 ky thuc te), khien ty le "toan he thong" tinh ra tu mau rat nho
    # ma cac o dem "n_banks_reported" cu KHONG phan anh dieu nay (chi dem "co entry", khong dem "co
    # dong gop vao tong"). Theo dung nguyen tac da dung cho cau truc ky han (structural_funding_
    # coverage_pct) - cong tong tai san CHI cua ngan hang thuc su dong gop vao sum_ir_abs/sum_liq_1m.
    sum_ta_ir_valid = 0.0
    n_banks_ir_valid = 0
    sum_ta_liq_valid = 0.0
    n_banks_liq_valid = 0
    missing_structural = []  # ma NH da "reported"/"patched" (co du lieu gap) nhung THIEU rieng
    # phan "No phai tra theo bucket" can cho Rollover Dependency - de nguoi dung biet CAN backfill
    # gi (user 2026-09-21, xem "Do phu du lieu" trong muc Rui ro he thong ngan hang tren web).
    all_ta_known = 0.0
    worst_ir = None
    worst_liq = None
    worst_rollover = None
    by_bank = {}

    for ticker in sorted(BANKING_TICKERS):
        store = bank_alm_store.load_bank_store(ticker)
        qbs = store.get("quarterly_balance_sheet", {})
        if qbs:
            latest_q = max(qbs.keys(), key=lambda k: (int(k.split("-Q")[0]), int(k.split("-Q")[1])))
            latest_ta = qbs[latest_q].get("total_assets")
            if latest_ta:
                all_ta_known += latest_ta

        entry, actual_pk = _ticker_gap_entry(store, period_key)
        if not entry or entry.get("status") == "missing":
            missing.append(ticker)
            by_bank[ticker] = {"status": "missing"}
            continue

        status = entry["status"]
        bs_snap = qbs.get(period_key)
        m = _bank_period_metrics(entry, bs_snap)
        if not m:
            by_bank[ticker] = {"status": status, "note": "thieu snapshot bang can doi cung ky"}
            continue

        if status == "reported":
            reported.append(ticker)
        else:
            patched.append(ticker)

        # KIEM TRA DU LIEU HONG truoc khi cong vao tong he thong (bug that phat hien 2026-09-19 qua
        # OCB 2026-Q2: interest_rate_gap luu SAI DON VI - VND thuc thay vi trieu dong, vd bucket
        # "tu_1_3_thang"=62.558 TY trieu dong, khien cum_gap_1y_ratio rieng OCB = +18.585.446% (185854
        # LAN tong tai san) - 1 ngan hang duy nhat lam sai lech HOAN TOAN ty le toan he thong (+3.095%
        # thay vi vai % nhu binh thuong).
        #
        # Chan tren MANG TINH TOAN HOC (khong phai nguong tuy y): gap 1 bucket = tai san bucket - no
        # bucket, ma tai san bucket luon <= TONG tai san (moi bucket la 1 TAP CON cua bang can doi) -
        # nen |gap 1 bucket| KHONG THE VUOT QUA tong tai san trong du lieu dung. Ho so 1.5x de tru
        # sai so lam tron/OCR nho, van du hep de bat ca truong hop nhe hon OCB - phat hien THEM qua
        # NVB 2026-Q2: bucket "tu_3_6_thang" = 304.699 ty, VUOT tong tai san 198.896 ty (1.53x) - ratio
        # tong +205% tuy khong do bang OCB nhung van la 1 con so KHONG THE THAT, cung bi loai o day.
        def _has_bucket_over_assets(gap_key, ta_ty):
            gap = entry.get(gap_key) or {}
            return any(abs((v or 0) / 1000) > ta_ty * 1.5 for v in gap.values())

        ta_ty = m["total_assets"]
        ir_ratio_valid = not _has_bucket_over_assets("interest_rate_gap", ta_ty)
        liq_ratio_valid = not _has_bucket_over_assets("liquidity_gap", ta_ty)
        if not ir_ratio_valid:
            print(f"  [WARN] {ticker} {period_key}: co bucket interest_rate_gap vuot qua tong tai san "
                  f"(ratio he thong tinh duoc: {(m.get('cum_gap_1y_ratio') or 0)*100:+.0f}%) - nghi ngo "
                  f"sai don vi, LOAI khoi tong hop lai suat he thong (van tinh thanh khoan binh thuong)")
        if not liq_ratio_valid:
            print(f"  [WARN] {ticker} {period_key}: co bucket liquidity_gap vuot qua tong tai san "
                  f"(ratio he thong tinh duoc: {(m.get('cum_gap_1m_ratio') or 0)*100:+.0f}%) - nghi ngo "
                  f"sai don vi, LOAI khoi tong hop thanh khoan he thong (van tinh lai suat binh thuong)")

        # KIEM TRA CHEO lai suat vs thanh khoan (user 2026-09-23, sau khi lo ngai "sai vi tri, lay
        # nham so o vi tri khac" khong bi bat boi guard bucket-vs-tong-tai-san tren): 2 bang gap lai
        # suat + thanh khoan CUNG 1 ky, CUNG 1 ngan hang deu = Tong tai san - Tong no phai tra tai
        # DUNG 1 thoi diem (chi chia theo 2 kieu ky han khac nhau) - nen TONG cua 2 bang PHAI xap xi
        # bang nhau. Verify qua hang chuc anh chup thuc te (ABB/ACB/BID/HDB/MBB/MSB/NVB/PGB/SGB/STB/
        # VBB/VIB...): tong 2 bang luon KHOP CHINH XAC hoac chi lech ~1-2% (nguyen nhan lam tron/phan
        # loai khac nhau giua 2 thuyet minh cua 1 so bank). Lech > 20% (dac biet gan 100% hoac dau
        # nguoc nhau) la dau hieu ro rang 1 trong 2 bang bi doc sai (lech cot, gop nham dong, doc
        # nham trang khac) MA KHONG lo ra qua guard bucket-vs-tong-tai-san (vi moi bucket rieng le
        # van "hop ly" ve do lon, chi SAI VI TRI/GOM NHAM). KHONG biet chac bang nao sai nen loai CA
        # HAI khoi tong hop (an toan hon giu nham 1 ben).
        ir_sum_ty = sum((v or 0) / 1000 for v in (entry.get("interest_rate_gap") or {}).values())
        liq_sum_ty = sum((v or 0) / 1000 for v in (entry.get("liquidity_gap") or {}).values())
        ir_liq_cross_valid = True
        if (ir_sum_ty or liq_sum_ty):
            denom = max(abs(ir_sum_ty), abs(liq_sum_ty), 0.01)
            if abs(ir_sum_ty - liq_sum_ty) / denom > 0.20:
                ir_liq_cross_valid = False
        if not ir_liq_cross_valid:
            print(f"  [WARN] {ticker} {period_key}: tong khe ho lai suat ({ir_sum_ty:+.0f} ty) lech qua "
                  f"nhieu so tong khe ho thanh khoan ({liq_sum_ty:+.0f} ty) - nghi ngo 1 trong 2 bang bi "
                  f"doc sai vi tri/gop nham dong, LOAI CA HAI khoi tong hop he thong ky nay")
            ir_ratio_valid = False
            liq_ratio_valid = False

        # KIEM TRA DU LIEU HONG rieng cho Rollover Dependency (bug that phat hien 2026-09-21 qua BAB
        # 2026-Q2): khong bucket nao vuot qua tong tai san (qua duoc guard tren) nhung "Tong no phai
        # tra" theo bucket bi LECH TAP TRUNG bat thuong (~93% dat vao 1 bucket "tu_1-5_nam" duy nhat,
        # gan nhu khong co gi o cac bucket ngan han - khong hop ly cho co cau huy dong ngan hang thuc
        # te), khien mau so (No den han luy ke) qua nho so voi tu so (Forward Gap), ra rollover
        # dependency 12M = 838% (khong the that). Rollover dependency KHONG co chan tren "mang tinh
        # toan hoc" don gian nhu bucket-vs-tong-tai-san (ve ly thuyet co the vuot 100% neu tai san
        # dao han qua it), nhung mot ngan hang thuc te KHONG THE can 8 LAN nghia vu den han moi du bu
        # dap - nguong 300% du rong de khong loai nham truong hop cang thang thuc su (vd du lieu that
        # trong file huong dan chi 7-23%), nhung du hep de bat duoc BAB.
        rollover_vals = [m.get("rollover_dependency_1m"), m.get("rollover_dependency_3m"), m.get("rollover_dependency_12m")]
        structural_funding_valid = liq_ratio_valid and not any(v is not None and abs(v) > 3.0 for v in rollover_vals)
        if liq_ratio_valid and not structural_funding_valid:
            print(f"  [WARN] {ticker} {period_key}: rollover dependency bat thuong "
                  f"(1M={m.get('rollover_dependency_1m')}, 3M={m.get('rollover_dependency_3m')}, "
                  f"12M={m.get('rollover_dependency_12m')}) - nghi ngo No phai tra theo bucket bi lech "
                  f"tap trung sai, LOAI khoi tong hop cau truc ky han he thong")
        if not structural_funding_valid or m.get("liab_due_12m") is None:
            missing_structural.append(ticker)

        sum_ta += m["total_assets"]
        if ir_ratio_valid:
            sum_ir_net += m["cum_gap_1y"]
            sum_ir_abs += abs(m["cum_gap_1y"])
            sum_weighted += m["weighted_gap_raw"]
            sum_ta_ir_valid += m["total_assets"]
            n_banks_ir_valid += 1
        sum_nii += m["nii"] or 0.0
        if liq_ratio_valid:
            sum_liq_1m += m["cum_gap_1m"]
            sum_ta_liq_valid += m["total_assets"]
            n_banks_liq_valid += 1
        if structural_funding_valid and m.get("liab_due_12m") is not None:
            sum_ta_structural += m["total_assets"]
            n_banks_structural += 1
        if structural_funding_valid:
            if m.get("liq_cum_gap_1m_conservative") is not None:
                sum_cons_gap_1m += m["liq_cum_gap_1m_conservative"]
            if m.get("liq_cum_gap_3m_conservative") is not None:
                sum_cons_gap_3m += m["liq_cum_gap_3m_conservative"]
            if m.get("liq_cum_gap_1y_conservative") is not None:
                sum_cons_gap_12m += m["liq_cum_gap_1y_conservative"]
            if m.get("liab_due_1m") is not None:
                sum_liab_due_1m += m["liab_due_1m"]
            if m.get("liab_due_3m") is not None:
                sum_liab_due_3m += m["liab_due_3m"]
            if m.get("liab_due_12m") is not None:
                sum_liab_due_12m += m["liab_due_12m"]
            if m.get("long_term_assets") is not None:
                sum_long_term_assets += m["long_term_assets"]
            if m.get("stable_funding") is not None:
                sum_stable_funding += m["stable_funding"]
        sum_liquid_assets += m["liquid_assets"]
        sum_cust_dep += m["customer_deposits"]

        bank_ir_ratio = m["cum_gap_1y_ratio"] if ir_ratio_valid else None
        bank_cov10 = m["deposit_run_coverage_10pct"]
        bank_rollover_12m = m.get("rollover_dependency_12m") if structural_funding_valid else None
        if bank_ir_ratio is not None and (worst_ir is None or abs(bank_ir_ratio) > abs(worst_ir[1])):
            worst_ir = (ticker, bank_ir_ratio)
        if bank_cov10 is not None and (worst_liq is None or bank_cov10 < worst_liq[1]):
            worst_liq = (ticker, bank_cov10)
        if bank_rollover_12m is not None and (worst_rollover is None or bank_rollover_12m > worst_rollover[1]):
            worst_rollover = (ticker, bank_rollover_12m)

        by_bank[ticker] = {
            "status": status,
            "period_used": actual_pk if status == "reported" else entry.get("patched_from"),
            "cum_gap_1y_ratio": bank_ir_ratio,
            "deposit_run_coverage_10pct": bank_cov10,
            "rollover_dependency_12m": bank_rollover_12m,
            "long_term_funding_coverage": m.get("long_term_funding_coverage") if structural_funding_valid else None,
        }
        if not ir_ratio_valid:
            by_bank[ticker]["ir_data_quality_note"] = "cum_gap_1y_ratio bat thuong, da loai khoi tong hop he thong"

    net_gap_ratio = (sum_ir_net / sum_ta) if sum_ta else None
    dispersion_ratio = (sum_ir_abs / sum_ta) if sum_ta else None
    stress_nii, stress_nii_ratio = {}, {}
    for bps in (100, 200):
        v = round(sum_weighted * (bps / 10000))
        stress_nii[f"+{bps}bp"] = v
        stress_nii[f"-{bps}bp"] = -v
        stress_nii_ratio[f"+{bps}bp"] = (v / sum_nii) if sum_nii else None
        stress_nii_ratio[f"-{bps}bp"] = (-v / sum_nii) if sum_nii else None

    liq_net_ratio = (sum_liq_1m / sum_ta) if sum_ta else None
    liquid_assets_ratio = (sum_liquid_assets / sum_ta) if sum_ta else None
    deposit_coverage = {}
    for pct in (5, 10, 20):
        outflow = sum_cust_dep * (pct / 100)
        deposit_coverage[f"-{pct}%"] = (sum_liquid_assets / outflow) if outflow else None

    assets_coverage_pct = (sum_ta / all_ta_known * 100) if all_ta_known else None
    # Do phu THUC SU (sau khi loai boi guard kiem tra cheo) cho tung ty le rieng - xem ghi chu o
    # cho khai bao sum_ta_ir_valid/sum_ta_liq_valid o tren.
    ir_coverage_pct = (sum_ta_ir_valid / sum_ta * 100) if sum_ta else None
    liq_coverage_pct = (sum_ta_liq_valid / sum_ta * 100) if sum_ta else None

    # ── Rollover Dependency + Long-term Funding Coverage HE THONG (xem "Danh gia rui ro thanh khoan
    # cau truc he thong.docx", user 2026-09-21) - cong TU/MAU rieng qua tung ngan hang (da lam o
    # vong lap tren) roi moi chia, GIONG nguyen tac cac ty le khac trong ham nay (KHONG lay trung
    # binh % tung ngan hang, tranh 1 ngan hang nho lam sai lech nhu ca ngan hang lon).
    rollover_dep_1m_sys = (abs(sum_cons_gap_1m) / sum_liab_due_1m) if sum_liab_due_1m else None
    rollover_dep_3m_sys = (abs(sum_cons_gap_3m) / sum_liab_due_3m) if sum_liab_due_3m else None
    rollover_dep_12m_sys = (abs(sum_cons_gap_12m) / sum_liab_due_12m) if sum_liab_due_12m else None
    long_term_funding_coverage_sys = (sum_stable_funding / sum_long_term_assets) if sum_long_term_assets else None
    # Do phu RIENG cho cau truc ky han (khac han assets_coverage_pct chung o tren) — bug that phat
    # hien 2026-09-21 (user nghi ngo dung, xem thao luan): "liabilities_by_bucket" thuong cham co du
    # lieu hon "liquidity_gap" (nhieu ngan hang moi co gap nhung chua co bang no theo bucket ky moi
    # nhat) - ky moi nhat co the tut xuong CHI 12/26 ngan hang (thieu han MBB/VCB/CTG...) trong khi
    # ky truoc co 20/26, khien ty le "tut manh" chi vi MAU SO khac nhau giua 2 ky, KHONG phai tin
    # hieu cau truc thuc su cai thien. Dung de CANH BAO/CHAN ket luan xu huong khi do phu qua thap
    # hoac lech qua nhieu so ky truoc (xem classify_structural_funding_phase).
    structural_funding_coverage_pct = (sum_ta_structural / sum_ta * 100) if sum_ta else None

    return {
        "period": period_key,
        "n_banks_total": len(BANKING_TICKERS),
        "n_banks_reported": len(reported), "n_banks_patched": len(patched), "n_banks_missing": len(missing),
        "reported_tickers": reported, "patched_tickers": patched, "missing_tickers": missing,
        "assets_coverage_pct": assets_coverage_pct,
        "total_assets": sum_ta if sum_ta else None,
        "interest_rate_risk": {
            "net_gap_ratio": net_gap_ratio, "dispersion_gap_ratio": dispersion_ratio,
            "stress_nii": stress_nii, "stress_nii_ratio": stress_nii_ratio,
            "worst_bank": {"ticker": worst_ir[0], "ratio": worst_ir[1]} if worst_ir else None,
            "coverage_pct": ir_coverage_pct, "n_banks_included": n_banks_ir_valid,
        },
        "liquidity_risk": {
            "net_gap_1m_ratio": liq_net_ratio, "liquid_assets_ratio": liquid_assets_ratio,
            "deposit_run_coverage": deposit_coverage,
            "weakest_bank": {"ticker": worst_liq[0], "coverage": worst_liq[1]} if worst_liq else None,
            "coverage_pct": liq_coverage_pct, "n_banks_included": n_banks_liq_valid,
        },
        # ── Cau truc ky han nguon von he thong (xem "Danh gia rui ro thanh khoan cau truc he
        # thong.docx") - tra loi cau hoi "he thong dang can bao nhieu % nghia vu den han phai
        # rollover" (ap luc canh tranh huy dong/day lai suat ky han dai len) VA "tai san dai han da
        # duoc nguon von on dinh tai tro bao nhieu %".
        "structural_funding": {
            "rollover_dependency_1m": rollover_dep_1m_sys, "rollover_dependency_3m": rollover_dep_3m_sys,
            "rollover_dependency_12m": rollover_dep_12m_sys,
            "long_term_funding_coverage": long_term_funding_coverage_sys,
            "long_term_structural_gap": (1 - long_term_funding_coverage_sys)
                                         if long_term_funding_coverage_sys is not None else None,
            "most_dependent_bank": {"ticker": worst_rollover[0], "rollover_dependency_12m": worst_rollover[1]}
                                    if worst_rollover else None,
            "coverage_pct": structural_funding_coverage_pct, "n_banks_included": n_banks_structural,
            "missing_tickers": sorted(missing_structural),
        },
        "by_bank": by_bank,
    }


def recompute_system_aggregate_all_periods():
    """Trả về dict {period_key: aggregate_dict} cho MỌI kỳ QUÝ CHUẨN HÓA từng xuất hiện ở BẤT KỲ
    ngân hàng nào trong store — LUÔN tính lại từ đầu (rẻ, thuần cộng/chia trên vài chục file JSON
    nhỏ), không cache riêng kết quả tổng hợp.

    Quy đổi MỌI period_key thô (có thể là "YYYY-Qn" HOẶC "YYYY-FY") sang quý chuẩn hóa qua
    gap_period_to_quarter() TRƯỚC khi gom — nếu không, "2025-Q4" (ngân hàng A, báo cáo quý thường)
    và "2025-FY" (ngân hàng B, báo cáo năm kiểm toán) sẽ bị coi là 2 KỲ HỆ THỐNG KHÁC NHAU dù cùng
    phản ánh 1 thời điểm, khiến mỗi kỳ "thiếu" đúng những ngân hàng đang dùng kiểu báo cáo còn lại."""
    from bank_universe import BANKING_TICKERS
    all_periods = set()
    for ticker in BANKING_TICKERS:
        store = bank_alm_store.load_bank_store(ticker)
        for pk in store.get("gap_periods", {}).keys():
            canon = bank_alm_store.gap_period_to_quarter(pk)
            if canon:
                all_periods.add(canon)
    return {pk: _aggregate_for_period(pk) for pk in sorted(all_periods, key=bank_alm_store._period_sort_key)}


def _classify_resilience_ratio(ratio_abs):
    if ratio_abs is None:
        return "Không xác định"
    if ratio_abs < 0.05:
        return "TỐT"
    if ratio_abs < 0.15:
        return "VỪA PHẢI"
    return "CẦN LƯU Ý"


def _classify_resilience_coverage(coverage):
    if coverage is None:
        return "Không xác định"
    if coverage >= 1.0:
        return "TỐT"
    if coverage >= 0.5:
        return "VỪA PHẢI"
    return "CẦN LƯU Ý"


def classify_structural_funding_phase(history):
    """Đọc "pha" áp lực cấu trúc kỳ hạn nguồn vốn theo "Danh gia rui ro thanh khoan cau truc he
    thong.docx" (user 2026-09-21) — KHÔNG nhìn mức tuyệt đối 1 kỳ (1 kỳ chỉ cho biết trạng thái TẠI
    thời điểm đó, không cho biết HƯỚNG ĐI, xem mục XII file hướng dẫn), mà nhìn HƯỚNG THAY ĐỔI của 2
    biến rollover_dependency_12m và long_term_funding_coverage qua các kỳ liên tiếp gần nhất — ý
    tưởng cốt lõi: ngân hàng/hệ thống đang thiếu nguồn vốn dài hạn → phải cạnh tranh kéo dài kỳ hạn
    tiền gửi → áp lực đẩy lãi suất huy động kỳ hạn dài tăng; ngược lại khi mismatch thu hẹp thì áp
    lực đó giảm (mục XI, XVII).

    `history`: list [{"period", "rollover_dependency_12m", "long_term_funding_coverage",
    "coverage_pct"}, ...] ĐÃ SẮP XẾP tăng dần theo thời gian (period cũ nhất trước) — dùng TỐI ĐA 3
    điểm CUỐI (mục XII file hướng dẫn: "ông phải lấy ít nhất Q4/2025 -> Q1/2026 -> Q2/2026"). Trả về
    None nếu < 2 điểm hợp lệ (chưa đủ để biết hướng đi).

    BUG THẬT phát hiện 2026-09-21 (user nghi ngờ đúng khi thấy rollover 12M "rớt" 28% -> 1,4% chỉ
    trong 1 quý — LS huy động thực tế vẫn cao, không khớp câu chuyện "đã qua đỉnh"): kỳ MỚI NHẤT
    (2026-Q2) chỉ có 12/26 ngân hàng có đủ "liabilities_by_bucket" (thiếu hẳn MBB/VCB/CTG/HDB/VIB...
    — các NH lớn CHƯA backfill kịp bảng thanh khoản cho kỳ mới nhất), so với 20/26 của kỳ trước —
    tỷ lệ "giảm mạnh" đó là do SO SÁNH 2 MẪU NGÂN HÀNG KHÁC NHAU (coverage_pct sụt), KHÔNG PHẢI cấu
    trúc hệ thống thực sự cải thiện. Giờ LOẠI các điểm coverage quá thấp (< 60% tổng tài sản đã biết
    — mẫu quá nhỏ để đại diện hệ thống) VÀ cảnh báo RÕ khi coverage lệch quá nhiều (>15 điểm %) giữa
    điểm đầu/cuối chuỗi dùng để so sánh — 2 mẫu khác nhau không thể dùng để kết luận HƯỚNG ĐI."""
    MIN_COVERAGE_PCT = 60.0
    usable = [h for h in history if h.get("rollover_dependency_12m") is not None
              and h.get("long_term_funding_coverage") is not None
              and (h.get("coverage_pct") or 0) >= MIN_COVERAGE_PCT]
    pts = usable[-3:]
    if len(pts) < 2:
        latest = history[-1] if history else None
        if latest and latest.get("coverage_pct") is not None and latest["coverage_pct"] < MIN_COVERAGE_PCT:
            return {"phase": "low_coverage", "phaseLabel": "Chưa đủ dữ liệu để xác định pha",
                    "narrative": (f"Kỳ {latest['period']} mới có {latest['coverage_pct']:.0f}% tổng tài sản hệ "
                                  f"thống có đủ dữ liệu bảng Nợ phải trả theo kỳ hạn (nhiều ngân hàng lớn chưa "
                                  f"backfill kịp) — CHƯA ĐỦ đại diện để so sánh xu hướng, cần đợi backfill đầy đủ "
                                  f"hơn thay vì kết luận từ 1 mẫu nhỏ/khác kỳ trước."),
                    "periodsUsed": [latest["period"]]}
        return None
    roll_seq = [p["rollover_dependency_12m"] for p in pts]
    ltfc_seq = [p["long_term_funding_coverage"] for p in pts]
    cov_seq = [p.get("coverage_pct") for p in pts]
    roll_delta = roll_seq[-1] - roll_seq[0]
    ltfc_delta = ltfc_seq[-1] - ltfc_seq[0]
    # Xu huong tung buoc gan nhat (kỳ cuối so kỳ ngay truoc) — dung phan biet Pha 2 (dinh, moi bat
    # dau dao chieu) voi Pha 1 (van con xau di) khi so 3 diem chi cho xu huong tong the.
    roll_last_step = roll_seq[-1] - roll_seq[-2]
    # Canh bao lech coverage giua diem dau/cuoi dung de so sanh (2 mau ngan hang khac nhau -> huong
    # di tinh duoc KHONG dang tin, du tung diem rieng le da qua MIN_COVERAGE_PCT).
    cov_known = [c for c in cov_seq if c is not None]
    coverage_mismatch_note = ""
    if len(cov_known) >= 2 and abs(cov_known[-1] - cov_known[0]) > 15.0:
        coverage_mismatch_note = (f" (Lưu ý: độ phủ dữ liệu lệch khá nhiều giữa các kỳ so sánh — "
                                   f"{cov_known[0]:.0f}% -> {cov_known[-1]:.0f}% tổng tài sản — nên đọc "
                                   f"hướng đi này với mức độ tin cậy VỪA PHẢI, không phải chắc chắn.)")

    if roll_delta > 0.02 and ltfc_delta < -0.02:
        phase, label = "deterioration", "Pha 1 — Áp lực cấu trúc ĐANG TĂNG"
        narrative = (f"Rollover dependency 12M tăng từ {roll_seq[0]*100:.1f}% lên {roll_seq[-1]*100:.1f}%, "
                     f"Long-term Funding Coverage giảm từ {ltfc_seq[0]*100:.1f}% xuống {ltfc_seq[-1]*100:.1f}% "
                     f"— hệ thống ngày càng cần rollover/huy động thêm nguồn vốn dài hạn, áp lực cạnh tranh "
                     f"lãi suất huy động kỳ hạn dài có xu hướng TĂNG.")
    elif roll_delta > 0 and roll_last_step < 0:
        phase, label = "peak", "Pha 2 — Có dấu hiệu ĐÃ QUA ĐỈNH căng thẳng"
        narrative = (f"Rollover dependency 12M vẫn ở mức {roll_seq[-1]*100:.1f}% (cao hơn {roll_seq[0]*100:.1f}% "
                     f"của kỳ đầu chuỗi) nhưng đã GIẢM so với kỳ ngay trước — dấu hiệu SỚM cho thấy hệ thống "
                     f"đang đi qua đỉnh căng thẳng cấu trúc và bắt đầu tái cân bằng. Chưa thể khẳng định lãi "
                     f"suất huy động kỳ hạn dài sẽ giảm ngay, nhưng động lượng tăng thêm đã chững lại.")
    elif roll_delta < -0.02 and ltfc_delta > -0.02:
        phase, label = "normalization", "Pha 3 — Áp lực cấu trúc ĐANG GIẢM (chuẩn hóa)"
        narrative = (f"Rollover dependency 12M giảm từ {roll_seq[0]*100:.1f}% xuống {roll_seq[-1]*100:.1f}%, "
                     f"Long-term Funding Coverage {'tăng' if ltfc_delta >= 0 else 'ổn định'} — hệ thống không "
                     f"còn cần rollover ngắn hạn nhiều như trước, có điều kiện để giảm mức độ cạnh tranh huy "
                     f"động kỳ hạn dài. Đây là tín hiệu có thể dùng để nhận diện đỉnh lãi suất huy động đã đi qua.")
    else:
        phase, label = "mixed", "Chưa rõ xu hướng (tín hiệu hỗn hợp)"
        narrative = (f"Rollover dependency 12M: {roll_seq[0]*100:.1f}% → {roll_seq[-1]*100:.1f}%; "
                     f"Long-term Funding Coverage: {ltfc_seq[0]*100:.1f}% → {ltfc_seq[-1]*100:.1f}% — chưa đủ "
                     f"rõ ràng để xác định pha, cần theo dõi thêm ít nhất 1 kỳ nữa.")
    return {"phase": phase, "phaseLabel": label, "narrative": narrative + coverage_mismatch_note,
            "periodsUsed": [p["period"] for p in pts], "coveragePctByPeriod": cov_seq}


def build_system_risk_summary_text(agg, phase_info=None):
    """Bản tóm tắt 2-3 câu cho 1 kỳ đã tổng hợp (agg = kết quả _aggregate_for_period) — cùng văn
    phong build_risk_summary() của bank_risk_notes.py (đánh giá theo KẾT QUẢ STRESS TEST thực tế,
    không chỉ nhìn gap ròng thô), dùng cho cả mục riêng "Rủi ro hệ thống ngân hàng" lẫn câu tóm tắt
    lồng trong watch_points của phần tổng hợp vĩ mô.

    `phase_info`: dict trả về từ classify_structural_funding_phase() (có thể None nếu chưa đủ ≥2 kỳ
    lịch sử) — thêm 1 câu về PHA áp lực cấu trúc kỳ hạn nguồn vốn/lãi suất huy động nếu có."""
    if not agg:
        return None
    ir, liq = agg["interest_rate_risk"], agg["liquidity_risk"]
    sf = agg.get("structural_funding") or {}
    parts = []
    if ir.get("net_gap_ratio") is not None:
        worst_ratio = ir["stress_nii_ratio"].get("+200bp")
        resil = _classify_resilience_ratio(abs(worst_ratio)) if worst_ratio is not None else "Không xác định"
        extra = f" (sốc +200bp làm NII hệ thống đổi khoảng {abs(worst_ratio)*100:.1f}%)" if worst_ratio is not None else ""
        wb = ir.get("worst_bank")
        wb_s = f" Ngân hàng lệch nhiều nhất: {wb['ticker']} ({wb['ratio']*100:+.1f}%)." if wb else ""
        # Canh bao khi do phu THUC SU (sau khi loai boi guard kiem tra cheo lai suat/thanh khoan)
        # qua thap (user 2026-09-23) - cung nguyen tac cov_s cua structural_funding o duoi.
        ir_cov = ir.get("coverage_pct")
        ir_cov_s = (f" (LƯU Ý: sau khi loại các NH có dữ liệu bất thường, chỉ {ir_cov:.0f}% tổng tài "
                    f"sản hệ thống còn được tính vào tỷ lệ này — CHƯA đại diện đầy đủ toàn hệ thống.)"
                    ) if ir_cov is not None and ir_cov < 60 else ""
        parts.append(
            f"Rủi ro lãi suất hệ thống: gap ròng ≤1 năm = {ir['net_gap_ratio']*100:+.2f}% tổng tài sản "
            f"(mức phân tán {ir['dispersion_gap_ratio']*100:.2f}% — phần bù trừ giữa các ngân hàng KHÔNG "
            f"thực sự phòng hộ lẫn nhau vì là các pháp nhân riêng biệt), khả năng chống chịu {resil}{extra}."
            f"{wb_s}{ir_cov_s}"
        )
    if liq.get("net_gap_1m_ratio") is not None:
        cov20 = liq["deposit_run_coverage"].get("-20%")
        resil = _classify_resilience_coverage(cov20) if cov20 is not None else "Không xác định"
        extra = f" (che phủ {cov20*100:.0f}% ở kịch bản rút -20% tiền gửi)" if cov20 is not None else ""
        wb = liq.get("weakest_bank")
        wb_s = f" Ngân hàng thanh khoản yếu nhất: {wb['ticker']} (che phủ -10%: {wb['coverage']*100:.0f}%)." if wb else ""
        liq_cov = liq.get("coverage_pct")
        liq_cov_s = (f" (LƯU Ý: sau khi loại các NH có dữ liệu bất thường, chỉ {liq_cov:.0f}% tổng tài "
                     f"sản hệ thống còn được tính vào tỷ lệ này — CHƯA đại diện đầy đủ toàn hệ thống.)"
                     ) if liq_cov is not None and liq_cov < 60 else ""
        parts.append(
            f"Rủi ro thanh khoản hệ thống: Liquid Assets/Tổng TS = {liq['liquid_assets_ratio']*100:.1f}%, "
            f"khả năng chống chịu {resil}{extra}.{wb_s}{liq_cov_s}"
        )
    if sf.get("rollover_dependency_12m") is not None:
        mdb = sf.get("most_dependent_bank")
        mdb_s = (f" NH phụ thuộc rollover nhiều nhất: {mdb['ticker']} "
                 f"({mdb['rollover_dependency_12m']*100:.0f}%).") if mdb else ""
        # Canh bao RO khi do phu rieng cho cau truc ky han qua thap (bug that phat hien 2026-09-21) —
        # so nay du tinh dung TOAN HOC nhung dua tren 1 mau ngan hang chua day du (nhieu NH lon chua
        # backfill kip bang "No phai tra" theo bucket cho ky moi nhat), KHONG nen doc nhu so lieu
        # dai dien toan he thong.
        cov_pct = sf.get("coverage_pct")
        cov_s = (f" (LƯU Ý: chỉ {cov_pct:.0f}% tổng tài sản hệ thống có đủ dữ liệu bảng Nợ phải trả "
                 f"theo kỳ hạn cho kỳ này — số liệu 2 chỉ số này CHƯA đại diện đầy đủ toàn hệ thống.)"
                 ) if cov_pct is not None and cov_pct < 60 else ""
        parts.append(
            f"Cấu trúc kỳ hạn nguồn vốn: Rollover Dependency 12 tháng = {sf['rollover_dependency_12m']*100:.1f}% "
            f"(tỷ lệ nghĩa vụ đến hạn ≤12 tháng KHÔNG được tài sản cùng kỳ hạn tự tài trợ, buộc phải huy "
            f"động mới/rollover), Long-term Funding Coverage = {sf['long_term_funding_coverage']*100:.1f}%."
            f"{mdb_s}{cov_s}"
        )
    if phase_info:
        parts.append(f"{phase_info['phaseLabel']}: {phase_info['narrative']}")
    coverage_bits = [f"{agg['n_banks_reported']}/{agg['n_banks_total']} ngân hàng đã công bố kỳ {agg['period']}"]
    if agg["n_banks_patched"]:
        coverage_bits.append(f"{agg['n_banks_patched']} đang dùng số liệu kỳ trước (chưa công bố)")
    if agg["n_banks_missing"]:
        coverage_bits.append(f"{agg['n_banks_missing']} chưa có dữ liệu")
    cov_pct = agg.get("assets_coverage_pct")
    cov_pct_s = f" — che phủ {cov_pct:.0f}% tổng tài sản hệ thống" if cov_pct is not None else ""
    parts.append("Độ phủ dữ liệu: " + ", ".join(coverage_bits) + cov_pct_s + ".")
    return " ".join(parts)


def build_banking_system_risk_section(agg, history=None):
    """Xây dict cho field JSON top-level "bankingSystemRisk" (mục RIÊNG trong data/vimo.json) từ 1
    kỳ đã tổng hợp. Trả về None nếu agg rỗng/không có ngân hàng nào có dữ liệu.

    `history`: list các agg CŨ HƠN (kể cả agg hiện tại ở cuối, xem classify_structural_funding_phase)
    — tuỳ chọn, dùng để tính PHA áp lực cấu trúc kỳ hạn nguồn vốn/lãi suất huy động (xem "Danh gia
    rui ro thanh khoan cau truc he thong.docx", user 2026-09-21)."""
    if not agg or (agg["n_banks_reported"] + agg["n_banks_patched"]) == 0:
        return None
    ir, liq, sf = agg["interest_rate_risk"], agg["liquidity_risk"], agg.get("structural_funding") or {}
    phase_info = None
    if history:
        hist_points = [{"period": h["period"], **(h.get("structural_funding") or {})} for h in history]
        phase_info = classify_structural_funding_phase(hist_points)
    return {
        "asOf": agg["period"],
        "summaryText": build_system_risk_summary_text(agg, phase_info=phase_info),
        "interestRateRisk": {
            "netGapRatio": ir["net_gap_ratio"], "dispersionGapRatio": ir["dispersion_gap_ratio"],
            "stressNiiByShock": ir["stress_nii"], "stressNiiRatioByShock": ir["stress_nii_ratio"],
            "worstBank": ir["worst_bank"],
            "coveragePct": ir.get("coverage_pct"), "nBanksIncluded": ir.get("n_banks_included"),
        },
        "liquidityRisk": {
            "netGapRatio": liq["net_gap_1m_ratio"], "liquidAssetsRatio": liq["liquid_assets_ratio"],
            "depositRunCoverageByStress": liq["deposit_run_coverage"], "weakestBank": liq["weakest_bank"],
            "coveragePct": liq.get("coverage_pct"), "nBanksIncluded": liq.get("n_banks_included"),
        },
        "structuralFunding": {
            "rolloverDependency1m": sf.get("rollover_dependency_1m"),
            "rolloverDependency3m": sf.get("rollover_dependency_3m"),
            "rolloverDependency12m": sf.get("rollover_dependency_12m"),
            "longTermFundingCoverage": sf.get("long_term_funding_coverage"),
            "longTermStructuralGap": sf.get("long_term_structural_gap"),
            "mostDependentBank": sf.get("most_dependent_bank"),
            "coveragePct": sf.get("coverage_pct"), "nBanksIncluded": sf.get("n_banks_included"),
            "missingTickers": sf.get("missing_tickers"),
            "phase": phase_info,
        },
        "coverage": {
            "nBanksTotal": agg["n_banks_total"], "nBanksReported": agg["n_banks_reported"],
            "nBanksPatched": agg["n_banks_patched"], "nBanksMissing": agg["n_banks_missing"],
            "assetsCoveragePct": agg["assets_coverage_pct"],
            "reportedTickers": agg["reported_tickers"], "patchedTickers": agg["patched_tickers"],
            "missingTickers": agg["missing_tickers"],
        },
        "byBank": agg["by_bank"],
    }


# ── Sheet Excel riêng lưu dữ liệu THÔ per-bank/per-kỳ (yêu cầu user 2026-08-30, mục 7 kế hoạch) ──

_ALM_SHEET_NAME = "ALM_NganHang_Raw"
# SUA (user 2026-09-18): cot "Trang thai" chung TRUOC DAY co the gay hieu lam — status="reported"
# chi can 1 TRONG 2 bang (lai suat/thanh khoan) thanh cong la du (xem is_fully_reported()), khien
# nguoi dung tuong da co du du lieu du that ra chi co 1 nua (bug that phat hien qua VIB/VAB). Them 2
# cot trang thai RIENG BIET theo tung loai du lieu (lai suat/thanh khoan), tinh TRUC TIEP tu viec co
# du lieu hay khong (khong phu thuoc status tong the) - nguoi dung nhin thang vao sheet la biet
# CHINH XAC dang thieu gi, khong can doan qua status chung.
_ALM_SHEET_HEADERS = [
    "Ma", "Ky", "Trang thai", "Trang thai Lai suat", "Trang thai Thanh khoan",
    # SUA (user 2026-09-23): "Trang thai" o tren CHI cho biet CO du lieu hay khong, KHONG cho biet
    # du lieu do co KHOP NOI BO khong - nhieu hang "reported" van dang bi loai khoi tong hop he
    # thong vi 1 trong 2 bang sai (xem guard kiem tra cheo trong _aggregate_for_period), nhung nhin
    # vao sheet chi thay toan "reported" gay hieu nham. Them cot rieng hien TRUC TIEP ket qua kiem
    # tra nay (tong lai suat vs tong thanh khoan cung ky PHAI xap xi bang nhau).
    "Kiem tra cheo LS/TK",
    # 3 cot GIA TRI SO rieng (user 2026-09-23, muon so sanh truc quan khong can doc chuoi chu trong
    # cot Kiem tra cheo o tren) - TS suy ra tu bang thanh khoan, TS suy ra tu bang lai suat, TTS
    # THAT lay tu Vietcap (bang can doi) - ca 3 cung don vi TY DONG.
    "TS TK (ty)", "TS LS (ty)", "TTS thuc - Vietcap (ty)",
    "Va tu ky", "Tong tai san (ty)", "VCSH (ty)", "NII (ty)",
    "Tien gui KH (ty)", "Cho vay KH (ty)", "LDR (%)", "Liquid Assets (ty)",
    # -- Rui ro thanh khoan --
    "Gap thanh khoan rong <=1thang (ty)", "Gap thanh khoan/TTS <=1thang (%)",
    "Gap thanh khoan/TTS <=3thang (%)", "Gap thanh khoan/TTS <=1nam (%)",
    "Gap thanh khoan/TTS <=1nam THAN TRONG (%)",
    "Short-term Funding (ty)", "Long-term Assets (ty)", "ST Funding/LT Assets (%)",
    "NSFR proxy (%)", "LMI (%)",
    # -- Cau truc ky han nguon von (Rollover Dependency, xem "Danh gia rui ro thanh khoan cau truc
    # he thong.docx") - NSFR proxy o tren da chinh la Long-term Funding Coverage cua file huong dan,
    # KHONG them cot trung lap, chi them 3 cot Rollover Dependency thuc su MOI.
    "Rollover Dependency 1thang (%)", "Rollover Dependency 3thang (%)", "Rollover Dependency 12thang (%)",
    "A/L <=1thang (%)", "A/L 1-3thang (%)", "A/L 3-12thang (%)",
    "Buffer thanh khoan/No <=1thang (lan)", "Tai san gan tien (Tien mat+NHNN)/No <=1thang (lan)",
    "Tien gui KH <=1thang/Tong tien gui (%)", "Tien gui KH <=1nam/Tong tien gui (%)",
    "Che phu rut -5% tien gui (lan)", "Che phu rut -10% tien gui (lan)", "Che phu rut -20% tien gui (lan)",
    # -- Rui ro lai suat --
    "Gap lai suat rong <=1nam (ty)", "Gap lai suat/TTS <=1thang (%)", "Gap lai suat/TTS <=3thang (%)",
    "Gap lai suat/TTS <=6thang (%)", "Gap lai suat/TTS <=1nam (%)", "Gap lai suat/TTS <=5nam (%)",
    "RSA (ty)", "RSL (ty)", "RSA/RSL (%)",
    "Stress NII +100bp (ty)", "Stress NII +100bp/NII (%)", "Stress NII +100bp/VCSH (%)",
    "Stress NII +200bp (ty)", "Stress NII +200bp/NII (%)",
    "Nguon (tieu de)", "Nguon (url)", "Cap nhat luc",
]


def _alm_implied_total_assets(gap_d, liab_d):
    """Tong tai san SUY RA tu 1 bang (Gap + No phai tra, ca 2 deu tinh tu CHINH tai lieu BCTC that -
    khong qua don vi/quy uoc ben ngoai) - tra ve (gia_tri_ty_dong, None) hoac (None, None) neu thieu
    du lieu. Tra ve don vi TY DONG (da chia 1000 tu trieu dong luu trong store)."""
    if not gap_d or not liab_d:
        return None
    gap_sum = sum(v or 0 for v in gap_d.values()) / 1000
    liab_sum = sum(v or 0 for v in liab_d.values()) / 1000
    return gap_sum + liab_sum


def _alm_cross_check(entry, status_ls, status_tk, ta_ty):
    """Tra ve (label, ts_tk, ts_ls) cho cot "Kiem tra cheo LS/TK" + 2 cot gia tri so rieng "TS TK
    (ty)"/"TS LS (ty)" - dung CHUNG cho ca 2 sheet Excel (Raw va RawBuckets). ts_tk/ts_ls la None neu
    khong tinh duoc (thieu No phai tra theo bucket o ben do).

    SUA (user 2026-09-23, phan hoi qua vi du HDB Q2/2026): kiem tra CHINH phai la so TONG TAI SAN
    SUY RA tu bang Lai suat vs SUY RA tu bang Thanh khoan VOI NHAU (ca 2 deu tu CHINH 1 tai lieu BCTC,
    cung don vi/quy uoc trinh bay - so sanh nay dang tin cay nhat, KHONG le thuoc nguon ben ngoai).
    Tong tai san THAT tu Vietcap (ta_ty) CHI dung lam THAM KHAO PHU de doan ben nao co kha nang dung
    hon khi 2 ben lech nhau (KHONG dung lam chuan chinh vi co the lech ~1-2% do quy uoc gop/rong du
    phong rui ro khac nhau giua Vietcap va thuyet minh - da xac nhan qua vi du ACB truoc do), tuyet
    doi KHONG tu bao "SAI" chi vi lech so voi Vietcap khi ca 2 bang LS/TK van khop nhau tot."""
    if status_ls != "reported" or status_tk != "reported":
        return "", None, None
    ir_gap = entry.get("interest_rate_gap") or {}
    liq_gap = entry.get("liquidity_gap") or {}
    ta_ls = _alm_implied_total_assets(ir_gap, entry.get("interest_rate_liabilities_by_bucket"))
    ta_tk = _alm_implied_total_assets(liq_gap, entry.get("liabilities_by_bucket"))

    if ta_ls is not None and ta_tk is not None:
        denom = max(abs(ta_ls), abs(ta_tk), 0.01)
        diff_pct = abs(ta_ls - ta_tk) / denom * 100
        if diff_pct <= 10:
            return f"OK (TS LS={ta_ls:,.0f} ~ TS TK={ta_tk:,.0f} ty)", ta_tk, ta_ls
        # Lech nhau - dung ta_ty (Vietcap) CHI de GOI Y ben nao gan thuc te hon, khong khang dinh
        hint = ""
        if ta_ty:
            d_ls = abs(ta_ls - ta_ty)
            d_tk = abs(ta_tk - ta_ty)
            hint = " -> nghi TK sai" if d_ls < d_tk else (" -> nghi LS sai" if d_tk < d_ls else "")
        label = f"LECH {diff_pct:.0f}%: TS LS={ta_ls:,.0f} vs TS TK={ta_tk:,.0f} ty (TS thuc~{ta_ty:,.0f} ty){hint}"
        return label, ta_tk, ta_ls

    # Thieu No phai tra theo bucket o >=1 ben - khong tinh duoc Tong tai san suy ra, fallback ve so
    # tong Gap don (kem canh bao ro la khong chac chan bang nao sai).
    ir_sum_ty = sum((v or 0) / 1000 for v in ir_gap.values())
    liq_sum_ty = sum((v or 0) / 1000 for v in liq_gap.values())
    if not (ir_sum_ty or liq_sum_ty):
        return "", ta_tk, ta_ls
    _denom = max(abs(ir_sum_ty), abs(liq_sum_ty), 0.01)
    _diff_pct = abs(ir_sum_ty - liq_sum_ty) / _denom * 100
    label = ("OK" if _diff_pct <= 20 else
             f"LECH {_diff_pct:.0f}% Gap LS/TK - thieu No phai tra theo bucket nen chua tinh duoc Tong tai san de xac dinh ben nao sai")
    return label, ta_tk, ta_ls


def update_bank_alm_excel_sheet(out_dir):
    """Ghi sheet "ALM_NganHang_Raw" (dạng tidy/long: 1 hàng = 1 (ngân hàng, kỳ)) trong CÙNG workbook
    VIMO_Lich_Su_Chi_So.xlsx (out_dir/VIMO_Lich_Su_Chi_So.xlsx) trực tiếp từ data/bank_alm/, để dữ
    liệu THÔ (gap buckets quy đổi tỷ đồng + snapshot bảng cân đối cùng kỳ) có thể tái sử dụng cho
    phân tích từng mã lẻ mà KHÔNG cần OCR lại (yêu cầu user 2026-08-30).

    GHI ĐÈ TOÀN BỘ sheet mỗi lần chạy (khác update_excel_history_vimo chỉ APPEND CỘT MỚI) — đây là
    bảng tra cứu theo hàng (ticker, kỳ), không phải chuỗi thời gian theo cột, và luôn dựng lại từ
    trạng thái MỚI NHẤT của store (rẻ — vài chục file JSON nhỏ) để không sót hàng cũ/lệch khi 1 kỳ
    "patched" sau đó chuyển thành "reported" thật (cùng 1 (ticker, period_key), giá trị đổi nhưng
    key không đổi — ghi đè cả sheet đảm bảo luôn phản ánh trạng thái mới nhất, không cần dò-sửa
    từng ô như update_excel_history_vimo phải làm với chuỗi append-cột)."""
    import openpyxl
    from bank_universe import BANKING_TICKERS

    xlsx_path = os.path.join(out_dir, "VIMO_Lich_Su_Chi_So.xlsx")
    if os.path.exists(xlsx_path):
        wb = openpyxl.load_workbook(xlsx_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    if _ALM_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[_ALM_SHEET_NAME])
    ws = wb.create_sheet(title=_ALM_SHEET_NAME)
    for c, h in enumerate(_ALM_SHEET_HEADERS, start=1):
        ws.cell(row=1, column=c, value=h)

    row_idx = 2
    for ticker in sorted(BANKING_TICKERS):
        store = bank_alm_store.load_bank_store(ticker)
        gap_periods = store.get("gap_periods", {})
        qbs = store.get("quarterly_balance_sheet", {})
        # SUA (user 2026-09-23, phat hien qua sheet OCB: hang "2025-FY" hien rieng biet voi "2025-Q4"
        # khien nhin nham thanh 2 ky khac nhau, 1 ben "missing" 1 ben co du lieu that): "-FY" (bao cao
        # nam kiem toan) va "-Q4" (bao cao quy thuong) CUNG la 1 thoi diem cuoi nam, chi khac LAN
        # CONG BO (da co _ticker_gap_entry() gop 2 cai nay lam 1 khi tong hop he thong - o day ap
        # dung DUNG nguyen tac do cho tung hang chi tiet: chi hien DUY NHAT 1 hang "Q4" moi nam, uu
        # tien ban FY neu da co (dang tin hon), roi moi den ban Q4 thuong, khong con hang "-FY" rieng
        # nua). Quy Q1-Q3 giu nguyen, khong doi.
        years = sorted({int(pk.split("-")[0]) for pk in gap_periods.keys()})
        for year in years:
            for q in (1, 2, 3, 4):
                canonical_pk = f"{year}-Q{q}"
                if q == 4:
                    fy_entry = gap_periods.get(f"{year}-FY")
                    q4_entry = gap_periods.get(canonical_pk)
                    if fy_entry and fy_entry.get("status") != "missing":
                        entry = fy_entry
                    elif q4_entry is not None:
                        entry = q4_entry
                    else:
                        entry = fy_entry  # co the None (khong co du lieu ky nay) hoac "missing"
                else:
                    entry = gap_periods.get(canonical_pk)
                if entry is None:
                    continue
                period_key = canonical_pk
                status = entry.get("status")
                bs_snap = qbs.get(period_key)
                m = _bank_period_metrics(entry, bs_snap)
                source = entry.get("source") or {}

                def _pct(key):
                    v = m.get(key)
                    return round(v * 100, 3) if v is not None else None

                def _rnd(key, nd=3):
                    v = m.get(key)
                    return round(v, nd) if v is not None else None

                status_ls = "reported" if entry.get("interest_rate_gap") else "missing"
                status_tk = "reported" if entry.get("liquidity_gap") else "missing"
                # Cot C tong hop: chi xet Lai suat + Thanh khoan. "patched" giu nguyen rieng (du lieu ke
                # thua tu ky truoc, khong phai dang thieu can OCR lai).
                if status == "patched":
                    status_overall = "patched"
                elif status_ls == "reported" and status_tk == "reported":
                    status_overall = "reported"
                else:
                    status_overall = "missing"

                # Kiem tra cheo: so Tong tai san suy ra tu bang LS vs bang TK - xem _alm_cross_check().
                # Hien TRUC TIEP o day de biet dong nao "reported" nhung THUC RA dang bi loai khoi
                # tong hop he thong.
                cross_check, ts_tk, ts_ls = _alm_cross_check(entry, status_ls, status_tk, m.get("total_assets"))

                row = [
                    ticker, period_key, status_overall, status_ls, status_tk, cross_check,
                    round(ts_tk, 3) if ts_tk is not None else None, round(ts_ls, 3) if ts_ls is not None else None,
                    m.get("total_assets"), entry.get("patched_from"),
                    m.get("total_assets"), m.get("equity"), m.get("nii"), m.get("customer_deposits"),
                    m.get("loans"), _pct("ldr"), m.get("liquid_assets"),
                    # -- Rui ro thanh khoan --
                    m.get("cum_gap_1m"), _pct("cum_gap_1m_ratio"), _pct("liq_cum_gap_3m_ratio"),
                    _pct("liq_cum_gap_1y_ratio"), _pct("liq_cum_gap_1y_conservative_ratio"),
                    _rnd("short_term_funding"), _rnd("long_term_assets"),
                    _pct("st_funding_lt_assets_ratio"), _pct("nsfr_proxy"), _pct("lmi"),
                    _pct("rollover_dependency_1m"), _pct("rollover_dependency_3m"), _pct("rollover_dependency_12m"),
                    _pct("liq_al_ratio_1m"), _pct("liq_al_ratio_3m"), _pct("liq_al_ratio_12m"),
                    _rnd("liquidity_buffer_coverage_1m"), _rnd("near_cash_coverage_1m_precise"),
                    _pct("customer_deposits_pct_1m"), _pct("customer_deposits_pct_1y"),
                    _rnd("deposit_run_coverage_5pct"), _rnd("deposit_run_coverage_10pct"), _rnd("deposit_run_coverage_20pct"),
                    # -- Rui ro lai suat --
                    m.get("cum_gap_1y"), _pct("ir_cum_gap_1m_ratio"), _pct("ir_cum_gap_3m_ratio"),
                    _pct("ir_cum_gap_6m_ratio"), _pct("cum_gap_1y_ratio"), _pct("ir_cum_gap_5y_ratio"),
                    _rnd("rsa"), _rnd("rsl"), _pct("rsa_rsl_ratio"),
                    m.get("stress_nii_100bp"), _pct("stress_nii_ratio_100bp"), _pct("stress_nii_ratio_equity_100bp"),
                    m.get("stress_nii_200bp"), _pct("stress_nii_ratio_200bp"),
                    source.get("title"), source.get("url"), entry.get("fetched_at"),
                ]
                for c, val in enumerate(row, start=1):
                    ws.cell(row=row_idx, column=c, value=val)
                row_idx += 1

    for c in range(1, len(_ALM_SHEET_HEADERS) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 16

    os.makedirs(out_dir, exist_ok=True)
    wb.save(xlsx_path)
    print(f"  [OK] Sheet {_ALM_SHEET_NAME}: {row_idx - 2} hang")

    _update_bank_alm_raw_buckets_sheet(xlsx_path)


# ── Sheet DU LIEU THO THEO TUNG BUCKET (user 2026-09-18: sheet ALM_NganHang_Raw hien co chi hien ty
# le/chi so DA TINH — muon xem duoc DUNG gia tri tho OCR trich duoc cho tung ky han, de tu kiem soat
# bucket nao con thieu du lieu can bo sung, khong chi nhin ket qua tinh cuoi cung) ────────────────

_ALM_RAW_BUCKETS_SHEET_NAME = "ALM_NganHang_RawBuckets"

# Nhan ngan gon cho tung bucket (dung lam hau to ten cot) — GIU NGUYEN ten bucket goc (khong dich)
# de doi chieu truc tiep voi key trong data/bank_alm/*.json khi can tra cuu.
_LIQ_BUCKETS_ALL = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang",
                    "tu_3_12_thang", "tu_1_5_nam", "tren_5_nam"]
_IR_BUCKETS_ALL = ["qua_han", "khong_anh_huong_lai_suat", "den_1_thang", "tu_1_3_thang",
                   "tu_3_6_thang", "tu_6_12_thang", "tu_1_5_nam", "tren_5_nam"]


def _update_bank_alm_raw_buckets_sheet(xlsx_path):
    """Ghi sheet "ALM_NganHang_RawBuckets" (tidy: 1 hang = 1 (ngan hang, ky)) — hien THANG gia tri THO
    da trich duoc cho TUNG BUCKET rieng le (khac han sheet ALM_NganHang_Raw chi hien ty le/chi so DA
    TINH), de nguoi dung tu kiem tra duoc bucket/dong nao con thieu du lieu OCR can bo sung thu cong.
    Assets theo bucket = Gap + No phai tra (suy ra, KHONG phai OCR truc tiep) — de TRONG (None) neu
    thieu 1 trong 2 phia thay vi doan, dung nguyen tac "khong doan mu" xuyen suot he thong."""
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path)
    if _ALM_RAW_BUCKETS_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[_ALM_RAW_BUCKETS_SHEET_NAME])
    ws = wb.create_sheet(title=_ALM_RAW_BUCKETS_SHEET_NAME)

    # Cot "Kiem tra cheo LS/TK" + 3 cot gia tri so GIONG HET sheet ALM_NganHang_Raw (xem giai thich
    # chi tiet o _alm_cross_check()) - user (2026-09-23) chu yeu xem sheet NAY (RawBuckets) nen phai
    # co CA 2 noi, khong chi 1. Dat ten "...Tong" de khong nham voi cac cot "TS TK <bucket>"/"TS LS
    # <bucket>" rieng le da co san (Gap+No PER BUCKET, khac voi tong toan bang o day).
    headers = ["Ma", "Ky", "Trang thai", "Trang thai Lai suat", "Trang thai Thanh khoan", "Kiem tra cheo LS/TK",
               "TS TK Tong (ty)", "TS LS Tong (ty)", "TTS thuc - Vietcap (ty)"]
    for prefix, buckets in (("Gap TK", _LIQ_BUCKETS_ALL), ("No TK", _LIQ_BUCKETS_ALL), ("TS TK", _LIQ_BUCKETS_ALL),
                            ("Gap LS", _IR_BUCKETS_ALL), ("No LS", _IR_BUCKETS_ALL), ("TS LS", _IR_BUCKETS_ALL),
                            ("Tien mat", _LIQ_BUCKETS_ALL), ("Tien gui NHNN", _LIQ_BUCKETS_ALL),
                            ("Tien gui KH", _LIQ_BUCKETS_ALL)):
        headers += [f"{prefix} {b}" for b in buckets]
    headers += ["Nguon (tieu de)", "Cap nhat luc"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)

    from bank_universe import BANKING_TICKERS
    row_idx = 2
    for ticker in sorted(BANKING_TICKERS):
        store = bank_alm_store.load_bank_store(ticker)
        gap_periods = store.get("gap_periods", {})
        qbs = store.get("quarterly_balance_sheet", {})
        # SUA (user 2026-09-23): gop "-FY"/"-Q4" thanh DUY NHAT 1 hang "Q4" moi nam - cung nguyen tac
        # da ap dung o update_bank_alm_excel_sheet(), xem ghi chu chi tiet o do. Van GIU nguyen tac
        # "khong bo qua status=missing" (dong ngay duoi day) - chi khong con tach rieng "-FY" thanh
        # 1 hang nua thoi.
        years = sorted({int(pk.split("-")[0]) for pk in gap_periods.keys()})
        for year in years:
            for q in (1, 2, 3, 4):
                canonical_pk = f"{year}-Q{q}"
                if q == 4:
                    fy_entry = gap_periods.get(f"{year}-FY")
                    q4_entry = gap_periods.get(canonical_pk)
                    if fy_entry and fy_entry.get("status") != "missing":
                        entry = fy_entry
                    elif q4_entry is not None:
                        entry = q4_entry
                    else:
                        entry = fy_entry
                else:
                    entry = gap_periods.get(canonical_pk)
                if entry is None:
                    continue
                period_key = canonical_pk
                # KHONG bo qua status="missing" — chinh nhung hang nay moi la tin hieu ro nhat cho biet
                # ngan hang/ky nao dang HOAN TOAN thieu du lieu tho, dung muc dich chinh cua sheet nay.
                liq_gap = entry.get("liquidity_gap") or {}
                liq_liab = entry.get("liabilities_by_bucket") or {}
                ir_gap = entry.get("interest_rate_gap") or {}
                ir_liab = entry.get("interest_rate_liabilities_by_bucket") or {}
                cash_b = entry.get("cash_by_bucket") or {}
                sbv_b = entry.get("sbv_dep_by_bucket") or {}
                cust_dep_b = entry.get("customer_deposits_by_bucket") or {}
                source = entry.get("source") or {}

                def _ty(d, k):
                    v = d.get(k)
                    return round(v / 1000, 3) if v is not None else None

                def _assets_ty(gap_d, liab_d, k):
                    g, l = gap_d.get(k), liab_d.get(k)
                    return round((g + l) / 1000, 3) if (g is not None and l is not None) else None

                status_ls = "reported" if entry.get("interest_rate_gap") else "missing"
                status_tk = "reported" if entry.get("liquidity_gap") else "missing"
                # Cot C tong hop: xem giai thich chi tiet o update_bank_alm_excel_sheet() - chi xet Lai
                # suat + Thanh khoan.
                raw_status = entry.get("status")
                if raw_status == "patched":
                    status_overall = "patched"
                elif status_ls == "reported" and status_tk == "reported":
                    status_overall = "reported"
                else:
                    status_overall = "missing"
                bs_snap = qbs.get(period_key)
                ta_ty = bs_snap.get("total_assets") if bs_snap else None
                cross_check, ts_tk, ts_ls = _alm_cross_check(entry, status_ls, status_tk, ta_ty)
                row = [ticker, period_key, status_overall, status_ls, status_tk, cross_check,
                       round(ts_tk, 3) if ts_tk is not None else None,
                       round(ts_ls, 3) if ts_ls is not None else None, ta_ty]
                row += [_ty(liq_gap, k) for k in _LIQ_BUCKETS_ALL]
                row += [_ty(liq_liab, k) for k in _LIQ_BUCKETS_ALL]
                row += [_assets_ty(liq_gap, liq_liab, k) for k in _LIQ_BUCKETS_ALL]
                row += [_ty(ir_gap, k) for k in _IR_BUCKETS_ALL]
                row += [_ty(ir_liab, k) for k in _IR_BUCKETS_ALL]
                row += [_assets_ty(ir_gap, ir_liab, k) for k in _IR_BUCKETS_ALL]
                row += [_ty(cash_b, k) for k in _LIQ_BUCKETS_ALL]
                row += [_ty(sbv_b, k) for k in _LIQ_BUCKETS_ALL]
                row += [_ty(cust_dep_b, k) for k in _LIQ_BUCKETS_ALL]
                row += [source.get("title"), entry.get("fetched_at")]

                for c, val in enumerate(row, start=1):
                    ws.cell(row=row_idx, column=c, value=val)
                row_idx += 1

    for c in range(1, len(headers) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 14

    wb.save(xlsx_path)
    print(f"  [OK] Sheet {_ALM_RAW_BUCKETS_SHEET_NAME}: {row_idx - 2} hang")
