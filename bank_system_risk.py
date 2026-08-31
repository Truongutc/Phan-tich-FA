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

def backfill_period(period_key):
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
       _ticker_gap_entry), giờ áp dụng luôn từ bước BACKFILL/FETCH thay vì chỉ ở bước tổng hợp."""
    from bank_universe import BANKING_TICKERS
    from bank_risk_notes import fetch_bank_risk_gaps_for_period

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

    results = {}
    for ticker in sorted(BANKING_TICKERS):
        try:
            found = False
            for try_period in candidates:
                existing = bank_alm_store.get_period_entry(ticker, try_period)
                if existing and existing.get("status") == "reported":
                    print(f"  [SKIP] {ticker} {try_period}: da co du lieu that, bo qua")
                    results[ticker] = f"da_co ({try_period})"
                    found = True
                    break
                gaps = fetch_bank_risk_gaps_for_period(ticker, try_period)
                if gaps and (gaps.get("interest_rate_gap") or gaps.get("liquidity_gap")):
                    source = {"title": gaps.get("source_title"), "url": gaps.get("source_url"),
                              "fetched_year": gaps.get("fetched_year")}
                    bank_alm_store.upsert_reported_period(ticker, try_period, gaps, source)
                    results[ticker] = f"da_co_du_lieu ({try_period})"
                    found = True
                    break
            if not found:
                bank_alm_store.mark_missing_period(ticker, period_key)
                results[ticker] = "thieu"
        except Exception as e:
            print(f"  [WARN] Backfill {ticker} {period_key}: loi ({e})")
            results[ticker] = f"loi: {e}"
    print(f"[DONE] Backfill {period_key}: {results}")
    return results


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
    else:
        short_term_funding = long_term_assets_liq = stable_funding = None
        st_funding_lt_assets_ratio = nsfr_proxy = lmi = None

    if liab_ir_raw:
        liab_ir_ty = {k: (v or 0) / 1000 for k, v in liab_ir_raw.items()}
        assets_ir_ty = {k: ir_gap_ty.get(k, 0.0) + liab_ir_ty.get(k, 0.0)
                         for k in set(ir_gap_ty) | set(liab_ir_ty)}
        rsa = sum(assets_ir_ty.get(k, 0.0) for k in _IR_CUMULATIVE_ORDER)
        rsl = sum(liab_ir_ty.get(k, 0.0) for k in _IR_CUMULATIVE_ORDER)
        rsa_rsl_ratio = (rsa / rsl) if rsl else None
    else:
        rsa = rsl = rsa_rsl_ratio = None

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
        "ir_cum_gap_curve_ratio": ir_curve_ratio,
        "short_term_funding": short_term_funding, "long_term_assets": long_term_assets_liq,
        "st_funding_lt_assets_ratio": st_funding_lt_assets_ratio,
        "stable_funding": stable_funding, "nsfr_proxy": nsfr_proxy, "lmi": lmi,
        "rsa": rsa, "rsl": rsl, "rsa_rsl_ratio": rsa_rsl_ratio,
        "weighted_gap_raw": weighted,
        "stress_nii_100bp": stress_nii_100,
        "stress_nii_ratio_100bp": (stress_nii_100 / nii) if nii else None,
        "stress_nii_ratio_equity_100bp": (stress_nii_100 / equity) if equity else None,
        "deposit_run_coverage_10pct": (liquid_assets / (cust_dep * 0.10)) if cust_dep else None,
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
    all_ta_known = 0.0
    worst_ir = None
    worst_liq = None
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

        sum_ta += m["total_assets"]
        sum_ir_net += m["cum_gap_1y"]
        sum_ir_abs += abs(m["cum_gap_1y"])
        sum_weighted += m["weighted_gap_raw"]
        sum_nii += m["nii"] or 0.0
        sum_liq_1m += m["cum_gap_1m"]
        sum_liquid_assets += m["liquid_assets"]
        sum_cust_dep += m["customer_deposits"]

        bank_ir_ratio = m["cum_gap_1y_ratio"]
        bank_cov10 = m["deposit_run_coverage_10pct"]
        if bank_ir_ratio is not None and (worst_ir is None or abs(bank_ir_ratio) > abs(worst_ir[1])):
            worst_ir = (ticker, bank_ir_ratio)
        if bank_cov10 is not None and (worst_liq is None or bank_cov10 < worst_liq[1]):
            worst_liq = (ticker, bank_cov10)

        by_bank[ticker] = {
            "status": status,
            "period_used": actual_pk if status == "reported" else entry.get("patched_from"),
            "cum_gap_1y_ratio": bank_ir_ratio,
            "deposit_run_coverage_10pct": bank_cov10,
        }

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
        },
        "liquidity_risk": {
            "net_gap_1m_ratio": liq_net_ratio, "liquid_assets_ratio": liquid_assets_ratio,
            "deposit_run_coverage": deposit_coverage,
            "weakest_bank": {"ticker": worst_liq[0], "coverage": worst_liq[1]} if worst_liq else None,
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


def build_system_risk_summary_text(agg):
    """Bản tóm tắt 2-3 câu cho 1 kỳ đã tổng hợp (agg = kết quả _aggregate_for_period) — cùng văn
    phong build_risk_summary() của bank_risk_notes.py (đánh giá theo KẾT QUẢ STRESS TEST thực tế,
    không chỉ nhìn gap ròng thô), dùng cho cả mục riêng "Rủi ro hệ thống ngân hàng" lẫn câu tóm tắt
    lồng trong watch_points của phần tổng hợp vĩ mô."""
    if not agg:
        return None
    ir, liq = agg["interest_rate_risk"], agg["liquidity_risk"]
    parts = []
    if ir.get("net_gap_ratio") is not None:
        worst_ratio = ir["stress_nii_ratio"].get("+200bp")
        resil = _classify_resilience_ratio(abs(worst_ratio)) if worst_ratio is not None else "Không xác định"
        extra = f" (sốc +200bp làm NII hệ thống đổi khoảng {abs(worst_ratio)*100:.1f}%)" if worst_ratio is not None else ""
        wb = ir.get("worst_bank")
        wb_s = f" Ngân hàng lệch nhiều nhất: {wb['ticker']} ({wb['ratio']*100:+.1f}%)." if wb else ""
        parts.append(
            f"Rủi ro lãi suất hệ thống: gap ròng ≤1 năm = {ir['net_gap_ratio']*100:+.2f}% tổng tài sản "
            f"(mức phân tán {ir['dispersion_gap_ratio']*100:.2f}% — phần bù trừ giữa các ngân hàng KHÔNG "
            f"thực sự phòng hộ lẫn nhau vì là các pháp nhân riêng biệt), khả năng chống chịu {resil}{extra}."
            f"{wb_s}"
        )
    if liq.get("net_gap_1m_ratio") is not None:
        cov20 = liq["deposit_run_coverage"].get("-20%")
        resil = _classify_resilience_coverage(cov20) if cov20 is not None else "Không xác định"
        extra = f" (che phủ {cov20*100:.0f}% ở kịch bản rút -20% tiền gửi)" if cov20 is not None else ""
        wb = liq.get("weakest_bank")
        wb_s = f" Ngân hàng thanh khoản yếu nhất: {wb['ticker']} (che phủ -10%: {wb['coverage']*100:.0f}%)." if wb else ""
        parts.append(
            f"Rủi ro thanh khoản hệ thống: Liquid Assets/Tổng TS = {liq['liquid_assets_ratio']*100:.1f}%, "
            f"khả năng chống chịu {resil}{extra}.{wb_s}"
        )
    coverage_bits = [f"{agg['n_banks_reported']}/{agg['n_banks_total']} ngân hàng đã công bố kỳ {agg['period']}"]
    if agg["n_banks_patched"]:
        coverage_bits.append(f"{agg['n_banks_patched']} đang dùng số liệu kỳ trước (chưa công bố)")
    if agg["n_banks_missing"]:
        coverage_bits.append(f"{agg['n_banks_missing']} chưa có dữ liệu")
    cov_pct = agg.get("assets_coverage_pct")
    cov_pct_s = f" — che phủ {cov_pct:.0f}% tổng tài sản hệ thống" if cov_pct is not None else ""
    parts.append("Độ phủ dữ liệu: " + ", ".join(coverage_bits) + cov_pct_s + ".")
    return " ".join(parts)


def build_banking_system_risk_section(agg):
    """Xây dict cho field JSON top-level "bankingSystemRisk" (mục RIÊNG trong data/vimo.json) từ 1
    kỳ đã tổng hợp. Trả về None nếu agg rỗng/không có ngân hàng nào có dữ liệu."""
    if not agg or (agg["n_banks_reported"] + agg["n_banks_patched"]) == 0:
        return None
    ir, liq = agg["interest_rate_risk"], agg["liquidity_risk"]
    return {
        "asOf": agg["period"],
        "summaryText": build_system_risk_summary_text(agg),
        "interestRateRisk": {
            "netGapRatio": ir["net_gap_ratio"], "dispersionGapRatio": ir["dispersion_gap_ratio"],
            "stressNiiByShock": ir["stress_nii"], "stressNiiRatioByShock": ir["stress_nii_ratio"],
            "worstBank": ir["worst_bank"],
        },
        "liquidityRisk": {
            "netGapRatio": liq["net_gap_1m_ratio"], "liquidAssetsRatio": liq["liquid_assets_ratio"],
            "depositRunCoverageByStress": liq["deposit_run_coverage"], "weakestBank": liq["weakest_bank"],
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
_ALM_SHEET_HEADERS = [
    "Ma", "Ky", "Trang thai", "Va tu ky", "Tong tai san (ty)", "VCSH (ty)", "NII (ty)",
    "Tien gui KH (ty)", "Cho vay KH (ty)", "LDR (%)", "Liquid Assets (ty)",
    # -- Rui ro thanh khoan --
    "Gap thanh khoan rong <=1thang (ty)", "Gap thanh khoan/TTS <=1thang (%)",
    "Gap thanh khoan/TTS <=3thang (%)", "Gap thanh khoan/TTS <=1nam (%)",
    "Short-term Funding (ty)", "Long-term Assets (ty)", "ST Funding/LT Assets (%)",
    "NSFR proxy (%)", "LMI (%)", "Che phu rut -10% tien gui (lan)",
    # -- Rui ro lai suat --
    "Gap lai suat rong <=1nam (ty)", "Gap lai suat/TTS <=1thang (%)", "Gap lai suat/TTS <=3thang (%)",
    "Gap lai suat/TTS <=1nam (%)", "RSA (ty)", "RSL (ty)", "RSA/RSL (%)",
    "Stress NII +100bp (ty)", "Stress NII +100bp/NII (%)", "Stress NII +100bp/VCSH (%)",
    "Nguon (tieu de)", "Nguon (url)", "Cap nhat luc",
]


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
        for period_key in sorted(gap_periods.keys(), key=bank_alm_store._period_sort_key):
            entry = gap_periods[period_key]
            status = entry.get("status")
            qkey = bank_alm_store.gap_period_to_quarter(period_key)
            bs_snap = qbs.get(qkey) if qkey else None
            m = _bank_period_metrics(entry, bs_snap)
            source = entry.get("source") or {}

            def _pct(key):
                v = m.get(key)
                return round(v * 100, 3) if v is not None else None

            def _rnd(key, nd=3):
                v = m.get(key)
                return round(v, nd) if v is not None else None

            row = [
                ticker, period_key, status, entry.get("patched_from"),
                m.get("total_assets"), m.get("equity"), m.get("nii"), m.get("customer_deposits"),
                m.get("loans"), _pct("ldr"), m.get("liquid_assets"),
                # -- Rui ro thanh khoan --
                m.get("cum_gap_1m"), _pct("cum_gap_1m_ratio"), _pct("liq_cum_gap_3m_ratio"),
                _pct("liq_cum_gap_1y_ratio"), _rnd("short_term_funding"), _rnd("long_term_assets"),
                _pct("st_funding_lt_assets_ratio"), _pct("nsfr_proxy"), _pct("lmi"),
                _rnd("deposit_run_coverage_10pct"),
                # -- Rui ro lai suat --
                m.get("cum_gap_1y"), _pct("ir_cum_gap_1m_ratio"), _pct("ir_cum_gap_3m_ratio"),
                _pct("cum_gap_1y_ratio"), _rnd("rsa"), _rnd("rsl"), _pct("rsa_rsl_ratio"),
                m.get("stress_nii_100bp"), _pct("stress_nii_ratio_100bp"), _pct("stress_nii_ratio_equity_100bp"),
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
