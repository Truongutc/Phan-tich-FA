#!/usr/bin/env python3
"""
bank_system_risk.py — Tổng hợp rủi ro lãi suất/thanh khoản TOÀN HỆ THỐNG ngân hàng niêm yết/UPCoM
(26 mã, xem bank_universe.py), dùng cho phần "Rủi ro hệ thống ngân hàng" trong báo cáo Vĩ mô.

Kiến trúc (xem plan đã duyệt 2026-08-30):
- Dữ liệu THÔ (gap buckets từ 2 bảng thuyết minh BCTC + snapshot bảng cân đối) lưu RIÊNG TỪNG NGÂN
  HÀNG, TỪNG KỲ trong data/bank_alm/<TICKER>.json (bank_alm_store.py) — KHÔNG lưu số liệu tổng hợp
  hệ thống dạng lịch sử riêng, để tránh 2 nơi lưu cùng 1 con số có thể lệch nhau. Số liệu hệ thống
  LUÔN được tính lại (recompute_system_aggregate_all_periods) từ các file per-bank này.
- 2 bảng gap (lãi suất, thanh khoản) CHỈ có ở BCTC đã kiểm toán năm (FY) hoặc soát xét bán niên
  (H1) — KHÔNG có ở báo cáo quý thường. Vì vậy chỉ có ~2 điểm dữ liệu/năm cho 2 chỉ số này, khác
  với các tỷ lệ bảng cân đối thuần (LDR/CASA/NIM/Loan-Assets) có đủ số liệu theo quý thật từ Vietcap
  (không cần OCR).
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

import bank_alm_store

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

# Field code Vietcap (xem template_banking.py get_yr — CHIA /1e9 để ra tỷ đồng, khớp đơn vị dùng
# xuyên suốt template_banking.py/bank_risk_notes.py).
_BS_FIELD_MAP = {
    "total_assets": "bsa53", "equity": "bsa78", "customer_deposits": "bsb113",
    "cash": "bsa2", "sbv_dep": "bsb97", "bank_dep": "bsb98", "interbank_liab": "bsb112",
    "loans": "bsb103", "bonds_issued": "bsb116",
}

_IR_HORIZON_1Y_KEYS = ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang"]
_IR_WEIGHTS = {"den_1_thang": 11.5 / 12, "tu_1_3_thang": 10 / 12, "tu_3_6_thang": 7.5 / 12, "tu_6_12_thang": 3 / 12}
_LIQ_ST_KEYS = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang"]


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
    print(f"  [INFO] He thong ALM: ky muc tieu hien tai = {target_period}")

    actions = {}
    changed = False
    for ticker in sorted(BANKING_TICKERS):
        try:
            entry = bank_alm_store.get_period_entry(ticker, target_period)
            if entry and entry.get("status") == "reported":
                actions[ticker] = "skip_da_co"
                continue

            cp = cheap_periods.get(ticker)
            if cp == target_period:
                gaps = fetch_bank_risk_gaps_for_period(ticker, target_period)
                if gaps and (gaps.get("interest_rate_gap") or gaps.get("liquidity_gap")):
                    source = {"title": gaps.get("source_title"), "url": gaps.get("source_url"),
                              "fetched_year": gaps.get("fetched_year")}
                    bank_alm_store.upsert_reported_period(ticker, target_period, gaps, source)
                    actions[ticker] = "da_OCR_thanh_cong"
                    changed = True
                else:
                    actions[ticker] = "OCR_that_bai"
                continue

            latest_reported = bank_alm_store.latest_reported_period(ticker)
            if latest_reported:
                if entry and entry.get("status") == "patched" and entry.get("patched_from") == latest_reported:
                    actions[ticker] = "skip_da_va_dung_nguon"
                else:
                    bank_alm_store.upsert_patched_period(ticker, target_period, latest_reported)
                    actions[ticker] = f"da_va_tu_{latest_reported}"
                    changed = True
            else:
                if not entry:
                    bank_alm_store.mark_missing_period(ticker, target_period)
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
    CafeF/24hmoney không còn lưu bản cũ, xem fetch_bank_risk_gaps_for_period)."""
    from bank_universe import BANKING_TICKERS
    from bank_risk_notes import fetch_bank_risk_gaps_for_period

    results = {}
    for ticker in sorted(BANKING_TICKERS):
        try:
            existing = bank_alm_store.get_period_entry(ticker, period_key)
            if existing and existing.get("status") == "reported":
                print(f"  [SKIP] {ticker} {period_key}: da co du lieu that, bo qua")
                results[ticker] = "da_co"
                continue
            gaps = fetch_bank_risk_gaps_for_period(ticker, period_key)
            if gaps and (gaps.get("interest_rate_gap") or gaps.get("liquidity_gap")):
                source = {"title": gaps.get("source_title"), "url": gaps.get("source_url"),
                          "fetched_year": gaps.get("fetched_year")}
                bank_alm_store.upsert_reported_period(ticker, period_key, gaps, source)
                results[ticker] = "da_co_du_lieu"
            else:
                bank_alm_store.mark_missing_period(ticker, period_key)
                results[ticker] = "thieu"
        except Exception as e:
            print(f"  [WARN] Backfill {ticker} {period_key}: loi ({e})")
            results[ticker] = f"loi: {e}"
    print(f"[DONE] Backfill {period_key}: {results}")
    return results


# ── Tổng hợp hệ thống (arithmetic thuần, luôn tính lại từ đầu, không cache riêng) ────────────────

def _aggregate_for_period(period_key):
    """Tổng hợp CÓ TRỌNG SỐ THEO QUY MÔ (cộng dồn số tuyệt đối tỷ VND rồi mới chia ra tỷ lệ) — KHÔNG
    lấy trung bình cộng % của từng ngân hàng, vì cách đó coi 1 ngân hàng nhỏ ngang 1 ngân hàng lớn,
    sai lệch nghiêm trọng ý nghĩa "rủi ro của TOÀN HỆ THỐNG". Chỉ tính trên ngân hàng có
    status in ("reported","patched") VÀ có snapshot bảng cân đối cùng kỳ (quý tương ứng qua
    gap_period_to_quarter) — thiếu 1 trong 2 thì coi như thiếu dữ liệu cho kỳ này, không đoán."""
    from bank_universe import BANKING_TICKERS

    reported, patched, missing = [], [], []
    sum_ta = sum_ir_net = sum_ir_abs = sum_weighted = sum_nii = 0.0
    sum_liq_1m = sum_liquid_assets = sum_cust_dep = 0.0
    all_ta_known = 0.0
    worst_ir = None
    worst_liq = None
    by_bank = {}

    qkey = bank_alm_store.gap_period_to_quarter(period_key)

    for ticker in sorted(BANKING_TICKERS):
        store = bank_alm_store.load_bank_store(ticker)
        qbs = store.get("quarterly_balance_sheet", {})
        if qbs:
            latest_q = max(qbs.keys(), key=lambda k: (int(k.split("-Q")[0]), int(k.split("-Q")[1])))
            latest_ta = qbs[latest_q].get("total_assets")
            if latest_ta:
                all_ta_known += latest_ta

        entry = store.get("gap_periods", {}).get(period_key)
        if not entry or entry.get("status") == "missing":
            missing.append(ticker)
            by_bank[ticker] = {"status": "missing"}
            continue

        status = entry["status"]
        bs_snap = qbs.get(qkey) if qkey else None
        if not bs_snap or not bs_snap.get("total_assets"):
            by_bank[ticker] = {"status": status, "note": "thieu snapshot bang can doi cung ky"}
            continue

        if status == "reported":
            reported.append(ticker)
        else:
            patched.append(ticker)

        ta = bs_snap["total_assets"]
        ir_gap_ty = {k: v / 1000 for k, v in (entry.get("interest_rate_gap") or {}).items()}
        cum_1y = sum(ir_gap_ty.get(k, 0.0) for k in _IR_HORIZON_1Y_KEYS)
        weighted = sum(ir_gap_ty.get(k, 0.0) * _IR_WEIGHTS[k] for k in _IR_HORIZON_1Y_KEYS)
        nii = bs_snap.get("nii") or 0.0

        liq_gap_ty = {k: v / 1000 for k, v in (entry.get("liquidity_gap") or {}).items()}
        cum_1m = sum(liq_gap_ty.get(k, 0.0) for k in _LIQ_ST_KEYS)
        liquid_assets = (bs_snap.get("cash") or 0) + (bs_snap.get("sbv_dep") or 0) + (bs_snap.get("bank_dep") or 0)
        cust_dep = bs_snap.get("customer_deposits") or 0.0

        sum_ta += ta
        sum_ir_net += cum_1y
        sum_ir_abs += abs(cum_1y)
        sum_weighted += weighted
        sum_nii += nii
        sum_liq_1m += cum_1m
        sum_liquid_assets += liquid_assets
        sum_cust_dep += cust_dep

        bank_ir_ratio = (cum_1y / ta) if ta else None
        bank_cov10 = (liquid_assets / (cust_dep * 0.10)) if cust_dep else None
        if bank_ir_ratio is not None and (worst_ir is None or abs(bank_ir_ratio) > abs(worst_ir[1])):
            worst_ir = (ticker, bank_ir_ratio)
        if bank_cov10 is not None and (worst_liq is None or bank_cov10 < worst_liq[1]):
            worst_liq = (ticker, bank_cov10)

        by_bank[ticker] = {
            "status": status,
            "period_used": period_key if status == "reported" else entry.get("patched_from"),
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
    """Trả về dict {period_key: aggregate_dict} cho MỌI kỳ gap từng xuất hiện ở BẤT KỲ ngân hàng nào
    trong store — LUÔN tính lại từ đầu (rẻ, thuần cộng/chia trên vài chục file JSON nhỏ), không cache
    riêng kết quả tổng hợp."""
    from bank_universe import BANKING_TICKERS
    all_periods = set()
    for ticker in BANKING_TICKERS:
        store = bank_alm_store.load_bank_store(ticker)
        all_periods.update(store.get("gap_periods", {}).keys())
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
