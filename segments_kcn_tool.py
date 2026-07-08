#!/usr/bin/env python3
"""
segments_kcn_tool.py — Quản lý kho dữ liệu doanh thu/giá vốn theo mảng cho nhóm BĐS KCN.

Kho dữ liệu: data/segments_kcn/<TICKER>.json (xem schema trong skill bds-kcn).
Công cụ này KHÔNG tự parse OCR/PDF — dữ liệu được agent đọc từ ảnh/PDF rồi ghi thành
patch JSON, sau đó merge vào kho qua lệnh `merge`. Công cụ chỉ đảm nhiệm phần
deterministic: validate schema, đối chiếu tổng với Vietcap, merge an toàn, suy Q4.

Usage:
    python segments_kcn_tool.py validate <TICKER>
    python segments_kcn_tool.py merge <TICKER> <patch.json> [--force]
    python segments_kcn_tool.py derive-q4 <TICKER> <year> [--force]
"""
import os
import sys
import json
import argparse

# Fix Windows console encoding (cp1252 không hỗ trợ tiếng Việt)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
STORE_DIR = os.path.join(PROJECT_ROOT, "data", "segments_kcn")
CACHE_DIR = os.path.join(PROJECT_ROOT, ".cache")


def store_path(ticker):
    return os.path.join(STORE_DIR, f"{ticker.upper()}.json")


def load_store(ticker):
    path = store_path(ticker)
    if not os.path.exists(path):
        print(f"[ERROR] Không tìm thấy kho dữ liệu mảng: {path}")
        print("        Hãy tạo file này trước (xem schema trong skill bds-kcn) rồi chạy lại.")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_store(ticker, store):
    path = store_path(ticker)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    print(f"[OK] Đã lưu {path}")


def load_vietcap_cache(ticker):
    path = os.path.join(CACHE_DIR, f"{ticker.upper()}_bctc.json")
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _period_totals(period_data):
    rev = sum((v.get("revenue") or 0) for v in period_data.values())
    cogs = sum((v.get("cogs") or 0) for v in period_data.values() if v.get("cogs") is not None)
    return rev, cogs


def _vietcap_year_totals(raw, year):
    """Tổng DT thuần (isa3) và GVHB (isa4) theo năm từ Vietcap cache, đơn vị tỷ VND."""
    recs = raw["sections"]["INCOME_STATEMENT"].get("years", [])
    for r in recs:
        if r.get("yearReport") == year:
            rev = r.get("isa3")
            cogs = r.get("isa4")
            return (rev / 1e9 if rev is not None else None,
                    cogs / 1e9 if cogs is not None else None)
    return None, None


def _vietcap_quarter_totals(raw, year, quarter):
    recs = raw["sections"]["INCOME_STATEMENT"].get("quarters", [])
    for r in recs:
        if r.get("yearReport") == year and r.get("lengthReport") == quarter:
            rev = r.get("isa3")
            cogs = r.get("isa4")
            return (rev / 1e9 if rev is not None else None,
                    cogs / 1e9 if cogs is not None else None)
    return None, None


def cmd_validate(args):
    ticker = args.ticker.upper()
    store = load_store(ticker)
    raw = load_vietcap_cache(ticker)
    if raw is None:
        print(f"[WARN] Không có cache Vietcap tại .cache/{ticker}_bctc.json — chỉ kiểm tra schema, không đối chiếu tổng.")

    problems = 0
    for period_type, periods in (("yearly", store.get("yearly", {})), ("quarterly", store.get("quarterly", {}))):
        for period_key, seg_data in sorted(periods.items()):
            for seg, v in seg_data.items():
                if seg not in store.get("segments", {}):
                    print(f"[FAIL] {period_type}/{period_key}: mảng '{seg}' không có trong 'segments' định nghĩa")
                    problems += 1
                rev, cogs = v.get("revenue"), v.get("cogs")
                if rev is not None and cogs is not None and cogs > rev * 1.5:
                    print(f"[WARN] {period_type}/{period_key}/{seg}: giá vốn ({cogs}) > 1.5x doanh thu ({rev}) — kiểm tra lại")
                if "source" not in v or not v["source"]:
                    print(f"[FAIL] {period_type}/{period_key}/{seg}: thiếu trường 'source'")
                    problems += 1

            rev_sum, cogs_sum = _period_totals(seg_data)
            if raw is not None:
                if period_type == "yearly":
                    year = int(period_key)
                    vc_rev, vc_cogs = _vietcap_year_totals(raw, year)
                else:
                    year = int(period_key[:4])
                    q = int(period_key[-1])
                    vc_rev, vc_cogs = _vietcap_quarter_totals(raw, year, q)

                if vc_rev:
                    gap_pct = abs(rev_sum - vc_rev) / vc_rev * 100
                    status = "OK" if gap_pct <= 3 else ("WARN" if gap_pct <= 10 else "FAIL")
                    print(f"[{status}] {period_type}/{period_key}: Σmảng DT={rev_sum:.1f} vs Vietcap DT={vc_rev:.1f} (lệch {gap_pct:.1f}%)")
                    if status == "FAIL":
                        problems += 1
                else:
                    print(f"[INFO] {period_type}/{period_key}: không có dữ liệu Vietcap DT để đối chiếu")

                if vc_cogs and cogs_sum:
                    gap_pct = abs(cogs_sum - vc_cogs) / vc_cogs * 100
                    status = "OK" if gap_pct <= 3 else ("WARN" if gap_pct <= 10 else "FAIL")
                    print(f"[{status}] {period_type}/{period_key}: Σmảng GV={cogs_sum:.1f} vs Vietcap GV={vc_cogs:.1f} (lệch {gap_pct:.1f}%)")
                    if status == "FAIL":
                        problems += 1

    if problems:
        print(f"\n[RESULT] {problems} lỗi cần xử lý trước khi dùng dữ liệu này.")
        sys.exit(1)
    else:
        print("\n[RESULT] Kho dữ liệu hợp lệ.")


def cmd_merge(args):
    ticker = args.ticker.upper()
    store = load_store(ticker)
    with open(args.patch_file, "r", encoding="utf-8") as f:
        patch = json.load(f)

    for period_type in ("yearly", "quarterly"):
        patch_periods = patch.get(period_type, {})
        store.setdefault(period_type, {})
        for period_key, seg_data in patch_periods.items():
            existing = store[period_type].get(period_key, {})
            for seg, v in seg_data.items():
                existing_v = existing.get(seg)
                if existing_v and existing_v.get("sourceType") == "manual" and not args.force:
                    print(f"[SKIP] {period_type}/{period_key}/{seg}: giữ nguyên entry 'manual' (dùng --force để ghi đè)")
                    continue
                if existing_v and not args.force:
                    print(f"[SKIP] {period_type}/{period_key}/{seg}: đã có dữ liệu (dùng --force để ghi đè)")
                    continue
                existing[seg] = v
                print(f"[MERGE] {period_type}/{period_key}/{seg} <- {v.get('revenue')}/{v.get('cogs')} ({v.get('sourceType')})")
            store[period_type][period_key] = existing

    if "segments" in patch:
        store.setdefault("segments", {})
        for seg, meta in patch["segments"].items():
            if seg not in store["segments"]:
                store["segments"][seg] = meta

    save_store(ticker, store)


def cmd_derive_q4(args):
    ticker = args.ticker.upper()
    year = int(args.year)
    store = load_store(ticker)
    quarterly = store.get("quarterly", {})
    yearly = store.get("yearly", {})

    fy = yearly.get(str(year))
    if not fy:
        print(f"[ERROR] Chưa có dữ liệu 'yearly/{year}' trong kho — cần BCTC năm kiểm toán trước.")
        sys.exit(1)

    q_keys = [f"{year}Q{i}" for i in (1, 2, 3)]
    missing = [qk for qk in q_keys if qk not in quarterly]
    if missing:
        print(f"[ERROR] Thiếu quý {missing} — cần đủ Q1,Q2,Q3 mới suy được Q4.")
        sys.exit(1)

    q4_key = f"{year}Q4"
    existing_q4 = quarterly.get(q4_key, {})
    q4 = {}
    for seg in store.get("segments", {}):
        fy_seg = fy.get(seg, {})
        if not fy_seg:
            continue
        q_sum_rev = sum((quarterly[qk].get(seg, {}).get("revenue") or 0) for qk in q_keys)
        q_sum_cogs = sum((quarterly[qk].get(seg, {}).get("cogs") or 0) for qk in q_keys)
        rev4 = round((fy_seg.get("revenue") or 0) - q_sum_rev, 2)
        cogs4 = round((fy_seg.get("cogs") or 0) - q_sum_cogs, 2) if fy_seg.get("cogs") is not None else None

        if seg in existing_q4 and existing_q4[seg].get("sourceType") == "manual" and not args.force:
            print(f"[SKIP] {q4_key}/{seg}: giữ nguyên entry 'manual'")
            continue
        q4[seg] = {
            "revenue": rev4, "cogs": cogs4,
            "source": f"Suy từ BCTC năm {year} kiểm toán trừ Q1+Q2+Q3 (derive-q4)",
            "sourceType": "derived", "derived": True,
        }
        print(f"[DERIVE] {q4_key}/{seg}: DT={rev4}, GV={cogs4}")

    quarterly[q4_key] = q4
    store["quarterly"] = quarterly
    save_store(ticker, store)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_val = sub.add_parser("validate")
    p_val.add_argument("ticker")
    p_val.set_defaults(func=cmd_validate)

    p_merge = sub.add_parser("merge")
    p_merge.add_argument("ticker")
    p_merge.add_argument("patch_file")
    p_merge.add_argument("--force", action="store_true")
    p_merge.set_defaults(func=cmd_merge)

    p_q4 = sub.add_parser("derive-q4")
    p_q4.add_argument("ticker")
    p_q4.add_argument("year")
    p_q4.add_argument("--force", action="store_true")
    p_q4.set_defaults(func=cmd_derive_q4)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
