"""Kiem tra tinh hop ly cua 2 bang gap (lai suat + thanh khoan) cua 1 ky/1 ngan hang.

Viet sau dot doc lai thu cong toan bo 26 ngan hang x 10 ky tu PDF that (2026-09) - moi loi tim thay
deu thuoc 1 trong cac nhom duoi day, nen moi nhom co 1 chot kiem tra RIENG:

  1. Doc nham dong "Muc chenh lech" vao cot "No phai tra" (loi pho bien nhat - hay dinh vao cot
     "Khong chiu lai", "1-3 thang"). Bang 2 phep kiem tra: (a) dong No phai tra theo bucket trung voi
     dong gap; (b) tong No phai tra cua bang lai suat lech tong No phai tra cua bang thanh khoan (cung
     1 bang can doi -> phai xap xi bang nhau).
  2. Gan quá han cua TAI SAN vao mang NO PHAI TRA (no phai tra khong co qua han).
  3. Lech cot/gop nham dong -> tong tai san suy ra (Gap + No) cua 2 bang khong khop nhau / khong khop
     tong tai san that (Vietcap).
  4. Nham don vi (ty/trieu/VND) -> bucket vuot tong tai san.

Module thuan Python (khong phu thuoc OCR/mang) de chay duoc ca trong workflow, test va script tay.
Don vi: TRIEU dong (dung theo bank_alm_store); tong tai san Vietcap (`ta_ty`) tinh bang TY dong.
"""
import json
import os
from datetime import datetime, timezone

REVIEW_QUEUE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "bank_alm",
                                 "_needs_review.json")

IR_KEYS = ["qua_han", "khong_anh_huong_lai_suat", "den_1_thang", "tu_1_3_thang", "tu_3_6_thang",
           "tu_6_12_thang", "tu_1_5_nam", "tren_5_nam"]
LIQ_KEYS = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang", "tu_3_12_thang",
            "tu_1_5_nam", "tren_5_nam"]

# Nguong (ty le). Chon tu du lieu that: hai bang cua CUNG 1 bao cao thuong khop CHINH XAC, toi da
# lech ~0,03% (lam tron) - rieng vai ngan hang tu lech 2 bang ~0,01-0,5% vi phan loai phai sinh khac
# nhau; sai doc so thuong lech > 3%. Vietcap lech them ~1-2% do quy uoc gop/rong du phong.
TA_IR_VS_TK_FAIL = 0.005
NO_IR_VS_TK_FAIL = 0.01
NO_IR_VS_TK_WARN = 0.002
TA_VS_VIETCAP_FAIL = 0.10
NEG_LIAB_TOL = 0.005      # no phai tra 1 bucket am toi da 0,5% tong (phai sinh co the am nhe)
NEG_ASSET_TOL = 0.01      # tai san suy ra 1 bucket am toi da 1% tong


def _vals(d):
    return [v or 0 for v in (d or {}).values()]


def _implied_total(gap_d, liab_d):
    if not gap_d or not liab_d:
        return None
    return sum(_vals(gap_d)) + sum(_vals(liab_d))


def _rel(a, b):
    return abs(a - b) / max(abs(a), abs(b), 1e-9)


def validate_entry(entry, ta_ty=None):
    """Tra ve {"ok": bool, "level": "ok"|"warn"|"fail", "issues": [{"code","msg"}]}.
    `entry` la 1 phan tu cua store["gap_periods"]; `ta_ty` la tong tai san Vietcap (ty dong) neu co.
    Chi kiem tra nhung gi tinh duoc - thieu du lieu 1 ben KHONG bi coi la loi."""
    issues = []

    def add(level, code, msg):
        issues.append({"level": level, "code": code, "msg": msg})

    ir_gap, ir_liab = entry.get("interest_rate_gap"), entry.get("interest_rate_liabilities_by_bucket")
    lq_gap, lq_liab = entry.get("liquidity_gap"), entry.get("liabilities_by_bucket")

    ta_ir = _implied_total(ir_gap, ir_liab)
    ta_lq = _implied_total(lq_gap, lq_liab)

    # (3) tong tai san suy ra: 2 bang voi nhau, moi bang voi Vietcap
    if ta_ir is not None and ta_lq is not None and _rel(ta_ir, ta_lq) > TA_IR_VS_TK_FAIL:
        add("fail", "ta_ir_vs_tk", f"tong tai san suy ra: lai suat {ta_ir/1e3:,.0f} ty vs thanh khoan "
            f"{ta_lq/1e3:,.0f} ty lech {_rel(ta_ir, ta_lq)*100:.1f}%")
    if ta_ty:
        for name, ta in (("lai suat", ta_ir), ("thanh khoan", ta_lq)):
            if ta is not None and _rel(ta / 1e3, ta_ty) > TA_VS_VIETCAP_FAIL:
                add("fail", "ta_vs_vietcap", f"tong tai san suy ra bang {name} {ta/1e3:,.0f} ty lech "
                    f"{_rel(ta/1e3, ta_ty)*100:.0f}% so voi tong tai san that {ta_ty:,.0f} ty")

    # (1b) tong No phai tra 2 bang - bat loi doc nham dong Muc chenh lech vao cot no
    if ir_liab and lq_liab:
        li, ll = sum(_vals(ir_liab)), sum(_vals(lq_liab))
        if li and ll:
            d = _rel(li, ll)
            if d > NO_IR_VS_TK_FAIL:
                add("fail", "no_ir_vs_tk", f"tong no phai tra: lai suat {li/1e3:,.0f} ty vs thanh khoan "
                    f"{ll/1e3:,.0f} ty lech {d*100:.1f}% (nghi doc nham dong Muc chenh lech/lech cot)")
            elif d > NO_IR_VS_TK_WARN:
                add("warn", "no_ir_vs_tk_small", f"tong no phai tra 2 bang lech {d*100:.2f}%")

    for label, gap_d, liab_d, keys_overdue in (("lai suat", ir_gap, ir_liab, 1), ("thanh khoan", lq_gap, lq_liab, 2)):
        if not liab_d:
            continue
        lv = list(_vals(liab_d))
        tot = sum(lv) or 1
        # (2) no phai tra khong co qua han
        if any(abs(x) > 0.001 * abs(tot) for x in lv[:keys_overdue]):
            add("fail", "no_qua_han", f"bang {label}: no phai tra co gia tri o cot qua han (no khong the qua "
                f"han - nghi gan nham so cua tai san)")
        # (1a) dong no trung dong gap
        if gap_d:
            gv = list(_vals(gap_d))
            same = sum(1 for a, b in zip(lv, gv) if a and abs(a - b) < 1)
            if same >= 3:
                add("fail", "no_equals_gap", f"bang {label}: {same} bucket no phai tra trung tuyet doi voi "
                    f"gap (nghi doc nham dong)")
            # tai san suy ra khong the am / khong the vuot tong tai san
            if tot > 0:
                assets = [g + l for g, l in zip(gv, lv)]
                if any(a < -NEG_ASSET_TOL * tot for a in assets):
                    add("fail", "asset_negative", f"bang {label}: tai san suy ra (gap + no) am o >=1 bucket")
                if any(a > 1.5 * (sum(assets) or 1) for a in assets):
                    add("fail", "bucket_over_total", f"bang {label}: 1 bucket lon hon 1,5 lan tong tai san")
        if any(x < -NEG_LIAB_TOL * abs(tot) for x in lv):
            add("fail", "liab_negative", f"bang {label}: no phai tra am o >=1 bucket (ngoai dung sai phai sinh)")

    # Ky bu MOT PHAN (vd STB Q2/2024: thanh khoan that, lai suat mang tu ky truoc vi ngan hang khong
    # cong bo) - 2 bang KHONG cung 1 thoi diem nen cac phep so cheo giua 2 bang / voi Vietcap khong ap dung
    if entry.get("patched_from") and entry.get("status") != "patched":
        issues = [i for i in issues if i["code"] not in ("ta_ir_vs_tk", "no_ir_vs_tk", "no_ir_vs_tk_small",
                                                         "ta_vs_vietcap")]
    level = "fail" if any(i["level"] == "fail" for i in issues) else ("warn" if issues else "ok")
    return {"ok": level != "fail", "level": level, "issues": issues}


def reconcile_liabilities(assets, liab, gap, n_overdue):
    """Doi chieu 3 dong doc duoc tu CUNG 1 bang: Tong tai san, Tong no phai tra, Muc chenh lech.
    Phep dong nhat: gap[i] = assets[i] - liab[i]. Dung ngay TRONG luc trich xuat (truoc khi luu).

    - Cot qua han cua no phai tra ep ve 0 (no khong co qua han).
    - Bucket nao lech dong nhat (hoac no trung dong gap - dau hieu gom nham dong) -> thay bang
      assets - gap NEU chi lech cuc bo (<= 1/3 so bucket, va gia tri thay khong am).
    Tra ve (liab_moi, [canh_bao]). Thieu assets/gap -> chi ep qua han = 0."""
    warns = []
    liab = list(liab)
    for i in range(min(n_overdue, len(liab))):
        if liab[i]:
            warns.append(f"no phai tra cot qua han [{i}]={liab[i]:,.0f} ep ve 0")
            liab[i] = 0
    if not assets or not gap or len(assets) != len(liab) or len(gap) != len(liab):
        return liab, warns
    total = sum(abs(a) for a in assets) or 1
    tol = max(2.0, 0.0005 * total)
    bad = [i for i in range(len(liab)) if abs(assets[i] - liab[i] - gap[i]) > tol]
    if not bad:
        return liab, warns
    if len(bad) > max(1, len(liab) // 3):
        warns.append(f"{len(bad)}/{len(liab)} bucket khong thoa gap = tai san - no (khong sua duoc tu dong)")
        return liab, warns
    for i in bad:
        alt = assets[i] - gap[i]
        if alt >= -NEG_LIAB_TOL * total and (i >= n_overdue or abs(alt) <= tol):
            warns.append(f"bucket [{i}]: no phai tra {liab[i]:,.0f} -> {alt:,.0f} (= tai san - gap)")
            liab[i] = alt
        else:
            warns.append(f"bucket [{i}] lech dong nhat nhung gia tri thay ({alt:,.0f}) khong hop ly - giu nguyen")
    return liab, warns


def queue_review(ticker, period_key, result, source=None, path=None):
    """Ghi/cap nhat hang doi can xem lai thu cong (data/bank_alm/_needs_review.json). Khoa (ticker,
    period) - loi da het (result ok) thi go khoi hang doi."""
    path = path or REVIEW_QUEUE_PATH
    try:
        q = json.load(open(path, encoding="utf-8")) if os.path.exists(path) else []
    except Exception:
        q = []
    q = [x for x in q if not (x.get("ticker") == ticker and x.get("period") == period_key)]
    if result and not result.get("ok"):
        q.append({"ticker": ticker, "period": period_key, "issues": result.get("issues", []),
                  "url": (source or {}).get("url"), "at": datetime.now(timezone.utc).isoformat(timespec="seconds")})
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(q, f, ensure_ascii=False, indent=2)
    return q


def validate_all(write_queue=True, mark_verified=None):
    """Quet TOAN BO data/bank_alm: chay validate_entry, ghi `validation` vao tung entry (kem
    `verified` neu mark_verified duoc truyen), cap nhat hang doi can xem lai. Tra ve list ket qua."""
    import bank_alm_store as store
    from bank_universe import BANKING_TICKERS
    rows = []
    for t in sorted(BANKING_TICKERS):
        s = store.load_bank_store(t)
        qbs = s.get("quarterly_balance_sheet", {})
        dirty = False
        for pk, e in s.get("gap_periods", {}).items():
            if not e or e.get("status") == "missing":
                continue
            snap = qbs.get(store.gap_period_to_quarter(pk)) or {}
            res = validate_entry(e, snap.get("total_assets"))
            # ky "patched" (carry-forward) tu ky khac: lech Vietcap la dieu DUOC - chi xet tinh nhat quan noi bo
            if e.get("status") == "patched":
                res["issues"] = [i for i in res["issues"] if i["code"] != "ta_vs_vietcap"]
                res["level"] = "fail" if any(i["level"] == "fail" for i in res["issues"]) else \
                    ("warn" if res["issues"] else "ok")
                res["ok"] = res["level"] != "fail"
            e["validation"] = {"ok": res["ok"], "level": res["level"],
                               "issues": [i["code"] for i in res["issues"]]}
            if mark_verified and res["ok"] and not e.get("verified"):
                e["verified"] = mark_verified
            dirty = True
            rows.append((t, pk, e.get("status"), res))
            if write_queue:
                queue_review(t, pk, res, e.get("source"))
        if dirty:
            store.save_bank_store(t, s)
    return rows


if __name__ == "__main__":
    import sys
    rows = validate_all(write_queue=True, mark_verified=("pdf_visual_2026-09" if "--mark-verified" in sys.argv else None))
    bad = [r for r in rows if r[3]["level"] == "fail"]
    warn = [r for r in rows if r[3]["level"] == "warn"]
    lines = [f"Kiem tra ALM: {len(rows)} ky, {len(bad)} loi, {len(warn)} canh bao"]
    for t, pk, st, res in bad:
        lines.append(f"  [FAIL] {t} {pk} ({st}): " + "; ".join(i["msg"] for i in res["issues"] if i["level"] == "fail"))
    for t, pk, st, res in warn:
        lines.append(f"  [WARN] {t} {pk} ({st}): " + "; ".join(i["msg"] for i in res["issues"]))
    print("\n".join(lines))
    summ = os.environ.get("GITHUB_STEP_SUMMARY")
    if summ:
        with open(summ, "a", encoding="utf-8") as f:
            f.write("### Kiem tra du lieu ALM\n```\n" + "\n".join(lines) + "\n```\n")
    if "--strict" in sys.argv and bad:
        sys.exit(1)
