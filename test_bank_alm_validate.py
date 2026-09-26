"""Test bo kiem tra ALM - moi ca la 1 loi THAT da gap khi doc lai 26 ngan hang tu PDF (2026-09).
Chay: python test_bank_alm_validate.py  (hoac pytest)."""
import bank_alm_validate as v

IR = ["qua_han", "khong_anh_huong_lai_suat", "den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang",
      "tu_1_5_nam", "tren_5_nam"]
LQ = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang", "tu_3_12_thang", "tu_1_5_nam",
      "tren_5_nam"]


def entry(irA, irL, lqA, lqL):
    return {
        "interest_rate_gap": dict(zip(IR, [a - l for a, l in zip(irA, irL)])),
        "interest_rate_liabilities_by_bucket": dict(zip(IR, irL)),
        "liquidity_gap": dict(zip(LQ, [a - l for a, l in zip(lqA, lqL)])),
        "liabilities_by_bucket": dict(zip(LQ, lqL)),
    }


# VAB FY-2024 (dung, doc tu PDF)
VAB_FY24 = dict(
    irA=[1424730, 7963815, 38580884, 18916913, 7339819, 30449903, 5947676, 10348854],
    irL=[0, 2337582, 16332754, 4314556, 1129509, 59297365, 26092986, 1470608],
    lqA=[431416, 993314, 38068857, 5057508, 53764828, 15074177, 7582495],
    lqL=[0, 0, 35774012, 21737443, 49795228, 2668676, 1000000],
)


def _codes(res):
    return {i["code"] for i in res["issues"]}


def test_correct_data_passes():
    r = v.validate_entry(entry(**VAB_FY24), ta_ty=120972.6)
    assert r["ok"] and r["level"] == "ok", r


def test_gap_row_read_as_liability_is_caught():
    # loi VAB FY-2024: cot Khong chiu lai / 1-3 thang cua no phai tra lay tu dong "Muc chenh lech"
    d = dict(VAB_FY24)
    d["irL"] = [0, 5626234, 22248130, 14602357, 1129509, 59297365, 26092986, 1470608]  # 3 cot sai
    r = v.validate_entry(entry(**d))
    assert not r["ok"] and "no_ir_vs_tk" in _codes(r), r


def test_asset_overdue_in_liabilities_is_caught():
    # loi VAB Q3/2024 / FY-2024: qua han cua TAI SAN (993.314) gan vao mang NO thanh khoan
    d = dict(VAB_FY24)
    d["lqL"] = [0, 993314, 35774012, 21737443, 49795228, 2668676, 1000000]
    r = v.validate_entry(entry(**d))
    assert not r["ok"] and "no_qua_han" in _codes(r), r


def test_column_shift_total_assets_mismatch_is_caught():
    d = dict(VAB_FY24)
    d["lqA"] = [431416, 993314, 38068857, 5057508, 53764828, 15074177, 9582495]  # lech 2 trieu
    r = v.validate_entry(entry(**d))
    assert not r["ok"] and "ta_ir_vs_tk" in _codes(r), r


def test_wrong_unit_vs_vietcap_is_caught():
    # OCB ghi bang VND (chua chia 1e6) -> tong tai san suy ra lech Vietcap hang trieu lan
    d = {k: [x * 1e6 for x in v_] for k, v_ in VAB_FY24.items()}
    r = v.validate_entry(entry(**d), ta_ty=120972.6)
    assert not r["ok"] and "ta_vs_vietcap" in _codes(r), r


def test_partial_patched_entry_skips_cross_checks():
    # STB Q2/2024: thanh khoan that + lai suat carry ky truoc -> khong so cheo 2 bang
    e = entry(**VAB_FY24)
    e["interest_rate_gap"] = {k: x * 0.97 for k, x in e["interest_rate_gap"].items()}
    e["patched_from"] = "2024-Q1 (chi phan lai suat)"
    e["status"] = "reported"
    assert v.validate_entry(e, ta_ty=120972.6)["ok"]


def test_missing_side_is_not_an_error():
    e = entry(**VAB_FY24)
    e["liquidity_gap"] = None
    e["liabilities_by_bucket"] = None
    assert v.validate_entry(e)["ok"]


def test_reconcile_fixes_gap_row_in_liability_column():
    A = VAB_FY24["irA"]
    L = VAB_FY24["irL"]
    gap = [a - l for a, l in zip(A, L)]
    bad = list(L)
    bad[1] = gap[1]  # copy nham dong chenh lech vao 1 cot no
    fixed, warns = v.reconcile_liabilities(A, bad, gap, 1)
    assert fixed == L and warns, (fixed, warns)


def test_reconcile_zeroes_overdue_liability():
    A = VAB_FY24["lqA"]
    L = [0, 993314, 35774012, 21737443, 49795228, 2668676, 1000000]
    gap = [a - l for a, l in zip(A, VAB_FY24["lqL"])]
    fixed, _ = v.reconcile_liabilities(A, L, gap, 2)
    assert fixed[:2] == [0, 0]


def test_reconcile_leaves_clean_data_alone():
    A, L = VAB_FY24["irA"], VAB_FY24["irL"]
    gap = [a - l for a, l in zip(A, L)]
    fixed, warns = v.reconcile_liabilities(A, L, gap, 1)
    assert fixed == L and not warns


def test_reconcile_gives_up_when_many_buckets_bad():
    A, L = VAB_FY24["irA"], VAB_FY24["irL"]
    gap = [a - l + 5_000_000 for a, l in zip(A, L)]  # gap sai o moi cot
    fixed, warns = v.reconcile_liabilities(A, L, gap, 1)
    assert fixed == L and warns


if __name__ == "__main__":
    n = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            n += 1
    print(f"OK {n} test")
