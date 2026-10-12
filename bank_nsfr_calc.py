# -*- coding: utf-8 -*-
"""Tinh NSFR (Ty le nguon von on dinh rong) XAP XI theo khung Thong tu 50/2026/TT-NHNN.

NGUON GOC BANG HE SO: trich nguyen van tu Phu luc II (Huong dan cach xac dinh ty le
nguon von on dinh rong), kem theo Thong tu so 50/2026/TT-NHNN ngay 30/9/2026 cua NHNN
- file PDF goc nguoi dung cung cap (Thong_tu_50_2026_TT.nhnn_ceea.pdf, doc truc tiep
tung trang anh de tranh loi OCR/font, da doi chieu ca 2 Phan A (ASF) va Phan B (RSF)
day du tung dong). KHONG dung so lieu suy doan/bao chi cho bang he so chinh - chi dung
bang goc.

GIOI HAN QUAN TRONG (da bao truoc voi user, user dong y xu ly theo huong proxy ro rang,
neu thieu thi theo chuan Basel III quoc te): BCTC CONG KHAI cua ngan hang VN KHONG co du
chi tiet de ap dung dung 100% bang he so nay, vi thieu:
  - Phan loai tien gui theo muc do "on dinh/kem on dinh" (phu thuoc han muc bao hiem
    tien gui tung khach hang - khong cong khai).
  - Phan loai cho vay theo he so rui ro tin dung (<=35% hay >35%) theo Thong tu 14/2025
    hay 41/2016 - chi co trong bao cao Tru cot 3 (Pillar 3) rieng, khong co trong BCTC
    chuan.
  - Phan tach HQLA cap 1/2A/2B trong "Chung khoan dau tu" - BCTC chi cho tong so.
Vi vay day la 1 UOC TINH (proxy) bam sat NHAT CO THE theo dung % cua TT50, dung du lieu
GOP THEO KY HAN CON LAI (da co san trong gap_periods cua tung bank, xem bank_alm_store)
+ cac dong tong BCTC (quarterly_balance_sheet) - KHONG PHAI so NSFR chinh thuc ngan hang
tu bao cao len NHNN. Moi gia dinh don gian hoa deu ghi chu ro trong code va trong sheet
"Gia dinh & Gioi han" cua file Excel xuat ra.

CAU TRUC: moi field BCTC duoc gan vao 1 "nhom cong cu" (instrument group), moi nhom co
1 "lich he so" rieng theo dung % cua khoan tuong ung trong Phu luc II (xem ASF_SCHEDULES/
RSF_SCHEDULES ben duoi). Ty trong ky han cua TUNG nhom duoc gia dinh BANG ty trong ky han
TRUNG BINH cua toan bo Tai san/No phai tra (lay tu gap_periods) - vi BCTC khong cho ky han
rieng tung dong, day la gia dinh don gian hoa CAN THIET, ghi ro trong docstring.
"""

import json
import os

import bank_universe

DATA_DIR = "data/bank_alm"

# Cac bucket ky han "tho" co san trong gap_periods (xem bank_alm_store/fetch cap nhat) -
# anh xa ve 4 cot ky han cua TT50: Duoi 6 thang / 6 thang den duoi 1 nam / Tu 1 nam tro len.
# "tu_3_12_thang" trai dai qua moc 6 thang (3-12 thang) - KHONG the tach chinh xac tu du
# lieu hien co, gia dinh chia theo ty trong THOI GIAN: 3/9 thuoc "duoi 6 thang" (phan 3-6
# thang), 6/9 thuoc "6 thang den duoi 1 nam" (phan 6-12 thang) - gia dinh TUYEN TINH, ghi
# ro trong sheet gia dinh.
_BUCKET_WEIGHTS = {
    "under_6m": {"qua_han_tren_3t": 1.0, "qua_han_den_3t": 1.0, "den_1_thang": 1.0,
                 "tu_1_3_thang": 1.0, "tu_3_12_thang": 3.0 / 9.0},
    "6m_to_1y": {"tu_3_12_thang": 6.0 / 9.0},
    "over_1y": {"tu_1_5_nam": 1.0, "tren_5_nam": 1.0},
}


def _bucket_shares(bucket_dict):
    """Tra {'under_6m':.., '6m_to_1y':.., 'over_1y':..} ty trong (tong=1) tu 1 dict bucket
    tho (liquidity_gap hoac liabilities_by_bucket), hoac None neu thieu/tong<=0."""
    if not bucket_dict:
        return None
    total = sum(v or 0 for v in bucket_dict.values())
    if total <= 0:
        return None
    out = {}
    for key, weights in _BUCKET_WEIGHTS.items():
        out[key] = sum((bucket_dict.get(b) or 0) * w for b, w in weights.items()) / total
    return out


# (duoi_6_thang, 6_thang_den_1_nam, tu_1_nam_tro_len) - dung cho phia NGUON VON (ASF),
# theo dung Bang he so ASF, Phu lục II Phan A (da doc tu PDF goc, xem docstring dau file).
ASF_SCHEDULES = {
    # Von chu so huu (khoan 1.1-1.3 gop lam 1, vi BCTC khong tach von cap 1/cap 2/cong cu
    # khac) - uu tien dung % cua khoan 1.1 (von cap 1, 100% moi ky han) vi day la phan
    # chiem da so trong VCSH 1 ngan hang thuong mai thong thuong.
    "equity": (1.00, 1.00, 1.00),
    # Tien gui khach hang (customer_deposits, gop ban le+to chuc) - BCTC KHONG tach duoc
    # "on dinh" (khoan 9.1/10.1, 95%) voi "kem on dinh" (khoan 9.2/10.2, 90%) theo han muc
    # bao hiem tien gui - dung khoan 9.2/10.2 "kem on dinh" lam muc GIUA THAN TRONG (thap
    # hon muc "on dinh" 1 chut, cao hon tien gui to chuc/doanh nghiep 50% nhieu) cho TOAN
    # BO tien gui khach hang.
    "customer_deposits": (0.90, 0.90, 1.00),
    # Tien gui + vay nhan tu TCTD khac (interbank_liab, bsb112 - DA GOM ca tien gui VA vay,
    # KHONG cong them tctd_dep vi la SUBSET cua chinh interbank_liab) - khoan 11.2.
    "interbank_liab": (0.00, 0.50, 1.00),
    # Phat hanh trai phieu (bonds_issued) - khoan 6.
    "bonds_issued": (0.00, 0.50, 1.00),
    # Tien gui + vay KBNN (kbnn_dep + kbnn_loan) - khoan 7 (Chinh phu/KBNN...).
    "kbnn": (0.50, 0.50, 1.00),
    # Phan du (No phai tra khac khong xac dinh duoc nguon cu the) - khoan 14.1 (co ky han
    # cu the, dung lich tuong tu TCTD de than trong - khong gan nham vao muc 95% tien gui).
    "other_liab": (0.00, 0.50, 1.00),
}

# (duoi_6_thang, 6_thang_den_1_nam, tu_1_nam_tro_len) - dung cho phia TAI SAN (RSF), theo
# dung Bang he so RSF, Phu lục II Phan B.
RSF_SCHEDULES = {
    # Cho vay khach hang (loans, bsb103) - BCTC KHONG tach duoc theo he so rui ro tin
    # dung <=35% (khoan 18.x.1, 65% tu 1 nam) hay >35% (khoan 18.x.2, 85% tu 1 nam) - dung
    # muc <=35% (65%) lam dai dien, vi phan lon du no cho vay ca nhan/the chap tai san o
    # VN thuoc nhom nay theo Thong tu 14/2025 phuong phap tieu chuan - CO THE LAM THAP
    # UOC LUONG RSF (tuc LAM CAO uoc luong NSFR) so voi thuc te neu ty trong cho vay rui
    # ro cao (>35%, kinh doanh BDS, tieu dung khong TSBD...) lon hon gia dinh nay.
    "loans": (0.50, 0.50, 0.65),
    # Chung khoan dau tu (inv_sec_bs, bsb106, KHONG cong them tpdn - tpdn la so KHAI THAC
    # TU GHI CHU thuyet minh, nhieu kha nang la 1 SUB-SET da nam trong inv_sec_bs, cong
    # them se double-count) - khong tach duoc HQLA cap 1 (TPCP, 5%) khoi "chung khoan
    # khac" (khoan 11.3, 50-100%) - dung khoan 11.3 "chung khoan khac" lam dai dien THAN
    # TRONG (RSF cao hon, tuc NSFR uoc tinh THAP hon thuc te neu ngan hang giu nhieu TPCP).
    "inv_sec_bs": (0.50, 0.50, 1.00),
    # Tien gui/cho vay tai TCTD khac (bank_dep, bsb98 - tai san, KHAC interbank_liab la no
    # phai tra) - khoan 15.3 "tien gui, tru 15.1/15.2".
    "bank_dep": (0.50, 0.50, 1.00),
    # Phan du (Tai san khac khong xac dinh duoc nguon cu the: TSCD, phai thu khac...) -
    # khoan 20 "tai san khac khong xac dinh duoc ky han" = 100% CO DINH moi ky han (ap
    # dung dung nhu ban goc, khong phan theo bucket).
    "other_asset": (1.00, 1.00, 1.00),
}

# Tien mat + du tru bat buoc tai NHNN - khoan 8.1/8.2, LUON 0% RSF bat ke ky han.
_ZERO_RSF_FIELDS = ("cash", "sbv_dep")


def _weighted_schedule(amount, shares, schedule):
    """amount: gia tri VND (ty); shares: {'under_6m','6m_to_1y','over_1y'} ty trong tong=1;
    schedule: (pct_duoi_6m, pct_6m_1y, pct_tu_1nam) - tra gia tri da nhan he so (ty)."""
    if amount is None or shares is None:
        return 0.0
    pct_u6, pct_6_1, pct_1y = schedule
    return amount * (shares["under_6m"] * pct_u6 + shares["6m_to_1y"] * pct_6_1
                      + shares["over_1y"] * pct_1y)


def compute_nsfr_for_period(snap, gap_entry):
    """snap: 1 entry cua quarterly_balance_sheet (tong BCTC); gap_entry: 1 entry cua
    gap_periods (co liquidity_gap/liabilities_by_bucket). Tra dict chi tiet hoac None
    neu thieu du lieu bat buoc (liquidity_gap/liabilities_by_bucket rong)."""
    liab_bucket = gap_entry.get("liabilities_by_bucket") if gap_entry else None
    liq_gap = gap_entry.get("liquidity_gap") if gap_entry else None
    if not liab_bucket or not liq_gap:
        return None
    liab_shares = _bucket_shares(liab_bucket)
    asset_bucket_raw = {k: (liq_gap.get(k) or 0) + (liab_bucket.get(k) or 0) for k in liab_bucket}
    asset_shares = _bucket_shares(asset_bucket_raw)
    if liab_shares is None or asset_shares is None:
        return None

    total_assets = snap.get("total_assets") or 0
    equity = snap.get("equity") or 0
    total_liab = total_assets - equity

    customer_deposits = snap.get("customer_deposits") or 0
    interbank_liab = snap.get("interbank_liab") or 0
    bonds_issued = snap.get("bonds_issued") or 0
    kbnn = (snap.get("kbnn_dep") or 0) + (snap.get("kbnn_loan") or 0)
    other_liab = total_liab - equity - customer_deposits - interbank_liab - bonds_issued - kbnn

    cash = snap.get("cash") or 0
    sbv_dep = snap.get("sbv_dep") or 0
    bank_dep = snap.get("bank_dep") or 0
    loans = snap.get("loans") or 0
    inv_sec_bs = snap.get("inv_sec_bs") or 0
    other_asset = total_assets - cash - sbv_dep - bank_dep - loans - inv_sec_bs

    asf_components = {
        "equity": equity * ASF_SCHEDULES["equity"][2],  # VCSH luon 100%, khong theo bucket
        "customer_deposits": _weighted_schedule(customer_deposits, liab_shares, ASF_SCHEDULES["customer_deposits"]),
        "interbank_liab": _weighted_schedule(interbank_liab, liab_shares, ASF_SCHEDULES["interbank_liab"]),
        "bonds_issued": _weighted_schedule(bonds_issued, liab_shares, ASF_SCHEDULES["bonds_issued"]),
        "kbnn": _weighted_schedule(kbnn, liab_shares, ASF_SCHEDULES["kbnn"]),
        "other_liab": _weighted_schedule(other_liab, liab_shares, ASF_SCHEDULES["other_liab"]),
    }
    asf_total = sum(asf_components.values())

    rsf_components = {
        "cash": 0.0,
        "sbv_dep": 0.0,
        "bank_dep": _weighted_schedule(bank_dep, asset_shares, RSF_SCHEDULES["bank_dep"]),
        "loans": _weighted_schedule(loans, asset_shares, RSF_SCHEDULES["loans"]),
        "inv_sec_bs": _weighted_schedule(inv_sec_bs, asset_shares, RSF_SCHEDULES["inv_sec_bs"]),
        "other_asset": other_asset * RSF_SCHEDULES["other_asset"][2],
    }
    rsf_total = sum(rsf_components.values())

    nsfr_pct = (asf_total / rsf_total * 100.0) if rsf_total else None
    return {
        "asf_total": asf_total, "rsf_total": rsf_total, "nsfr_pct": nsfr_pct,
        "asf_components": asf_components, "rsf_components": rsf_components,
        "total_assets": total_assets, "equity": equity,
    }


def compute_nsfr_all_banks():
    """Chay cho toan bo 26 bank x tat ca ky co du lieu. Tra
    {ticker: {period: result_dict_or_None}}."""
    results = {}
    for ticker in sorted(bank_universe.BANKING_TICKERS):
        path = os.path.join(DATA_DIR, f"{ticker}.json")
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        qbs = d.get("quarterly_balance_sheet", {}) or {}
        gp = d.get("gap_periods", {}) or {}
        bank_results = {}
        for period, gap_entry in gp.items():
            # gap_periods dung nhan "YYYY-FY" cho bao cao nam (kiem toan) - trong khi
            # quarterly_balance_sheet dung thang nhat "YYYY-Q4" cho CUNG 1 ky do (BCTC
            # kiem toan nam CHINH LA so cuoi Q4) - quy doi nhan truoc khi tra cuu.
            qbs_period = period.replace("-FY", "-Q4") if period.endswith("-FY") else period
            snap = qbs.get(qbs_period)
            if not snap:
                continue
            r = compute_nsfr_for_period(snap, gap_entry)
            if r is not None:
                bank_results[qbs_period] = r
        if bank_results:
            results[ticker] = bank_results
    return results


_NSFR_SHEET_NAME = "NSFR_NganHang_Raw"
_NSFR_SHEET_HEADERS = [
    "Ma NH", "Ky", "Tong tai san (ty)", "VCSH (ty)",
    "ASF - Nguon von on dinh kha dung (ty)", "RSF - Nguon von on dinh yeu cau (ty)",
    "NSFR (%, >=100% la dat)",
    "ASF - Von chu so huu (ty)", "ASF - Tien gui khach hang (ty)", "ASF - TCTD khac (ty)",
    "ASF - Trai phieu phat hanh (ty)", "ASF - KBNN (ty)", "ASF - No khac (ty)",
    "RSF - Tien mat+DTBB (ty)", "RSF - Tien gui tai TCTD khac (ty)", "RSF - Cho vay KH (ty)",
    "RSF - Chung khoan dau tu (ty)", "RSF - Tai san khac (ty)",
]

_NSFR_SUMMARY_SHEET_NAME = "NSFR_TongHop_HeThong"
_NSFR_SUMMARY_HEADERS = [
    "Ky", "So NH co du lieu", "Tong ASF toan HT (ty)", "Tong RSF toan HT (ty)",
    "NSFR toan he thong (%)", "So NH duoi 100%", "So NH tu 100-110%", "So NH tren 110%",
]

_NSFR_ASSUMPTIONS_SHEET_NAME = "NSFR_GiaDinh_GioiHan"


def update_bank_nsfr_excel_sheet(out_dir):
    """Ghi sheet "NSFR_NganHang_Raw" (1 hang = 1 (ngan hang, quy)), "NSFR_TongHop_HeThong" (1
    hang = 1 quy, cong don ASF/RSF ca 26 ngan hang - KHONG PHAI trung binh % tung bank, giong
    dung cach tong hop LDR toan he thong da lam o update_bank_ldr_excel_sheet) va 1 sheet
    "NSFR_GiaDinh_GioiHan" ghi ro phuong phap + gioi han, vao CUNG workbook VIMO_Lich_Su_Chi_So.xlsx
    (user: "tinh toan cho toan bo thi truong va tinh cho tung bank, luu lai o file excel").
    GHI DE toan bo sheet moi lan chay."""
    import openpyxl

    xlsx_path = os.path.join(out_dir, "VIMO_Lich_Su_Chi_So.xlsx")
    if os.path.exists(xlsx_path):
        wb = openpyxl.load_workbook(xlsx_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    results = compute_nsfr_all_banks()

    if _NSFR_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[_NSFR_SHEET_NAME])
    ws = wb.create_sheet(title=_NSFR_SHEET_NAME)
    for c, h in enumerate(_NSFR_SHEET_HEADERS, start=1):
        ws.cell(row=1, column=c, value=h)
    row_idx = 2
    for ticker in sorted(results.keys()):
        for period in sorted(results[ticker].keys()):
            r = results[ticker][period]
            ac, rc = r["asf_components"], r["rsf_components"]
            row = [
                ticker, period, round(r["total_assets"], 1), round(r["equity"], 1),
                round(r["asf_total"], 1), round(r["rsf_total"], 1),
                round(r["nsfr_pct"], 2) if r["nsfr_pct"] is not None else None,
                round(ac["equity"], 1), round(ac["customer_deposits"], 1), round(ac["interbank_liab"], 1),
                round(ac["bonds_issued"], 1), round(ac["kbnn"], 1), round(ac["other_liab"], 1),
                round(rc["cash"] + rc["sbv_dep"], 1), round(rc["bank_dep"], 1), round(rc["loans"], 1),
                round(rc["inv_sec_bs"], 1), round(rc["other_asset"], 1),
            ]
            for c, val in enumerate(row, start=1):
                ws.cell(row=row_idx, column=c, value=val)
            row_idx += 1
    for c in range(1, len(_NSFR_SHEET_HEADERS) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 18
    print(f"  [OK] Sheet {_NSFR_SHEET_NAME}: {row_idx - 2} hang")

    all_periods = sorted({p for periods in results.values() for p in periods})
    if _NSFR_SUMMARY_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[_NSFR_SUMMARY_SHEET_NAME])
    ws2 = wb.create_sheet(title=_NSFR_SUMMARY_SHEET_NAME)
    for c, h in enumerate(_NSFR_SUMMARY_HEADERS, start=1):
        ws2.cell(row=1, column=c, value=h)
    for i, period in enumerate(all_periods):
        asf_sum = rsf_sum = 0.0
        n_banks = n_under100 = n_100_110 = n_over110 = 0
        for ticker, periods in results.items():
            r = periods.get(period)
            if not r:
                continue
            n_banks += 1
            asf_sum += r["asf_total"]
            rsf_sum += r["rsf_total"]
            if r["nsfr_pct"] is not None:
                if r["nsfr_pct"] < 100:
                    n_under100 += 1
                elif r["nsfr_pct"] <= 110:
                    n_100_110 += 1
                else:
                    n_over110 += 1
        nsfr_system = round(asf_sum / rsf_sum * 100, 2) if rsf_sum else None
        row = [period, n_banks, round(asf_sum, 1), round(rsf_sum, 1), nsfr_system,
               n_under100, n_100_110, n_over110]
        for c, val in enumerate(row, start=1):
            ws2.cell(row=i + 2, column=c, value=val)
    for c in range(1, len(_NSFR_SUMMARY_HEADERS) + 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 22
    print(f"  [OK] Sheet {_NSFR_SUMMARY_SHEET_NAME}: {len(all_periods)} hang")

    if _NSFR_ASSUMPTIONS_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[_NSFR_ASSUMPTIONS_SHEET_NAME])
    ws3 = wb.create_sheet(title=_NSFR_ASSUMPTIONS_SHEET_NAME)
    notes = [
        "NSFR UOC TINH (PROXY) theo khung Thong tu 50/2026/TT-NHNN - KHONG PHAI so NSFR",
        "chinh thuc ngan hang tu bao cao len NHNN.",
        "",
        "NGUON BANG HE SO: Phu luc II (Huong dan cach xac dinh ty le nguon von on dinh rong),",
        "kem theo Thong tu so 50/2026/TT-NHNN ngay 30/9/2026, doc truc tiep tu van ban goc.",
        "",
        "LY DO LA PROXY, KHONG PHAI SO CHINH THUC:",
        "BCTC cong khai cua ngan hang KHONG cong bo du chi tiet de ap dung dung 100% bang he so:",
        "  - Khong tach tien gui theo muc do on dinh (phu thuoc han muc bao hiem tien gui tung",
        "    khach hang - thong tin noi bo, khong cong khai).",
        "  - Khong tach cho vay theo he so rui ro tin dung <=35% hay >35% (chi co trong bao cao",
        "    Tru cot 3 rieng, khong co trong BCTC chuan).",
        "  - Khong tach HQLA cap 1/2A/2B trong tong Chung khoan dau tu.",
        "  - Khong co ky han chi tiet cho tung dong (tien gui/cho vay/trai phieu...) - chi co",
        "    TONG Tai san/No phai tra theo 7 bucket ky han (du lieu 'gap_periods' co san).",
        "",
        "GIA DINH DUNG DE UOC TINH (xem chi tiet trong bank_nsfr_calc.py):",
        "  1. Ty trong ky han cua TUNG dong (tien gui, cho vay, trai phieu...) = ty trong ky han",
        "     TRUNG BINH cua toan bo Tai san/No phai tra ngan hang do (tu gap_periods).",
        "  2. Bucket 'tu 3-12 thang' (khong khop chinh xac moc 6 thang cua TT50) duoc chia",
        "     tuyen tinh: 3/9 vao 'duoi 6 thang', 6/9 vao '6 thang - duoi 1 nam'.",
        "  3. Tien gui khach hang: dung he so 'kem on dinh' (khoan 9.2/10.2 TT50: 90%/90%/100%)",
        "     cho TOAN BO (khong tach duoc ban le/to chuc, on dinh/kem on dinh).",
        "  4. Cho vay khach hang: dung he so rui ro <=35% (khoan 18.x.1 TT50: 50%/50%/65%) -",
        "     neu ngan hang co ty trong cho vay rui ro cao (BDS, tieu dung khong TSBD...) lon,",
        "     NSFR UOC TINH se CAO HON NSFR THAT (vi RSF thuc te se cao hon).",
        "  5. Chung khoan dau tu: dung he so 'chung khoan khac' (khoan 11.3 TT50: 50%/50%/100%)",
        "     cho TOAN BO - neu ngan hang giu nhieu TPCP (HQLA cap 1, RSF chi 5%), NSFR UOC TINH",
        "     se THAP HON NSFR THAT.",
        "  6. tpdn (trai phieu DN, trich tu ghi chu BCTC) KHONG duoc cong rieng - gia dinh da",
        "     nam trong 'inv_sec_bs' (tong chung khoan dau tu tren BCTC), tranh double-count.",
        "  7. Von chu so huu: 100% ASF moi ky han (dung khoan 1.1, khong tru von cap 2).",
        "",
        "KET LUAN SU DUNG: con so NSFR nay chi nen doc theo XU HUONG (tang/giam qua cac quy,",
        "so sanh TUONG DOI giua cac ngan hang) - KHONG nen trich dan nhu so NSFR CHINH THUC da",
        "bao cao len NHNN, vi cac gia dinh (4), (5) co the lam LECH ca 2 chieu tuy co cau tai",
        "san/nguon von thuc te cua tung ngan hang.",
    ]
    for i, line in enumerate(notes, start=1):
        ws3.cell(row=i, column=1, value=line)
    ws3.column_dimensions["A"].width = 100
    print(f"  [OK] Sheet {_NSFR_ASSUMPTIONS_SHEET_NAME}: {len(notes)} dong ghi chu")

    os.makedirs(out_dir, exist_ok=True)
    wb.save(xlsx_path)


if __name__ == "__main__":
    import sys
    out = compute_nsfr_all_banks()
    print(f"Tinh duoc NSFR cho {len(out)} bank", file=sys.stderr)
    for ticker, periods in list(out.items())[:1]:
        for period, r in sorted(periods.items()):
            print(ticker, period, f"NSFR={r['nsfr_pct']:.1f}%" if r['nsfr_pct'] else "N/A",
                  f"ASF={r['asf_total']:.0f}", f"RSF={r['rsf_total']:.0f}")
