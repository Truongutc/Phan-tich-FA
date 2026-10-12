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
chi tiet de ap dung dung 100% bang he so nay, vi thieu phan loai tien gui theo HAN MUC
BAO HIEM tung khach hang, phan loai cho vay theo HE SO RUI RO TIN DUNG chinh xac (Tru cot
3/Pillar 3 rieng), va phan tach HQLA cap 1/2A/2B trong tong "Chung khoan dau tu".

CAP NHAT 2026-10-12 (user: "cac ty le nay chung chung qua, tu tim nguon de tinh chinh xac
hon" — xem _load_bank_notes() ben duoi): da THAY 2 gia dinh "chung cho ca cuc" quan trong
nhat bang DU LIEU THAT cua tung bank, lay tu thuyet minh BCTC (Vietcap API, section NOTE):
  - Tien gui khach hang: tach THAT CASA (khong ky han)/Tiet kiem (ban le)/Co ky han khac
    (nghieng to chuc-DN) thay vi ap 1 he so ban le chung cho toan bo so du.
  - Cho vay khach hang: dung THANG co cau Ngan/Trung/Dai han THAT cua tung bank, thay vi
    suy tu ty trong bucket TAI SAN GOP CHUNG (gap_periods, von di tron ca cho vay lan
    chung khoan/tien gui TCTD khac vao 1 ty trong duy nhat).
Phan con lai (chung khoan dau tu, lien ngan hang, trai phieu, KBNN...) VAN dung ty trong
ky han TRUNG BINH cua toan bo Tai san/No phai tra (tu gap_periods, bank_alm_store) vi BCTC
khong cho ky han rieng cho cac dong nay — day la UOC TINH (proxy), KHONG PHAI so NSFR
chinh thuc ngan hang bao cao len NHNN. Moi gia dinh con lai ghi chu ro trong code va sheet
"Gia dinh & Gioi han" cua file Excel xuat ra.

CAU TRUC: moi field BCTC duoc gan vao 1 "nhom cong cu" (instrument group), moi nhom co
1 "lich he so" rieng theo dung % cua khoan tuong ung trong Phu luc II (xem ASF_SCHEDULES/
RSF_SCHEDULES ben duoi). Nhom nao co du lieu THAT theo ky han/loai (tien gui, cho vay) thi
dung thang; nhom con lai dung ty trong bucket TRUNG BINH tu gap_periods.
"""

import json
import os

import bank_universe
import fetch_data

DATA_DIR = "data/bank_alm"
_BCTC_CACHE_DIR = ".cache"

# THEM 2026-10-12 (user: "cac phan ty le nay no cung chung chung qua... tu tim cac nguon de danh
# gia tinh NSFR cho chinh xac nhe" — roi user paste 1 huong dan chi tiet goi y dung chinh thuyet
# minh BCTC "bang rui ro thanh khoan" + phan loai tien gui/cho vay de thay the gia dinh chung).
# Phat hien: Vietcap API (section=NOTE, da dung san cho .cache/*.json va template_banking.py) co
# SAN dung 7 dong thuyet minh RAT CU THE cho ca 26/26 bank (chi SSB thieu trong cache cu, phai
# fetch song moi):
#   nob65 = "Cac khoan tien gui phan theo loai tien gui" (= TONG, doi chieu khop voi
#           customer_deposits trong quarterly_balance_sheet)
#   nob66 = "Tien gui khong ky han" (CASA that cua TUNG bank, KHONG PHAI gia dinh)
#   nob67 = "Tien gui co ky han" (THUONG nghieng ve to chuc/doanh nghiep - KHAC "tiet kiem")
#   nob68 = "Tien gui tiet kiem" (THUONG nghieng ve ca nhan/ban le)
#   nob46/47/48 = "Cho vay ngan/trung/dai han" (CO CAU KY HAN THAT cua tung bank, thay cho
#           gia dinh dung ty trong bucket TAI SAN GOP chung tu gap_periods).
# Dung THAT cac dong nay thay vi ap 1 gia dinh "chung chung" cho moi bank nhu ban dau.
def _load_bank_notes(ticker):
    """Tra {period('YYYY-Qn'): {casa, term_other, savings, loans_short, loans_medium,
    loans_long}} (don vi TY dong) cho 1 ticker, uu tien doc tu cache .cache/{ticker}_bctc.json
    (co san tu cac lan chay truoc), fetch song qua Vietcap API neu cache thieu/rong (vd SSB)."""
    cache_path = os.path.join(_BCTC_CACHE_DIR, f"{ticker}_bctc.json")
    quarters = []
    if os.path.exists(cache_path):
        try:
            with open(cache_path, encoding="utf-8") as f:
                cached = json.load(f)
            quarters = cached.get("sections", {}).get("NOTE", {}).get("quarters", []) or []
        except Exception:
            quarters = []
    if not quarters:
        try:
            q_data = fetch_data.fetch_section(ticker, "NOTE", quarterly=True)
            quarters = q_data.get("quarters") or q_data.get("years") or []
        except Exception as e:
            print(f"  [WARN] NSFR: khong fetch duoc thuyet minh NOTE cho {ticker}: {e}")
            return {}

    out = {}
    for rec in quarters:
        year, q = rec.get("yearReport"), rec.get("lengthReport")
        if not year or not q or q not in (1, 2, 3, 4):
            continue
        period = f"{year}-Q{q}"
        casa = (rec.get("nob66") or 0) / 1e9
        term_other = (rec.get("nob67") or 0) / 1e9
        savings = (rec.get("nob68") or 0) / 1e9
        loans_short = (rec.get("nob46") or 0) / 1e9
        loans_medium = (rec.get("nob47") or 0) / 1e9
        loans_long = (rec.get("nob48") or 0) / 1e9
        if casa or term_other or savings or loans_short or loans_medium or loans_long:
            out[period] = {"casa": casa, "term_other": term_other, "savings": savings,
                            "loans_short": loans_short, "loans_medium": loans_medium,
                            "loans_long": loans_long}
    return out

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
    # THEM 2026-10-12: khi co du lieu thuyet minh THAT (_load_bank_notes) - tach rieng "Tien gui
    # co ky han" KHONG PHAI tiet kiem (nob67, thuong nghieng ve to chuc/doanh nghiep gui ky han,
    # khac han CASA/tiet kiem ca nhan) - ap he so THAP HON nhu tien gui doanh nghiep (khoan 12:
    # 50%/50%/100%) thay vi gop chung vao muc ban le 90%/90%/100% o tren.
    "customer_deposits_term_other": (0.50, 0.50, 1.00),
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


def compute_nsfr_for_period(snap, gap_entry, notes=None):
    """snap: 1 entry cua quarterly_balance_sheet (tong BCTC); gap_entry: 1 entry cua
    gap_periods (co liquidity_gap/liabilities_by_bucket); notes: entry tuong ung tu
    _load_bank_notes() (casa/savings/term_other/loans_short/medium/long) hoac None neu
    khong co - khi co, dung THAY cho gia dinh bucket chung de tinh tien gui KH + cho vay
    CHINH XAC hon theo tung bank (xem ghi chu 2026-10-12 o dau file). Tra dict chi tiet
    hoac None neu thieu du lieu bat buoc (liquidity_gap/liabilities_by_bucket rong)."""
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

    if notes and (notes["casa"] or notes["savings"] or notes["term_other"]):
        # Tien gui KH THAT: CASA (khong ky han, ap thang he so "khong ky han" 90% - khong can
        # chia bucket vi CASA ban chat KHONG CO ky han) + Tiet kiem (ban le, van chia bucket vi
        # co ky han) + Co ky han KHAC (nghieng to chuc/DN, he so thap hon) + phan du (ky quy/von
        # chuyen dung nho, ap mac dinh ban le).
        residual_dep = max(0.0, customer_deposits - notes["casa"] - notes["savings"] - notes["term_other"])
        customer_deposits_asf = (
            notes["casa"] * ASF_SCHEDULES["customer_deposits"][0]
            + _weighted_schedule(notes["savings"], liab_shares, ASF_SCHEDULES["customer_deposits"])
            + _weighted_schedule(notes["term_other"], liab_shares, ASF_SCHEDULES["customer_deposits_term_other"])
            + _weighted_schedule(residual_dep, liab_shares, ASF_SCHEDULES["customer_deposits"])
        )
    else:
        customer_deposits_asf = _weighted_schedule(customer_deposits, liab_shares, ASF_SCHEDULES["customer_deposits"])

    asf_components = {
        "equity": equity * ASF_SCHEDULES["equity"][2],  # VCSH luon 100%, khong theo bucket
        "customer_deposits": customer_deposits_asf,
        "interbank_liab": _weighted_schedule(interbank_liab, liab_shares, ASF_SCHEDULES["interbank_liab"]),
        "bonds_issued": _weighted_schedule(bonds_issued, liab_shares, ASF_SCHEDULES["bonds_issued"]),
        "kbnn": _weighted_schedule(kbnn, liab_shares, ASF_SCHEDULES["kbnn"]),
        "other_liab": _weighted_schedule(other_liab, liab_shares, ASF_SCHEDULES["other_liab"]),
    }
    asf_total = sum(asf_components.values())

    if notes and (notes["loans_short"] or notes["loans_medium"] or notes["loans_long"]):
        # Cho vay THAT theo ky han ngan/trung/dai han (thuyet minh) - ngan han (<=12 thang) VA
        # 6m-1y trong bang TT50 CUNG he so 50% nen gop lam 1; trung+dai han (>12 thang) ap he so
        # tu 1 nam (0.65, gia dinh rui ro tin dung <=35% - xem ghi chu RSF_SCHEDULES["loans"]).
        loans_rsf = (notes["loans_short"] * RSF_SCHEDULES["loans"][0]
                     + (notes["loans_medium"] + notes["loans_long"]) * RSF_SCHEDULES["loans"][2])
    else:
        loans_rsf = _weighted_schedule(loans, asset_shares, RSF_SCHEDULES["loans"])

    rsf_components = {
        "cash": 0.0,
        "sbv_dep": 0.0,
        "bank_dep": _weighted_schedule(bank_dep, asset_shares, RSF_SCHEDULES["bank_dep"]),
        "loans": loans_rsf,
        "inv_sec_bs": _weighted_schedule(inv_sec_bs, asset_shares, RSF_SCHEDULES["inv_sec_bs"]),
        "other_asset": other_asset * RSF_SCHEDULES["other_asset"][2],
    }
    rsf_total = sum(rsf_components.values())

    nsfr_pct = (asf_total / rsf_total * 100.0) if rsf_total else None
    return {
        "asf_total": asf_total, "rsf_total": rsf_total, "nsfr_pct": nsfr_pct,
        "asf_components": asf_components, "rsf_components": rsf_components,
        "total_assets": total_assets, "equity": equity,
        # Gia tri THO (chua nhan he so) - dung de tinh ty trong co cau cho phan nhan dinh dinh
        # tinh (qualitative_driver_notes) ben duoi, KHONG dung de tinh ASF/RSF (da tinh o tren).
        "raw": {"total_liab": total_liab, "customer_deposits": customer_deposits,
                "interbank_liab": interbank_liab, "loans": loans,
                "cash_sbv": cash + sbv_dep},
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
        notes_by_period = _load_bank_notes(ticker)
        bank_results = {}
        for period, gap_entry in gp.items():
            # gap_periods dung nhan "YYYY-FY" cho bao cao nam (kiem toan) - trong khi
            # quarterly_balance_sheet dung thang nhat "YYYY-Q4" cho CUNG 1 ky do (BCTC
            # kiem toan nam CHINH LA so cuoi Q4) - quy doi nhan truoc khi tra cuu.
            qbs_period = period.replace("-FY", "-Q4") if period.endswith("-FY") else period
            snap = qbs.get(qbs_period)
            if not snap:
                continue
            r = compute_nsfr_for_period(snap, gap_entry, notes_by_period.get(qbs_period))
            if r is not None:
                bank_results[qbs_period] = r
        if bank_results:
            results[ticker] = bank_results
    return results


# THEM 2026-10-12 (user gui bai viet HSC "Son Nguyen" ve NSFR vs SFL: "cac ngan hang co von chu so
# huu day, ty trong nguon von dai han cao, cung co cau tien gui thien ve ban le/SME se co loi the
# ro ret [ve ASF]. Nguoc lai, cac ngan hang phu thuoc nhieu vao nguon von tu doanh nghiep lon va
# lien ngan hang ngan han se chiu bat loi... Ve phia RSF, cac ngan hang co ty trong cho vay ngan
# han va rui ro thap se duoc yeu cau it nguon von on dinh hon. Nguoc lai, cac ngan hang co thien
# huong cho vay trung dai han lon... se bi ap dung trong so on dinh cao hon" - roi user: "ok them
# card danh gia do di nhe" - sinh NHAN DINH RULE-BASED (ghep cau tu nguong %, KHONG phai van AI tu
# do) bang cach SO TUNG BANK VOI TRUNG BINH HE THONG cung ky, dung dung 4 ty le ma bai viet nhac
# toi: VCSH/TongTS, TienGuiKH/NguonVon, LienNH/NguonVon, ChoVay/TongTS.
_DRIVER_THRESHOLD = 1.10  # lech >=10% so TB he thong moi coi la dang chu y (tranh nhieu/qua vun vat)


def qualitative_driver_notes(results, period):
    """results: output cua compute_nsfr_all_banks(); period: ky can danh gia (vd ky gan nhat).
    Tra {ticker: {'asf_note': str, 'rsf_note': str}} - rong neu bank khong co du lieu ky do."""
    rows = []
    for ticker, periods in results.items():
        r = periods.get(period)
        if not r or not r.get("raw"):
            continue
        raw = r["raw"]
        total_assets = r["total_assets"] or 0
        total_liab = raw["total_liab"] or 0
        if total_assets <= 0 or total_liab <= 0:
            continue
        rows.append({
            "ticker": ticker,
            "equity_pct": r["equity"] / total_assets * 100,
            "cust_dep_pct": raw["customer_deposits"] / total_liab * 100,
            "interbank_pct": raw["interbank_liab"] / total_liab * 100,
            "loans_pct": raw["loans"] / total_assets * 100,
        })
    if not rows:
        return {}

    avg = {k: sum(row[k] for row in rows) / len(rows)
           for k in ("equity_pct", "cust_dep_pct", "interbank_pct", "loans_pct")}

    out = {}
    for row in rows:
        asf_parts = []
        if row["equity_pct"] >= avg["equity_pct"] * _DRIVER_THRESHOLD:
            asf_parts.append(f"VCSH/Tổng TS {row['equity_pct']:.1f}% (TB hệ thống {avg['equity_pct']:.1f}%) — đệm vốn dày hơn TB, lợi thế ASF (vốn chủ sở hữu luôn 100%)")
        elif row["equity_pct"] <= avg["equity_pct"] / _DRIVER_THRESHOLD:
            asf_parts.append(f"VCSH/Tổng TS {row['equity_pct']:.1f}% (TB hệ thống {avg['equity_pct']:.1f}%) — đệm vốn mỏng hơn TB")
        if row["cust_dep_pct"] >= avg["cust_dep_pct"] * _DRIVER_THRESHOLD:
            asf_parts.append(f"Tiền gửi KH/Nguồn vốn {row['cust_dep_pct']:.1f}% (TB {avg['cust_dep_pct']:.1f}%) — lợi thế ASF (hệ số ổn định cao 90-95%)")
        if row["interbank_pct"] >= avg["interbank_pct"] * _DRIVER_THRESHOLD:
            asf_parts.append(f"Liên NH/Nguồn vốn {row['interbank_pct']:.1f}% (TB {avg['interbank_pct']:.1f}%) — bất lợi ASF (hệ số thấp 0-50%, phụ thuộc vốn ngắn hạn nhiều hơn TB)")
        asf_note = "; ".join(asf_parts) if asf_parts else "Cơ cấu nguồn vốn gần với TB hệ thống, không có yếu tố nổi bật."

        rsf_parts = []
        if row["loans_pct"] >= avg["loans_pct"] * _DRIVER_THRESHOLD:
            rsf_parts.append(f"Cho vay/Tổng TS {row['loans_pct']:.1f}% (TB {avg['loans_pct']:.1f}%) — áp lực RSF cao hơn TB (cho vay ≥1 năm 65-85% so với tiền mặt/HQLA 0-15%)")
        elif row["loans_pct"] <= avg["loans_pct"] / _DRIVER_THRESHOLD:
            rsf_parts.append(f"Cho vay/Tổng TS {row['loans_pct']:.1f}% (TB {avg['loans_pct']:.1f}%) — áp lực RSF thấp hơn TB (tỷ trọng tài sản thanh khoản/khác cao hơn)")
        rsf_note = "; ".join(rsf_parts) if rsf_parts else "Cơ cấu tài sản gần với TB hệ thống, không có yếu tố nổi bật."

        out[row["ticker"]] = {"asf_note": asf_note, "rsf_note": rsf_note}
    return out


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
        "CAP NHAT 2026-10-12 (user: 'cac ty le nay chung chung qua, tu tim nguon de tinh chinh",
        "xac hon'): da thay 2 gia dinh 'chung' quan trong nhat bang DU LIEU THAT cua TUNG bank,",
        "lay tu thuyet minh BCTC (Vietcap, section NOTE, da co san o .cache/*.json cho ca 26/26",
        "bank):",
        "  - Tien gui khach hang: KHONG con gia dinh 1 he so chung cho ca cuc - da TACH THAT",
        "    thanh CASA/khong ky han (nob66), Tien gui tiet kiem - ban le (nob68, ap he so ban",
        "    le 90-95%), va Tien gui co ky han KHAC - nghieng to chuc/DN (nob67, ap he so thap",
        "    hon nhu tien gui DN 50%/50%/100%).",
        "  - Cho vay khach hang: KHONG con dung ty trong bucket TAI SAN GOP chung tu gap_periods",
        "    - da dung THANG co cau Ngan/Trung/Dai han THAT cua tung bank (nob46/47/48).",
        "Nho 2 thay doi nay, ket qua NSFR gio PHAN HOA RO RET hon giua cac bank theo dung co cau",
        "huy dong/cho vay thuc te (truoc day dai da so bank ra 96-135%, khoang hep; nay 67-121%,",
        "nhieu bank duoi 100% - gan voi canh bao cua gioi phan tich rang NSFR se la thach thuc",
        "that cho khong it ngan hang VN).",
        "",
        "LY DO VAN LA PROXY (chua phai so chinh thuc), BCTC cong khai van con thieu:",
        "  - Khong tach tien gui theo muc do on dinh THEO HAN MUC BAO HIEM tung khach hang (vd",
        "    CASA/tiet kiem co the la 'on dinh' 95% hay 'kem on dinh' 90% tuy quy mo tung so -",
        "    dang dung muc 90% (kem on dinh) lam gia dinh THAN TRONG cho toan bo).",
        "  - Khong tach cho vay theo he so rui ro tin dung <=35% hay >35% (chi co trong bao cao",
        "    Tru cot 3 rieng, khong co trong BCTC chuan) - dang dung muc <=35% (65%) cho phan",
        "    cho vay trung+dai han.",
        "  - Khong tach HQLA cap 1/2A/2B trong tong Chung khoan dau tu (inv_sec_bs).",
        "",
        "GIA DINH CON LAI (xem chi tiet trong bank_nsfr_calc.py):",
        "  1. Tien gui tiet kiem (nob68) + phan du (ky quy/von chuyen dung nho) + Lien NH/Trai",
        "     phieu/KBNN/No khac: VAN chia theo ty trong ky han TRUNG BINH cua toan bo Tai",
        "     san/No phai tra ngan hang do (tu gap_periods) - vi BCTC KHONG cho ky han rieng",
        "     cho tung dong nay.",
        "  2. Bucket 'tu 3-12 thang' (khong khop chinh xac moc 6 thang cua TT50) duoc chia",
        "     tuyen tinh: 3/9 vao 'duoi 6 thang', 6/9 vao '6 thang - duoi 1 nam'.",
        "  3. Cho vay 'ngan han' (nob46, <=12 thang) ap chung 1 he so 50% (TT50 cho ca 'duoi 6",
        "     thang' va '6 thang-1 nam' CUNG 50% nen khong can tach them); 'trung han'+'dai han'",
        "     (nob47+48, >12 thang) ap he so 'tu 1 nam' = 65% (gia dinh rui ro tin dung <=35%).",
        "  4. Chung khoan dau tu: dung he so 'chung khoan khac' (khoan 11.3 TT50: 50%/50%/100%)",
        "     cho TOAN BO - neu ngan hang giu nhieu TPCP (HQLA cap 1, RSF chi 5%), NSFR UOC TINH",
        "     se THAP HON NSFR THAT.",
        "  5. tpdn (trai phieu DN, trich tu ghi chu BCTC) KHONG duoc cong rieng - gia dinh da",
        "     nam trong 'inv_sec_bs' (tong chung khoan dau tu tren BCTC), tranh double-count.",
        "  6. Von chu so huu: 100% ASF moi ky han (dung khoan 1.1, khong tru von cap 2).",
        "",
        "KET LUAN SU DUNG: con so NSFR nay chinh xac hon ban dau nhung VAN chi nen doc theo XU",
        "HUONG (tang/giam qua cac quy, so sanh TUONG DOI giua cac ngan hang) - KHONG nen trich",
        "dan nhu so NSFR CHINH THUC da bao cao len NHNN.",
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
