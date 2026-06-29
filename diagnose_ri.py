"""
Diagnostic script: trace toan bo luong tinh RI valuation cho VIB va TCB
Kiem tra tung buoc: EPS, BVPS, RI, CV, PV(CV) va so sanh voi Book/Share
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import json, os, statistics as stats

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

def load_cache(ticker):
    path = os.path.join(PROJECT_ROOT, ".cache", f"{ticker}_bctc.json")
    with open(path, encoding='utf-8') as f:
        return json.load(f)

def get_yr(records, year, field):
    for r in records:
        if r.get("yearReport") == year:
            v = r.get(field)
            if v is not None:
                return v / 1e9
    return 0

def diagnose(ticker):
    print(f"\n{'='*60}")
    print(f"  DIAGNOSTIC: {ticker} — Residual Income Valuation Trace")
    print(f"{'='*60}")
    
    raw = load_cache(ticker)
    is_recs = raw["sections"]["INCOME_STATEMENT"].get("years", [])
    bs_recs = raw["sections"]["BALANCE_SHEET"].get("years", [])
    
    available_years = sorted(set(r.get("yearReport") for r in is_recs if r.get("yearReport")))
    years_hist = available_years[-5:]
    years_fc = [years_hist[-1]+1, years_hist[-1]+2, years_hist[-1]+3]
    
    # Lay so lieu lich su
    np_hist = [get_yr(is_recs, y, "isa20") for y in years_hist]
    equity_hist = [get_yr(bs_recs, y, "bsa78") for y in years_hist]
    loans_hist = [get_yr(bs_recs, y, "bsb103") for y in years_hist]
    cust_dep_hist = [get_yr(bs_recs, y, "bsb113") for y in years_hist]
    bonds_hist = [get_yr(bs_recs, y, "bsb116") for y in years_hist]
    charter_hist = [get_yr(bs_recs, y, "bsa80") for y in years_hist]
    
    # Tinh shares tu charter capital
    charter_capital = charter_hist[-1] * 1e9  # VND
    shares = int(charter_capital / 10000)
    print(f"\n[Data] Charter Capital (ty): {charter_hist[-1]:.0f}")
    print(f"[Data] Shares Outstanding: {shares:,}")
    
    # Lich su 5 nam
    print(f"\n[History] LNST (ty):")
    for y, v in zip(years_hist, np_hist):
        print(f"  {y}: {v:,.1f} ty")
    print(f"[History] VCSH (ty):")
    for y, v in zip(years_hist, equity_hist):
        print(f"  {y}: {v:,.1f} ty")
    
    # BVPS hien tai (cuoi nam lich su cuoi)
    bvps_hist_calc = [equity_hist[i] * 1e9 / shares for i in range(len(years_hist))]
    eps_hist_calc = [np_hist[i] * 1e9 / shares for i in range(len(years_hist))]
    
    print(f"\n[BVPS] BVPS lich su (VND):")
    for y, b in zip(years_hist, bvps_hist_calc):
        print(f"  {y}: {b:,.0f}")
    print(f"\n[EPS] EPS lich su (VND):")
    for y, e in zip(years_hist, eps_hist_calc):
        print(f"  {y}: {e:,.0f}")
    
    bvps_base = bvps_hist_calc[-1]
    print(f"\n[KEY] BVPS hien tai (base): {bvps_base:,.0f} VND")
    
    # Forecast assumptions (giong template_banking.py)
    loans_growth_fc = [0.14, 0.13, 0.12]
    dep_growth_fc = [0.12, 0.11, 0.10]
    iea_growth_fc = [0.13, 0.12, 0.11]
    nim_fc = [0.0340, 0.0345, 0.0350]
    cir_fc = [0.33, 0.32, 0.32]
    coc_fc = [0.013, 0.012, 0.011]
    non_int_growth_fc = [0.15, 0.15, 0.14]
    tax_rate = 0.20
    
    # Lay so lieu can thiet cho du phong
    cash_hist = [get_yr(bs_recs, y, "bsa2") for y in years_hist]
    sbv_dep_hist = [get_yr(bs_recs, y, "bsb97") for y in years_hist]
    bank_dep_hist = [get_yr(bs_recs, y, "bsb98") for y in years_hist]
    inv_sec_bs_hist = [get_yr(bs_recs, y, "bsb106") for y in years_hist]
    nii_hist = [get_yr(is_recs, y, "isb27") for y in years_hist]
    toi_hist = [get_yr(is_recs, y, "isb38") for y in years_hist]
    
    # Build forecast
    loans_fc = []
    dep_fc = []
    iea_end_hist = [loans_hist[i] + bank_dep_hist[i] + inv_sec_bs_hist[i] + cash_hist[i] + sbv_dep_hist[i] for i in range(len(years_hist))]
    iea_fc = []
    for i in range(3):
        loans_fc.append(loans_hist[-1] * (1+loans_growth_fc[i]) if i==0 else loans_fc[i-1]*(1+loans_growth_fc[i]))
        dep_fc.append(cust_dep_hist[-1] * (1+dep_growth_fc[i]) if i==0 else dep_fc[i-1]*(1+dep_growth_fc[i]))
        prev_iea = iea_end_hist[-1]
        iea_fc.append(prev_iea * (1+iea_growth_fc[i]) if i==0 else iea_fc[i-1]*(1+iea_growth_fc[i]))
    
    iea_avg_fc = [(iea_end_hist[-1] + iea_fc[0])/2]
    for i in range(1, 3):
        iea_avg_fc.append((iea_fc[i-1] + iea_fc[i])/2)
    
    nii_fc = [iea_avg_fc[i] * nim_fc[i] for i in range(3)]
    base_non_int = toi_hist[-1] - nii_hist[-1]
    non_int_fc = []
    for i in range(3):
        non_int_fc.append(base_non_int * (1+non_int_growth_fc[i]) if i==0 else non_int_fc[i-1]*(1+non_int_growth_fc[i]))
    toi_fc = [nii_fc[i] + non_int_fc[i] for i in range(3)]
    opex_fc = [toi_fc[i] * cir_fc[i] for i in range(3)]
    ppop_fc = [toi_fc[i] - opex_fc[i] for i in range(3)]
    avg_loans = [(loans_hist[-1]+loans_fc[0])/2, (loans_fc[0]+loans_fc[1])/2, (loans_fc[1]+loans_fc[2])/2]
    prov_fc = [avg_loans[i] * coc_fc[i] for i in range(3)]
    pbt_fc = [ppop_fc[i] - prov_fc[i] for i in range(3)]
    np_fc = [pbt_fc[i] * (1 - tax_rate) for i in range(3)]
    
    eps_fc_calc = [np_fc[i] * 1e9 / shares for i in range(3)]
    
    print(f"\n[Forecast NII] (ty):")
    for y, v in zip(years_fc, nii_fc):
        print(f"  {y}F: {v:,.1f}")
    print(f"[Forecast Non-II] (ty): base={base_non_int:.1f}")
    for y, v in zip(years_fc, non_int_fc):
        print(f"  {y}F: {v:,.1f}")
    print(f"[Forecast LNST] (ty):")
    for y, v in zip(years_fc, np_fc):
        print(f"  {y}F: {v:,.1f}")
    print(f"[Forecast EPS] (VND):")
    for y, v in zip(years_fc, eps_fc_calc):
        print(f"  {y}F: {v:,.0f}")
    
    # COE
    COE = 0.045 + 1.0 * 0.07  # fallback: Rf=4.5%, beta=1.0, ERP=7%
    terminal_growth = 0.03
    
    print(f"\n[COE] {COE*100:.2f}% (fallback estimation)")
    print(f"[g]   {terminal_growth*100:.2f}%")
    
    # RI Model - step by step
    print(f"\n{'─'*50}")
    print(f"  RESIDUAL INCOME STEP-BY-STEP")
    print(f"{'─'*50}")
    print(f"  BVPS base (book/share current): {bvps_base:,.0f} VND")
    
    ri_results = []
    bv = bvps_base
    for i in range(3):
        bv_start = bv
        eps_i = eps_fc_calc[i]
        capital_charge = bv_start * COE
        ri = eps_i - capital_charge
        ri_results.append(ri)
        bv = bv_start + eps_i
        
        print(f"\n  Year {years_fc[i]}F:")
        print(f"    EPS             = {eps_i:>10,.0f} VND")
        print(f"    BVPS đầu kỳ     = {bv_start:>10,.0f} VND")
        print(f"    Capital Charge  = {capital_charge:>10,.0f} VND  ({bv_start:,.0f} × {COE*100:.2f}%)")
        print(f"    Residual Income = {ri:>10,.0f} VND")
        roe_implied = eps_i / bv_start * 100
        print(f"    Implied ROE     = {roe_implied:.2f}%  (vs COE={COE*100:.2f}%  RI {'> 0 OK' if ri > 0 else '< 0 WARN'})")
    
    # PV calculations
    pv_ri = sum(ri_results[i] / (1+COE)**(i+1) for i in range(3))
    cv = ri_results[-1] * (1+terminal_growth) / (COE - terminal_growth)
    pv_cv = cv / (1+COE)**3
    ri_value = bvps_base + pv_ri + pv_cv
    
    print(f"\n{'─'*50}")
    print(f"  SUMMARY VALUATION")
    print(f"{'─'*50}")
    print(f"  PV của RI 3 năm:          {pv_ri:>10,.0f} VND")
    print(f"  Continuing Value (CV):    {cv:>10,.0f} VND  = {ri_results[-1]:,.0f} × 1.03 / ({COE:.4f} - {terminal_growth})")
    print(f"  PV của CV:                {pv_cv:>10,.0f} VND")
    print(f"  Book/Share (BVPS base):   {bvps_base:>10,.0f} VND")
    print(f"")
    ratio = pv_cv / bvps_base
    flag = "⚠️  PV(CV) > BVPS!" if pv_cv > bvps_base else "✅ PV(CV) < BVPS (OK)"
    print(f"  PV(CV) / BVPS = {ratio:.2%}  {flag}")
    print(f"")
    print(f"  RI VALUE = {ri_value:>12,.0f} VND")
    
    # Check formula
    print(f"\n{'─'*50}")
    print(f"  EXCEL FORMULA CHECK (B8 = PV of Continuing Value)")
    print(f"{'─'*50}")
    print(f"  Formula: =(D29*(1+B3)/(B2-B3))*D30")
    print(f"  D29 = RI year 3    = {ri_results[2]:,.0f}")
    print(f"  B3  = g            = {terminal_growth}")
    print(f"  B2  = COE          = {COE:.4f}")
    print(f"  D30 = 1/(1+COE)^3  = {1/(1+COE)**3:.4f}")
    print(f"  = ({ri_results[2]:,.0f} × {1+terminal_growth} / ({COE:.4f}-{terminal_growth})) × {1/(1+COE)**3:.4f}")
    print(f"  = {cv:,.0f} × {1/(1+COE)**3:.4f}")
    print(f"  = {pv_cv:,.0f}  {'✅ Correct' if abs(pv_cv - cv/(1+COE)**3) < 1 else '❌ Mismatch'}")

# Run for VIB and TCB
for t in ["VIB", "TCB"]:
    try:
        diagnose(t)
    except Exception as e:
        print(f"\n[ERROR] {t}: {e}")
