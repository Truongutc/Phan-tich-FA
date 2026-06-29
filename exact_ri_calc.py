def calc_ri(ticker, coe, g):
    # VIB data from diagnostic output
    if ticker == "VIB":
        bvps_base = 13779
        eps_fc = [2585, 3137, 3684]
    else: # TCB
        bvps_base = 25331
        eps_fc = [3160, 3804, 4441]
        
    print(f"\n--- {ticker} (COE={coe*100:.2f}%, g={g*100:.2f}%) ---")
    bv = bvps_base
    ri_list = []
    pv_ri = 0
    for i in range(3):
        bv_start = bv
        eps = eps_fc[i]
        cap_charge = bv_start * coe
        ri = eps - cap_charge
        ri_list.append(ri)
        df = 1 / (1 + coe) ** (i + 1)
        pv_ri += ri * df
        bv = bv_start + eps
        print(f"  Year {i+1}: BV_start={bv_start:.0f}, EPS={eps:.0f}, CapCharge={cap_charge:.0f}, RI={ri:.0f}, PV(RI)={ri*df:.0f}")
        
    cv = ri_list[-1] * (1 + g) / (coe - g)
    pv_cv = cv / (1 + coe) ** 3
    ri_val = bvps_base + pv_ri + pv_cv
    print(f"  PV RI 3 nam: {pv_ri:.0f}")
    print(f"  Continuing Value (CV): {cv:.0f}")
    print(f"  PV of CV: {pv_cv:.0f}")
    print(f"  Book Value (BVPS): {bvps_base:.0f}")
    print(f"  PV(CV) / BVPS: {pv_cv / bvps_base:.2%}")
    print(f"  RI Value: {ri_val:.0f}")

calc_ri("VIB", 0.1010, 0.03)
calc_ri("TCB", 0.1136, 0.03)
