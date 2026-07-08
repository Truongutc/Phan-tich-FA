#!/usr/bin/env python3
"""
fetch_kcn_drivers.py — Tu dong lay du lieu dac thu KCN:
  quy dat, ty le lap day, gia cho thue, doanh thu chua thuc hien.

Pipeline (3 tang fallback):
  T1: Tim PDF bai phan tich CTCK tu cac source non-SPA
      -> Download -> Convert sang MD -> Extract regex
  T2: Parse BCTC PDF markdown da co (thuyet minh BCTC)
      -> Extract so lieu quy dat neu co
  T3: Balance Sheet analysis (luon kha dung)
      -> Deferred revenue / advance payments -> implied g%

Output: dict cho build_excel_kcn (sheet A3_KCN_Drivers)
"""
import os
import re
import sys
import json
import glob
import subprocess
import statistics as stats

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
RESEARCH_DIR = os.path.join(PROJECT_ROOT, "Research_PDF")

# ── BS FIELD CODES ─────────────────────────────────────────────────────
BS_ADV_ST    = "bsa58"   # Nguoi mua tra tien truoc (ngan han)
BS_ADV_LT    = "bsa170"  # Nguoi mua tra tien truoc (dai han)
BS_DEFER_REV = "bsa76"   # Doanh thu chua thuc hien
IS_REVENUE   = "isa1"    # Tong doanh thu thuan

# Danh sach broker non-SPA co the scrape truc tiep
BROKER_SEARCH_URLS = [
    "https://agriseco.com.vn/bao-cao-phan-tich?symbol={ticker}",
    "https://www.vdsc.com.vn/bao-cao.rv?keyword={ticker}",
    "https://yuanta.com.vn/vi/research/stock-report?keyword={ticker}",
    "https://pinetree.vn/research/search?q={ticker}",
    "https://www.vps.com.vn/research/reports?keyword={ticker}",
    "https://www.mirae-asset.com.vn/research?q={ticker}",
    "https://www.bsc.com.vn/Category/Index/85?search={ticker}",
    "https://www.phs.vn/noi-dung/bao-cao-phan-tich?q={ticker}",
]


def _parse_num(s):
    """Parse '1.234,5' hoac '1,234.5' -> float."""
    if not s:
        return None
    s = str(s).strip().replace(" ", "")
    if "," in s and "." in s:
        if s.index(",") > s.index("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _extract_kcn_metrics(text, min_ha=10, max_ha=50000):
    """
    Extract KCN metrics tu 1 doan text/markdown.
    Tra ve: {areas_ha, occupancy_pct, price_usd, new_ha}
    """
    text_l = text.lower()
    r = {"areas_ha": [], "occupancy_pct": [], "price_usd": [], "new_ha": []}

    # Dien tich (ha) - cac mau tim kiem
    ha_patterns = [
        r"(?:tong\s*)?dien\s*tich[^:\n]{0,30}[:]\s*([\d\.,]+)\s*ha",
        r"quy\s*dat[^:\n]{0,30}[:]\s*([\d\.,]+)\s*ha",
        r"kcn[^:\n]{0,30}[:]\s*([\d\.,]+)\s*ha",
        r"([\d]{3,4}(?:[,\.]\d+)?)\s*ha(?:\s|,|\.|;)",
        r"rong\s*([\d\.,]+)\s*ha",
        r"area[^:\n]{0,20}[:]\s*([\d\.,]+)\s*ha",
    ]
    for pat in ha_patterns:
        for m in re.finditer(pat, text_l, re.IGNORECASE):
            v = _parse_num(m.group(1))
            if v and min_ha <= v <= max_ha:
                r["areas_ha"].append(v)

    # Ty le lap day (%)
    occ_patterns = [
        r"(?:ty\s*le\s*)?lap\s*day[^:\n%\d]{0,20}[:?]?\s*([\d\.,]+)\s*%",
        r"occupancy[^:\n%\d]{0,20}[:?]?\s*([\d\.,]+)\s*%",
        r"cho\s*thue[^:\n%\d]{0,30}([\d\.,]+)\s*%",
        r"filled[^:\n%\d]{0,20}[:?]?\s*([\d\.,]+)\s*%",
        r"([\d\.,]+)\s*%\s*(?:lap\s*day|cho\s*thue)",
    ]
    for pat in occ_patterns:
        for m in re.finditer(pat, text_l, re.IGNORECASE):
            v = _parse_num(m.group(1))
            if v and 0 < v <= 100:
                r["occupancy_pct"].append(v)

    # Gia cho thue (USD/m2)
    price_patterns = [
        r"([\d]+(?:[\.]\d+)?)\s*usd\s*/\s*m[2]",
        r"gia\s*(?:cho\s*)?thue[^:\n\d]{0,20}[:]\s*([\d\.,]+)\s*usd",
        r"([\d]+(?:[\.]\d+)?)\s*usd(?:\s*/\s*(?:m[2]|sqm|sqm\s*|m2))?",
        r"lease\s*price[^:\n\d]{0,20}[:]\s*([\d\.,]+)",
    ]
    for pat in price_patterns:
        for m in re.finditer(pat, text_l, re.IGNORECASE):
            v = _parse_num(m.group(1))
            if v and 10 <= v <= 600:
                r["price_usd"].append(v)

    # Dien tich cho thue MOI trong nam (ha)
    new_ha_patterns = [
        r"(?:cho\s*thue|ban)\s*(?:moi|them)[^:\n\d]{0,20}[:]\s*([\d\.,]+)\s*ha",
        r"([\d\.,]+)\s*ha\s*(?:moi|trong\s*nam|leased)",
        r"newly\s*leased[^:\n\d]{0,20}[:]\s*([\d\.,]+)\s*ha",
        r"signed[^:\n\d]{0,20}[:]\s*([\d\.,]+)\s*ha",
    ]
    for pat in new_ha_patterns:
        for m in re.finditer(pat, text_l, re.IGNORECASE):
            v = _parse_num(m.group(1))
            if v and 0 < v <= 5000:
                r["new_ha"].append(v)

    return r


# ══════════════════════════════════════════════════════════════════════
# TANG 1: Tim va download PDF bai phan tich CTCK
# ══════════════════════════════════════════════════════════════════════

def _scrape_broker_pdf_links(ticker, timeout=10):
    """
    Quet cac website CTCK non-SPA tim link PDF bai phan tich.
    Tra ve list URL PDF co lien quan den ticker.
    """
    try:
        import requests
    except ImportError:
        return []

    found = []
    hdrs = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept-Language": "vi-VN,vi;q=0.9",
    }
    for url_tpl in BROKER_SEARCH_URLS:
        url = url_tpl.format(ticker=ticker)
        try:
            r = requests.get(url, headers=hdrs, timeout=timeout)
            if r.status_code != 200 or len(r.text) < 300:
                continue
            text = r.text
            # Tim tat ca PDF links
            pdf_links = re.findall(r'https?://[^\s"\'<>]+\.pdf', text, re.IGNORECASE)
            for link in pdf_links:
                # Uu tien PDF co ten ticker
                if ticker.upper() in link.upper():
                    found.insert(0, link)
                else:
                    # Kiem tra context xung quanh link
                    pos = text.find(link)
                    ctx = text[max(0, pos-300):pos+100] if pos >= 0 else ""
                    if (ticker.upper() in ctx.upper() or
                            any(k in ctx.lower() for k in ["khu cong", "kcn", "industrial"])):
                        found.append(link)
        except Exception:
            continue

    return list(dict.fromkeys(found))  # deduplicate


def _download_pdf(url, save_path, timeout=30):
    """Download PDF, tra ve True neu thanh cong."""
    try:
        import requests
        r = requests.get(url, timeout=timeout,
                         headers={"User-Agent": "Mozilla/5.0"},
                         stream=True)
        if r.status_code == 200:
            ct = r.headers.get("content-type", "")
            if "pdf" in ct.lower() or url.lower().endswith(".pdf"):
                with open(save_path, "wb") as f:
                    for chunk in r.iter_content(65536):
                        f.write(chunk)
                return os.path.getsize(save_path) > 5000
    except Exception:
        pass
    return False


def _convert_pdf_to_md(pdf_path, out_dir):
    """Convert PDF sang markdown dung opendataloader (giong BCTC pipeline)."""
    try:
        res = subprocess.run(
            [sys.executable, "bctc_pdf_tool.py", "convert-single",
             "--pdf", pdf_path, "--outdir", out_dir],
            cwd=PROJECT_ROOT, capture_output=True, text=True,
            encoding="utf-8", timeout=120
        )
        mds = sorted(
            glob.glob(os.path.join(out_dir, "*.md")),
            key=os.path.getmtime, reverse=True
        )
        return mds[0] if mds else None
    except Exception:
        pass

    # Fallback: goi truc tiep opendataloader
    try:
        res = subprocess.run(
            [sys.executable, "-c",
             f"from opendataloader_pdf import convert; convert('{pdf_path}', '{out_dir}')"],
            cwd=PROJECT_ROOT, capture_output=True, timeout=120
        )
        mds = sorted(glob.glob(os.path.join(out_dir, "*.md")),
                     key=os.path.getmtime, reverse=True)
        return mds[0] if mds else None
    except Exception:
        return None


def fetch_research_pdfs(ticker, verbose=True):
    """
    Tier 1: Tim, download, parse PDF bai phan tich CTCK.
    Tra ve dict KCN metrics (co the rong neu khong tim duoc).
    """
    aggregated = {"areas_ha": [], "occupancy_pct": [], "price_usd": [], "new_ha": []}

    pdf_links = _scrape_broker_pdf_links(ticker)
    if not pdf_links:
        if verbose:
            print(f"    [T1] Khong tim duoc PDF bao cao CTCK cho {ticker}")
        return aggregated

    if verbose:
        print(f"    [T1] Tim duoc {len(pdf_links)} PDF links tu CTCK")

    os.makedirs(RESEARCH_DIR, exist_ok=True)
    ticker_dir = os.path.join(RESEARCH_DIR, ticker)
    md_dir = os.path.join(ticker_dir, "md")
    os.makedirs(md_dir, exist_ok=True)

    for i, pdf_url in enumerate(pdf_links[:4]):
        fname = f"{ticker}_broker_{i+1}.pdf"
        pdf_path = os.path.join(ticker_dir, fname)
        if not os.path.exists(pdf_path):
            ok = _download_pdf(pdf_url, pdf_path)
            if not ok:
                continue
            if verbose:
                print(f"      Downloaded: {fname}")

        md_path = _convert_pdf_to_md(pdf_path, md_dir)
        if not md_path:
            continue

        try:
            text = open(md_path, encoding="utf-8", errors="ignore").read()
            found = _extract_kcn_metrics(text)
            for k in aggregated:
                aggregated[k].extend(found[k])
        except Exception:
            continue

    for k in aggregated:
        aggregated[k] = sorted(
            set(round(v, 1) for v in aggregated[k]),
            reverse=(k != "price_usd")
        )

    if verbose and any(aggregated.values()):
        print(f"    [T1] Ket qua: areas={aggregated['areas_ha'][:3]} ha, "
              f"occ={aggregated['occupancy_pct'][:3]}%, "
              f"price={aggregated['price_usd'][:3]} USD/m2")
    return aggregated


# ══════════════════════════════════════════════════════════════════════
# TANG 2: Parse BCTC PDF Markdown da co
# ══════════════════════════════════════════════════════════════════════

def parse_bctc_markdowns(ticker, verbose=True):
    """
    Tier 2: Quet cac file .md extract tu BCTC PDF, tim du lieu KCN.
    Tra ve dict KCN metrics.
    """
    aggregated = {"areas_ha": [], "occupancy_pct": [], "price_usd": [], "new_ha": []}

    md_dir = os.path.join(PROJECT_ROOT, "BCTC_PDF", ticker, "extracted_md")
    if not os.path.isdir(md_dir):
        if verbose:
            print(f"    [T2] Khong tim thay thu muc markdown BCTC: {md_dir}")
        return aggregated

    mds = sorted(glob.glob(os.path.join(md_dir, "*.md")), reverse=True)[:8]
    if verbose:
        print(f"    [T2] Quet {len(mds)} file markdown BCTC...")

    for md_path in mds:
        try:
            text = open(md_path, encoding="utf-8", errors="ignore").read()
            found = _extract_kcn_metrics(text)
            for k in aggregated:
                aggregated[k].extend(found[k])
        except Exception:
            continue

    for k in aggregated:
        aggregated[k] = sorted(
            set(round(v, 1) for v in aggregated[k]),
            reverse=(k != "price_usd")
        )

    if verbose:
        if aggregated["areas_ha"]:
            print(f"    [T2] areas_ha={aggregated['areas_ha'][:4]} ha")
        if aggregated["occupancy_pct"]:
            print(f"    [T2] occupancy={aggregated['occupancy_pct'][:4]}%")
        if aggregated["price_usd"]:
            print(f"    [T2] price={aggregated['price_usd'][:4]} USD/m2")
        if not any(aggregated.values()):
            print(f"    [T2] Khong trich xuat duoc du lieu KCN tu BCTC markdown")

    return aggregated


# ══════════════════════════════════════════════════════════════════════
# TANG 3: Balance Sheet — Deferred Revenue & Advance Analysis
# ══════════════════════════════════════════════════════════════════════

def bs_deferred_revenue_analysis(bs_recs_y, is_recs_y, hist_years, verbose=True):
    """
    Tier 3: Phan tich doanh thu chua thuc hien va nguoi mua tra tien truoc.

    Logic du bao tang truong:
    ─────────────────────────────────────────────────────────────────
    advance_st[Y] = tien da nhan, se ghi nhan DT trong nam Y+1
    => CAGR(advance_st) = toc do ky hop dong moi => g_leasing

    defer_rev[Y] = tong DT chua ghi nhan (phan bo 50 nam)
    => recognition_rate ≈ advance_st[Y] / defer_rev[Y]
    => implied_annual ≈ defer_last x avg_rate (ty/nam)

    adv_yoy[Y] = tang truong advance_st YoY
    => leading indicator cho DT nam toi
    ─────────────────────────────────────────────────────────────────
    """
    def get_bs(year, field):
        for r in bs_recs_y:
            if r.get("yearReport") == year:
                v = r.get(field)
                return v / 1e9 if v else 0.0
        return 0.0

    def get_is(year, field):
        for r in is_recs_y:
            if r.get("yearReport") == year:
                v = r.get(field)
                return v / 1e9 if v else 0.0
        return 0.0

    years = sorted(hist_years)

    adv_st    = {y: get_bs(y, BS_ADV_ST)    for y in years}
    adv_lt    = {y: get_bs(y, BS_ADV_LT)    for y in years}
    defer_rev = {y: get_bs(y, BS_DEFER_REV) for y in years}
    revenue   = {y: get_is(y, IS_REVENUE)   for y in years}

    # YoY growth cua advance_st
    adv_yoy = {}
    for i in range(1, len(years)):
        y, yp = years[i], years[i-1]
        prev = adv_st.get(yp, 0)
        if prev > 0:
            adv_yoy[y] = round((adv_st[y] - prev) / prev, 4)

    # CAGR advance_st 3-4 nam
    adv_series = [adv_st[y] for y in years[-4:] if adv_st.get(y, 0) > 0]
    if len(adv_series) >= 2:
        n = len(adv_series) - 1
        g_raw = (adv_series[-1] / adv_series[0]) ** (1 / n) - 1
        g_cagr = round(max(-0.15, min(0.35, g_raw)), 4)
    else:
        g_cagr = 0.07

    # Blend: 60% CAGR + 40% YoY trung binh gan nhat
    recent_yoy = [v for v in list(adv_yoy.values())[-2:] if v is not None]
    if recent_yoy:
        g_blend = round(0.6 * g_cagr + 0.4 * stats.mean(recent_yoy), 4)
        g_blend = max(-0.15, min(0.35, g_blend))
    else:
        g_blend = g_cagr

    # Ty le advance_st / revenue
    adv_to_rev = {}
    for y in years:
        rev = revenue.get(y, 0)
        adv = adv_st.get(y, 0)
        adv_to_rev[y] = round(adv / rev, 4) if rev > 0 else None

    # Ty le ghi nhan (recognition rate)
    recog_rates = []
    for y in years:
        d = defer_rev.get(y, 0)
        a = adv_st.get(y, 0)
        if d > 0 and a > 0:
            recog_rates.append(min(max(a / d, 0.01), 0.30))

    avg_recog_rate = stats.mean(recog_rates) if recog_rates else 0.05

    last_y = years[-1]
    defer_last  = defer_rev.get(last_y, 0)
    adv_st_last = adv_st.get(last_y, 0)
    adv_lt_last = adv_lt.get(last_y, 0)
    implied_annual = round(defer_last * avg_recog_rate, 1)

    if verbose:
        print(f"    [T3] Advance ST: "
              f"{[round(adv_st[y],0) for y in years]} ty")
        print(f"    [T3] Defer Rev:  "
              f"{[round(defer_rev[y],0) for y in years]} ty")
        print(f"    [T3] g_CAGR={g_cagr*100:.1f}% | g_blend={g_blend*100:.1f}% | "
              f"adv_st_last={adv_st_last:.0f}ty | defer_last={defer_last:.0f}ty "
              f"=> ~{implied_annual:.0f}ty ghi nhan/nam")

    return {
        "hist_years":       years,
        "advance_st":       adv_st,
        "advance_lt":       adv_lt,
        "deferred_rev":     defer_rev,
        "revenue":          revenue,
        "adv_yoy":          adv_yoy,
        "adv_to_rev":       adv_to_rev,
        "g_leasing_cagr":   g_cagr,
        "g_leasing_blend":  g_blend,
        "avg_recog_rate":   round(avg_recog_rate, 4),
        "implied_annual":   implied_annual,
        "adv_st_last":      adv_st_last,
        "adv_lt_last":      adv_lt_last,
        "defer_last":       defer_last,
        "note": (
            f"g={g_blend*100:.1f}% (blend CAGR+YoY advance_st); "
            f"adv_st={adv_st_last:.0f}ty=floor DT thue dat Y+1; "
            f"defer={defer_last:.0f}ty~>{implied_annual:.0f}ty ghi nhan/nam"
        ),
    }


# ══════════════════════════════════════════════════════════════════════
# PUBLIC ENTRY POINT
# ══════════════════════════════════════════════════════════════════════

def fetch_kcn_drivers(ticker, bs_recs_y, is_recs_y, hist_years, verbose=True):
    """
    Entry point chinh: lay du lieu dac thu KCN tu 3 tang fallback.
    Tra ve dict day du cho build_excel_kcn -> sheet A3_KCN_Drivers.
    """
    if verbose:
        print(f"  [KCN Drivers] Auto-fetching operational data for {ticker}...")

    result = {"ticker": ticker}

    # Tang 1: Broker research PDFs
    t1 = fetch_research_pdfs(ticker, verbose)
    result["research_pdf"] = t1

    # Tang 2: BCTC PDF markdowns
    t2 = parse_bctc_markdowns(ticker, verbose)
    result["bctc_pdf"] = t2

    # Merge T1 + T2 (uu tien T2 - BCTC chinh thuc)
    def merge_dedup(a, b, reverse=True):
        merged = sorted(set(round(v, 1) for v in (a + b)), reverse=reverse)
        return merged

    merged_areas  = merge_dedup(t2["areas_ha"],     t1["areas_ha"],     reverse=True)
    merged_occ    = merge_dedup(t2["occupancy_pct"],t1["occupancy_pct"],reverse=True)
    merged_price  = merge_dedup(t2["price_usd"],    t1["price_usd"],    reverse=False)
    merged_new_ha = merge_dedup(t2["new_ha"],       t1["new_ha"],       reverse=True)

    # Tang 3: BS analysis (luon co)
    t3 = bs_deferred_revenue_analysis(bs_recs_y, is_recs_y, hist_years, verbose)
    result["bs"] = t3

    # Tong hop
    total_area  = merged_areas[0]   if merged_areas   else None
    sub_areas   = merged_areas[1:5] if len(merged_areas) > 1 else []
    occ_latest  = round(stats.mean(merged_occ[:3]), 1) if merged_occ else None
    usd_prices  = [p for p in merged_price if 20 <= p <= 500]
    price_ref   = round(stats.median(usd_prices), 0) if usd_prices else 200.0
    new_ha_ref  = merged_new_ha[0] if merged_new_ha else None

    result["summary"] = {
        # KCN physical data (tu PDF neu co)
        "total_area_ha":       total_area,
        "sub_areas_ha":        sub_areas,
        "occupancy_pct":       occ_latest,
        "lease_price_usd":     price_ref,
        "new_ha_recent":       new_ha_ref,

        # BS-derived (luon co)
        "advance_st_last":     t3["adv_st_last"],
        "advance_lt_last":     t3["adv_lt_last"],
        "defer_rev_last":      t3["defer_last"],
        "implied_annual_rev":  t3["implied_annual"],
        "avg_recog_rate":      t3["avg_recog_rate"],
        "g_leasing_blend":     t3["g_leasing_blend"],
        "g_leasing_cagr":      t3["g_leasing_cagr"],
        "adv_yoy":             t3["adv_yoy"],
        "adv_to_rev":          t3["adv_to_rev"],
        "advance_st_series":   t3["advance_st"],
        "advance_lt_series":   t3["advance_lt"],
        "deferred_rev_series": t3["deferred_rev"],
        "revenue_series":      t3["revenue"],

        # Data sources traceability
        "sources": {
            "physical":  "T2-BCTC" if t2["areas_ha"]       else ("T1" if t1["areas_ha"]       else "N/A"),
            "occupancy": "T2-BCTC" if t2["occupancy_pct"]  else ("T1" if t1["occupancy_pct"]  else "N/A"),
            "price":     "T2-BCTC" if t2["price_usd"]      else ("T1" if t1["price_usd"]      else "Default 200 USD/m2"),
            "growth":    "T3-BS advance_st CAGR+YoY blend",
        },
        "note": t3["note"],
    }

    if verbose:
        s = result["summary"]
        print(f"  [KCN Drivers] Summary for {ticker}:")
        print(f"    area={s['total_area_ha']} ha | occ={s['occupancy_pct']}% | "
              f"price={s['lease_price_usd']} USD/m2 | new_ha={s['new_ha_recent']} ha/nam")
        print(f"    adv_st={s['advance_st_last']:.0f}ty | defer={s['defer_rev_last']:.0f}ty | "
              f"g_fc={s['g_leasing_blend']*100:.1f}% | implied={s['implied_annual_rev']:.0f}ty/nam")

    return result


# ── STANDALONE TEST ────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys as _sys
    ticker = _sys.argv[1] if len(_sys.argv) > 1 else "SIP"
    cache_path = os.path.join(PROJECT_ROOT, ".cache", f"{ticker}_bctc.json")
    raw = json.load(open(cache_path, encoding="utf-8"))
    bs_y = raw["sections"]["BALANCE_SHEET"]["years"]
    is_y = raw["sections"]["INCOME_STATEMENT"]["years"]
    hist_years = sorted({r["yearReport"] for r in is_y if r.get("yearReport")})[-5:]
    result = fetch_kcn_drivers(ticker, bs_y, is_y, hist_years, verbose=True)
    print("\n=== SUMMARY JSON ===")
    print(json.dumps(result["summary"], ensure_ascii=False, indent=2, default=str))