#!/usr/bin/env python3
"""
fetch_data.py — Fetch financial statements from Vietcap IQ API.
Can be used standalone:  python fetch_data.py HPG
Returns cached JSON + DataFrame for Excel model integration.
"""

import requests, json, os, sys

VIETCAP_BASE = "https://iq.vietcap.com.vn/api/iq-insight-service/v1"
CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache")

SECTIONS = ["BALANCE_SHEET", "INCOME_STATEMENT", "CASH_FLOW", "NOTE"]

def _session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://trading.vietcap.com.vn/",
    })
    s.get("https://trading.vietcap.com.vn/iq/company?ticker=HPG&tab=financial&isIndex=false&financialTab=financialStatement", timeout=10)
    return s

def fetch_metrics(ticker):
    url = f"{VIETCAP_BASE}/company/{ticker}/financial-statement/metrics"
    r = _session().get(url, timeout=15)
    r.raise_for_status()
    return r.json()["data"]

def fetch_section(ticker, section):
    url = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section={section}"
    r = _session().get(url, timeout=15)
    r.raise_for_status()
    return r.json()["data"]

def fetch_all(ticker, use_cache=True):
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"{ticker}_bctc.json")

    if use_cache and os.path.exists(cache_file):
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)

    print(f"[Fetch] Loading {ticker} financial data from Vietcap...")
    data = {"ticker": ticker, "metrics": {}, "sections": {}}

    data["metrics"] = fetch_metrics(ticker)
    for sec in SECTIONS:
        print(f"  -> {sec}...")
        data["sections"][sec] = fetch_section(ticker, sec)

    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, default=str)
    print(f"[OK] Cached to {cache_file}")
    return data

def section_to_years(data, section):
    """Return list of yearly records for a section."""
    return data["sections"][section].get("years", [])

def section_to_quarters(data, section):
    """Return list of quarterly records for a section."""
    return data["sections"][section].get("quarters", [])

def get_field_map(data, section):
    """Return {field_id: title_vi} for a section."""
    return {m["field"]: m["titleVi"] for m in data["metrics"].get(section, [])}

def get_val(records, year, field):
    """Get a value from yearly records for a given year and field."""
    for r in records:
        if r.get("yearReport") == year:
            v = r.get(field)
            if v is not None:
                return v / 1e9  # convert to billions
    return None

if __name__ == "__main__":
    ticker = sys.argv[1] if len(sys.argv) > 1 else "HPG"
    data = fetch_all(ticker, use_cache=False)

    print(f"\n=== {ticker} FINANCIAL DATA ===")
    for sec in SECTIONS:
        years = section_to_years(data, sec)
        quarters = section_to_quarters(data, sec)
        field_map = get_field_map(data, sec)
        print(f"\n{sec}: {len(years)} years, {len(quarters)} quarters, {len(field_map)} fields")

        if years:
            yr = years[-1]
            print(f"  Latest year ({yr.get('yearReport')}) sample fields:")
            shown = 0
            for k, v in sorted(yr.items()):
                if k in field_map and v not in (None, 0, "0") and shown < 10:
                    name = field_map[k][:40]
                    print(f"    {k}: {round(v/1e9,2)} ty  ({name})")
                    shown += 1
