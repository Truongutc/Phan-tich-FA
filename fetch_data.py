#!/usr/bin/env python3
"""
fetch_data.py — Fetch financial statements from Vietcap IQ API.
Can be used standalone:  python fetch_data.py HPG
Returns cached JSON + DataFrame for Excel model integration.
- Retry logic: 3 attempts with exponential backoff
- Timeout: 30s per request
- Parallel fetching for all sections
"""

import requests
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

VIETCAP_BASE = "https://iq.vietcap.com.vn/api/iq-insight-service/v1"
CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache")

SECTIONS = ["BALANCE_SHEET", "INCOME_STATEMENT", "CASH_FLOW", "NOTE"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Origin": "https://trading.vietcap.com.vn",
    "Referer": "https://trading.vietcap.com.vn/",
}

TIMEOUT = 30
MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds between retries


def _get_with_retry(url, retries=MAX_RETRIES):
    """GET request with retry logic and exponential backoff."""
    last_error = None
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except requests.exceptions.Timeout as e:
            last_error = e
            wait = RETRY_DELAY * (attempt + 1)
            print(f"  [Retry {attempt+1}/{retries}] Timeout on {url} — waiting {wait}s...")
            time.sleep(wait)
        except requests.exceptions.HTTPError as e:
            # 4xx errors won't benefit from retry
            raise
        except requests.exceptions.ConnectionError as e:
            last_error = e
            wait = RETRY_DELAY * (attempt + 1)
            print(f"  [Retry {attempt+1}/{retries}] Connection error on {url} — waiting {wait}s...")
            time.sleep(wait)
        except Exception as e:
            last_error = e
            wait = RETRY_DELAY * (attempt + 1)
            print(f"  [Retry {attempt+1}/{retries}] Error on {url}: {e} — waiting {wait}s...")
            time.sleep(wait)
    raise Exception(f"Failed after {retries} retries: {last_error}")


def fetch_metrics(ticker):
    url = f"{VIETCAP_BASE}/company/{ticker}/financial-statement/metrics"
    result = _get_with_retry(url)
    return result.get("data", {})


def fetch_section(ticker, section, quarterly=False):
    url = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section={section}"
    if quarterly:
        url += "&quarterly=true"
    result = _get_with_retry(url)
    return result.get("data", {})


def fetch_all(ticker, use_cache=True):
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"{ticker}_bctc.json")

    if use_cache and os.path.exists(cache_file):
        print(f"[Cache] Loading {ticker} from cache...")
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)

    print(f"[Fetch] Loading {ticker} financial data from Vietcap API...")
    data = {"ticker": ticker, "metrics": {}, "sections": {}, "companyName": ticker, "currentPrice": 0}

    # Fetch metrics first (needed to map fields)
    try:
        data["metrics"] = fetch_metrics(ticker)
    except Exception as e:
        print(f"  [WARN] Could not fetch metrics: {e}")

    # Fetch all sections in parallel for speed (both annual and quarterly)
    def fetch_one(section):
        try:
            print(f"  -> {section} (Annual)...")
            annual = fetch_section(ticker, section, quarterly=False)
            
            quarters = []
            if section != "NOTE":
                print(f"  -> {section} (Quarterly)...")
                q_data = fetch_section(ticker, section, quarterly=True)
                quarters = q_data.get("quarters") or q_data.get("years") or []
                
            return section, {
                "years": annual.get("years") or [],
                "quarters": quarters
            }
        except Exception as e:
            print(f"  [WARN] Could not fetch {section}: {e}")
            return section, {"years": [], "quarters": []}

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(fetch_one, sec): sec for sec in SECTIONS}
        for future in as_completed(futures):
            section, result = future.result()
            data["sections"][section] = result

    # Try to get company name and price
    try:
        details_url = f"{VIETCAP_BASE}/company/details?ticker={ticker}"
        details_resp = _get_with_retry(details_url)
        details = details_resp.get("data", {})
        if details:
            data["companyName"] = details.get("viOrganName") or details.get("enOrganName") or ticker
            data["currentPrice"] = details.get("currentPrice") or 0
    except Exception as e:
        print(f"  [WARN] Could not fetch company details: {e}")

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
