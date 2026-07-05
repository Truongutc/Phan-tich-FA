import os
import json
import requests
import urllib3
import datetime
from concurrent.futures import ThreadPoolExecutor

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SECURITIES_TICKERS = ['SSI', 'VND', 'HCM', 'VCI', 'FTS', 'SHS', 'BSI', 'VIX', 'MBS',
                      'CTS', 'AGR', 'BVS', 'APG', 'ORS', 'TVS', 'VDS', 'TCX', 'VCK', 'VPX']

VIETCAP_BASE = "https://iq.vietcap.com.vn/api/iq-insight-service/v1"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Origin": "https://trading.vietcap.com.vn",
    "Referer": "https://trading.vietcap.com.vn/",
}

NAMES = {
    'SSI': 'SSI', 'VND': 'VNDIRECT', 'HCM': 'HSC', 'VCI': 'Vietcap', 'FTS': 'FPTS',
    'SHS': 'SHS', 'BSI': 'BSC', 'VIX': 'VIX', 'MBS': 'MBS', 'CTS': 'VietinBank Securities',
    'AGR': 'Agriseco', 'BVS': 'BVSC', 'APG': 'APG', 'ORS': 'Tien Phong Securities',
    'TVS': 'Thien Viet Securities', 'VDS': 'Rong Viet Securities',
    'TCX': 'TCBS', 'VCK': 'VPS', 'VPX': 'VPBankS'
}


def fetch_financial_data(ticker):
    print(f"Fetching Vietcap IQ data for peer CTCK: {ticker}...")
    price, shares, mcap = 20000.0, 1, 0.0
    try:
        url_det = f"{VIETCAP_BASE}/company/details?ticker={ticker}"
        r = requests.get(url_det, headers=HEADERS, verify=False, timeout=10)
        if r.status_code == 200:
            det = r.json().get("data", {})
            price = det.get("currentPrice") or price
            shares = det.get("numberOfSharesMktCap") or 1
            mcap = (price * shares) / 1e9
    except Exception as e:
        print(f"Error price/mcap for {ticker}: {e}")

    pb = pe = roe = margin_to_equity = None
    pct_margin = pct_brokerage = pct_tudoanh = None
    charter_capital = equity_b = None
    try:
        url_bs = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=BALANCE_SHEET&quarterly=true"
        url_is = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=INCOME_STATEMENT&quarterly=true"
        r_bs = requests.get(url_bs, headers=HEADERS, verify=False, timeout=10)
        r_is = requests.get(url_is, headers=HEADERS, verify=False, timeout=10)

        if r_bs.status_code == 200 and r_is.status_code == 200:
            bs = r_bs.json().get("data", {}).get("quarters", [])
            is_st = r_is.json().get("data", {}).get("quarters", [])
            if bs and is_st:
                bs_sorted = sorted(bs, key=lambda x: (x.get("yearReport", 0), x.get("lengthReport", 0)))
                is_sorted = sorted(is_st, key=lambda x: (x.get("yearReport", 0), x.get("lengthReport", 0)))
                q_bs = bs_sorted[-1]
                q_is = is_sorted[-1]

                equity = q_bs.get("bsa78", 1) or 1
                equity_b = equity / 1e9
                charter_capital = (q_bs.get("bsa80", 0) or 0) / 1e9
                margin_loans_b = (q_bs.get("bss215", 0) or 0) / 1e9

                total_rev = q_is.get("isa1", 1) or 1
                margin_rev = q_is.get("iss120", 0) or 0
                brokerage_rev = q_is.get("iss42", 0) or 0
                fvtpl_gain = q_is.get("iss115", 0) or 0
                fvtpl_loss = q_is.get("iss124", 0) or 0
                npat = q_is.get("isa20", 0) or 0

                pct_margin = round(margin_rev / total_rev * 100, 1)
                pct_brokerage = round(brokerage_rev / total_rev * 100, 1)
                pct_tudoanh = round((fvtpl_gain + fvtpl_loss) / total_rev * 100, 1)

                roe = round((npat * 4 / equity) * 100, 2)
                margin_to_equity = round(margin_loans_b / max(equity_b, 0.001), 2)
                pb = round(mcap / max(equity_b, 0.001), 2)
                eps_ttm = npat * 4 / max(shares, 1)
                pe = round(price / eps_ttm, 2) if eps_ttm and eps_ttm > 0 else None
    except Exception as e:
        print(f"Error fetching statements for {ticker}: {e}")

    return {
        "ticker": ticker,
        "name": NAMES.get(ticker, ticker),
        "price": price,
        "mcap": round(mcap, 1),
        "charter_capital": round(charter_capital, 1) if charter_capital is not None else None,
        "equity": round(equity_b, 1) if equity_b is not None else None,
        "pb": pb,
        "pe": pe,
        "roe": roe,
        "margin_to_equity": margin_to_equity,
        "pct_margin_rev": pct_margin,
        "pct_brokerage_rev": pct_brokerage,
        "pct_tudoanh_rev": pct_tudoanh,
    }


def main():
    with ThreadPoolExecutor(max_workers=16) as executor:
        results = list(executor.map(fetch_financial_data, SECURITIES_TICKERS))

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    output = {
        "_meta": {"updated": today, "source": "Vietcap Live API Engine",
                  "note": "Bảng thống kê THUẦN TÚY diễn giải/so sánh — KHÔNG dùng làm input định giá cho bất kỳ mã nào."},
        "peers": results,
    }
    os.makedirs("data", exist_ok=True)
    with open("data/peer_benchmark_securities.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print("[OK] Peer benchmark (CTCK) saved to data/peer_benchmark_securities.json")


if __name__ == "__main__":
    main()
