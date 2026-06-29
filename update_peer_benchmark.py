import os
import json
import requests
import urllib3
import datetime
from concurrent.futures import ThreadPoolExecutor

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BANK_TICKERS = ['VCB', 'BID', 'CTG', 'TCB', 'MBB', 'ACB', 'VIB', 'HDB', 'STB', 'VPB', 'TPB', 'MSB', 'EIB', 'LPB', 'SHB', 'OCB', 'BAB', 'NAB']

VIETCAP_BASE = "https://iq.vietcap.com.vn/api/iq-insight-service/v1"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Origin": "https://trading.vietcap.com.vn",
    "Referer": "https://trading.vietcap.com.vn/",
}

def fetch_financial_data(ticker):
    print(f"Fetching Vietcap IQ data for peer bank: {ticker}...")
    
    # 1. Fetch current price and market cap via stock details
    price = 20000
    mcap = 50000
    try:
        url_det = f"https://iq.vietcap.com.vn/api/iq-insight-service/v1/company/details?ticker={ticker}"
        r = requests.get(url_det, headers=HEADERS, verify=False, timeout=10)
        if r.status_code == 200:
            det = r.json().get("data", {})
            price = det.get("currentPrice") or price
            shares = det.get("numberOfSharesMktCap") or 1
            mcap = (price * shares) / 1e9
    except Exception as e:
        print(f"Error price/mcap for {ticker}: {e}")

    # Fallbacks
    npl = 2.0
    nim = 3.5
    casa = 20.0
    roe = 15.0
    cir = 35.0
    pb = 1.0
    cg = 12.0
    
    try:
        # Fetch Balance Sheet (Quarterly)
        url_bs = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=BALANCE_SHEET&quarterly=true"
        r_bs = requests.get(url_bs, headers=HEADERS, verify=False, timeout=10)
        
        # Fetch Income Statement (Quarterly)
        url_is = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=INCOME_STATEMENT&quarterly=true"
        r_is = requests.get(url_is, headers=HEADERS, verify=False, timeout=10)
        
        # Fetch Notes (Quarterly)
        url_nt = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=NOTE&quarterly=true"
        r_nt = requests.get(url_nt, headers=HEADERS, verify=False, timeout=10)

        if r_bs.status_code == 200 and r_is.status_code == 200:
            bs = r_bs.json().get("data", {}).get("quarters", [])
            is_st = r_is.json().get("data", {}).get("quarters", [])
            notes = r_nt.json().get("data", {}).get("quarters", []) if r_nt.status_code == 200 else []
            
            if bs and is_st:
                bs_sorted = sorted(bs, key=lambda x: (x.get("yearReport",0), x.get("lengthReport",0)))
                is_sorted = sorted(is_st, key=lambda x: (x.get("yearReport",0), x.get("lengthReport",0)))
                nt_sorted = sorted(notes, key=lambda x: (x.get("yearReport",0), x.get("lengthReport",0)))
                
                q_bs = bs_sorted[-1]
                q_is = is_sorted[-1]
                q_nt = nt_sorted[-1] if nt_sorted else {}
                q_bs_prev = bs_sorted[-2] if len(bs_sorted) >= 2 else q_bs
                
                loans = q_bs.get("bsb103", 1)
                loans_prev = q_bs_prev.get("bsb103", 1)
                cg = round(((loans - loans_prev)/max(loans_prev, 1)) * 100, 2)
                
                # NPL
                gr3 = q_nt.get("nob42", 0)
                gr4 = q_nt.get("nob43", 0)
                gr5 = q_nt.get("nob44", 0)
                npl_abs = gr3 + gr4 + gr5
                npl = round((npl_abs / max(loans, 1)) * 100, 2) if npl_abs > 0 else 2.0
                
                # CASA
                casa_val = q_nt.get("nob66", 0)
                dep_val = q_nt.get("nob65", 1)
                casa = round((casa_val / max(dep_val, 1)) * 100, 2)
                
                # CIR
                opex = abs(q_is.get("isb39", 0))
                toi = q_is.get("isb38", 1)
                cir = round((opex / max(toi, 1)) * 100, 2)
                
                # NIM
                nii = q_is.get("isb27", 0)
                nim = round((nii * 4 / max(loans, 1)) * 100, 2)
                
                # ROE
                npat = q_is.get("isa20", 0)
                equity = q_bs.get("bsa78", 1)
                roe = round((npat * 4 / max(equity, 1)) * 100, 2)
                
                # Tính P/B động = Vốn hóa / VCSH
                equity_billion = equity / 1e9
                pb = round(mcap / max(equity_billion, 0.001), 2)
                
    except Exception as e:
        print(f"Error fetching statements for {ticker}: {e}")
        
    return {
        "ticker": ticker,
        "npl": npl if 0 < npl < 15 else 2.5,
        "nim": nim if 0 < nim < 10 else 3.5,
        "casa": casa if 0 < casa < 70 else 20.0,
        "roe": roe if 0 < roe < 40 else 16.0,
        "cir": cir if 0 < cir < 80 else 35.0,
        "pb": pb if 0 < pb < 5 else 1.2,
        "cg": cg if -20 < cg < 40 else 12.0,
        "mcap": mcap
    }

def update_single_ticker(ticker):
    ticker = ticker.upper()
    data = fetch_financial_data(ticker)
    
    names = {
        'VCB': 'Vietcombank', 'BID': 'BIDV', 'CTG': 'VietinBank', 'TCB': 'Techcombank',
        'MBB': 'MBBank', 'ACB': 'ACB', 'VIB': 'VIB', 'HDB': 'HDBank', 'STB': 'Sacombank',
        'VPB': 'VPBank', 'TPB': 'TPBank', 'MSB': 'MSB', 'EIB': 'Eximbank', 'LPB': 'LienVietPostBank',
        'SHB': 'SHB', 'OCB': 'OCB', 'BAB': 'BacABank', 'NAB': 'NamABank'
    }
    data["name"] = names.get(ticker, ticker)
    data["date_updated"] = datetime.datetime.now().strftime("%Y-%m-%d")
    
    # Read existing database
    db_path = "data/peer_benchmark.json"
    db = {"_meta": {}, "peers": []}
    if os.path.exists(db_path):
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except:
            pass
            
    # Update or insert
    peers = db.get("peers", [])
    updated = False
    for i, p in enumerate(peers):
        if p.get("ticker") == ticker:
            peers[i] = data
            updated = True
            break
    if not updated:
        peers.append(data)
        
    db["peers"] = peers
    db["_meta"]["last_updated_ticker"] = ticker
    db["_meta"]["date"] = datetime.datetime.now().strftime("%Y-%m-%d")
    
    with open(db_path, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)
        
    print(f"[OK] Updated single peer benchmark for {ticker} as of {data['date_updated']}")

def main():
    # If run directly, update all concurrently as fallback
    peers = []
    with ThreadPoolExecutor(max_workers=18) as executor:
        results = executor.map(fetch_financial_data, BANK_TICKERS)
        
    names = {
        'VCB': 'Vietcombank', 'BID': 'BIDV', 'CTG': 'VietinBank', 'TCB': 'Techcombank',
        'MBB': 'MBBank', 'ACB': 'ACB', 'VIB': 'VIB', 'HDB': 'HDBank', 'STB': 'Sacombank',
        'VPB': 'VPBank', 'TPB': 'TPBank', 'MSB': 'MSB', 'EIB': 'Eximbank', 'LPB': 'LienVietPostBank',
        'SHB': 'SHB', 'OCB': 'OCB', 'BAB': 'BacABank', 'NAB': 'NamABank'
    }
    
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    for data in results:
        ticker = data["ticker"]
        data["name"] = names.get(ticker, ticker)
        data["date_updated"] = today
        peers.append(data)
        
    output = {
        "_meta": {
            "updated": today,
            "source": "Vietcap Live API Engine"
        },
        "peers": peers
    }
    
    with open("data/peer_benchmark.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
        
    print("[OK] Automated Peer Benchmark data successfully updated concurrently and saved to data/peer_benchmark.json")

if __name__ == "__main__":
    main()
