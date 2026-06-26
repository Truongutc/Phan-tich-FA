import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import requests, json

s = requests.Session()
s.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://trading.vietcap.com.vn/",
})
s.get("https://trading.vietcap.com.vn/iq/company?ticker=HPG", timeout=10)

base = "https://trading.vietcap.com.vn/api/iq-insight-service/v1"

# 1. Statistics-Financial
r = s.get(f"{base}/company/HPG/statistics-financial", timeout=15)
data = r.json()["data"]
print(f"=== statistics-financial: {len(data)} records ===")
for d in data[:5]:
    print(f'  {d["year"]} Q{d["quarter"]}: PE={d["pe"]:.1f}  PB={d["pb"]:.2f}  EV/EBITDA={d["evToEbitda"]:.1f}  ROE={d["roe"]*100:.1f}%')
print(f"  ... ({len(data)} records total)")

# 2. Income Statement
r2 = s.get(f"{base}/company/HPG/financial-statement?section=INCOME_STATEMENT", timeout=15)
inc = r2.json()["data"]["years"]
print(f"\n=== INCOME_STATEMENT ({len(inc)} years) ===")
for y in inc[-5:]:
    print(f'  {y["yearReport"]}: Rev={y.get("isa1",0)/1e9:>8.0f}t GP={y.get("isa5",0)/1e9:>7.0f}t NPAT={y.get("isa20",0)/1e9:>7.0f}t EPS={y.get("isa23",0):>6.0f}')

# 3. Balance Sheet
r3 = s.get(f"{base}/company/HPG/financial-statement?section=BALANCE_SHEET", timeout=15)
bs = r3.json()["data"]["years"]
print(f"\n=== BALANCE_SHEET ({len(bs)} years) ===")
for y in bs[-5:]:
    print(f'  {y["yearReport"]}: Assets={y.get("bsa53",0)/1e9:>8.0f}t Equity={y.get("bsa78",0)/1e9:>7.0f}t Debt={y.get("bsa56",0)/1e9+y.get("bsa71",0)/1e9:>6.0f}t Cash={y.get("bsa2",0)/1e9:>6.0f}t')

# 4. Company details
r4 = s.get(f"{base}/company/details?ticker=HPG", timeout=15)
d = r4.json()["data"]
print(f"\n=== Company Details ===")
print(f'  Name: {d.get("enOrganName","")}')
print(f'  Ticker: {d["ticker"]} | Sector: {d.get("sector","")}')
print(f'  Price: {d["currentPrice"]:,.0f} VND')
print(f'  MarketCap: {d["marketCap"]/1e9:,.0f} ty')
print(f'  Shares: {d["numberOfSharesMktCap"]:,.0f}')
print(f'  Rating: {d["rating"]} (as of {d["ratingAsOf"]})')
print(f'  Target: {d["targetPrice"]:,.0f} VND')
print(f'  Upside: {d["upsideToTargetPercent"]*100:.1f}%')
print(f'  52W High/Low: {d["highestPrice1Year"]:,.0f} / {d["lowestPrice1Year"]:,.0f}')
print(f'  FreeFloat: {d["freeFloat"]:,} ({d["freeFloatPercentage"]*100:.0f}%)')
print(f'  Foreign: {d["foreignerPercentage"]*100:.1f}%')
print(f'  Analyst: {d.get("analyst","")}')
print(f'  ComGroup: {d.get("comGroupCode","")}')
