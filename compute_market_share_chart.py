import os
import json
import requests
import statistics
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Danh sách các CTCK niêm yết và TCBS
SECURITIES_TICKERS = ['SSI', 'VND', 'HCM', 'VCI', 'FTS', 'SHS', 'BSI', 'VIX', 'MBS',
                      'CTS', 'AGR', 'BVS', 'APG', 'ORS', 'TVS', 'VDS']

# Thị phần môi giới thực tế công bố của TCBS theo từng năm (HOSE) để tự động nội suy doanh thu
TCBS_HIST_SHARE = {
    2018: 0.035,
    2019: 0.041,
    2020: 0.048,
    2021: 0.052,
    2022: 0.058,
    2023: 0.063,
    2024: 0.068,
    2025: 0.072,
    2026: 0.075
}

VIETCAP_BASE = "https://iq.vietcap.com.vn/api/iq-insight-service/v1"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://trading.vietcap.com.vn/",
}

# ADTV thực tế lịch sử theo năm (tỷ VND/phiên)
ADTV_ANNUAL = {
    2018: 6500.0,
    2019: 4700.0,
    2020: 7400.0,
    2021: 23500.0,
    2022: 16800.0,
    2023: 14800.0,
    2024: 20500.0,
    2025: 28500.0,
    2026: 28000.0
}
TRADING_DAYS_PER_QUARTER = 62.5 # ~250 phiên / 4
ASSUMED_FEE_BPS = 15.0 # 0.15%

def fetch_quarterly_brokerage_rev(ticker):
    print(f"Fetching quarterly brokerage revenue for {ticker}...")
    try:
        url = f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=INCOME_STATEMENT&quarterly=true"
        r = requests.get(url, headers=HEADERS, verify=False, timeout=10)
        if r.status_code == 200:
            quarters = r.json().get("data", {}).get("quarters", [])
            # Lọc các bản ghi quý hợp lệ
            data = []
            for q in quarters:
                y = q.get("yearReport")
                qn = q.get("lengthReport")
                if y and qn in (1, 2, 3, 4):
                    # iss42 là doanh thu nghiệp vụ môi giới chứng khoán
                    brok_rev = q.get("iss42")
                    if brok_rev is not None:
                        data.append({
                            "year": y,
                            "quarter": qn,
                            "label": f"{y}Q{qn}",
                            "brok_rev": brok_rev / 1e9 # đổi sang tỷ VND
                        })
            # Sắp xếp theo trình tự thời gian tăng dần
            data = sorted(data, key=lambda x: (x["year"], x["quarter"]))
            return ticker, data
    except Exception as e:
        print(f"Error fetching quarterly brokerage for {ticker}: {e}")
    return ticker, []

def main():
    raw_data = {}
    for ticker in SECURITIES_TICKERS:
        tk, quarters = fetch_quarterly_brokerage_rev(ticker)
        if quarters:
            raw_data[tk] = quarters

    # Xác định các quý chung có đủ dữ liệu nhất của 16 cổ phiếu (hoặc ít nhất các cổ phiếu lớn)
    # Lấy danh sách tất cả các label quý xuất hiện
    all_quarter_labels = set()
    for quarters in raw_data.values():
        for q in quarters:
            all_quarter_labels.add(q["label"])
    
    sorted_labels = sorted(list(all_quarter_labels))
    
    # Tính thị phần ngụ ý cho từng CTCK tại từng quý
    # Market Share = Doanh thu môi giới quý / (ADTV năm/4 * 62.5 * Fee)
    market_shares = {} # {label: {ticker: share}}
    
    for label in sorted_labels:
        # Tách năm từ label dạng '2025Q1'
        try:
            year = int(label[:4])
        except:
            continue
        
        adtv = ADTV_ANNUAL.get(year, 20000.0)
        total_market_brok_capacity = adtv * TRADING_DAYS_PER_QUARTER * ASSUMED_FEE_BPS / 10000
        
        market_shares[label] = {}
        for ticker, quarters in raw_data.items():
            q_rec = next((x for x in quarters if x["label"] == label), None)
            if q_rec:
                # Thị phần = Doanh thu môi giới của CTCK / Quy mô doanh thu môi giới giả định toàn thị trường
                share = q_rec["brok_rev"] / total_market_brok_capacity if total_market_brok_capacity > 0 else 0.0
                market_shares[label][ticker] = round(share, 4)
        
        # Thêm TCBS vào danh sách thị phần
        tcbs_share = TCBS_HIST_SHARE.get(year, 0.06)
        market_shares[label]['TCBS'] = tcbs_share
        
        # Tạo thêm bản ghi doanh thu môi giới cho TCBS trong raw_data để các thuật toán so sánh tổng doanh thu hoạt động
        if 'TCBS' not in raw_data:
            raw_data['TCBS'] = []
        raw_data['TCBS'].append({
            "year": year,
            "quarter": int(label[-1]),
            "label": label,
            "brok_rev": tcbs_share * total_market_brok_capacity
        })

    # Tìm Top 8 CTCK có doanh thu môi giới lớn nhất trong 4 quý gần nhất
    recent_4_quarters = sorted_labels[-4:] if len(sorted_labels) >= 4 else sorted_labels
    ticker_recent_revs = {}
    for ticker, quarters in raw_data.items():
        recent_rev = sum(q["brok_rev"] for q in quarters if q["label"] in recent_4_quarters)
        ticker_recent_revs[ticker] = recent_rev
    
    top_8_tickers = sorted(ticker_recent_revs.keys(), key=lambda x: ticker_recent_revs[x], reverse=True)[:8]
    print("Top 8 CTCK có thị phần lớn nhất dựa trên 4 quý gần nhất:", top_8_tickers)

    # Chọn khoảng thời gian tối đa mà cả 8 cổ phiếu Top 8 đều có đủ dữ liệu
    common_labels = []
    for label in sorted_labels:
        has_all = True
        for ticker in top_8_tickers:
            if ticker not in market_shares[label] or market_shares[label][ticker] == 0.0:
                has_all = False
                break
        if has_all:
            common_labels.append(label)
            
    print(f"Khoảng thời gian tối đa đủ dữ liệu cho Top 8: {len(common_labels)} quý (từ {common_labels[0]} đến {common_labels[-1]})")

    # Giới hạn lấy tối đa 16 quý gần nhất để biểu đồ không quá dày đặc
    if len(common_labels) > 16:
        common_labels = common_labels[-16:]

    # Build output payload
    output = {
        "top_5": top_8_tickers, # Giữ key top_5 để tương thích với app_securities.js
        "quarters": common_labels,
        "data": {ticker: [market_shares[q].get(ticker, 0.0) for q in common_labels] for ticker in top_8_tickers}
    }
    
    os.makedirs("data", exist_ok=True)
    with open("data/market_share_history.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
        
    print("[OK] Đã lưu lịch sử thị phần vào data/market_share_history.json")

if __name__ == "__main__":
    main()
