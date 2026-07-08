#!/usr/bin/env python3
"""
run_analysis.py — Main orchestrator to run stock analysis, upload files, and update registry.

Flow for any ticker:
  1. Check if build_<ticker>_model.py exists
  2. If NOT, call generate_stock_model_builder.py to create it (requires GEMINI_API_KEY)
  3. Run build_<ticker>_model.py
  4. Upload Excel & PDF to Google Drive
  5. Update data/<TICKER>.json and data/index.json
"""
import os
import sys
import json
import subprocess
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

import google_drive_uploader
from generate_stock_model_builder import generate_model_builder


def update_registry(ticker, company_name, sector, excel_url, pdf_url):
    """Updates the index file listing all analyzed stocks for the dashboard."""
    data_dir = os.path.join(PROJECT_ROOT, "data")
    os.makedirs(data_dir, exist_ok=True)
    registry_path = os.path.join(data_dir, "index.json")

    registry = []
    if os.path.exists(registry_path):
        try:
            with open(registry_path, "r", encoding="utf-8") as f:
                registry = json.load(f)
        except Exception:
            registry = []

    # Remove existing entry if any (refresh)
    registry = [entry for entry in registry if entry.get("ticker") != ticker]

    registry.append({
        "ticker": ticker,
        "companyName": company_name,
        "sector": sector,
        "lastUpdated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "excelUrl": excel_url,
        "pdfUrl": pdf_url,
    })

    with open(registry_path, "w", encoding="utf-8") as f:
        json.dump(registry, f, ensure_ascii=False, indent=2)
    print(f"[OK] Updated registry at: {registry_path}")


def run_builder_script(script_path: str) -> bool:
    """Execute a builder script and return True if successful."""
    print(f"\n[Runner] Executing: {os.path.basename(script_path)}")
    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=False,  # stream output live
            check=True,
            cwd=PROJECT_ROOT
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Builder script failed with exit code {e.returncode}")
        return False


def load_ticker_json(ticker: str) -> dict:
    """Load the generated data/<TICKER>.json file."""
    json_path = os.path.join(PROJECT_ROOT, "data", f"{ticker}.json")
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def find_latest_output_files(ticker: str):
    """Scan Bao cao/<TICKER>/ for the latest Excel and PDF files."""
    out_dir = os.path.join(PROJECT_ROOT, "Bao cao", ticker)
    if not os.path.isdir(out_dir):
        return None, None

    xlsx_files = sorted([f for f in os.listdir(out_dir) if f.endswith(".xlsx")])
    pdf_files  = sorted([f for f in os.listdir(out_dir) if f.endswith(".pdf")])

    excel_path = os.path.join(out_dir, xlsx_files[-1]) if xlsx_files else None
    pdf_path   = os.path.join(out_dir, pdf_files[-1])  if pdf_files  else None
    return excel_path, pdf_path


def run_analysis(ticker: str):
    ticker = ticker.upper()
    print(f"\n{'='*60}")
    print(f"  STOCK ANALYSIS PIPELINE — {ticker}")
    print(f"{'='*60}\n")

    # Danh sách mã ngân hàng
    BANKING_TICKERS = {"VCB", "BID", "TCB", "MBB", "VPB", "ACB", "HDB", "TPB", "STB",
                       "LPB", "ABB", "SHB", "VAB", "VIB", "BAB", "KLB", "NAB", "NVB",
                       "SGB", "OCB", "EIB", "MSB"}

    # Danh sách mã công ty chứng khoán (CTCK)
    SECURITIES_TICKERS = {"SSI", "VND", "HCM", "VCI", "FTS", "SHS", "BSI", "VIX", "MBS",
                          "CTS", "AGR", "BVS", "APG", "ORS", "TVS", "VDS", "TCX", "VCK", "VPX"}

    # Danh sách mã BĐS Khu Công nghiệp (KCN) — dùng BCTC doanh nghiệp thường (isa/bsa),
    # KHÔNG có field đặc thù để tự nhận diện như bank/CTCK nên phân loại theo danh sách tĩnh
    KCN_TICKERS = {"SIP", "IDC", "SZC", "SZL", "KBC", "NTC", "DPR", "BCM", "PHR",
                   "LHG", "VGC", "D2D", "TIP"}

    # ── STEP 1: Fetch raw data ───────────────────────────────────────────────
    import fetch_data
    print(f"[Step 1] Fetching raw financial data for {ticker}...")
    try:
        raw_data = fetch_data.fetch_all(ticker, use_cache=True)
    except Exception as e:
        print(f"[ERROR] Could not fetch data for {ticker}: {e}")
        sys.exit(1)

    # Tự động nhận diện Ngân hàng dựa trên cấu trúc tài khoản BCTC thực tế
    bs_recs = raw_data["sections"]["BALANCE_SHEET"].get("years", [])
    has_bank_accounts = False
    has_securities_accounts = False
    if bs_recs:
        # Kiểm tra xem có tài khoản Cho vay khách hàng (bsb103) hoặc Tiền gửi (bsb113) không
        latest_bs = bs_recs[-1]
        if latest_bs.get("bsb103") or latest_bs.get("bsb113") or latest_bs.get("bsb116"):
            has_bank_accounts = True
        # Tự động nhận diện CTCK: có tài khoản dư nợ cho vay Margin (bss215, riêng của mẫu BCTC CTCK).
        if latest_bs.get("bss215"):
            has_securities_accounts = True

    # ── STEP 2: Run specialized builder (AI-generated) or template engines ─────────
    builder_path = os.path.join(PROJECT_ROOT, f"build_{ticker.lower()}_model.py")
    is_bank = ticker in BANKING_TICKERS or has_bank_accounts
    # Ticker nằm trong danh sách KCN tĩnh được ƯU TIÊN hơn auto-detect CTCK
    is_kcn = not is_bank and ticker in KCN_TICKERS
    is_securities = not is_bank and not is_kcn and (ticker in SECURITIES_TICKERS or has_securities_accounts)

    if is_bank:
        print(f"\n[Step 2] Ticker {ticker} classified as Bank. Directly running upgraded template_banking.py...")
        try:
            print("[Runner] Updating all peer benchmark data from Live API...")
            import update_peer_benchmark
            update_peer_benchmark.main()
        except Exception as e:
            print(f"[WARN] Failed to update peer benchmark dynamically: {e}")
            
        import template_banking
        success = template_banking.run_banking_analysis(ticker, raw_data)
    elif is_securities:
        print(f"\n[Step 2] Ticker {ticker} classified as Securities (CTCK). Directly running template_securities.py...")
        try:
            print("[Runner] Updating peer benchmark (CTCK) data from Live API...")
            import update_peer_benchmark_securities
            update_peer_benchmark_securities.main()
        except Exception as e:
            print(f"[WARN] Failed to update securities peer benchmark dynamically: {e}")

        import template_securities
        success = template_securities.run_securities_analysis(ticker, raw_data)
    elif is_kcn:
        print(f"\n[Step 2] Ticker {ticker} classified as KCN (BĐS Khu Công nghiệp). Directly running template_kcn.py...")
        try:
            print("[Runner] Updating peer benchmark (KCN) data from Live API...")
            import update_peer_benchmark_kcn
            update_peer_benchmark_kcn.main()
        except Exception as e:
            print(f"[WARN] Failed to update KCN peer benchmark dynamically: {e}")

        import template_kcn
        success = template_kcn.run_kcn_analysis(ticker, raw_data)
    else:
        # 1. Try to generate specialized builder via Gemini if key is present and builder doesn't exist yet
        if not os.path.exists(builder_path) and os.environ.get("GEMINI_API_KEY"):
            print(f"\n[Step 2] Generating build_{ticker.lower()}_model.py via Gemini AI...")
            generate_model_builder(ticker)
            
        # 2. Run specialized builder if it exists (either pre-existing or just generated)
        if os.path.exists(builder_path):
            print(f"\n[Step 2] Running specialized builder: {os.path.basename(builder_path)}...")
            success = run_builder_script(builder_path)
        else:
            # 3. Otherwise fall back to static template engines
            print(f"\n[Step 2] No specialized builder. Non-bank stock. Running template_generic.py...")
            import template_generic
            success = template_generic.run_generic_analysis(ticker, raw_data)

    if not success:
        print(f"[ERROR] Analysis step failed for {ticker}.")
        sys.exit(1)

    # ── STEP 3: Find generated output files ───────────────────────────────────
    excel_path, pdf_path = find_latest_output_files(ticker)

    if not excel_path and not pdf_path:
        print(f"[WARN] No output files found in Bao cao/{ticker}/")
        print("  The builder may not have saved files to the expected location.")

    print(f"\n[Step 3] Output files:")
    print(f"  Excel: {excel_path or 'NOT FOUND'}")
    print(f"  PDF:   {pdf_path or 'NOT FOUND'}")


    # ── STEP 4: Upload to Google Drive ────────────────────────────────────────
    excel_url = None
    pdf_url = None

    print(f"\n[Step 4] Uploading to Google Drive...")
    # Load ticker_json early to find the sector name
    ticker_json = load_ticker_json(ticker)
    upload_sector = "Ngân hàng" if is_bank else (ticker_json.get("sector") if (ticker_json and ticker_json.get("sector")) else "Unknown")
    
    if excel_path and os.path.exists(excel_path):
        _, excel_url = google_drive_uploader.upload_file(excel_path, sector=upload_sector, ticker=ticker)
    if pdf_path and os.path.exists(pdf_path):
        _, pdf_url = google_drive_uploader.upload_file(pdf_path, sector=upload_sector, ticker=ticker)

    # ── STEP 5: Read or build JSON summary ────────────────────────────────────
    print(f"\n[Step 5] Updating data registry...")

    # Patch the upload URLs into the JSON file
    if ticker_json:
        ticker_json["gdriveExcelUrl"] = excel_url
        ticker_json["gdrivePdfUrl"] = pdf_url
        ticker_json["lastUpdated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        json_path = os.path.join(PROJECT_ROOT, "data", f"{ticker}.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(ticker_json, f, ensure_ascii=False, indent=2)
        print(f"[OK] Patched GDrive URLs and lastUpdated into data/{ticker}.json")
    else:
        # Build a basic JSON if the builder didn't create one
        print(f"[WARN] data/{ticker}.json not found, creating minimal entry...")
        ticker_json = {
            "ticker": ticker,
            "companyName": ticker,
            "sector": "Unknown",
            "currentPrice": 0,
            "marketCap": 0,
            "shares": 0,
            "gdriveExcelUrl": excel_url,
            "gdrivePdfUrl": pdf_url,
            "lastUpdated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "data": {"years": [], "revenue": [], "npat": [], "eps": [], "equity": []}
        }
        os.makedirs(os.path.join(PROJECT_ROOT, "data"), exist_ok=True)
        json_path = os.path.join(PROJECT_ROOT, "data", f"{ticker}.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(ticker_json, f, ensure_ascii=False, indent=2)

    update_registry(
        ticker,
        ticker_json.get("companyName", ticker),
        ticker_json.get("sector", "Unknown"),
        excel_url,
        pdf_url
    )

    # ── DONE ──────────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"  ANALYSIS COMPLETE — {ticker}")
    print(f"{'='*60}")
    print(f"  Excel: {excel_path}")
    print(f"  PDF:   {pdf_path}")
    if excel_url:
        print(f"  GDrive Excel: {excel_url}")
    if pdf_url:
        print(f"  GDrive PDF:   {pdf_url}")
    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python run_analysis.py <TICKER>")
        print("Example: python run_analysis.py TCB")
        sys.exit(1)
    run_analysis(sys.argv[1])
