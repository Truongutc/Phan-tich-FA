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

    # ── STEP 1: Check for or generate a specialized builder ───────────────────
    builder_path = os.path.join(PROJECT_ROOT, f"build_{ticker.lower()}_model.py")

    if not os.path.exists(builder_path):
        print(f"[Step 1] No specialized builder found for {ticker}.")
        print(f"         Generating build_{ticker.lower()}_model.py via Gemini AI...")
        generated = generate_model_builder(ticker)

        if not generated or not os.path.exists(builder_path):
            print(f"[ERROR] Could not generate builder for {ticker}. Aborting.")
            print("  Possible causes:")
            print("  - GEMINI_API_KEY is not set")
            print("  - Network connectivity issue")
            sys.exit(1)

        print(f"[Step 1] Builder generated successfully: {os.path.basename(builder_path)}")
    else:
        print(f"[Step 1] Found existing builder: {os.path.basename(builder_path)}")

    # ── STEP 2: Run the builder ────────────────────────────────────────────────
    print(f"\n[Step 2] Running {os.path.basename(builder_path)}...")
    success = run_builder_script(builder_path)

    if not success:
        print(f"[ERROR] Builder failed for {ticker}. Check the script for errors.")
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
    if excel_path and os.path.exists(excel_path):
        _, excel_url = google_drive_uploader.upload_file(excel_path)
    if pdf_path and os.path.exists(pdf_path):
        _, pdf_url = google_drive_uploader.upload_file(pdf_path)

    # ── STEP 5: Read or build JSON summary ────────────────────────────────────
    print(f"\n[Step 5] Updating data registry...")
    ticker_json = load_ticker_json(ticker)

    # Patch the upload URLs into the JSON file
    if ticker_json:
        ticker_json["gdriveExcelUrl"] = excel_url
        ticker_json["gdrivePdfUrl"] = pdf_url
        json_path = os.path.join(PROJECT_ROOT, "data", f"{ticker}.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(ticker_json, f, ensure_ascii=False, indent=2)
        print(f"[OK] Patched GDrive URLs into data/{ticker}.json")
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
