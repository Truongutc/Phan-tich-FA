#!/usr/bin/env python3
"""
generate_stock_model_builder.py
————————————————————————————————————————————
Given any Vietnamese stock ticker, this script:
1. Fetches financial data + company metadata from Vietcap API
2. Loads the analysis rules from INSTRUCTIONS_PHAN_TICH_CO_PHIEU.md + SKILL files
3. Constructs a rich, structured prompt for the Gemini API
4. Asks Gemini to write a complete, customized `build_<ticker>_model.py` script
5. Saves the generated Python file into the project root

Usage:
    python generate_stock_model_builder.py TCB
    GEMINI_API_KEY=... python generate_stock_model_builder.py SSI
"""

import os
import sys
import json
import requests
import re

# ─── ENSURE UTF-8 OUTPUT ON ALL PLATFORMS ─────────────────────────────────────
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass
from datetime import datetime

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
INSTRUCTIONS_PATH = os.path.join(PROJECT_ROOT, "mau", "INSTRUCTIONS_PHAN_TICH_CO_PHIEU.md")
XUAT_BAO_CAO_SKILL_PATH = os.path.join(PROJECT_ROOT, ".opencode", "skills", "xuat-bao-cao", "SKILL.md")
HPG_TEMPLATE_PATH = os.path.join(PROJECT_ROOT, "build_hpg_model.py")
VIETCAP_BASE = "https://trading.vietcap.com.vn/api/iq-insight-service/v1"

# Industry classification map (from phan-loai-nganh skill)
INDUSTRY_MAP = {
    "VCB": "Ngân hàng", "BID": "Ngân hàng", "TCB": "Ngân hàng", "MBB": "Ngân hàng",
    "VPB": "Ngân hàng", "ACB": "Ngân hàng", "HDB": "Ngân hàng", "TPB": "Ngân hàng",
    "STB": "Ngân hàng", "LPB": "Ngân hàng", "ABB": "Ngân hàng", "SHB": "Ngân hàng",
    "VAB": "Ngân hàng", "VIB": "Ngân hàng", "BAB": "Ngân hàng", "KLB": "Ngân hàng",
    "NAB": "Ngân hàng", "NVB": "Ngân hàng", "SGB": "Ngân hàng", "OCB": "Ngân hàng",
    "EIB": "Ngân hàng", "MSB": "Ngân hàng",
    "SSI": "Chứng khoán", "VND": "Chứng khoán", "HCM": "Chứng khoán",
    "VCI": "Chứng khoán", "FTS": "Chứng khoán", "SHS": "Chứng khoán",
    "VHM": "Bất động sản dân cư", "NLG": "Bất động sản dân cư", "KDH": "Bất động sản dân cư",
    "DXG": "Bất động sản dân cư", "PDR": "Bất động sản dân cư", "NVL": "Bất động sản dân cư",
    "IDC": "Bất động sản KCN", "KBC": "Bất động sản KCN", "BCM": "Bất động sản KCN",
    "SZC": "Bất động sản KCN", "VGC": "Bất động sản KCN",
    "HPG": "Thép", "HSG": "Thép", "NKG": "Thép", "TVN": "Thép",
    "MWG": "Bán lẻ", "FRT": "Bán lẻ", "PNJ": "Bán lẻ", "DGW": "Phân phối",
    "FPT": "Công nghệ", "CMG": "Công nghệ",
    "GMD": "Logistics", "VSC": "Cảng biển",
    "VNM": "Thực phẩm FMCG", "MSN": "Tập đoàn đa ngành",
    "DGC": "Hóa chất", "DCM": "Phân bón", "DPM": "Phân bón",
    "DHG": "Dược phẩm", "TRA": "Dược phẩm", "IMP": "Dược phẩm",
    "HAG": "Nông nghiệp", "BAF": "Nông nghiệp",
    "POW": "Điện", "NT2": "Điện", "REE": "Điện",
    "KSV": "Khoáng sản",
}

BANKING_TICKERS = {
    "VCB", "BID", "TCB", "MBB", "VPB", "ACB", "HDB", "TPB", "STB", "LPB",
    "ABB", "SHB", "VAB", "VIB", "BAB", "KLB", "NAB", "NVB", "SGB", "OCB",
    "EIB", "MSB",
}
SECURITIES_TICKERS = {"SSI", "VND", "HCM", "VCI", "FTS", "SHS"}


def _vietcap_session():
    """Create and warm up a Vietcap session."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://trading.vietcap.com.vn/",
    })
    try:
        s.get("https://trading.vietcap.com.vn/", timeout=8)
    except Exception:
        pass
    return s


def fetch_vietcap_metadata(ticker: str) -> dict:
    """Fetch company details, key ratios, and income statement snapshot."""
    s = _vietcap_session()
    result = {
        "ticker": ticker,
        "companyName": ticker,
        "sector": INDUSTRY_MAP.get(ticker, "General"),
        "currentPrice": 10000,
        "shares": 1_000_000_000,
        "marketCap": 10_000_000_000_000,
        "financial_summary": {},
        "years_available": [],
    }

    # 1. Company details
    try:
        r = s.get(f"{VIETCAP_BASE}/company/details?ticker={ticker}", timeout=10)
        r.raise_for_status()
        d = r.json().get("data", {})
        result["companyName"] = d.get("organName", d.get("enOrganName", ticker))
        result["currentPrice"] = d.get("currentPrice", 10000)
        result["shares"] = d.get("numberOfSharesMktCap", 1_000_000_000)
        result["marketCap"] = d.get("marketCap", 10_000_000_000_000)
        result["exchange"] = d.get("comGroupCode", "HOSE")
        result["sector_api"] = d.get("sector", "")
        mcap_t = result['marketCap'] / 1e12 if result['marketCap'] else 0
        print(f"[Meta] {ticker}: {result['companyName']} | Price: {result['currentPrice']:,} | MCap: {mcap_t:.1f} nghin ty")
    except Exception as e:
        print(f"[WARN] Could not fetch company details: {e}")

    # 2. IS: last 5 years
    try:
        r = s.get(f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=INCOME_STATEMENT", timeout=15)
        r.raise_for_status()
        years = r.json().get("data", {}).get("years", [])
        years = sorted(years, key=lambda x: x.get("yearReport", 0))[-5:]
        result["years_available"] = [y.get("yearReport") for y in years]
        # Revenue: isa3 (non-bank net revenue), fallback to isb38 (TOI for banks)
        def _rev(y):
            v = y.get("isa3") or y.get("isa1") or y.get("isb38") or y.get("isb22") or 0
            return round(v / 1e9, 1)
        result["financial_summary"]["revenue"] = [_rev(y) for y in years]
        result["financial_summary"]["gross_profit"] = [round((y.get("isa5") or 0) / 1e9, 1) for y in years]
        # NPAT: isa22 (non-bank), isa20 fallback, then isb46 (banking)
        def _npat(y):
            v = y.get("isa22") or y.get("isa20") or y.get("isb46") or 0
            return round(v / 1e9, 1)
        result["financial_summary"]["npat"] = [_npat(y) for y in years]
        result["financial_summary"]["eps"] = [y.get("isa23", 0) for y in years]
        # Banking specific (NII = isb27, TOI = isb38, PPOP = isb40)
        result["financial_summary"]["nii"]  = [round((y.get("isb27") or 0) / 1e9, 1) for y in years]
        result["financial_summary"]["toi"]  = [round((y.get("isb38") or 0) / 1e9, 1) for y in years]
        result["financial_summary"]["ppop"] = [round((y.get("isb40") or 0) / 1e9, 1) for y in years]
        result["financial_summary"]["pbt"]  = [round((y.get("isa16") or 0) / 1e9, 1) for y in years]
        print(f"[Meta] Available years: {result['years_available']}")
        print(f"[Meta] Revenue (ty): {result['financial_summary']['revenue']}")
        print(f"[Meta] NPAT (ty):    {result['financial_summary']['npat']}")
        if any(result["financial_summary"]["nii"]):
            print(f"[Meta] NII (ty):     {result['financial_summary']['nii']}")
    except Exception as e:
        print(f"[WARN] Could not fetch IS data: {e}")

    # 3. Balance Sheet: total assets, equity, loans
    try:
        r = s.get(f"{VIETCAP_BASE}/company/{ticker}/financial-statement?section=BALANCE_SHEET", timeout=15)
        r.raise_for_status()
        bs_years = r.json().get("data", {}).get("years", [])
        bs_years = sorted(bs_years, key=lambda x: x.get("yearReport", 0))[-5:]
        result["financial_summary"]["total_assets"] = [round(y.get("bsa53", y.get("bsa50", 0)) / 1e9, 1) for y in bs_years]
        result["financial_summary"]["equity"] = [round(y.get("bsa78", y.get("bsa75", 0)) / 1e9, 1) for y in bs_years]
        result["financial_summary"]["cash"] = [round(y.get("bsa2", y.get("bsa1", 0)) / 1e9, 1) for y in bs_years]
        result["financial_summary"]["short_debt"] = [round(y.get("bsa56", 0) / 1e9, 1) for y in bs_years]
        result["financial_summary"]["long_debt"] = [round(y.get("bsa71", 0) / 1e9, 1) for y in bs_years]
        result["financial_summary"]["loans"] = [round(y.get("bsb24", 0) / 1e9, 1) for y in bs_years]
        result["financial_summary"]["deposits"] = [round(y.get("bsb33", 0) / 1e9, 1) for y in bs_years]
    except Exception as e:
        print(f"[WARN] Could not fetch BS data: {e}")

        r = s.get(f"{VIETCAP_BASE}/company/{ticker}/statistics-financial", timeout=10)
        r.raise_for_status()
        ratios = r.json().get("data", [])
        # All quarters sorted chronologically
        all_quarters = sorted([x for x in ratios if x.get("quarter") in (1,2,3,4)], key=lambda x: (x.get("year", 0), x.get("quarter", 0)))
        result["financial_summary"]["pe_quarters"] = [round(x.get("pe", 0) or 0, 2) for x in all_quarters]
        result["financial_summary"]["pb_quarters"] = [round(x.get("pb", 0) or 0, 2) for x in all_quarters]
        result["financial_summary"]["quarter_labels"] = [f"{x.get('year')}-Q{x.get('quarter')}" for x in all_quarters]
        
        # Keep last 5 annual entries
        annual_ratios = sorted([x for x in ratios if x.get("quarter") == 4], key=lambda x: x.get("year", 0))[-5:]
        result["financial_summary"]["pe_hist"] = [round(x.get("pe", 0) or 0, 1) for x in annual_ratios]
        result["financial_summary"]["pb_hist"] = [round(x.get("pb", 0) or 0, 2) for x in annual_ratios]
        result["financial_summary"]["roe_hist"] = [round((x.get("roe", 0) or 0) * 100, 1) for x in annual_ratios]
    except Exception as e:
        print(f"[WARN] Could not fetch ratios data: {e}")

    return result



def load_instructions_excerpt() -> str:
    """Load key analysis rules from the instructions file (first 2000 chars for prompt size)."""
    excerpt = ""
    if os.path.exists(INSTRUCTIONS_PATH):
        with open(INSTRUCTIONS_PATH, "r", encoding="utf-8") as f:
            content = f.read()
        # Just take the first ~3000 chars (overview + quy trinh)
        excerpt = content[:3000]
    return excerpt


def load_export_skill_excerpt() -> str:
    """Load key export/model rules from the xuat-bao-cao skill."""
    excerpt = ""
    if os.path.exists(XUAT_BAO_CAO_SKILL_PATH):
        with open(XUAT_BAO_CAO_SKILL_PATH, "r", encoding="utf-8") as f:
            content = f.read()
        # Take first 2500 chars
        excerpt = content[:2500]
    return excerpt


def load_hpg_template_excerpt() -> str:
    """Load the first ~300 lines of build_hpg_model.py as code template reference."""
    if not os.path.exists(HPG_TEMPLATE_PATH):
        return ""
    with open(HPG_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        lines = f.readlines()
    # Take first 250 lines as structural reference
    return "".join(lines[:250])


def build_prompt(ticker: str, meta: dict, is_bank: bool, is_securities: bool) -> str:
    instructions_excerpt = load_instructions_excerpt()
    export_skill = load_export_skill_excerpt()
    hpg_template = load_hpg_template_excerpt()

    sector = meta.get("sector", "General")
    company = meta.get("companyName", ticker)
    price = meta.get("currentPrice", 10000)
    shares = meta.get("shares", 1_000_000_000)
    market_cap = meta.get("marketCap", 10_000_000_000_000)
    years = meta.get("years_available", [2021, 2022, 2023, 2024, 2025])
    fs = meta.get("financial_summary", {})

    # Format data for prompt
    fin_table = f"""
Historical Financial Data for {ticker} (tỷ VND unless noted):

Years available: {years}
Revenue:      {fs.get('revenue', [])}
Gross Profit: {fs.get('gross_profit', [])}
NPAT:         {fs.get('npat', [])}
EPS (VND):    {fs.get('eps', [])}
Total Assets: {fs.get('total_assets', [])}
Equity (VCSH):{fs.get('equity', [])}
Cash:         {fs.get('cash', [])}
Short Debt:   {fs.get('short_debt', [])}
Long Debt:    {fs.get('long_debt', [])}
"""
    if is_bank:
        fin_table += f"""
--- BANKING SPECIFIC ---
NII (Net Interest Income): {fs.get('nii', [])}
PBT:         {fs.get('pbt', [])}
Gross Loans: {fs.get('loans', [])}
Deposits:    {fs.get('deposits', [])}
"""

    special_notes = ""
    if is_bank:
        special_notes = """
IMPORTANT - THIS IS A BANK:
- Use banking-specific P&L model: NII + Non-Interest Income → TOI → PPOP → Provision → LNTT → LNST
- Key assumptions: NIM (%), Credit Growth (%), CIR (%), Credit Cost (CoC %), CASA %, NPL %
- Use Residual Income (RI) model for valuation (P/B as secondary)
- Calculate banking ratios: NIM, LDR, CASA, CIR, NPL, LLR, CoC, ROE, ROA, CAR
- Excel sheets MUST include: 01_Cover, 02_Assumptions, 03_Income_Model, 04_PnL, 05_Balance_Sheet, 06_Ratios, 07_Valuation, 08_Sensitivity, 09_PESTLE, 10_Leading_Indicators, 11_Investment_Thesis, 12_Summary_Snapshot, 13_PE_PB_History
"""
    elif is_securities:
        special_notes = """
IMPORTANT - THIS IS A SECURITIES COMPANY:
- Revenue drivers: Brokerage (market volume × market share × fee rate), Margin Lending, Proprietary Trading, IB
- Key metrics: Market share (%), KLGD thị trường, VN-Index correlation
- Valuation: P/E + P/B (market-linked)
"""
    else:
        special_notes = f"""
IMPORTANT - THIS IS A {sector.upper()} COMPANY:
- Use appropriate revenue drivers for sector: {sector}
- Use EV/EBITDA + P/E + DCF for valuation
- Excel sheets: 01_Cover, 02_Assumptions, 03_Revenue_Model, 04_PnL, 05_Balance_Sheet, 06_Cash_Flow, 07_Valuation, 08_Sensitivity, 09_PESTLE, 10_Leading_Indicators, 11_Investment_Thesis, 12_Summary_Snapshot
"""

    prompt = f"""You are an expert Vietnamese financial analyst and Python developer. Your task is to write a complete, working Python script called `build_{ticker.lower()}_model.py` that generates a professional Excel model and PDF analysis report for {ticker} ({company}).

=== CONTEXT ===
Stock: {ticker}
Company: {company}
Sector: {sector}
Current Price: {price:,} VND
Market Cap: {market_cap/1e9:,.0f} tỷ VND
Shares Outstanding: {shares:,}
Analysis Month: {datetime.now().strftime('%Y-%m-%d')}
{special_notes}

=== FINANCIAL DATA TO USE ===
{fin_table}

=== INSTRUCTIONS & RULES (from system documentation) ===
{instructions_excerpt[:1500]}

=== EXPORT SKILL RULES ===
{export_skill[:2000]}

=== TEMPLATE STRUCTURE (from build_hpg_model.py) ===
```python
{hpg_template[:2500]}
```

=== YOUR TASK ===
Write a complete Python script `build_{ticker.lower()}_model.py` that:

1. **Imports & Configuration**:
   - Import: os, sys, json, math, datetime, openpyxl (with styles, charts), reportlab (SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak), matplotlib (Agg backend), requests, statistics
   - Set OUT_DIR = `os.path.join(os.path.dirname(__file__), "Bao cao", "{ticker}")`
   - Set TICKER = "{ticker}", COMPANY = "{company}", PRICE = {price}, SHARES = {shares}
   - MONTH = datetime.now().strftime("%Y-%m-%d")

2. **Data Fetching** (MANDATORY):
   - Call `fetch_all("{ticker}")` from fetch_data.py (already exists in project)
   - Use Vietcap API details endpoint: `https://trading.vietcap.com.vn/api/iq-insight-service/v1/company/details?ticker={ticker}`
   - Hardcode the historical data arrays from the financial data above as fallback
   - Define years_hist = {years}
   - Extract all standard financial metrics by year

3. **Forecast Model** (2 years forward minimum):
   - Build sector-appropriate revenue forecast with named drivers
   - Forecast P&L: Revenue → Gross Profit → EBIT → LNST → EPS
   - Forecast Balance Sheet key items (equity growth, net debt)
   - All forecasts must use realistic assumptions based on historical trends

4. **Excel Model** (openpyxl):
   - Create workbook with properly named sheets per sector rules
   - Cover sheet with key metrics, thesis, and the analysis date and time in the format "Ngày lập: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
   - Assumptions sheet (all parameters in one place)
   - Revenue/Income Model sheet with driver-based calculations
   - P&L sheet (5Y historical + 2-3Y forecast)
   - Balance Sheet sheet
   - Valuation sheet (Bear/Base/Bull with 3 methods)
   - Sensitivity matrix sheet (5×5)
   - PESTLE sheet (6 factors)
   - Leading Indicators sheet
   - Apply professional styling: blue headers, alternating row colors, number formats (#,##0 for billions, % for margins)
   - Save to: `os.path.join(OUT_DIR, f"{{TICKER}}_Model_{{MONTH}}.xlsx")`

5. **PDF Report** (reportlab):
   - Create matplotlib charts first (save as PNG), then embed in PDF
   - Chart 1: Bar chart of Revenue & NPAT by year (historical + forecast)
   - Chart 2: Line chart of margin trends (GP%, NPAT%)
   - Chart 3: Valuation sensitivity heatmap or bar chart
   - PDF sections: Cover (must display the analysis date and time as "Ngày lập: %d/%m/%Y %H:%M") → Investment Summary → Financial Analysis → Valuation → Risks → Conclusion
   - Investment Summary must include: Rating (BUY/HOLD/SELL), Target Price, Upside %, 3 bull points, 3 risk points, snapshot table
   - Write all analysis in Vietnamese, keep financial terms in English (EPS, P/E, NIM, EBITDA, etc.)
   - Save to: `os.path.join(OUT_DIR, f"{{TICKER}}_Phan_Tich_{{MONTH}}.pdf")`

6. **JSON Summary Export**:
   - Save `data/{ticker}.json` with: ticker, companyName, sector, currentPrice, marketCap, shares, gdrivePdfUrl=None, gdriveExcelUrl=None
   - data dict: years (hist+fc), revenue (hist+fc), npat (hist+fc), eps (hist+fc), equity (hist+fc)
   - Add rich evaluation data for dashboard:
     - `thesis`: list of 3 detailed investment thesis paragraphs (Vietnamese).
     - `risks`: list of 3 detailed investment risk points (Vietnamese).
     - `moats`: dict with competitive advantage scorecards (Network Effect, Cost Advantage, Switching Cost, Intangible Assets, Efficient Scale) from 1 to 5, and a brief description for each.
     - `pestle`: list of 6 dicts, each with keys: `factor` (Political, Economic, etc.), `content` (description), `impact` (Positive, Neutral, Negative).
     - `valuation`: dict with scenario target prices: `bear`, `base`, `bull`, plus corresponding target multiples.
     - `comments`: dict containing longer written text for: `businessModel` (Vietnamese, ~100 words), `financialPerformance` (~100 words), `valuationText` (~100 words).
     - `pe_hist`: list of historical P/E multiples (for the historical years).
     - `pb_hist`: list of historical P/B multiples (for the historical years).
     - `pe_quarters`: list of all quarterly P/E values from statistics-financial.
     - `pb_quarters`: list of all quarterly P/B values from statistics-financial.
     - `quarter_labels`: list of strings for quarter names corresponding to pe_quarters/pb_quarters (e.g. ["2021-Q1", "2021-Q2", ...]).
      - `income_quarterly`: list of quarterly records containing ``yearReport, quarter, nii, npat`` (nii represents Net Interest Income for banks, or net revenue for non-banks, both in billion VND). Use section_to_quarters.
     - `ratios_quarterly`: dict containing:
        - `quarters`: list of quarter labels (e.g., ["2024-Q1", ...])
        - `nim`: quarterly NIM for banks or Gross Margin for non-banks
        - `ldr`: quarterly LDR for banks or Debt/Equity for non-banks
        - `casa`: quarterly CASA for banks or ROA for non-banks
        - `npl`: quarterly NPL for banks or Asset Turnover for non-banks
     - **`ratios`**: A dictionary containing calculated historical and forecast ratio arrays. **FOR EVERY RATIO, YOU MUST CALCULATE IT YOURSELF FROM RAW BCTC DATA FOR ALL YEARS (both historical and forecast) INSTEAD OF TRUSTING VIETCAP PRE-CALCULATED RATIOS**.
       - For Banks:
         - `nim`: Net Interest Income / Average Earning Assets
         - `roe`: NPAT / Average Equity
         - `roa`: NPAT / Average Assets
         - `npl`: NPL / Gross Loans
         - `ldr`: Gross Loans / (Deposits + PaperValuable/Giấy tờ có giá). **Formula: LDR = Loans / (Deposits + Giấy tờ có giá) * 100**. For TCB, Deposits (~470k tỷ) + Giấy tờ có giá (~80k tỷ CD/Trái phiếu) = ~550k tỷ. Loans = ~480k tỷ. LDR must calculate to a realistic ~80-87%, NOT 132.9%.
         - `casa`: CASA / Deposits
       - For Non-Banks:
         - `gross_margin`: Gross Profit / Revenue
         - `roe`: NPAT / Average Equity
         - `roa`: NPAT / Average Assets
         - `debt_to_equity`: Net Debt / Equity



7. **Main execution**:
   - Create output directories
   - Build Excel, save
   - Build matplotlib charts, save PNG
   - Build PDF with embedded charts, save
   - Print success messages with file paths
   - If `__name__ == "__main__"`, call main()

=== CRITICAL REQUIREMENTS ===
- The script must be 100% self-contained and executable as: `python build_{ticker.lower()}_model.py`
- Use only libraries already in requirements.txt (openpyxl, reportlab, matplotlib, numpy, requests)
- All strings must be safe ASCII or UTF-8 encoded for reportlab (use Helvetica font family, NOT Unicode-only fonts)
- Handle missing API data gracefully with the hardcoded fallback values
- Output ONLY the complete Python code, no explanations, no markdown code fences
- The script must not import google_drive_uploader (the orchestrator handles uploads)
"""
    return prompt


def generate_model_builder(ticker: str) -> str | None:
    """Main function: generates and saves build_<ticker>_model.py via Gemini API."""
    ticker = ticker.upper()
    output_path = os.path.join(PROJECT_ROOT, f"build_{ticker.lower()}_model.py")

    print(f"\n{'='*60}")
    print(f"  AI CODE GENERATOR — {ticker}")
    print(f"{'='*60}")

    # 1. Check if already exists
    if os.path.exists(output_path):
        print(f"[Skip] {os.path.basename(output_path)} already exists. Delete it to regenerate.")
        return output_path

    # 2. Get Gemini API Key
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("[ERROR] GEMINI_API_KEY environment variable not set.")
        print("  Set it locally: set GEMINI_API_KEY=your_key_here")
        print("  On GitHub: add GEMINI_API_KEY as a Repository Secret")
        return None

    # 3. Fetch financial metadata
    print(f"[Step 1/3] Fetching financial data for {ticker}...")
    meta = fetch_vietcap_metadata(ticker)

    is_bank = ticker in BANKING_TICKERS
    is_securities = ticker in SECURITIES_TICKERS

    if is_bank:
        print(f"  → Classified as BANK. Will use banking model template.")
    elif is_securities:
        print(f"  → Classified as SECURITIES. Will use securities model template.")
    else:
        sector = INDUSTRY_MAP.get(ticker, "General")
        print(f"  → Classified as: {sector}")

    # 4. Build the prompt
    print("[Step 2/3] Building prompt for Gemini...")
    prompt = build_prompt(ticker, meta, is_bank, is_securities)

    # 5. Call Gemini API (using new google-genai SDK)
    print(f"[Step 3/3] Calling Gemini API to generate build_{ticker.lower()}_model.py ...")
    try:
        from google import genai
        from google.genai import types as genai_types

        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config=genai_types.GenerateContentConfig(
                temperature=0.2,
                max_output_tokens=16384,
            ),
        )
        generated_code = response.text

    except Exception as e:
        print(f"[ERROR] Gemini API call failed: {e}")
        return None


    # 6. Clean up the generated code
    # Remove any markdown code fences if present
    code = generated_code.strip()
    if code.startswith("```python"):
        code = code[len("```python"):].strip()
    elif code.startswith("```"):
        code = code[3:].strip()
    if code.endswith("```"):
        code = code[:-3].strip()

    # 7. Validate — must at least import openpyxl
    if "openpyxl" not in code or "reportlab" not in code:
        print("[WARN] Generated code may be incomplete — missing key imports.")
        print("  Saving anyway for manual review...")

    # 8. Prepend a header comment
    header = f"""#!/usr/bin/env python3
\"\"\"
{ticker} — Excel Model + PDF Report Generator
Auto-generated by generate_stock_model_builder.py on {datetime.now().strftime('%Y-%m-%d %H:%M')}
Company: {meta.get('companyName', ticker)}
Sector: {meta.get('sector', 'Unknown')}
Price at generation: {meta.get('currentPrice', 0):,} VND
\"\"\"
# ⚠ This file was auto-generated by AI. Review assumptions before using in production.

"""
    final_code = header + code

    # 9. Save
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(final_code)

    print(f"\n[OK] Generated successfully: {output_path}")
    print(f"     Lines: {len(final_code.splitlines())}")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_stock_model_builder.py <TICKER>")
        print("Example: python generate_stock_model_builder.py TCB")
        sys.exit(1)
    result = generate_model_builder(sys.argv[1])
    if result:
        print(f"\nNext step: python run_analysis.py {sys.argv[1].upper()}")
    else:
        sys.exit(1)
