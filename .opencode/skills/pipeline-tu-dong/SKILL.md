---
name: pipeline-tu-dong
description: Dùng skill này khi cần chạy phân tích cổ phiếu tự động bằng Python/GitHub Actions. Front-loaded keywords: chạy tự động, run analysis, python pipeline, GitHub Actions, generate model, build model, tạo báo cáo tự động, run_analysis, generate_stock_model_builder, build_generic_model.
---

# SKILL PIPELINE TỰ ĐỘNG — PHÂN TÍCH CỔ PHIẾU

> **Mục đích**: Mô tả toàn bộ quy trình tự động để tạo báo cáo cổ phiếu (Excel + PDF + JSON) từ lệnh `python run_analysis.py <TICKER>` hoặc GitHub Actions workflow.

---

## Kiến trúc tổng thể

```
Nhận TICKER (từ CLI hoặc GitHub Actions)
        │
        ▼
┌─────────────────────────────────────────────────────┐
│ run_analysis.py — ORCHESTRATOR (5 bước)             │
│                                                     │
│ Bước 1: Kiểm tra build_<ticker>_model.py            │
│   → Nếu KHÔNG có: gọi generate_stock_model_builder │
│   → Nếu CÓ: dùng luôn (hoặc force_regenerate=true) │
│                                                     │
│ Bước 2: Chạy build_<ticker>_model.py                │
│   → Tạo Bao cao/<TICKER>/<TICKER>_Model_<MM>.xlsx   │
│   → Tạo Bao cao/<TICKER>/<TICKER>_Phan_Tich_<MM>.pdf│
│   → Tạo data/<TICKER>.json                         │
│                                                     │
│ Bước 3: Tìm file output                             │
│   → Scan Bao cao/<TICKER>/ để lấy path xlsx + pdf  │
│                                                     │
│ Bước 4: Upload Google Drive                         │
│   → google_drive_uploader.upload_file()             │
│   → Trả về gdriveExcelUrl + gdrivePdfUrl            │
│                                                     │
│ Bước 5: Cập nhật registry                          │
│   → Patch URLs vào data/<TICKER>.json               │
│   → Update data/index.json (dashboard list)         │
└─────────────────────────────────────────────────────┘
```

---

## Chi tiết từng script

### 1. `generate_stock_model_builder.py` — AI Code Generator

**Khi nào dùng**: Khi chưa có `build_<ticker>_model.py` cho ticker mới.

**Quy trình**:
1. Fetch metadata từ Vietcap API (company details, BCTC, statistics-financial)
2. Phân loại ngành: `BANKING_TICKERS` / `SECURITIES_TICKERS` / `INDUSTRY_MAP`
3. Build prompt cho Gemini API (bao gồm dữ liệu thực + rules từ skills)
4. Gọi `gemini-2.5-flash` với `max_output_tokens=16384`
5. Làm sạch output (bỏ markdown fences)
6. Lưu thành `build_<ticker>_model.py`

**Yêu cầu env**: `GEMINI_API_KEY`

**Phân loại ngành tự động**:
```python
BANKING_TICKERS = {"VCB", "BID", "TCB", "MBB", "VPB", "ACB", "HDB", "TPB", "STB",
                   "LPB", "ABB", "SHB", "VAB", "VIB", "BAB", "KLB", "NAB", "NVB",
                   "SGB", "OCB", "EIB", "MSB"}

SECURITIES_TICKERS = {"SSI", "VND", "HCM", "VCI", "FTS", "SHS"}

INDUSTRY_MAP = {
    "HPG": "Thép", "HSG": "Thép", "NKG": "Thép",
    "MWG": "Bán lẻ", "FRT": "Bán lẻ", "PNJ": "Bán lẻ",
    "FPT": "Công nghệ", "CMG": "Công nghệ",
    "VHM": "Bất động sản dân cư", "NLG": "Bất động sản dân cư",
    "IDC": "Bất động sản KCN", "KBC": "Bất động sản KCN",
    # ... (xem file đầy đủ)
}
```

**Prompt template** gửi Gemini bao gồm:
- Dữ liệu tài chính thực tế từ API (revenue, NPAT, NII, assets, equity...)
- Excerpt từ `INSTRUCTIONS_PHAN_TICH_CO_PHIEU.md`
- Excerpt từ skill `xuat-bao-cao/SKILL.md`
- 250 dòng đầu của `build_hpg_model.py` làm template code
- Yêu cầu output: complete Python script `build_<ticker>_model.py`

---

### 2. `build_<ticker>_model.py` — Specialized Builder (AI-generated)

**Yêu cầu bắt buộc trong script được generate**:

```python
# Cấu trúc bắt buộc của mọi builder script:

import os, sys, json, math, datetime, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import requests

# Constants
TICKER = "<TICKER>"
COMPANY = "<Company Name>"
PRICE = <current_price>
SHARES = <shares_outstanding>
OUT_DIR = os.path.join(os.path.dirname(__file__), "Bao cao", TICKER)
MONTH = datetime.datetime.now().strftime("%Y-%m")

def fetch_all(ticker):
    """Fetch từ Vietcap API, fallback về hardcoded historical data."""
    pass

def build_excel():
    """Tạo Excel model theo đúng cấu trúc sheets."""
    pass

def build_charts():
    """Vẽ matplotlib charts, lưu PNG tạm."""
    pass

def build_pdf():
    """Tạo PDF report, nhúng chart PNG."""
    pass

def save_json():
    """Lưu data/<TICKER>.json với đầy đủ schema dashboard."""
    pass

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    build_excel()
    build_charts()
    build_pdf()
    save_json()

if __name__ == "__main__":
    main()
```

---

### 3. `build_generic_model.py` — Fallback khi không có specialized builder

**Khi nào dùng**: Ticker chưa có specialized builder VÀ không thể generate (không có GEMINI_API_KEY).

**Đặc điểm**:
- Dùng `fetch_data.fetch_all(ticker, use_cache=True)` để lấy BCTC thực từ API
- Có sector content cho 5 nhóm ngành: Banking, Real Estate, Technology, Steel/Materials, Consumer/Retail
- Tạo forecast đơn giản (avg growth + avg margin từ lịch sử)
- Valuation: target P/E = 10x, target P/B = 1.2x (mặc định)
- Output đầy đủ: Excel + PDF + JSON

---

### 4. `fetch_data.py` — Data Fetching Layer

**Các hàm chính**:
```python
def fetch_all(ticker, use_cache=True) -> dict:
    """Lấy toàn bộ BCTC từ Vietcap API, có cache tại .cache/<ticker>.json"""

def section_to_years(data, section) -> list[dict]:
    """Trích xuất list records theo section: INCOME_STATEMENT, BALANCE_SHEET, CASH_FLOW"""

def get_field_map(data, section) -> dict:
    """Map field names (isa3, bsa53...) sang tên tiếng Việt"""
```

**API endpoints dùng**:
```
Base: https://trading.vietcap.com.vn/api/iq-insight-service/v1

GET /company/details?ticker={ticker}          → company info, giá, vốn hóa
GET /company/{ticker}/financial-statement?section=INCOME_STATEMENT  → KQKD
GET /company/{ticker}/financial-statement?section=BALANCE_SHEET     → CĐKT
GET /company/{ticker}/financial-statement?section=CASH_FLOW         → LCTT
GET /company/{ticker}/statistics-financial                           → PE/PB/EV lịch sử theo quý
```

**Handshake trước khi gọi API**:
```python
session.get("https://trading.vietcap.com.vn/", timeout=8)  # warm up
# Sau đó mới gọi các endpoints khác
```

---

### 5. `google_drive_uploader.py` — Upload Files

**Yêu cầu env**: `GDRIVE_SERVICE_ACCOUNT_JSON` (JSON string của service account)

**Hàm chính**:
```python
def upload_file(file_path, folder_id=None) -> tuple[str, str]:
    """Upload file lên Google Drive. Return (file_id, shareable_url)"""
```

**Folder mặc định**: ID được hardcode trong file, tương ứng với thư mục "AIC FA Reports" trên Drive.

---

### 6. `data/index.json` — Dashboard Registry

Format:
```json
[
  {
    "ticker": "TCB",
    "companyName": "Ngân hàng TMCP Kỹ thương Việt Nam",
    "sector": "Ngân hàng",
    "lastUpdated": "2026-06-27 10:30",
    "excelUrl": "https://drive.google.com/...",
    "pdfUrl": "https://drive.google.com/..."
  },
  ...
]
```

---

## GitHub Actions Workflow

File: `.github/workflows/analyze_stock.yml`

**Trigger**: Manual dispatch (`workflow_dispatch`) với inputs:
- `ticker` (required): mã cổ phiếu
- `force_regenerate` (optional, default=false): xóa builder cũ để regenerate

**Secrets cần cấu hình trên GitHub**:
- `GEMINI_API_KEY` — Gemini API key
- `GDRIVE_SERVICE_ACCOUNT_JSON` — Google Drive service account JSON

**Quy trình workflow**:
```yaml
1. Checkout repo (fetch-depth: 1)
2. Setup Python 3.11
3. pip install -r requirements.txt
4. [Optional] Delete existing builder if force_regenerate=true
5. python run_analysis.py <ticker>
6. git add data/ build_<ticker>_model.py
7. git commit -m "Auto-analysis: <ticker> — <date>"
8. git push
```

**Sau khi push**: dashboard tự cập nhật vì `data/index.json` và `data/<TICKER>.json` đã được commit vào repo.

---

## Checklist vận hành pipeline

### Trước khi chạy lần đầu:
```
☐ GEMINI_API_KEY đã set (GitHub Secret hoặc env local)
☐ GDRIVE_SERVICE_ACCOUNT_JSON đã set
☐ Google Drive folder đã chia sẻ với service account email
☐ requirements.txt đủ: requests, openpyxl, reportlab, matplotlib, numpy, google-api-python-client, google-auth-httplib2, google-auth-oauthlib, google-genai
```

### Sau mỗi lần chạy:
```
☐ Bao cao/<TICKER>/<TICKER>_Model_<YYYY-MM>.xlsx đã tạo?
☐ Bao cao/<TICKER>/<TICKER>_Phan_Tich_<YYYY-MM>.pdf đã tạo?
☐ data/<TICKER>.json đã tạo và có đủ fields?
☐ data/index.json đã update với ticker mới?
☐ Files đã upload lên Google Drive (gdriveExcelUrl/gdrivePdfUrl != null)?
☐ GitHub Actions commit đã push thành công?
☐ Dashboard (index.html) hiển thị ticker mới?
```

### Xử lý lỗi phổ biến:

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| `GEMINI_API_KEY not set` | Chưa set secret | Thêm vào GitHub Secrets |
| `Could not fetch company details` | Vietcap API timeout | Retry, hoặc kiểm tra network |
| `Builder script failed` | Code AI generate bị lỗi | Bật `force_regenerate=true` để generate lại |
| `No output files found` | Builder chạy nhưng không save | Xem log builder script |
| `Google Drive upload failed` | Service account không có quyền | Check folder sharing |
| `data/index.json not updated` | JSON parse error | Xem log step 5 |

---

## Quy tắc bắt buộc khi AI generate builder script

> Những quy tắc này phải được đảm bảo trong prompt gửi Gemini:

1. **Self-contained**: Script phải chạy được hoàn toàn độc lập: `python build_<ticker>_model.py`
2. **Fallback hardcode**: Nếu API fail → dùng hardcoded historical data (đã nhúng vào prompt)
3. **Không import google_drive_uploader**: Upload do `run_analysis.py` xử lý
4. **Font reportlab**: Dùng Helvetica (không dùng Unicode-only fonts)
5. **Output path chuẩn**:
   - Excel: `Bao cao/<TICKER>/<TICKER>_Model_<YYYY-MM>.xlsx`
   - PDF: `Bao cao/<TICKER>/<TICKER>_Phan_Tich_<YYYY-MM>.pdf`
   - JSON: `data/<TICKER>.json`
6. **JSON phải có đủ fields**: Xem schema tại skill `xuat-bao-cao` phần "JSON Dashboard Export"
7. **Banking special**: Nếu `is_bank=True` → dùng banking P&L model (NII→TOI→PPOP→LNTT→LNST)
8. **Libraries**: Chỉ dùng thư viện có trong `requirements.txt`

---

## Pipeline Ngân Hàng — template_banking.py (QUAN TRỌNG)

Đây là luồng RIÊNG BIỆT dành cho cổ phiếu ngân hàng. Khi `run_analysis.py` nhận ra ticker là ngân hàng (kiểm tra `is_bank`), nó **KHÔNG** gọi `build_{ticker}_model.py` hay `build_generic_model.py`, mà gọi thẳng `template_banking.py`.

### Luồng Banking Pipeline

```
python run_analysis.py VIB
         │
         ├─ Bước 1: fetch_data.fetch_all("VIB") → .cache/VIB_bctc.json
         │
         ├─ Bước 2: is_bank = True (VIB ∈ BANKING_TICKERS)
         │     ├─ update_peer_benchmark.main()     ← CẬP NHẬT 18 NH SONG SONG
         │     └─ template_banking.run_banking_analysis("VIB", raw_data)
         │           ├─ Tính chỉ số (NIM, NPL, CASA, ROE, LLR, CIR...)
         │           ├─ Dự phóng 3 năm (NII, PPOP, LNST)
         │           ├─ build_excel_banking()       → VIB_Model_2026-06.xlsx
         │           │     └─ Sheet 15_Peer_Benchmark (18 NH từ peer_benchmark.json)
         │           ├─ generate_charts_banking()  → 13 chart PNG
         │           ├─ build_pdf_banking()         → VIB_Phan_Tich_2026-06.pdf
         │           │     ├─ Trang 4: Peer Benchmark Table
         │           │     └─ Footnote: "P/B cập nhật ngày: YYYY-MM-DD"
         │           └─ save_json_summary()         → data/VIB.json
         │
         ├─ Bước 3-5: Upload GDrive + Update registry (giống non-bank)
```

### Sections Vietcap API dùng trong Banking

| Section key | Nội dung | API endpoint |
|---|---|---|
| `BALANCE_SHEET` | Bảng cân đối kế toán | `financial-statement?section=BALANCE_SHEET` |
| `INCOME_STATEMENT` | Kết quả kinh doanh | `financial-statement?section=INCOME_STATEMENT` |
| `CASH_FLOW` | Lưu chuyển tiền tệ | `financial-statement?section=CASH_FLOW` |
| `NOTE` | Thuyết minh BCTC | `financial-statement?section=NOTE` |

> ⚠️ **BẮT BUỘC dùng key `"NOTE"`** (không phải `"NOTES"`). Đây là key chuẩn trong `SECTIONS` của `fetch_data.py`.

### Bộ 13 Chart Banking

| Chart | Nội dung |
|---|---|
| A | NIM Decomposition (YOEA / COF / NIM trend) |
| B | PE/PB/EVEBITDA lịch sử theo quý |
| C | Tăng trưởng tín dụng vs huy động theo quý |
| D | CASA ratio xu hướng |
| E | NPL xu hướng (annual) |
| F | PPOP / LNST / Dự phòng so sánh |
| G | Residual Income Sensitivity (COE × g 5×5) |
| H | Cơ cấu tài sản sinh lãi theo năm (Stacked Area) |
| I | NII vs NonII theo quý (Stacked Bar) |
| J | NPL / LLR theo quý (Dual Axis) |
| K | Tăng trưởng tín dụng QoQ |
| L | PE/PB lịch sử rộng (12×5) |
| M | SWOT radar |

---

## update_peer_benchmark.py — Chi tiết triển khai

### Mục đích
Cào dữ liệu 18 ngân hàng niêm yết đồng thời, tính toán các chỉ số tài chính và lưu vào `data/peer_benchmark.json`. **Bắt buộc chạy mỗi lần phân tích mã ngân hàng** để đảm bảo P/B và các chỉ số phản ánh đúng ngày phân tích.

### Danh sách 18 NH
```python
BANK_TICKERS = ['VCB', 'BID', 'CTG', 'TCB', 'MBB', 'ACB', 'VIB', 'HDB',
                'STB', 'VPB', 'TPB', 'MSB', 'EIB', 'LPB', 'SHB', 'OCB', 'BAB', 'NAB']
```

### API Endpoints chuẩn (KHÔNG dùng api.vietcap.vn — bị block)

```python
BASE_IQ = "https://iq.vietcap.com.vn/api/iq-insight-service/v1"

# Giá + Số CP lưu hành
url_det = f"{BASE_IQ}/company/details?ticker={ticker}"
# → det.get("currentPrice")            # VND thực tế, vd: 61400.0
# → det.get("numberOfSharesMktCap")    # Số CP lưu hành

# Balance Sheet quý
url_bs = f"{BASE_IQ}/company/{ticker}/financial-statement?section=BALANCE_SHEET&quarterly=true"
# → data["quarters"][-1].get("bsa78")  # VCSH quý gần nhất
# → data["quarters"][-1].get("bsb103") # Cho vay KH

# Income Statement quý
url_is = f"{BASE_IQ}/company/{ticker}/financial-statement?section=INCOME_STATEMENT&quarterly=true"
# → data["quarters"][-1].get("isa20")  # LNST
# → data["quarters"][-1].get("isb27")  # NII
# → data["quarters"][-1].get("isb38")  # TOI
# → data["quarters"][-1].get("isb39")  # OPEX (âm)

# Note (thuyết minh) quý
url_nt = f"{BASE_IQ}/company/{ticker}/financial-statement?section=NOTE&quarterly=true"
# → data["quarters"][-1].get("nob42")  # Nợ nhóm 3
# → data["quarters"][-1].get("nob43")  # Nợ nhóm 4
# → data["quarters"][-1].get("nob44")  # Nợ nhóm 5
# → data["quarters"][-1].get("nob65")  # Tổng tiền gửi KH
# → data["quarters"][-1].get("nob66")  # Tiền gửi không kỳ hạn (CASA)
```

### Tính P/B (CẤM dùng API định giá — hay bị 403)

```python
# Vốn hóa (tỷ VND)
price = det.get("currentPrice")           # VND (61400.0 là đúng, không phải 61.4)
shares = det.get("numberOfSharesMktCap")  # Số CP
mcap = (price * shares) / 1e9            # Tỷ VND

# VCSH quý gần nhất
equity = q_bs.get("bsa78", 1)            # VND
equity_billion = equity / 1e9            # Tỷ VND

# P/B = Vốn hóa / VCSH
pb = round(mcap / max(equity_billion, 0.001), 2)
```

### Cấu trúc `data/peer_benchmark.json`

```json
{
  "_meta": {
    "updated": "2026-06-29",
    "source": "Vietcap Live API Engine"
  },
  "peers": [
    {
      "ticker": "VCB",
      "name": "Vietcombank",
      "npl": 0.63,
      "nim": 4.09,
      "casa": 32.55,
      "roe": 16.17,
      "cir": 32.5,
      "pb": 2.19,
      "cg": 4.78,
      "mcap": 513038,
      "date_updated": "2026-06-29"
    }
  ]
}
```

### Tích hợp trong run_analysis.py

```python
if is_bank:
    try:
        import update_peer_benchmark
        update_peer_benchmark.main()  # Cào 18 NH song song — ~5-8s
    except Exception as e:
        print(f"[WARN] Peer benchmark update failed: {e}")
        # Tiếp tục phân tích với dữ liệu peer_benchmark.json cũ
```

### Ngày phân tích trong PDF

```python
# Trong template_banking.py — cuối bảng Peer Benchmark trang 4 PDF
today_str = datetime.datetime.now().strftime("%Y-%m-%d")
story.append(Paragraph(
    f"<i>Nguồn: BCTC Q1/2026 và Vietcap. "
    f"P/B được cập nhật theo thị giá đóng cửa ngày: {today_str}.</i>",
    body_style
))
```

### Lỗi phổ biến

| Lỗi | Nguyên nhân | Giải pháp |
|---|---|---|
| `KeyError: 'NOTES'` | Dùng key sai trong template_banking.py | Đổi thành `"NOTE"` (xem SECTIONS trong fetch_data.py) |
| `mcap: 6.14e-05` | Dùng key `outstandingShares` không tồn tại | Dùng `numberOfSharesMktCap` |
| P/B = 1.00x (tất cả NH) | API valuation-history bị 403 | Tự tính P/B = mcap / (bsa78/1e9) |
| `ThreadPoolExecutor not defined` | Import ở giữa file | Chuyển `from concurrent.futures import ThreadPoolExecutor` lên đầu file |
| `AttributeError: module datetime has no attribute 'now'` | Import `import datetime` nhưng dùng `datetime.now()` | Sửa thành `datetime.datetime.now()` |
