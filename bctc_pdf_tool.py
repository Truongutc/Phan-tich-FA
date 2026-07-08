#!/usr/bin/env python3
"""
bctc_pdf_tool.py — Tìm, tải, và render PDF báo cáo tài chính (BCTC) hợp nhất
cho nhóm BĐS KCN, phục vụ trích xuất doanh thu/giá vốn theo mảng.

Nguồn chính (đã verify hoạt động qua curl thường, không cần cookie/JS):
    GET https://cafef.vn/du-lieu/Ajax/PageNew/FileBCTC.ashx?Symbol=<mã thường>&Type=1&Year=0
    -> JSON {"Data": [{"Quarter":1-4|5(năm), "Year":YYYY, "Name": "...", "Link": "...pdf"}]}
    Quarter=5 nghĩa là báo cáo năm (kiểm toán). Name chứa "hợp nhất" hoặc "công ty mẹ".

Dự phòng: 24hmoney.vn/stock/<mã>/financial-report (HTML có sẵn link PDF CDN,
không phân nhãn hợp nhất/riêng rõ ràng — phải tự phân biệt bằng số trang).

Usage:
    python bctc_pdf_tool.py list <TICKER>
    python bctc_pdf_tool.py plan-downloads <TICKER> [--years 5] [--quarters 4]
    python bctc_pdf_tool.py download <url> --out <path.pdf>
    python bctc_pdf_tool.py render <pdf> --pages 10-15 --out <dir>
"""
import os
import sys

# Fix Windows console encoding (cp1252 không hỗ trợ tiếng Việt)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import json
import argparse
import unicodedata
import requests

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
STORE_DIR = os.path.join(PROJECT_ROOT, "data", "segments_kcn")
PDF_DIR = os.path.join(PROJECT_ROOT, "BCTC_PDF")

CAFEF_API = "https://cafef.vn/du-lieu/Ajax/PageNew/FileBCTC.ashx"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"}


def _strip_diacritics(s):
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def is_consolidated(name):
    n = _strip_diacritics(name or "").lower()
    return "hop nhat" in n


def fetch_cafef_list(ticker):
    r = requests.get(CAFEF_API, params={"Symbol": ticker.lower(), "Type": 1, "Year": 0},
                      headers=HEADERS, timeout=20)
    r.raise_for_status()
    data = r.json()
    if not data.get("Success", True) and not data.get("Data"):
        raise RuntimeError(f"CafeF API trả lỗi: {data}")
    return data.get("Data", [])


def select_best_reports(items):
    # Nhóm tài liệu theo (Year, Quarter) và chọn tài liệu tối ưu nhất
    grouped = {}
    for x in items:
        y = x.get("Year")
        q = x.get("Quarter")
        if y is None or q is None:
            continue
        key = (y, q)
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(x)
        
    consolidated = []
    for key, cands in grouped.items():
        valid_cands = [x for x in cands if "cong ty me" not in _strip_diacritics(x.get("Name", "")).lower() 
                       and "cty me" not in _strip_diacritics(x.get("Name", "")).lower()
                       and "c.ty me" not in _strip_diacritics(x.get("Name", "")).lower()]
        if not valid_cands:
            continue
            
        def get_priority(x):
            name_norm = _strip_diacritics(x.get("Name", "")).lower()
            p = 0
            if "hop nhat" in name_norm:
                p += 10
            if "kiem toan" in name_norm or "soat xet" in name_norm:
                p += 5
            return p
        best = max(valid_cands, key=get_priority)
        consolidated.append(best)
    return consolidated


def cmd_list(args):
    ticker = args.ticker.upper()
    items = fetch_cafef_list(ticker)
    consolidated = select_best_reports(items)
    print(f"[INFO] Tổng {len(items)} tài liệu, đã chọn {len(consolidated)} bản tối ưu cho {ticker}:")
    for x in sorted(consolidated, key=lambda x: (x.get("Year", 0), x.get("Quarter", 0)), reverse=True):
        label = f"Q{x['Quarter']}/{x['Year']}" if x.get("Quarter") in (1, 2, 3, 4, 6) else f"CN/{x['Year']}"
        print(f"  {label:10s} | {x['Name']} | {x['Link']}")


def _period_key(quarter, year):
    if quarter == 5:
        return str(year)  # năm kiểm toán
    return f"{year}Q{quarter}"


def cmd_plan_downloads(args):
    ticker = args.ticker.upper()
    items = fetch_cafef_list(ticker)
    consolidated = select_best_reports(items)

    store_file = os.path.join(STORE_DIR, f"{ticker}.json")
    store = {}
    have_yearly, have_quarterly = set(), set()
    if os.path.exists(store_file):
        with open(store_file, "r", encoding="utf-8") as f:
            store = json.load(f)
        
        # Chỉ coi là đã có nếu dữ liệu không chỉ chứa mảng "Khac" (hoặc nếu mảng Khac là hợp lệ duy nhất)
        # Đối với các công ty đã được thiết lập mảng chuẩn (như IDC, SIP, PHR, SZC),
        # nếu chỉ có mảng Khac thì coi như dữ liệu chưa đầy đủ.
        for pkey, seg_data in store.get("yearly", {}).items():
            non_khac = [s for s in seg_data if s != "Khac" and (seg_data[s].get("revenue", 0) > 0 or seg_data[s].get("cogs", 0) > 0)]
            if non_khac or ticker not in ("IDC", "SIP", "PHR", "SZC"):
                have_yearly.add(pkey)
                
        for pkey, seg_data in store.get("quarterly", {}).items():
            non_khac = [s for s in seg_data if s != "Khac" and (seg_data[s].get("revenue", 0) > 0 or seg_data[s].get("cogs", 0) > 0)]
            if non_khac or ticker not in ("IDC", "SIP", "PHR", "SZC"):
                have_quarterly.add(pkey)

    # Lấy năm cao nhất từ danh sách tài liệu CafeF
    current_year = max((x["Year"] for x in consolidated), default=2026)
    
    # 1. Yearly: Tải 5 năm gần nhất để phân tích lịch sử đầy đủ
    years_wanted = [current_year, current_year - 1, current_year - 2, current_year - 3, current_year - 4]
    
    # 2. Quarterly: Tải các quý của năm nay và năm ngoái
    quarters_wanted_years = [current_year, current_year - 1]

    to_fetch = []
    # Lập kế hoạch Yearly
    to_fetch_yearly = []
    for y in years_wanted:
        pkey = f"{y}(CN)"
        next_pkey = f"{y+1}(CN)"
        if next_pkey in have_yearly or next_pkey in [item[0] for item in to_fetch_yearly]:
            print(f"  [INFO] Bỏ qua lập lịch tải {pkey} vì dữ liệu năm này sẽ có trong thuyết minh của năm {y+1}")
            continue
            
        if pkey in have_yearly:
            continue
            
        cands = [x for x in consolidated if x.get("Quarter") == 5 and x.get("Year") == y]
        if cands:
            to_fetch_yearly.append((pkey, cands[0]))
        else:
            print(f"[MISS] Không tìm thấy BCTC năm kiểm toán {y} (hợp nhất) cho {ticker}")
            
    to_fetch.extend(to_fetch_yearly)

    # Lập kế hoạch Quarterly
    for y in quarters_wanted_years:
        for q in (1, 2, 3):
            pkey = f"{y}Q{q}"
            if pkey in have_quarterly:
                continue
            if q == 2:
                cands = [x for x in consolidated if x.get("Quarter") in (2, 6) and x.get("Year") == y]
            else:
                cands = [x for x in consolidated if x.get("Quarter") == q and x.get("Year") == y]
            if cands:
                to_fetch.append((pkey, cands[0]))
            else:
                print(f"[MISS] Không tìm thấy BCTC {pkey} (hợp nhất) cho {ticker}")

    print(f"\n[PLAN] Cần tải {len(to_fetch)} file PDF cho {ticker} (Yearly: {years_wanted}, Quarterly: {quarters_wanted_years}):")
    for pkey, item in to_fetch:
        print(f"  {pkey:10s} | {item['Name']} | {item['Link']}")
    return to_fetch


def slice_pdf_tail(input_path: str, output_path: str) -> int:
    """
    Cắt lấy 1/3 cuối PDF (bỏ 5 trang cuối cùng) và lưu thành file tạm.
    Trả về số trang của file sliced.
    Dùng pypdfium2 (đã cài sẵn qua pdfplumber dependency).
    """
    import pypdfium2 as pdfium

    doc = pdfium.PdfDocument(input_path)
    total = len(doc)

    start_page = int(total * 2 / 3)   # bắt đầu từ 2/3 chiều dài
    end_page   = max(start_page + 1, total - 1)  # chỉ bỏ đúng 1 trang chữ ký cuối cùng

    # Đảm bảo không slice ra tập rỗng hoặc không hợp lệ
    if start_page >= end_page or start_page < 0:
        start_page = max(0, total - 30)  # fallback: lấy tối đa 30 trang cuối
        end_page = total - 1

    new_doc = pdfium.PdfDocument.new()
    new_doc.import_pages(doc, list(range(start_page, end_page)))
    new_doc.save(output_path)

    sliced_count = end_page - start_page
    print(f"[SlicePDF] {total} trang -> cắt trang {start_page+1} đến {end_page} ({sliced_count} trang) -> {os.path.basename(output_path)}")
    return sliced_count


def cmd_download(args):
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    url = args.url.strip()
    # Chuẩn hóa khoảng trắng
    url = url.replace(" ", "%20")
    
    r = None
    try:
        r = requests.get(url, headers=HEADERS, timeout=60)
        r.raise_for_status()
    except Exception as e:
        print(f"[ERROR] Không thể tải URL {url}: {e}")
        # Thử fallback với URL rút gọn khoảng trắng kép
        cleaned_url = url.replace("%20%20", "%20").replace("  ", " ")
        if cleaned_url != url:
            try:
                print(f"  -> Thử tải lại với URL rút gọn: {cleaned_url}...")
                r = requests.get(cleaned_url, headers=HEADERS, timeout=60)
                r.raise_for_status()
            except Exception as e2:
                print(f"[ERROR] Tải lại với URL rút gọn thất bại: {e2}")
        
        if not r and not url.lower().endswith(".pdf"):
            pdf_url = url + ".pdf"
            try:
                print(f"  -> Thử tải lại với đuôi .pdf: {pdf_url}...")
                r = requests.get(pdf_url, headers=HEADERS, timeout=60)
                r.raise_for_status()
            except Exception as e3:
                print(f"[ERROR] Tải với đuôi .pdf thất bại: {e3}")
                return
        if not r:
            return

    if not r:
        return

    with open(args.out, "wb") as f:
        f.write(r.content)
    print(f"[OK] Đã tải {len(r.content)} bytes -> {args.out}")
    
    # Slice PDF lấy phần sau trước khi convert sang Markdown
    sliced_pdf = args.out.replace(".pdf", "_sliced.pdf")
    try:
        slice_pdf_tail(args.out, sliced_pdf)
        convert_target = sliced_pdf
    except Exception as e:
        print(f"[SlicePDF] [WARN] Không thể slice PDF: {e} - tiến hành convert toàn bộ file gốc")
        convert_target = args.out
        sliced_pdf = args.out

    # Tự động convert PDF sang Markdown (.md) offline bằng opendataloader-pdf
    md_out_dir = os.path.join(os.path.dirname(args.out), "extracted_md")
    try:
        import opendataloader_pdf
        os.makedirs(md_out_dir, exist_ok=True)
        print(f"[Opendataloader] Đang convert {os.path.basename(convert_target)} sang Markdown...")
        opendataloader_pdf.convert(
            input_path=convert_target,
            output_dir=md_out_dir,
            format="markdown",
            hybrid="docling-fast",
            hybrid_fallback=True
        )
        print(f"[Opendataloader] [OK] Đã convert thành công sang Markdown tại {md_out_dir}")
    except Exception as e:
        print(f"[Opendataloader] [WARN] Không thể convert sang Markdown: {e}")
        return
    finally:
        # Xóa file sliced tạm
        if os.path.exists(sliced_pdf) and sliced_pdf != args.out:
            try:
                os.remove(sliced_pdf)
            except Exception:
                pass

    # Tự động gọi segments_kcn_parser.py để trích xuất số liệu mảng và nạp vào kho JSON
    if hasattr(args, "ticker") and args.ticker and hasattr(args, "period") and args.period:
        try:
            print(f"[Parser] Kích hoạt segments_kcn_parser.py trích xuất mảng cho {args.ticker} kỳ {args.period}...")
            import subprocess
            subprocess.run(
                [sys.executable, "segments_kcn_parser.py", args.ticker, args.period],
                cwd=PROJECT_ROOT,
                check=True
            )
        except Exception as e:
            print(f"[Parser] [WARN] Trích xuất tự động thất bại: {e}")


def cmd_render(args):
    import fitz
    doc = fitz.open(args.pdf)
    os.makedirs(args.out, exist_ok=True)
    start, end = (args.pages.split("-") if "-" in args.pages else (args.pages, args.pages))
    start, end = int(start), int(end)
    base = os.path.splitext(os.path.basename(args.pdf))[0]
    saved = []
    for i in range(start - 1, min(end, len(doc))):
        out_path = os.path.join(args.out, f"{base}_p{i+1}.png")
        doc[i].get_pixmap(dpi=args.dpi).save(out_path)
        saved.append(out_path)
    print(f"[OK] Đã render {len(saved)} trang -> {args.out}")
    for p in saved:
        print(f"  {p}")


def cmd_clean(args):
    ticker = args.ticker.upper()
    ticker_dir = os.path.join(PDF_DIR, ticker)
    if not os.path.isdir(ticker_dir):
        print(f"[Clean] Thư mục BCTC cho {ticker} không tồn tại: {ticker_dir}")
        return
        
    import shutil
    print(f"[Clean] Đang dọn dẹp các tệp trung gian (PDF, Markdown, ảnh) của {ticker}...")
    try:
        shutil.rmtree(ticker_dir)
        print(f"[Clean] [OK] Đã xóa toàn bộ thư mục tệp trung gian: {ticker_dir}")
    except Exception as e:
        print(f"[Clean] [WARN] Có lỗi khi xóa thư mục {ticker_dir}: {e}")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list")
    p_list.add_argument("ticker")
    p_list.set_defaults(func=cmd_list)

    p_plan = sub.add_parser("plan-downloads")
    p_plan.add_argument("ticker")
    p_plan.add_argument("--years", type=int, default=5)
    p_plan.add_argument("--quarters", type=int, default=4)
    p_plan.set_defaults(func=cmd_plan_downloads)

    p_dl = sub.add_parser("download")
    p_dl.add_argument("url")
    p_dl.add_argument("--out", required=True)
    p_dl.add_argument("--ticker", required=False, help="Mã cổ phiếu")
    p_dl.add_argument("--period", required=False, help="Kỳ BCTC ví dụ: 2024(CN) hoặc 2024Q1")
    p_dl.set_defaults(func=cmd_download)

    p_render = sub.add_parser("render")
    p_render.add_argument("pdf")
    p_render.add_argument("--pages", required=True, help="vd 10-15 hoặc 10")
    p_render.add_argument("--out", required=True)
    p_render.add_argument("--dpi", type=int, default=110)
    p_render.set_defaults(func=cmd_render)

    p_clean = sub.add_parser("clean")
    p_clean.add_argument("ticker")
    p_clean.set_defaults(func=cmd_clean)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

