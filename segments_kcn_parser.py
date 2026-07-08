#!/usr/bin/env python3
"""
segments_kcn_parser.py — Tự động trích xuất doanh thu & giá vốn theo mảng của nhóm BĐS KCN
từ các file Markdown (.md) được xuất bằng opendataloader-pdf.
Hoạt động hoàn toàn offline & online deterministic không cần AI.
"""

import os
import re
import sys

# Fix Windows console encoding
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import json
import subprocess

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(PROJECT_ROOT, "BCTC_PDF")
STORE_DIR = os.path.join(PROJECT_ROOT, "data", "segments_kcn")

# Định nghĩa các từ khóa mảng của từng cổ phiếu để trích xuất chuẩn
MAPPING_KEYWORDS = {
    "IDC": {
        "DienNang":    ["điện", "tiện ích điện", "electricity"],
        "HaTangKCN":   ["hạ tầng", "cho thuê đất", "industrial park land", "leasing"],
        "DichVuKCN":   ["dịch vụ khu công nghiệp", "tiện ích kcn", "industrial park services"],
        "BOT":         ["thu phí đường bộ", "bot", "tollroad"],
        "XayDung":     ["xây dựng", "construction"],
        "BDS":         ["kinh doanh bất động sản", "nhà ở", "bất động sản dân cư"],
        "Khac":        ["dịch vụ khác", "khác"]
    },
    "SIP": {
        "TienIchDienNuoc": ["tiện ích điện, nước", "tiện ích điện", "điện nước"],
        "HangHoa":         ["bán hàng hóa", "thành phẩm", "goods"],
        "DichVuKCNKhac":   ["dịch vụ tiện ích kcn khác", "tiện ích khác", "dịch vụ kcn"],
        "ChoThueDat":      ["cho thuê đất", "cho thuê hạ tầng", "cho thuê đất đã phát triển csht"],
        "XayDung":         ["xây dựng", "construction"],
        "BDS":             ["bán bất động sản", "kinh doanh bất động sản"],
        "Khac":            ["dịch vụ khác", "khác"]
    },
    "PHR": {
        "CaoSu":       ["cao su", "mủ cao su", "rubber"],
        "KCN":         ["cho thuê đất", "hạ tầng kcn", "cho thuê hạ tầng"],
        "Go":          ["gỗ", "thanh lý cây cao su", "wood"],
        "Khac":        ["dịch vụ khác", "khác"]
    }
}

def clean_number(val_str):
    """Làm sạch chuỗi số từ bảng markdown."""
    if not val_str:
        return 0.0
    val_str = val_str.strip().replace(" ", "").replace(",", "").replace(".", "") # Loại bỏ dấu phân cách hàng nghìn
    # Chuyển đổi chuỗi thành số thực, nếu có lỗi thì trả về 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def parse_markdown_tables(md_content):
    """
    Trích xuất các dòng bảng trong Markdown.
    Trả về danh sách các list hàng cột.
    """
    tables = []
    lines = md_content.split("\n")
    current_table = []
    
    for line in lines:
        line_strip = line.strip()
        if line_strip.startswith("|"):
            # Tách các cột
            cols = [c.strip() for c in line_strip.split("|")[1:-1]]
            current_table.append(cols)
        else:
            if current_table:
                # Đã kết thúc một bảng
                tables.append(current_table)
                current_table = []
                
    if current_table:
        tables.append(current_table)
        
    return tables

def strip_accents(s):
    """Chuyển chuỗi tiếng Việt thành không dấu, viết thường, loại bỏ khoảng trắng thừa."""
    if not s:
        return ""
    import unicodedata
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().replace(" ", "").replace("\n", "").replace("\r", "")


def find_segment_values(ticker, tables, period_key=""):
    """
    Tìm kiếm và trích xuất số liệu doanh thu mảng từ các bảng thuyết minh.
    Hỗ trợ khớp thông minh không dấu để tránh lỗi font/OCR của các mảng.
    """
    keywords = MAPPING_KEYWORDS.get(ticker, {})
    curr_results = {}
    prior_results = {}
    
    # Chuẩn hóa bộ từ khóa của ticker thành không dấu để so khớp
    normalized_keywords = {}
    for seg_key, kw_list in keywords.items():
        normalized_keywords[seg_key] = [strip_accents(kw) for kw in kw_list]
    
    # Duyệt qua các mảng định nghĩa
    for seg_key, kw_list in normalized_keywords.items():
        found_rev_curr = None
        found_cogs_curr = None
        found_rev_prior = None
        found_cogs_prior = None
        
        for table in tables:
            if len(table) < 2:
                continue
            
            # Quét các dòng trong bảng
            for row in table:
                if not row or not row[0]:
                    continue
                
                # Chuẩn hóa cột 0 của dòng thành không dấu để so khớp
                row_text_normalized = strip_accents(row[0])
                
                # Check khớp từ khóa mảng
                matched = any(kw in row_text_normalized for kw in kw_list)
                if matched:
                    # Cột 1: Kỳ này (current), Cột 2 hoặc Cột 3: Kỳ trước (prior)
                    if len(row) > 1:
                        # Tùy thuộc vào bảng có cột Mã số (code column) hay không
                        # Cấu trúc phổ biến: Mảng | Mã số | Kỳ này | Kỳ trước hoặc Mảng | Kỳ này | Kỳ trước
                        val_idx_curr = 1
                        val_idx_prior = 2
                        
                        # Nếu cột 1 là cột mã số ngắn (ví dụ: "tm", "24", "26", "i", "ii", "1", "2")
                        if len(row[1].strip()) <= 4 and len(row) > 2:
                            val_idx_curr = 2
                            val_idx_prior = 3
                            
                        # Chuẩn hóa dòng text để xem có phải dòng giá vốn hay không
                        is_cogs = any(x in row_text_normalized for x in ["giavon", "giathanh", "cogs"])
                            
                        if len(row) > val_idx_curr:
                            val_c = clean_number(row[val_idx_curr])
                            if val_c > 1000000: val_c = round(val_c / 1e9, 2)
                            if is_cogs:
                                found_cogs_curr = val_c
                            else:
                                found_rev_curr = val_c
                                
                        if len(row) > val_idx_prior:
                            val_p = clean_number(row[val_idx_prior])
                            if val_p > 1000000: val_p = round(val_p / 1e9, 2)
                            if is_cogs:
                                found_cogs_prior = val_p
                            else:
                                found_rev_prior = val_p
                            
        if found_rev_curr is not None:
            curr_results[seg_key] = {
                "revenue": found_rev_curr,
                "cogs": found_cogs_curr if found_cogs_curr is not None else round(found_rev_curr * 0.8, 2),
                "source": f"Trích xuất tự động từ BCTC md ({period_key})",
                "sourceType": "auto",
                "derived": False
            }
        if found_rev_prior is not None:
            prior_results[seg_key] = {
                "revenue": found_rev_prior,
                "cogs": found_cogs_prior if found_cogs_prior is not None else round(found_rev_prior * 0.8, 2),
                "source": f"Trích xuất cột so sánh từ BCTC md ({period_key})",
                "sourceType": "auto",
                "derived": False
            }
            
    # --- NEW: Fallback dành riêng cho Bảng thuyết minh bộ phận (bảng ngang ma trận) ---
    # Nếu kết quả rỗng, có khả năng bảng biểu diễn dạng cột (Tên mảng ở dòng tiêu đề, chỉ tiêu ở dòng đầu)
    if not curr_results:
        print("[Parser] [INFO] Quét cột dọc rỗng. Chuyển sang quét bảng bộ phận (bảng ma trận ngang)...")
        for table in tables:
            if len(table) < 3:
                continue
            
            # 1. Tìm dòng tiêu đề chứa tên các mảng
            header_row_idx = -1
            col_mappings = {} # map column_index -> seg_key
            
            # Quét thử 3 dòng đầu để tìm tiêu đề mảng
            for r_idx in range(min(3, len(table))):
                row = table[r_idx]
                if not row:
                    continue
                
                # Check xem dòng này có chứa mảng nào của ticker không
                for col_idx, cell in enumerate(row):
                    if not cell:
                        continue
                    cell_norm = strip_accents(cell)
                    for seg_key, kw_list in normalized_keywords.items():
                        if any(kw in cell_norm for kw in kw_list):
                            col_mappings[col_idx] = seg_key
                            header_row_idx = r_idx
            
            if col_mappings:
                # 2. Đã map được các cột tương ứng với các mảng. Tìm dòng "Doanh thu" và dòng "Lợi nhuận gộp"/"Giá vốn"
                rev_row = None
                cogs_row = None
                
                for r_idx in range(header_row_idx + 1, len(table)):
                    row = table[r_idx]
                    if not row or not row[0]:
                        continue
                    
                    row_lbl_norm = strip_accents(row[0])
                    
                    # Tìm dòng doanh thu thuần hoặc doanh thu bên ngoài
                    if any(x in row_lbl_norm for x in ["doanhthuthuan", "doanhthutu", "doanhthubanhang", "revenue"]):
                        # Tránh nhầm với dòng doanh thu nội bộ
                        if "noibo" not in row_lbl_norm:
                            rev_row = row
                    
                    # Tìm dòng lợi nhuận gộp hoặc giá vốn
                    if any(x in row_lbl_norm for x in ["loinhuangop", "grossprofit", "giavon", "cogs"]):
                        cogs_row = row
                
                # 3. Trích xuất số liệu từ các cột tương ứng
                if rev_row:
                    for col_idx, seg_key in col_mappings.items():
                        if col_idx < len(rev_row):
                            val_rev = clean_number(rev_row[col_idx])
                            if val_rev > 1000000: val_rev = round(val_rev / 1e9, 2)
                            
                            val_cogs = None
                            if cogs_row and col_idx < len(cogs_row):
                                val_cogs_raw = clean_number(cogs_row[col_idx])
                                if val_cogs_raw > 1000000: val_cogs_raw = round(val_cogs_raw / 1e9, 2)
                                
                                # Nếu dòng là "Lợi nhuận gộp", ta tính Giá vốn = Doanh thu - Lợi nhuận gộp
                                if any(x in strip_accents(cogs_row[0]) for x in ["loinhuangop", "grossprofit"]):
                                    val_cogs = round(val_rev - val_cogs_raw, 2)
                                else:
                                    val_cogs = val_cogs_raw
                            
                            if val_rev > 0:
                                curr_results[seg_key] = {
                                    "revenue": val_rev,
                                    "cogs": val_cogs if val_cogs is not None else round(val_rev * 0.8, 2),
                                    "source": f"Trích xuất ma trận ngang BCTC md ({period_key})",
                                    "sourceType": "auto",
                                    "derived": False
                                }
                    
                    # Với kỳ trước (prior period), bảng ngang ma trận thường tách thành 2 bảng riêng biệt (như IDC trang 62 và 63).
                    # Do đó, số liệu prior_results sẽ được lấy từ file BCTC năm trước tương ứng khi lập kế hoạch tải.
                    # Nên ta chỉ cần return kết quả tìm thấy của kỳ hiện tại.
                    break
                    
    return curr_results, prior_results

def run_parse_and_merge(ticker, period_key):
    """
    Đoạn chạy chính: đọc file md -> parse -> tạo file patch -> merge.
    """
    ticker = ticker.upper()
    # Tìm tệp Markdown trong thư mục extracted_md
    ticker_md_dir = os.path.join(PDF_DIR, ticker, "extracted_md")
    if not os.path.exists(ticker_md_dir):
        print(f"[Parser] Không tìm thấy thư mục Markdown đã xuất cho {ticker}: {ticker_md_dir}")
        return False
        
    md_files = [f for f in os.listdir(ticker_md_dir) if f.endswith(".md")]
    if not md_files:
        print(f"[Parser] Không tìm thấy tệp .md nào trong {ticker_md_dir}")
        return False
        
    # Đọc tệp md mới nhất
    md_file_path = os.path.join(ticker_md_dir, md_files[-1])
    print(f"[Parser] Đang phân tích tệp Markdown: {md_file_path}")
    
    with open(md_file_path, "r", encoding="utf-8") as f:
        md_content = f.read()
        
    tables = parse_markdown_tables(md_content)
    print(f"[Parser] Đã tìm thấy {len(tables)} bảng dữ liệu trong Markdown.")
    
    is_q = "Q" in period_key
    curr_vals, prior_vals = find_segment_values(ticker, tables, period_key=period_key)
    
    # Xác định năm của kỳ báo cáo này
    try:
        if is_q:
            report_year = int(period_key[:4])
        else:
            report_year = int(period_key.split("(")[0])
    except Exception:
        report_year = 2026

    # Nếu là năm nay (report_year >= 2026), chỉ giữ lại dữ liệu năm nay (curr_vals), bỏ qua kỳ so sánh để tránh trùng lặp chéo
    from datetime import datetime
    current_active_year = datetime.now().year # 2026
    
    if report_year >= current_active_year:
        print(f"[Parser] Phát hiện BCTC {period_key} thuộc năm hiện tại ({report_year}). Bỏ qua trích xuất cột so sánh (prior period).")
        prior_vals = {}

    if not curr_vals:
        print(f"[Parser] [WARN] Không trích xuất được mảng nào cho {period_key}.")
        return False
        
    # Tính toán kỳ trước (prior period key)
    prior_period_key = ""
    try:
        if is_q:
            yr = int(period_key[:4])
            q = period_key[4:]
            prior_period_key = f"{yr - 1}{q}"
        else:
            yr = int(period_key.split("(")[0])
            prior_period_key = f"{yr - 1}(CN)"
    except Exception:
        pass

    # Tạo cấu trúc patch JSON
    patch = {}
    if is_q:
        patch["quarterly"] = {period_key: curr_vals}
        if prior_vals and prior_period_key:
            patch["quarterly"][prior_period_key] = prior_vals
    else:
        patch["yearly"] = {period_key: curr_vals}
        if prior_vals and prior_period_key:
            patch["yearly"][prior_period_key] = prior_vals
        
    # Lưu file patch tạm thời
    patch_path = os.path.join(PROJECT_ROOT, f"temp_{ticker}_patch.json")
    with open(patch_path, "w", encoding="utf-8") as f:
        json.dump(patch, f, ensure_ascii=False, indent=2)
        
    # Thực hiện merge vào kho dữ liệu chính qua segments_kcn_tool.py
    print(f"[Parser] Đang merge dữ liệu vừa trích xuất vào kho mảng của {ticker}...")
    try:
        subprocess.run(
            [sys.executable, "segments_kcn_tool.py", "merge", ticker, patch_path, "--force"],
            check=True
        )
        print(f"[Parser] [OK] Đã tự động cập nhật BCTC mảng cho {ticker} kỳ {period_key}.")
        success = True
    except Exception as e:
        print(f"[Parser] [ERROR] Không thể merge patch: {e}")
        success = False
    finally:
        # Xóa file patch tạm
        if os.path.exists(patch_path):
            os.remove(patch_path)
            
    return success

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python segments_kcn_parser.py <TICKER> <PERIOD_KEY>")
        print("Example: python segments_kcn_parser.py SIP 2024(CN)")
        sys.exit(1)
    run_parse_and_merge(sys.argv[1], sys.argv[2])
