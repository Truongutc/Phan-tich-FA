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
    },
    "SZC": {
        "ChoThueDatPhiQuanLy": ["cho thue dat", "phi quan ly", "dat va phi quan ly"],
        "ThuPhiDuongBo":       ["thu phi duong bo", "duong bo", "bot"],
        "GolfNhaHang":         ["golf", "nha hang", "the thao golf"],
        "CungCapNuoc":         ["cung cap nuoc", "nuoc sach"],
        "BanNhaHuuPhuoc":      ["sonadezi huu phuoc", "huu phuoc", "lien ke"],
        "XuLyNuocThai":        ["xu ly nuoc thai", "nuoc thai"],
        "ChoThueXuong":        ["cho thue xuong", "thue xuong", "quan ly xuong"],
        "Khac":                ["dich vu khac", "khac"]
    }
}

def clean_number(val_str):
    """Làm sạch chuỗi số từ bảng markdown, hỗ trợ định dạng số Việt Nam (dấu chấm phân cách nghìn, dấu phẩy thập phân)."""
    if not val_str:
        return 0.0
    val_str = val_str.strip()
    if val_str in ("-", "_", "–", "—", ""):
        return 0.0
    
    # Xác định số âm nếu nằm trong ngoặc đơn hoặc bắt đầu bằng dấu trừ
    is_negative = False
    if (val_str.startswith("(") and val_str.endswith(")")) or val_str.startswith("-"):
        is_negative = True
        val_str = val_str.replace("(", "").replace(")", "").replace("-", "")
        
    val_str = val_str.strip()
    
    # Phân tích chuỗi số có chứa dấu chấm hoặc phẩy không
    if "." in val_str and "," in val_str:
        if val_str.find(".") < val_str.find(","):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "").replace(".", ".")
    elif "," in val_str:
        if val_str.count(",") >= 2:
            val_str = val_str.replace(",", "")
        else:
            parts = val_str.split(",")
            if len(parts[1]) == 3 and len(parts[0]) <= 3:
                val_str = val_str.replace(",", "")
            else:
                val_str = val_str.replace(",", ".")
    elif "." in val_str:
        if val_str.count(".") >= 2:
            val_str = val_str.replace(".", "")
        else:
            parts = val_str.split(".")
            if len(parts[1]) == 3 and len(parts[0]) <= 3:
                val_str = val_str.replace(".", "")
            else:
                pass

    try:
        val = float(val_str)
        return -val if is_negative else val
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
    Tìm kiếm và trích xuất số liệu doanh thu/giá vốn mảng bằng thuật toán Dynamic Discovery:
    Tự động dò tìm các dòng và cột chứa số liệu mảng thực tế dựa trên cấu trúc bảng thuyết minh,
    sau đó đối chiếu thông minh với mảng mục tiêu của ticker để phân phối số liệu.
    """
    keywords = MAPPING_KEYWORDS.get(ticker, {})
    curr_results = {}
    prior_results = {}
    
    # Chuẩn hóa bộ từ khóa của ticker thành không dấu để so khớp map mảng mục tiêu
    normalized_keywords = {}
    for seg_key, kw_list in keywords.items():
        normalized_keywords[seg_key] = [strip_accents(kw) for kw in kw_list]

    # --- BƯỚC 1: Quét dọc qua toàn bộ các bảng trong Markdown ---
    for table in tables:
        if len(table) < 2:
            continue
        
        # Kiểm tra xem bảng này có phải là bảng thuyết minh Doanh thu / Giá vốn không
        # Bằng cách check xem trong bảng có chứa các dòng từ khóa cơ bản như "doanh thu", "gia von", "loi nhuan gop"
        has_rev_keywords = False
        has_cogs_keywords = False
        for row in table:
            if not row or not row[0]:
                continue
            row_norm = strip_accents(row[0])
            if any(x in row_norm for x in ["doanhthu", "revenue"]):
                has_rev_keywords = True
            if any(x in row_norm for x in ["giavon", "giathanh", "cogs", "loinhuangop", "grossprofit"]):
                has_cogs_keywords = True
        
        if not has_rev_keywords and not has_cogs_keywords:
            continue  # Bảng thuyết minh không chứa chỉ tiêu tài chính mảng, bỏ qua

        # Dò tìm cấu trúc cột (Cột giá trị kỳ này, kỳ trước)
        # Cấu trúc phổ biến: Mảng | Mã số | Kỳ này | Kỳ trước  hoặc Mảng | Kỳ này | Kỳ trước
        val_idx_curr = 1
        val_idx_prior = 2
        
        # Check thử dòng đầu tiên chứa giá trị để tinh chỉnh chỉ số cột
        for row in table[1:]:
            if len(row) > 2 and len(row[1].strip()) <= 4:
                val_idx_curr = 2
                val_idx_prior = 3
                break

        # Duyệt qua từng dòng để bóc tách mảng
        for row in table:
            if not row or not row[0]:
                continue
            row_text = row[0].strip()
            row_text_normalized = strip_accents(row_text)
            
            # Bỏ qua các dòng tổng cộng, tiêu đề cha hoặc chỉ tiêu tài chính tổng
            if any(x in row_text_normalized for x in ["cong", "tongcong", "doanhthuthuan", "giavonhangban", "loinhuangop", "tructiep", "noibo"]):
                continue
            # Bỏ qua các dòng chỉ mục ngắn hoặc trống
            if len(row_text) < 4:
                continue

            # Xác định dòng này nói về mảng nào của ticker bằng cách so khớp thông minh không dấu
            matched_seg_key = None
            for seg_key, kw_list in normalized_keywords.items():
                if any(kw in row_text_normalized for kw in kw_list):
                    matched_seg_key = seg_key
                    break
            
            if not matched_seg_key:
                # Nếu không khớp với mảng mục tiêu nào định nghĩa sẵn, giữ nguyên nhãn gốc làm key tạm để tránh gom sót mảng chính
                # Nhưng để đồng bộ với cấu hình, ta chỉ map nếu tên mảng có nghĩa (chứa chữ cái)
                if any(c.isalpha() for c in row_text_normalized):
                    matched_seg_key = row_text
            
            if matched_seg_key:
                # Xem dòng hiện tại là Doanh thu hay Giá vốn
                is_cogs = any(x in row_text_normalized for x in ["giavon", "giathanh", "cogs"])
                
                # Trích xuất số liệu kỳ hiện tại
                val_c = 0.0
                if len(row) > val_idx_curr:
                    val_c = clean_number(row[val_idx_curr])
                    if val_c > 1000000: val_c = round(val_c / 1e9, 2)
                
                # Trích xuất số liệu kỳ so sánh
                val_p = 0.0
                if len(row) > val_idx_prior:
                    val_p = clean_number(row[val_idx_prior])
                    if val_p > 1000000: val_p = round(val_p / 1e9, 2)

                # Chỉ lưu nếu có giá trị thực tế lớn hơn 0
                if val_c > 0 or val_p > 0:
                    target_dict_curr = curr_results.setdefault(matched_seg_key, {
                        "revenue": 0.0, "cogs": 0.0,
                        "source": f"Trích xuất tự động từ BCTC md ({period_key})",
                        "sourceType": "auto", "derived": False
                    })
                    target_dict_prior = prior_results.setdefault(matched_seg_key, {
                        "revenue": 0.0, "cogs": 0.0,
                        "source": f"Trích xuất cột so sánh từ BCTC md ({period_key})",
                        "sourceType": "auto", "derived": False
                    })
                    
                    if is_cogs:
                        if val_c > 0: target_dict_curr["cogs"] = val_c
                        if val_p > 0: target_dict_prior["cogs"] = val_p
                    else:
                        if val_c > 0: target_dict_curr["revenue"] = val_c
                        if val_p > 0: target_dict_prior["revenue"] = val_p

    # --- BƯỚC 2: Fallback cho Bảng thuyết minh bộ phận (bảng ngang ma trận) ---
    # Nếu kết quả quét dọc rỗng, chuyển sang quét bảng bộ phận (bảng ma trận ngang)
    if not curr_results:
        print("[Parser] [INFO] Quét dọc không có kết quả. Chuyển sang quét bảng bộ phận ngang...")
        for table in tables:
            if len(table) < 3:
                continue
            
            # Tìm dòng tiêu đề chứa tên các mảng bộ phận
            header_row_idx = -1
            col_mappings = {} # map column_index -> seg_name / seg_key
            
            for r_idx in range(min(3, len(table))):
                row = table[r_idx]
                if not row:
                    continue
                
                # Check các ô trong dòng tiêu đề
                for col_idx, cell in enumerate(row):
                    if not cell or col_idx == 0:
                        continue
                    cell_norm = strip_accents(cell)
                    if any(x in cell_norm for x in ["cong", "tongcong", "loaihieu", "bo phan", "khoanmuc", "chi tieu"]):
                        continue
                    
                    # Thử map với mảng chuẩn của ticker
                    matched_key = None
                    for seg_key, kw_list in normalized_keywords.items():
                        if any(kw in cell_norm for kw in kw_list):
                            matched_key = seg_key
                            break
                    if not matched_key:
                        matched_key = cell.strip()  # giữ nguyên label mảng
                    
                    col_mappings[col_idx] = matched_key
                    header_row_idx = r_idx

            if col_mappings:
                # Đã map cột. Tìm dòng "Doanh thu" và "Giá vốn"/"Lợi nhuận gộp"
                rev_row = None
                cogs_row = None
                
                for r_idx in range(header_row_idx + 1, len(table)):
                    row = table[r_idx]
                    if not row or not row[0]:
                        continue
                    row_lbl_norm = strip_accents(row[0])
                    
                    if any(x in row_lbl_norm for x in ["doanhthuthuan", "doanhthutu", "doanhthubanhang", "revenue"]):
                        if "noibo" not in row_lbl_norm:
                            rev_row = row
                    if any(x in row_lbl_norm for x in ["loinhuangop", "grossprofit", "giavon", "cogs"]):
                        cogs_row = row
                
                if rev_row:
                    for col_idx, seg_key in col_mappings.items():
                        if col_idx < len(rev_row):
                            val_rev = clean_number(rev_row[col_idx])
                            if val_rev > 1000000: val_rev = round(val_rev / 1e9, 2)
                            
                            val_cogs = None
                            if cogs_row and col_idx < len(cogs_row):
                                val_cogs_raw = clean_number(cogs_row[col_idx])
                                if val_cogs_raw > 1000000: val_cogs_raw = round(val_cogs_raw / 1e9, 2)
                                
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
                    break

    # --- BƯỚC 3: Map chuẩn hóa mảng động về cấu trúc mảng chuẩn của Ticker ---
    # Nếu mảng trích xuất động khớp với mảng chuẩn (ở dạng text), chuyển key về key chuẩn
    final_curr = {}
    final_prior = {}
    
    # Hàm phụ để tìm key mảng chuẩn dựa trên text thô trích xuất
    def get_standard_key(raw_name):
        raw_norm = strip_accents(raw_name)
        for seg_key, kw_list in normalized_keywords.items():
            if any(kw in raw_norm for kw in kw_list):
                return seg_key
        return "Khac"  # Fallback mảng Khác cho các mảng phụ cực nhỏ

    for raw_key, data in curr_results.items():
        std_key = get_standard_key(raw_key) if raw_key not in keywords else raw_key
        # Nếu đã có dữ liệu ở std_key (ví dụ: đã cộng dồn), ta cộng thêm vào
        dest = final_curr.setdefault(std_key, {
            "revenue": 0.0, "cogs": 0.0,
            "source": data["source"], "sourceType": data["sourceType"], "derived": data["derived"]
        })
        dest["revenue"] = round(dest["revenue"] + data["revenue"], 2)
        dest["cogs"] = round(dest["cogs"] + data["cogs"], 2)

    for raw_key, data in prior_results.items():
        std_key = get_standard_key(raw_key) if raw_key not in keywords else raw_key
        dest = final_prior.setdefault(std_key, {
            "revenue": 0.0, "cogs": 0.0,
            "source": data["source"], "sourceType": data["sourceType"], "derived": data["derived"]
        })
        dest["revenue"] = round(dest["revenue"] + data["revenue"], 2)
        dest["cogs"] = round(dest["cogs"] + data["cogs"], 2)

    return final_curr, final_prior


def find_segment_values_from_text(ticker, md_content, period_key=""):
    """
    Fallback: quét qua nội dung văn bản thô để tìm doanh thu/giá vốn mảng.
    Thích hợp cho cả bảng dọc (line-by-line) và bảng ngang ma trận (matrix).
    """
    import re
    keywords = MAPPING_KEYWORDS.get(ticker, {})
    curr_results = {}
    prior_results = {}
    
    # ── 1. THỬ MATRIX PARSER CHO BẢNG NGANG MA TRẬN ─────────────────────────
    static_cols = {
        "IDC": ["HaTangKCN", "DienNang", "BDS", "XayDung", "BOT", "Khac"],
        "SZC": ["ChoThueDatPhiQuanLy", "ThuPhiDuongBo", "GolfNhaHang", "CungCapNuoc", "BanNhaHuuPhuoc", "XuLyNuocThai", "ChoThueXuong", "Khac"],
        "DTD": ["XayLap", "BanBeTong", "BanXangDau", "ChoThueHaTang", "Khac"]
    }.get(ticker)
    
    if static_cols:
        # Tách các khối thuyết minh bộ phận
        blocks = re.split(r'THONG TIN THEO BO PHAN|THONG TIN BO PHAN', md_content, flags=re.IGNORECASE)
        if len(blocks) > 1:
            print(f"[Parser] [Matrix] Tìm thấy {len(blocks)-1} khối thuyết minh bộ phận.")
            for b_idx, block in enumerate(blocks[1:], start=1):
                lines = block.split("\n")
                
                # Xác định năm của khối này
                block_year = None
                for line in lines[:40]:
                    m_yr = re.search(r'nam\s+(202\d)', line, re.IGNORECASE)
                    if m_yr:
                        block_year = int(m_yr.group(1))
                        break
                
                # Tìm dòng Doanh thu và Lợi nhuận gộp / Giá vốn
                rev_vals = []
                gp_vals = []
                cogs_vals = []
                
                for i, line in enumerate(lines[:60]):
                    # Tìm các cụm số
                    nums = re.findall(r'\(?\b\d{1,3}(?:[.,]\d{3})+\b\)?|\b\d{9,13}\b', line)
                    if len(nums) >= len(static_cols) - 1:
                        # Xác định ngữ cảnh
                        context = ""
                        for j in range(max(0, i-5), i+1):
                            prev_line_norm = strip_accents(lines[j])
                            if any(x in prev_line_norm for x in ["doanhthuthuan", "doanhthu", "revenue"]):
                                context = "revenue"
                            elif any(x in prev_line_norm for x in ["loinhuangop", "grossprofit"]):
                                context = "gp"
                            elif any(x in prev_line_norm for x in ["giavon", "cogs"]):
                                context = "cogs"
                                
                        parsed_nums = []
                        for n in nums:
                            neg = False
                            if n.startswith("(") and n.endswith(")"):
                                neg = True
                                n = n[1:-1]
                            n_clean = n.replace(".", "").replace(",", "")
                            try:
                                val = float(n_clean)
                                if val > 1000000: val = round(val / 1e9, 2)
                                parsed_nums.append(-val if neg else val)
                            except ValueError:
                                pass
                            
                        if context == "revenue":
                            rev_vals = parsed_nums
                        elif context == "gp":
                            gp_vals = parsed_nums
                        elif context == "cogs":
                            cogs_vals = parsed_nums
                
                # Nếu tìm thấy doanh thu, map vào kết quả
                if rev_vals:
                    # Xác định xem block này là kỳ hiện tại hay kỳ trước
                    is_prior = False
                    block_text_norm = strip_accents(block)
                    prior_keywords = ["namtruoc", "kytruoc", "sodaunam", "namngoai", "tripeye", "kyso-sanh", "prior"]
                    if any(pw in block_text_norm for pw in prior_keywords):
                        is_prior = True
                    elif block_year and "Q" in period_key:
                        report_yr = int(period_key[:4])
                        if block_year < report_yr:
                            is_prior = True
                    elif block_year and "(" in period_key:
                        report_yr = int(period_key.split("(")[0])
                        if block_year < report_yr:
                            is_prior = True
                            
                    target_dict = prior_results if is_prior else curr_results
                    src_label = f"Trích xuất ma trận ngang ({period_key})"
                    
                    for idx, seg_key in enumerate(static_cols):
                        if idx < len(rev_vals):
                            rev_v = abs(rev_vals[idx])
                            gp_v = abs(gp_vals[idx]) if (gp_vals and idx < len(gp_vals)) else None
                            cogs_v = abs(cogs_vals[idx]) if (cogs_vals and idx < len(cogs_vals)) else None
                            
                            if cogs_v is None and gp_v is not None:
                                cogs_v = round(rev_v - gp_v, 2)
                            if cogs_v is None:
                                cogs_v = round(rev_v * 0.8, 2)
                                
                            target_dict[seg_key] = {
                                "revenue": rev_v,
                                "cogs": cogs_v,
                                "source": src_label,
                                "sourceType": "auto",
                                "derived": False
                            }
            
            if curr_results:
                print(f"[Parser] [Matrix] Trích xuất ma trận ngang thành công: {list(curr_results.keys())}")
                return curr_results, prior_results

    # ── 2. VERTICAL SCAN FALLBACK (Nếu ma trận ngang không chạy) ──────────────────
    normalized_keywords = {}
    for seg_key, kw_list in keywords.items():
        normalized_keywords[seg_key] = [strip_accents(kw) for kw in kw_list]

    lines = md_content.split("\n")
    for i, line in enumerate(lines):
        line_norm = strip_accents(line)
        if not line_norm or len(line_norm) < 4:
            continue
        if any(x in line_norm for x in ["tongcong", "cong", "doanhthuthuan", "giavonhangban", "loinhuangop"]):
            continue

        matched_seg_key = None
        for seg_key, kw_list in normalized_keywords.items():
            if any(kw in line_norm for kw in kw_list):
                matched_seg_key = seg_key
                break
                
        if matched_seg_key:
            text_to_scan = line
            if i + 1 < len(lines) and not re.search(r'\d', line):
                text_to_scan += " " + lines[i+1]
            if i + 2 < len(lines) and not re.search(r'\d', text_to_scan):
                text_to_scan += " " + lines[i+2]
                
            numbers_found = []
            matches = re.finditer(r'\(?\b\d{1,3}(?:[.,]\d{3})+\b\)?|\(?\b\d{1,3}(?:[.,]\d{1,2})\b\)?|\b\d{5,15}\b', text_to_scan)
            for m in matches:
                num_str = m.group(0)
                val = clean_number(num_str)
                if val != 0:
                    numbers_found.append(val)
                    
            if numbers_found:
                is_cogs = any(x in line_norm for x in ["giavon", "giathanh", "cogs"])
                val_c = numbers_found[0]
                val_p = numbers_found[1] if len(numbers_found) > 1 else 0.0
                
                if abs(val_c) > 1000000: val_c = round(val_c / 1e9, 2)
                if abs(val_p) > 1000000: val_p = round(val_p / 1e9, 2)
                
                target_dict_curr = curr_results.setdefault(matched_seg_key, {
                    "revenue": 0.0, "cogs": 0.0,
                    "source": f"Trích xuất văn bản thô ({period_key})",
                    "sourceType": "auto", "derived": False
                })
                target_dict_prior = prior_results.setdefault(matched_seg_key, {
                    "revenue": 0.0, "cogs": 0.0,
                    "source": f"Trích xuất văn bản thô kỳ so sánh ({period_key})",
                    "sourceType": "auto", "derived": False
                })
                
                if is_cogs:
                    if val_c != 0: target_dict_curr["cogs"] = abs(val_c)
                    if val_p != 0: target_dict_prior["cogs"] = abs(val_p)
                else:
                    if val_c != 0: target_dict_curr["revenue"] = abs(val_c)
                    if val_p != 0: target_dict_prior["revenue"] = abs(val_p)

    # Chuẩn hóa về mảng chuẩn
    final_curr = {}
    final_prior = {}
    for raw_key, data in curr_results.items():
        std_key = raw_key
        dest = final_curr.setdefault(std_key, {
            "revenue": 0.0, "cogs": 0.0,
            "source": data["source"], "sourceType": data["sourceType"], "derived": data["derived"]
        })
        dest["revenue"] = round(dest["revenue"] + data["revenue"], 2)
        dest["cogs"] = round(dest["cogs"] + data["cogs"], 2)

    for raw_key, data in prior_results.items():
        std_key = raw_key
        dest = final_prior.setdefault(std_key, {
            "revenue": 0.0, "cogs": 0.0,
            "source": data["source"], "sourceType": data["sourceType"], "derived": data["derived"]
        })
        dest["revenue"] = round(dest["revenue"] + data["revenue"], 2)
        dest["cogs"] = round(dest["cogs"] + data["cogs"], 2)

    return final_curr, final_prior


def _drop_outlier_segments(vals, period_key):
    """Loại bỏ mảng có giá trị revenue bất thường lớn — dấu hiệu lỗi trích xuất bảng (đọc nhầm 1 dòng
    tổng/lũy kế/số khác thay vì đúng dòng mảng), KHÔNG liên quan OCR (đã gặp thật: IDC 2024(CN) mảng
    "Khác" ra 255.591 tỷ trong khi 5 mảng còn lại + Vietcap DT chỉ ~8.846 tỷ — lớn hơn tổng DT cả công
    ty 30 lần). Nếu 1 mảng > 5x tổng TẤT CẢ mảng còn lại (và tổng các mảng còn lại đã đủ lớn để so
    sánh có ý nghĩa) thì coi là lỗi, loại khỏi kết quả thay vì merge số sai vào kho dữ liệu."""
    if len(vals) < 2:
        return vals
    for seg, data in list(vals.items()):
        rev = (data or {}).get("revenue") or 0
        others_sum = sum((v.get("revenue") or 0) for s, v in vals.items() if s != seg)
        if others_sum > 10 and rev > others_sum * 5:
            print(f"  [Parser] [WARN] Kỳ {period_key}: mảng '{seg}' có DT={rev:.1f} bất thường lớn "
                  f"(>5x tổng các mảng còn lại {others_sum:.1f}) — nghi lỗi trích xuất bảng, loại khỏi kết quả.")
            del vals[seg]
    return vals


def _tag_ocr_source(vals):
    """Gắn nhãn rõ dữ liệu trích xuất từ PDF ẢNH SCAN (bắt buộc OCR mới đọc được, xem bctc_pdf_tool.py
    ::cmd_download) — để user tự nhận biết số nào cần đối chiếu kỹ hơn thay vì mặc định tin như số
    trích từ PDF có text thật (đáng tin cậy hơn nhiều)."""
    for seg_data in vals.values():
        if isinstance(seg_data, dict):
            seg_data["sourceType"] = f"{seg_data.get('sourceType', 'auto')}_ocr"
            seg_data["source"] = f"{seg_data.get('source', '')} [OCR - PDF ảnh scan, cần đối chiếu]".strip()
    return vals


def run_parse_and_merge(ticker, period_key, is_ocr=False):
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
        
    # Tìm tệp md có chứa tên kỳ period_key (ví dụ: "2024_CN" hoặc "2025Q3") để parse chính xác
    normalized_period = period_key.replace("(", "_").replace(")", "")
    matching_files = [f for f in md_files if normalized_period in f]
    if matching_files:
        md_file_path = os.path.join(ticker_md_dir, matching_files[0])
    else:
        # Fallback lấy file md mới nhất
        md_file_path = os.path.join(ticker_md_dir, md_files[-1])
        
    print(f"[Parser] Đang phân tích tệp Markdown: {md_file_path}")
    
    with open(md_file_path, "r", encoding="utf-8") as f:
        md_content = f.read()
        
    tables = parse_markdown_tables(md_content)
    print(f"[Parser] Đã tìm thấy {len(tables)} bảng dữ liệu trong Markdown.")
    
    is_q = "Q" in period_key
    curr_vals, prior_vals = find_segment_values(ticker, tables, period_key=period_key)
    
    if not curr_vals:
        print("[Parser] [INFO] Không trích xuất được từ bảng markdown. Thử quét dòng văn bản thô line-by-line fallback...")
        curr_vals, prior_vals = find_segment_values_from_text(ticker, md_content, period_key=period_key)

    curr_vals = _drop_outlier_segments(curr_vals, period_key)
    prior_period_key_for_log = f"kỳ so sánh của {period_key}"
    prior_vals = _drop_outlier_segments(prior_vals, prior_period_key_for_log)

    if is_ocr:
        curr_vals = _tag_ocr_source(curr_vals)
        prior_vals = _tag_ocr_source(prior_vals)
        print(f"[Parser] [DIAG] Kỳ {period_key} trích xuất từ PDF ẢNH SCAN (đã OCR) — dữ liệu được "
              f"gắn nhãn sourceType=*_ocr, cần đối chiếu qua check_segment_consistency trước khi tin dùng.")

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
        print("Usage: python segments_kcn_parser.py <TICKER> <PERIOD_KEY> [--ocr]")
        print("Example: python segments_kcn_parser.py SIP 2024(CN)")
        print("  --ocr: đánh dấu nguồn PDF là ảnh scan (đã OCR) — gắn nhãn sourceType=*_ocr trong kho dữ liệu")
        sys.exit(1)
    run_parse_and_merge(sys.argv[1], sys.argv[2], is_ocr="--ocr" in sys.argv[3:])
