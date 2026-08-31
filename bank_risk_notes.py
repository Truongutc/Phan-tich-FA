#!/usr/bin/env python3
"""
bank_risk_notes.py — Tự động dò, tải, OCR nhẹ và phân tích 2 thuyết minh BCTC "RỦI RO LÃI SUẤT" và
"RỦI RO THANH KHOẢN" cho cổ phiếu NGÂN HÀNG. Chỉ gọi từ template_banking.py — ticker đã được phân
loại ngành ngân hàng từ trước (run_analysis.py), nên module này không tự kiểm tra lại ngành.

Cơ chế (đã verify bằng dữ liệu THẬT — BCTC hợp nhất kiểm toán MBB năm 2025, 2026-08-30, xem chi tiết
trong PR/commit message đi kèm):
- 2 thuyết minh này chỉ có ĐẦY ĐỦ trong BCTC HỢP NHẤT ĐÃ KIỂM TOÁN/SOÁT XÉT (Quarter=5 theo quy ước
  CafeF) — báo cáo quý thường (không soát xét) mới rút gọn thuyết minh. Báo cáo BÁN NIÊN đã soát xét
  CÓ ĐẦY ĐỦ 2 bảng này y hệt báo cáo năm (verify ảnh chụp thật TCB "Tại 30/6/2026", 2026-08 — ban đầu
  từng nhầm loại hẳn báo cáo bán niên, xem fetch_bank_risk_gaps()). Dùng lại đúng cơ chế dò+tải PDF đã
  có ở bctc_pdf_tool.py (CafeF + 24hmoney, generic cho mọi ticker).
- Vị trí trong tài liệu: nằm trong khoảng ~72-100% cuối tài liệu (verify MBB 2025: trang 90-91/103 và
  95-96/103, tức 87-93%) — KHÔNG dùng số thứ tự thuyết minh cố định (số này đổi theo từng ngân
  hàng/năm — MBB 2025 là note 49/51) mà tìm theo TIÊU ĐỀ chữ.
- Cả 2 bảng đều đã có SẴN dòng tổng hợp do chính ngân hàng/kiểm toán tính:
    "Mức chênh nhạy cảm với lãi suất nội, ngoại bảng (5) = (3) + (4)"  (rủi ro lãi suất)
    "Mức chênh thanh khoản ròng (3) = (1) - (2)"                        (rủi ro thanh khoản)
  → CHỈ cần định vị và trích ĐÚNG dòng này (không tự cộng trừ lại từng dòng tài sản/nợ — đáng tin cậy
  hơn nhiều so với tái dựng toàn bộ bảng qua OCR, vốn dễ sai từng ô lẻ).
- OCR dùng pytesseract (nhẹ, vài giây/trang) — ĐÃ THỬ và loại bỏ 2 phương án khác: (1)
  opendataloader-pdf-hybrid (docling+easyocr, cùng cơ chế dùng cho HPG/MWG) bị treo >70 phút không
  xong với 1 file chỉ 25-34 trang khi chạy trên máy không có GPU; (2) easyocr gọi trực tiếp (không
  qua hybrid server) mất ~78 giây/trang, quá chậm để quét dò nhiều trang. pytesseract cần binary
  tesseract-ocr cài sẵn trên máy (đã có trong requirements.txt + .github/workflows/analyze_stock.yml
  cài qua apt cho CI; máy Windows cá nhân cần tự cài qua `winget install UB-Mannheim.TesseractOCR`
  — lệnh này cần chạy trong phiên có thể xác nhận UAC, không chạy được từ tiến trình nền không tương
  tác). Module KHÔNG BAO GIỜ crash nếu thiếu tesseract — trả về None, template_banking.py tự bỏ qua
  phần này và không in cảnh báo gây nhiễu (in 1 dòng [SKIP] rõ ràng để user biết cần cài gì).
"""
import os
import re
import sys
import unicodedata
import tempfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import requests

from bctc_pdf_tool import fetch_cafef_list, fetch_24hmoney_list, select_ranked_reports, HEADERS

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(PROJECT_ROOT, ".cache")

# Cột (bucket) theo ĐÚNG thứ tự trái->phải trong 2 bảng — verify thật trên BCTC MBB 2025. Một số ngân
# hàng có thể đặt "Quá hạn" thành 1 cột duy nhất thay vì tách "Trên 3 tháng"/"Đến 3 tháng" như MBB —
# labels chỉ dùng để hiển thị, KHÔNG dùng để suy luận, số bucket thật lấy từ chính số lượng giá trị
# OCR trích được (xem _extract_number_row).
INTEREST_RATE_BUCKETS = ["qua_han", "khong_anh_huong_lai_suat", "den_1_thang", "tu_1_3_thang",
                          "tu_3_6_thang", "tu_6_12_thang", "tu_1_5_nam", "tren_5_nam"]
LIQUIDITY_BUCKETS = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang",
                      "tu_3_12_thang", "tu_1_5_nam", "tren_5_nam"]


def _strip_accents(s):
    s = unicodedata.normalize("NFD", s or "")
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _autorotate(img, pytesseract):
    """Dò và sửa hướng trang (90/180/270 độ) bằng OSD của tesseract TRƯỚC khi OCR chính — 1 vài trang
    BCTC (thường là bảng dàn ngang rộng — vd bảng khe hở lãi suất/thanh khoản) bị nhúng/scan LỘN NGƯỢC
    180 độ so với phần còn lại của tài liệu (dù các trang chữ xuôi khác trong CÙNG file vẫn đọc bình
    thường), khiến OCR đọc ra chuỗi ký tự vô nghĩa/lộn ngược thay vì số thật thay vì báo lỗi rõ ràng —
    bug thật phát hiện 2026-08 qua log CI TCB: 1 trang bảng gap đọc ra rác dạng "9S/ téE ¿LE S06...",
    verify bằng cách so chuỗi rác với chính footer trang (LUÔN có ở mọi trang) — "NH-q191/€S08" khi
    đọc ngược lại đúng là "B05a/TCTD-HN", tức trang bị lộn ngược chứ không phải OCR kém đơn thuần.
    Cần cài thêm gói apt "tesseract-ocr-osd" (dữ liệu ngôn ngữ riêng cho OSD, KHÁC "tesseract-ocr-vie")
    — nếu thiếu, image_to_osd() lỗi, hàm này lặng lẽ trả về ảnh gốc KHÔNG xoay (không đoán mù góc quay
    khi không chắc chắn — thà bỏ sót còn hơn xoay sai làm hỏng 1 trang vốn đã đúng chiều)."""
    try:
        from pytesseract import Output
        osd = pytesseract.image_to_osd(img, output_type=Output.DICT)
        angle = osd.get("rotate", 0)
        if angle:
            return img.rotate(-angle, expand=True)
    except Exception:
        pass
    return img


def _ocr_page_text(pdf_path, page_index, dpi=300):
    """OCR 1 trang (0-based index) bằng pytesseract. Trả về "" nếu thiếu pytesseract/tesseract binary
    hoặc lỗi bất kỳ bước nào — KHÔNG BAO GIỜ raise."""
    try:
        import pytesseract
        import pypdfium2 as pdfium
    except ImportError:
        return None  # None = "không có công cụ", khác "" = "có công cụ nhưng OCR ra rỗng"
    try:
        doc = pdfium.PdfDocument(pdf_path)
        if page_index >= len(doc):
            return ""
        img = doc[page_index].render(scale=dpi / 72).to_pil()
        img = _autorotate(img, pytesseract)
        return pytesseract.image_to_string(img, lang="vie")
    except Exception as e:
        print(f"  [WARN] OCR trang {page_index+1} loi: {e}")
        return ""


def _ocr_page_words(pdf_path, page_index, dpi=300):
    """OCR 1 trang, trả về list dict {"text","left","top","width","height"} — 1 phần tử/từ nhận diện
    được (tọa độ pixel, gốc trên-trái). None nếu thiếu pytesseract, [] nếu lỗi OCR/hết trang — KHÔNG
    BAO GIỜ raise. Dùng cho _extract_number_row_by_position() — xem docstring hàm đó lý do cần tọa độ
    thay vì text tuyến tính từ image_to_string()."""
    try:
        import pytesseract
        from pytesseract import Output
        import pypdfium2 as pdfium
    except ImportError:
        return None
    try:
        doc = pdfium.PdfDocument(pdf_path)
        if page_index >= len(doc):
            return []
        img = doc[page_index].render(scale=dpi / 72).to_pil()
        img = _autorotate(img, pytesseract)
        data = pytesseract.image_to_data(img, lang="vie", output_type=Output.DICT)
        words = []
        for i in range(len(data["text"])):
            t = (data["text"][i] or "").strip()
            if not t:
                continue
            words.append({"text": t, "left": data["left"][i], "top": data["top"][i],
                          "width": data["width"][i], "height": data["height"][i]})
        return words
    except Exception as e:
        print(f"  [WARN] OCR (toa do) trang {page_index+1} loi: {e}")
        return []


def _find_note_pages(pdf_path, start_frac=0.70, max_pages=40):
    """Quét từ start_frac*tổng_số_trang tới hết tài liệu, OCR TỪNG TRANG (dừng ngay khi đã tìm đủ cả
    2 tiêu đề — không OCR speculative cả vùng). Trả về dict {"lai_suat": page_idx|None,
    "thanh_khoan": page_idx|None} (0-based) — CÓ THỂ rỗng/thiếu 1 trong 2 key nếu tài liệu này thật
    sự không có mục đó (vd BCTC quý không soát xét, rút gọn thuyết minh — không phải lỗi, gọi nơi
    dùng cần thử BẢN KHÁC). Trả về None (khác {} rỗng — phân biệt RÕ 2 tình huống) NẾU THIẾU
    pytesseract/tesseract-ocr binary — lúc đó dừng hẳn toàn bộ tính năng, thử bản khác cũng vô ích."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        return None
    doc = pdfium.PdfDocument(pdf_path)
    total = len(doc)
    start = max(0, int(total * start_frac))
    pages_to_scan = list(range(start, total))[:max_pages]

    # Chỉ khớp khi cụm từ nằm trên 1 DÒNG NGẮN đứng riêng (kiểu tiêu đề mục, vd "45.1   Rủi ro lãi
    # suất") — KHÔNG khớp khi cụm từ chỉ xuất hiện giữa 1 câu văn xuôi dài (vd ghi chú tổng quan liệt
    # kê "...rủi ro tín dụng, rủi ro thanh khoản và rủi ro thị trường"). Bug thật phát hiện 2026-08 qua
    # log TCB Q2/2026: ghi chú "43. CHÍNH SÁCH QUẢN LÝ RỦI RO TÀI CHÍNH" liệt kê "rủi ro thanh khoản"
    # trong 1 câu tổng quan Ở TRANG SỚM HƠN 7 TRANG so với tiêu đề mục "45.3 Rủi ro thanh khoản" thật
    # — khớp nhầm câu văn xuôi này làm mốc trang, khiến vùng quét bảng số (+8 trang từ mốc) dừng lại
    # đúng 1 trang TRƯỚC khi tới được bảng thật. Ngưỡng 50 ký tự đủ rộng cho tiêu đề có số mục + tên
    # ("45.1   Rủi ro lãi suất" ~21 ký tự) nhưng đủ hẹp để loại câu văn xuôi (luôn dài hơn nhiều).
    _HEADING_MAX_LEN = 50

    def _has_heading_line(text, phrase):
        for line in text.split("\n"):
            line_flat = _strip_accents(line).strip()
            if phrase in line_flat and len(line_flat) <= _HEADING_MAX_LEN:
                return True
        return False

    found = {}
    for idx in pages_to_scan:
        text = _ocr_page_text(pdf_path, idx)
        if text is None:
            return None  # thiếu pytesseract - dừng hẳn, không quét tiếp vô ích
        if "lai_suat" not in found and _has_heading_line(text, "rui ro lai suat"):
            found["lai_suat"] = idx
        if "thanh_khoan" not in found and _has_heading_line(text, "rui ro thanh khoan"):
            found["thanh_khoan"] = idx
        if "lai_suat" in found and "thanh_khoan" in found:
            break
    return found


_ROW_LABEL_FLAT = {
    "lai_suat": "muc chenh nhay cam",
    "thanh_khoan": "muc chenh thanh khoan rong",
}
# Cặp từ neo (đã strip dấu) để định vị dòng gap theo TỌA ĐỘ — xem _extract_number_row_by_position().
# Chọn 2 từ khá riêng biệt trong cụm nhãn (không dùng "muc"/"chenh" vì quá phổ biến, dễ trùng chỗ
# khác trên trang) để giảm khớp nhầm: "nhạy"+"cảm" (rủi ro lãi suất), "khoản"+"ròng" (rủi ro thanh khoản).
_ROW_ANCHOR_WORDS = {
    "lai_suat": ("nhay", "cam"),
    "thanh_khoan": ("khoan", "rong"),
}
# Dòng "Tổng nợ phải trả" nằm NGAY TRÊN dòng "Mức chênh thanh khoản ròng" trong CÙNG bảng thanh khoản
# (đã verify ảnh chụp thật TCB) — trích thêm dòng này (tận dụng lại đúng text/tọa độ trang ĐÃ OCR cho
# dòng gap, không tốn thêm lượt OCR nào) để tính "Nợ phải trả ngắn hạn" phục vụ tỷ lệ Liquid
# Assets/Nợ phải trả ngắn hạn (bản ĐƠN GIẢN HÓA của LCR mà user yêu cầu — KHÁC LCR Basel chuẩn, chỉ
# cần tổng nợ phải trả theo kỳ hạn, không cần phân loại HQLA/outflow rate như LCR thật).
_ROW_LABEL_FLAT_LIAB = "tong no phai tra"
_ROW_ANCHOR_WORDS_LIAB = ("no", "phai")


# Ô số: số VN chuẩn (dấu chấm phân cách nghìn, ngoặc = âm) HOẶC dấu gạch ngang đơn (= 0) HOẶC — dự
# phòng — 1 dãy ≥4 chữ số THUẦN không dấu chấm (OCR thỉnh thoảng làm mất dấu chấm ở 1 vài ô riêng lẻ,
# bug thật 2026-08: "34.144.640" bị đọc ra "34144640"). BẮT BUỘC dùng CHUNG 1 regex duy nhất (không
# tách "thử chặt trước, lỏng sau" như bản đầu) — thử tách riêng đã gây bug KHÁC: khi 1 ô giữa hàng bị
# mất dấu chấm, phần "chặt" vẫn đủ ĐẾM ra n_buckets token (vì tình cờ nhặt được cột "Tổng cộng" ở cuối
# hàng bù vào chỗ trống), lấy nhầm cột forecast/tổng thay cho ô lỗi — 1 regex duy nhất giữ đúng THỨ TỰ
# trái→phải nên ô mất dấu chấm được điền lại ĐÚNG VỊ TRÍ của nó, không bị cột sau đó nhảy vào thế chỗ.
_CELL_RE = re.compile(r"\(?-?[\d]{1,3}(?:\.[\d]{3})+\)?|(?<![\w.])-(?![\w.])|\(?[\d]{4,}\)?")


def _extract_cell_tokens(text, n_buckets):
    """Trích đúng `n_buckets` "ô số" đầu tiên từ `text` theo ĐÚNG thứ tự trái→phải (xem _CELL_RE ở
    trên). Trả về None nếu không đủ số lượng — KHÔNG ĐOÁN số liệu thiếu."""
    toks = [m.group(0) for m in _CELL_RE.finditer(text)]
    return toks[:n_buckets] if len(toks) >= n_buckets else None


def _tokens_to_values(toks):
    """Chuyển list token chuỗi (từ _extract_cell_tokens) sang list float — "-" = 0.0, ngoặc = âm,
    bỏ dấu chấm phân cách nghìn. Trả về None nếu có token không parse được (không nên xảy ra vì
    toks đã qua 2 regex ở trên, nhưng phòng hờ)."""
    vals = []
    for tok in toks:
        if tok == "-":
            vals.append(0.0)
            continue
        neg = tok.startswith("(") and tok.endswith(")")
        clean = tok.strip("()").replace(".", "")
        try:
            v = float(clean)
        except ValueError:
            return None
        vals.append(-v if neg else v)
    return vals


def _extract_number_row(text, label_flat, n_buckets, lines_after=6, debug_tag=None):
    """Tìm DÒNG có nhãn `label_flat` (đã strip dấu) trong `text`, trích các "ô số" trong nhãn đó +
    `lines_after` dòng kế tiếp. Chỉ trả về kết quả nếu số lượng ô trích được >= n_buckets (lấy đúng
    n_buckets ô ĐẦU TIÊN — cột "Tổng cộng" dư nếu có sẽ bị bỏ qua vì không cần). Trả về None nếu
    không tìm thấy nhãn hoặc thiếu số — KHÔNG ĐOÁN số liệu thiếu.

    Xử lý THEO DÒNG (không phải theo vị trí ký tự trong text gốc): tránh bug lệch index giữa bản đã
    strip dấu (`flat`, NFD rồi bỏ dấu tổ hợp LUÔN làm text NGẮN ĐI so với bản gốc) và text gốc — dùng
    vị trí ký tự trong `flat` để cắt `text` gốc là SAI vì 2 chuỗi lệch độ dài, sẽ cắt trúng đoạn không
    liên quan trên trang (bug thật, phát hiện 2026-08 khi xem lại log TCB: SKIP "OCR chất lượng kém"
    trong khi rất có thể nhãn+bảng đọc ĐÚNG nhưng bị cắt sai vị trí do bug này).

    "Ô số" chấp nhận CẢ dấu gạch ngang "-" (băng 0 — rất phổ biến trong thuyết minh BCTC cho cột
    không phát sinh giao dịch, vd cột "Quá hạn" của nhiều ngân hàng) — nếu chỉ khớp số thật sẽ làm
    lệch cột so với nhãn (dòng có 1 dấu "-" thì các cột sau đó bị đếm lùi 1 vị trí)."""
    lines = text.split("\n")
    flat_lines = [_strip_accents(l) for l in lines]
    # Nhãn có thể trải dài qua NHIỀU dòng (vd MBB thật: "Mức chênh nhạy cảm với lãi suất" rồi mới
    # đến "nội, ngoại bảng (5) = (3) + (4)" ở dòng sau) — ghép cửa sổ trượt 3 dòng liên tiếp để tìm.
    label_line_idx = None
    for i in range(len(lines)):
        joined = " ".join(flat_lines[i:i + 3])
        if label_flat in joined:
            label_line_idx = i  # lấy dòng ĐẦU tiên của cửa sổ khớp cuối cùng tìm được (label ổn định)
    if label_line_idx is None:
        if debug_tag:
            # In thêm các dòng có chứa "chenh" (từ khoá chung của MỌI biến thể nhãn dòng gap: "muc
            # chenh", "chenh lech"...) để biết ngay năm nay ngân hàng viết nhãn khác đi thế nào, thay
            # vì chỉ báo "không thấy" — tránh phải đoán mù rồi chờ 1 vòng chạy CI nữa mới biết (bug
            # thật, 2026-08: TCB dùng "nội bảng" thay vì "nội, ngoại bảng" cho nhãn lãi suất, phát
            # hiện được nhờ preview tương tự; nhãn thanh khoản có thể cũng đổi cách viết tương tự).
            candidates = [lines[i] for i in range(len(lines)) if "chenh" in flat_lines[i]]
            preview = " | ".join(c.strip() for c in candidates[:3])[:250] if candidates else \
                " ".join(lines).strip()[:200]
            print(f"  [DIAG] {debug_tag}: khong tim thay nhan '{label_flat}' tren trang nay. "
                  f"Dong co 'chenh' tren trang (neu co): \"{preview}\"")
        return None
    window_text = "\n".join(lines[label_line_idx:label_line_idx + 1 + lines_after])
    # Nhãn dòng gap LUÔN kèm công thức tham chiếu kiểu "(3) = (1) - (2)" hoặc "(5) = (3) + (4)" ngay
    # trong/cạnh chính nó — dấu "-" hoặc "+" ở đây là TOÁN TỬ, không phải ô dữ liệu, phải loại trước
    # khi tìm "ô số" (bug thật, 2026-08: dấu "-" của công thức bị đếm nhầm thành cột đầu tiên, làm
    # lệch toàn bộ các cột phía sau 1 vị trí).
    window_text = re.sub(r"\(\s*\d\s*\)\s*=\s*\(\s*\d\s*\)(?:\s*[+\-]\s*\(\s*\d\s*\))*", " ", window_text)
    toks = _extract_cell_tokens(window_text, n_buckets)
    if toks is None:
        if debug_tag:
            preview = window_text.replace("\n", " | ")[:300]
            print(f"  [DIAG] {debug_tag}: tim thay nhan nhung khong du so lieu. Cua so OCR: \"{preview}\"")
        return None
    return _tokens_to_values(toks)


def _extract_number_row_by_position(words, anchor_words, n_buckets, debug_tag=None):
    """Định vị dòng gap bằng TỌA ĐỘ PIXEL thay vì thứ tự đọc tuyến tính của image_to_string() — BUG
    THẬT phát hiện 2026-08 qua log CI thật của TCB: với bảng rộng 9 cột trải hết bề ngang trang,
    tesseract đọc lộn xộn — nhãn dòng gap bị nối liền với footnote/tiêu đề cột nằm Ở VỊ TRÍ KHÁC trên
    trang thay vì 8 số thật nằm NGAY SAU nó theo hàng ngang. `_extract_number_row` (dựa vào
    image_to_string) chỉ đáng tin cho các đoạn văn xuôi hẹp, KHÔNG đáng tin cho bảng số rộng — hàm
    này thay thế bằng cách tự dựng lại đúng 1 HÀNG theo tọa độ:
    1. Tìm 2 "từ neo" (`anchor_words`, đã strip dấu, vd "nhay"+"cam") đứng gần nhau theo trục dọc
       (cùng 1 hàng) — xác nhận đây đúng là cụm nhãn cần tìm, không phải từ trùng ngẫu nhiên ở chỗ khác.
    2. Lấy tọa độ Y của từ neo ĐẦU làm mốc hàng, gom TẤT CẢ các từ khác trên trang có Y lệch trong
       ngưỡng chiều cao 1 dòng chữ (cùng hàng ngang thật, không phụ thuộc thứ tự OCR trả về).
    3. Sắp xếp lại các từ đó theo X (trái→phải) — đây chính là thứ tái tạo ĐÚNG thứ tự cột của bảng.
    4. Trích số VN từ chuỗi đã sắp xếp lại, giống `_extract_number_row`.
    Nếu có NHIỀU cặp neo khớp trên trang (nhãn dòng gap có thể lặp — VD (3) nội bảng và (5) nội+ngoại
    bảng), lấy cặp CUỐI (gần cuối trang hơn, giống lý do dùng rfind trước đây)."""
    a1, a2 = anchor_words
    flat = [(_strip_accents(w["text"]), w) for w in words]
    anchor = None
    for t1, w1 in flat:
        if t1 != a1:
            continue
        for t2, w2 in flat:
            if t2 == a2 and abs(w2["top"] - w1["top"]) < max(w1["height"], w2["height"], 1) * 1.5:
                anchor = w1  # ghi đè nếu tìm thấy khớp SAU (gần cuối trang hơn) — tương đương rfind
                break
    if anchor is None:
        if debug_tag:
            # Tương tự preview "chenh" ở _extract_number_row: in các từ OCR đọc được có chứa 1 phần
            # của từ neo (vd "khoan"/"rong" hoặc biến thể gần đúng) để biết OCR đọc nhãn thành gì thay
            # vì chỉ báo "không thấy" — không đoán mù, chờ preview này ở log lần chạy sau.
            near = [w["text"] for t, w in flat if a1[:3] in t or a2[:3] in t]
            preview = ", ".join(near[:10]) if near else "(khong co tu nao gan giong)"
            print(f"  [DIAG] {debug_tag}: (toa do) khong tim thay tu neo '{a1}'+'{a2}' gan nhau tren "
                  f"trang nay. Tu OCR gan giong: {preview}")
        return None
    row_top, row_h = anchor["top"], max(anchor["height"], 1)
    # Bảng trải rất rộng hết bề ngang trang (~1700px) — 1 vài ô ở xa neo có thể lệch Y vài pixel do
    # scan hơi nghiêng/rung nét, không có 1 ngưỡng dung sai duy nhất đúng cho mọi lần quét. Thử tăng
    # dần dung sai, DÙNG NGƯỠNG NHỎ NHẤT đã đủ đọc ra n_buckets ô để giảm rủi ro gom nhầm hàng bên
    # cạnh (bug thật, 2026-08: log TCB v1 thiếu 3/8 ô dù đã định vị đúng hàng neo).
    toks, row_text = None, ""
    for tol_mult in (1.2, 2.0, 3.0, 4.5):
        row_words = sorted((w for w in words if abs(w["top"] - row_top) < row_h * tol_mult),
                            key=lambda w: w["left"])
        row_text = " ".join(w["text"] for w in row_words)
        toks = _extract_cell_tokens(row_text, n_buckets)
        if toks is not None:
            break
    if toks is None:
        if debug_tag:
            preview = row_text[:300]
            print(f"  [DIAG] {debug_tag}: (toa do) tim thay hang neo nhung khong du so lieu du moi nguong "
                  f"dung sai da thu. Hang rong nhat: \"{preview}\"")
        return None
    return _tokens_to_values(toks)


_CURRENCY_ROW_RE = re.compile(
    r"\b(USD|VND|EUR|GBP|JPY|AUD|CHF|CNY|SGD|HKD)\b\s+([\d]+,[\d]+)\s*%\s*"
    r"(\(?-?[\d]{1,3}(?:\.[\d]{3})*\)?)\s+(\(?-?[\d]{1,3}(?:\.[\d]{3})*\)?)"
)


def _clean_vn_number(tok):
    if tok in ("-", ""):
        return 0.0
    neg = tok.startswith("(") and tok.endswith(")")
    v = float(tok.strip("()").replace(".", ""))
    return -v if neg else v


def _extract_rate_sensitivity_table(text):
    """Trích bảng "Độ nhạy đối với lãi suất" do CHÍNH ngân hàng công bố (không phải khe hở kỳ hạn) —
    dạng "USD  1,50%  (189.357)  (151.486)" mỗi dòng 1 loại tiền: % lãi suất tăng giả định, ảnh hưởng
    LNTT (triệu đồng), ảnh hưởng VCSH (triệu đồng). KHÔNG PHẢI mọi ngân hàng có mục này riêng (MBB
    2025 không có, TCB 2025 có ở "45.1 Rủi ro lãi suất" ngay trước bảng khe hở) — trả về {} nếu không
    tìm thấy dòng nào khớp, KHÔNG suy diễn. Đáng tin hơn NII sensitivity tự ước tính (static gap) vì
    là số ngân hàng tự tính có tính đến hành vi/giả định nội bộ, không phải xấp xỉ tuyến tính đơn giản.
    """
    result = {}
    for m in _CURRENCY_ROW_RE.finditer(text):
        ccy, pct_str, pbt_tok, equity_tok = m.groups()
        if ccy in result:
            continue
        try:
            pct = float(pct_str.replace(",", "."))
            pbt = _clean_vn_number(pbt_tok)
            equity = _clean_vn_number(equity_tok)
        except ValueError:
            continue
        result[ccy] = {"rate_increase_pct": pct, "pbt_impact": pbt, "equity_impact": equity}
    return result


def _is_half_year(name):
    n = _strip_accents(name)
    return "6 thang" in n or "ban nien" in n or "soat xet" in n


# Tháng kết thúc kỳ báo cáo theo Quarter (quy ước CafeF: 1-4 = quý thường, 5 = năm kiểm toán, 6 = bán
# niên đã soát xét) — dùng để so sánh ĐỘ MỚI THẬT giữa các loại kỳ khác nhau (vd Quý 3 (tháng 9) mới
# hơn bán niên (tháng 6) dù cùng 1 năm).
_QUARTER_END_MONTH = {1: 3, 2: 6, 3: 9, 4: 12, 5: 12, 6: 6}


def _extract_gaps_from_pdf(pdf_path):
    """Định vị + trích 2 bảng gap từ 1 file PDF cụ thể đã tải sẵn. Trả về ("missing_tool", None) nếu
    thiếu pytesseract/tesseract-ocr binary (dừng hẳn, thử file khác cũng vô ích), ("no_note", None)
    nếu tài liệu này THẬT SỰ không có 1 trong 2 tiêu đề/không đọc đủ số (vd BCTC quý không soát xét,
    rút gọn thuyết minh — không phải lỗi, bên gọi nên thử bản BCTC khác), hoặc ("ok", partial_dict)
    với partial_dict chứa các key đã trích được trong {"interest_rate_gap", "liquidity_gap",
    "interest_rate_sensitivity_disclosed", "liabilities_by_bucket"} (key cuối — tổng nợ phải trả theo
    kỳ hạn từ bảng thanh khoản — chỉ có nếu trích được, dùng tính Liquid Assets/Nợ phải trả ngắn hạn)."""
    pages = _find_note_pages(pdf_path)
    if pages is None:
        return "missing_tool", None

    result = {}
    for key, bucket_list, n in (("lai_suat", INTEREST_RATE_BUCKETS, len(INTEREST_RATE_BUCKETS)),
                                 ("thanh_khoan", LIQUIDITY_BUCKETS, len(LIQUIDITY_BUCKETS))):
        page_idx = pages.get(key)
        if page_idx is None:
            continue
        # Bảng số nằm CÁCH trang tiêu đề/phương pháp luận 1 SỐ trang KHÔNG CỐ ĐỊNH — verify thật: MBB
        # 2025 cách đúng 1 trang (heading trang 90 -> bảng trang 91), nhưng TCB 2025 cách 2 trang vì
        # có thêm 1 trang phụ "Độ nhạy đối với lãi suất" (bảng ảnh hưởng LNTT/VCSH theo % lãi suất
        # tăng — khác bảng khe hở lãi suất theo kỳ hạn) chen giữa (heading trang 92 -> bảng trang 94).
        # Mục thanh khoản TCB 2025 KHÔNG tìm thấy nhãn/từ neo trong +5 trang đầu (verify qua log CI
        # thật) — có thể còn nhiều trang phụ hơn xen giữa so với mục lãi suất; nới lên +7 trang, vẫn
        # DỪNG NGAY khi đọc đủ số, không OCR speculative quá xa.
        vals = None
        for p in range(page_idx, page_idx + 8):
            text = _ocr_page_text(pdf_path, p)
            if not text:
                continue
            # Bảng "Độ nhạy đối với lãi suất" (do NH tự công bố, USD/VND... % LS -> LNTT/VCSH) hay nằm
            # NGAY TRƯỚC bảng khe hở lãi suất trong cùng mục 45.1 — tranh thủ đọc luôn trên đường quét
            # tìm bảng khe hở, không tốn thêm lượt OCR trang nào. Không phải NH nào cũng công bố mục
            # này (MBB 2025 không có) nên chỉ ghi nếu tìm thấy, không bắt buộc.
            if key == "lai_suat" and "interest_rate_sensitivity_disclosed" not in result:
                sens = _extract_rate_sensitivity_table(text)
                if sens:
                    result["interest_rate_sensitivity_disclosed"] = sens
            wwords = None
            vals = _extract_number_row(text, _ROW_LABEL_FLAT[key], n, debug_tag=f"{key} trang {p+1}")
            if not vals:
                # Fallback theo TỌA ĐỘ (xem docstring _extract_number_row_by_position) — CHỈ chạy khi
                # cách đọc tuyến tính ở trên thất bại, vì OCR theo tọa độ tốn thêm 1 lượt OCR trang
                # (chậm hơn) — hầu hết trang không phải bảng gap sẽ bị loại ngay ở bước tuyến tính (rẻ)
                # phía trên mà không cần OCR lại theo tọa độ.
                wwords = _ocr_page_words(pdf_path, p)
                if wwords:
                    vals = _extract_number_row_by_position(wwords, _ROW_ANCHOR_WORDS[key], n,
                                                            debug_tag=f"{key} trang {p+1}")
            if vals and key == "thanh_khoan" and "liabilities_by_bucket" not in result:
                # Tranh thủ trích luôn dòng "Tổng nợ phải trả" trên CÙNG trang/text/wwords đã OCR cho
                # dòng gap (KHÔNG tốn thêm lượt OCR nào) — dùng cho Liquid Assets/Nợ phải trả ngắn hạn.
                liab_vals = _extract_number_row(text, _ROW_LABEL_FLAT_LIAB, n, debug_tag=f"no_phai_tra trang {p+1}")
                if not liab_vals:
                    if wwords is None:
                        wwords = _ocr_page_words(pdf_path, p)
                    if wwords:
                        liab_vals = _extract_number_row_by_position(wwords, _ROW_ANCHOR_WORDS_LIAB, n,
                                                                     debug_tag=f"no_phai_tra trang {p+1}")
                if liab_vals:
                    result["liabilities_by_bucket"] = dict(zip(bucket_list, liab_vals))
            if vals:
                break
        if vals:
            gap_key = "interest_rate_gap" if key == "lai_suat" else "liquidity_gap"
            result[gap_key] = dict(zip(bucket_list, vals))

    if "interest_rate_gap" not in result and "liquidity_gap" not in result:
        return "no_note", None
    return "ok", result


def _download_report_pdf(ticker, cand):
    """Tải 1 báo cáo (dict từ select_ranked_reports) về cache, dùng lại nếu đã tải trước đó. Tên file
    cache phân biệt theo (Year, loại kỳ) — KHÔNG chỉ theo Year — để 1 báo cáo bán niên và báo cáo cả
    năm CÙNG NĂM (vd bán niên 2026 rồi cuối năm có thêm báo cáo năm 2026) không bị đè/dùng nhầm cache
    của nhau."""
    q = cand.get("Quarter")
    if q in (5, 6):
        period_tag = "H1" if (q == 6 or _is_half_year(cand.get("Name", ""))) else "FY"
    else:
        period_tag = f"Q{q}"
    pdf_path = os.path.join(CACHE_DIR, f"{ticker}_{cand['Year']}_{period_tag}_CN_full.pdf")
    if os.path.exists(pdf_path):
        return pdf_path
    r = requests.get(cand["Link"].replace(" ", "%20"), headers=HEADERS, timeout=60)
    r.raise_for_status()
    with open(pdf_path, "wb") as f:
        f.write(r.content)
    return pdf_path


def _is_reviewed_candidate(cand):
    q = cand.get("Quarter")
    return q in (5, 6) or _is_half_year(cand.get("Name", ""))


def _period_key_for_candidate(cand):
    """Trả về "YYYY-Qn" (n=1-4) hoặc "YYYY-FY" cho MỌI báo cáo hợp nhất hợp lệ — KỂ CẢ báo cáo quý
    thường không soát xét — None nếu không map được vào 1 kỳ rõ ràng. Dùng LÀM KHÓA cho toàn bộ hệ
    thống lưu trữ theo kỳ (bank_alm_store.py) — CÙNG quy ước "YYYY-Qn" với dữ liệu bảng cân đối theo
    quý (không cần OCR), để 1 kỳ dù đến từ nguồn nào cũng gộp về đúng 1 khóa.

    SỬA LỚN 2026-08 (user cung cấp ảnh chụp thật báo cáo Quý 3/2025 CỦA CẢ MBB và TCB): trước đây
    hàm này chỉ coi báo cáo đã kiểm toán/soát xét là đáng tin cho 2 bảng khe hở lãi suất/thanh
    khoản, dựa trên giả định BCTC quý thường không có thuyết minh này. Giả định đó SAI — verify
    trực tiếp bằng cách tải THẬT 2 báo cáo Quý 3/2025 (KHÔNG soát xét) và render ẢNH từng trang
    (KHÔNG dùng pdfplumber.extract_text() — phương pháp cũ từng cho kết luận sai vì các trang bảng
    dạng này thường không trích xuất được bằng text, phải xem bằng ảnh/OCR): cả 2 ngân hàng đều CÓ
    ĐẦY ĐỦ mục "RỦI RO THỊ TRƯỜNG" (lãi suất/tiền tệ/thanh khoản) ngay trong báo cáo quý thường. Vì
    vậy giờ MỌI Quarter hợp lệ (1-6) đều là candidate khả dĩ — _extract_gaps_from_pdf tự trả về
    "no_note" an toàn nếu 1 báo cáo cụ thể thực sự không có, không giả định trước theo loại kỳ nữa."""
    q = cand.get("Quarter")
    name = cand.get("Name", "")
    is_half = (q == 6) or (q in (2, 5) and _is_half_year(name))
    if is_half:
        return f"{cand['Year']}-Q2"
    if q == 5:
        return f"{cand['Year']}-FY"
    if q in (1, 2, 3, 4):
        return f"{cand['Year']}-Q{q}"
    return None


def _select_candidate_reports(ticker):
    """Lấy + xếp hạng danh sách BCTC của 1 ticker — logic CHỌN dùng chung cho cả fetch_bank_risk_gaps
    (mới nhất), fetch_bank_risk_gaps_for_period (1 kỳ lịch sử cụ thể), và latest_reviewed_period
    (kiểm tra rẻ, không tải/OCR). Trả về (newest_overall, newest_reviewed, reviewed_reports):
    - newest_overall: BCTC MỚI NHẤT bất kể loại kỳ (None nếu không lấy được danh sách/không có gì).
    - newest_reviewed: BCTC MỚI NHẤT có period_key xác định được (tên biến giữ nguyên từ trước khi
      sửa 2026-08 — KHÔNG còn nghĩa "đã soát xét" nữa, giờ bao gồm CẢ báo cáo quý thường, xem
      _period_key_for_candidate).
    - reviewed_reports: list TẤT CẢ bản CÓ period_key xác định được (mọi Quarter 1-6), mỗi phần tử
      có thêm key "period_key", ĐÃ KHỬ TRÙNG theo period_key — khi 1 kỳ có CẢ bản đã kiểm toán/soát
      xét LẪN bản quý thường (hiếm, nhưng có thể xảy ra), ưu tiên giữ bản đã kiểm toán/soát xét
      (đáng tin cậy hơn — số liệu chính thức, ít khả năng sai lệch do soát xét lại); nếu cùng mức độ
      soát xét, ưu tiên Quarter==6 (gắn đúng từ đầu, không phải do 24hmoney gắn nhầm Quarter=5)."""
    try:
        items = fetch_cafef_list(ticker) + fetch_24hmoney_list(ticker)
    except Exception:
        return None, None, []
    # select_ranked_reports() (KHÔNG phải select_best_reports()) — GIỮ LẠI mọi ứng viên hợp lệ cho
    # mỗi (Year, Quarter) thay vì chỉ 1 bản thắng cuộc, để fetch_bank_risk_gaps_for_period() có bản
    # THAY THẾ nếu bản ưu tiên nhất tải/đọc thất bại (xem docstring select_ranked_reports — user
    # 2026-08-31 chứng minh bằng ảnh chụp thật ACB/BID/HDB rằng dữ liệu THỰC SỰ tồn tại dù trước đó
    # bị báo "thiếu" chỉ vì thử đúng 1 nguồn rồi bỏ cuộc).
    reports = [x for x in select_ranked_reports(items) if x.get("Quarter") in _QUARTER_END_MONTH]
    if not reports:
        return None, None, []

    def _recency_key(x):
        is_half = _is_half_year(x.get("Name", ""))
        # fetch_24hmoney_list()/_parse_24hmoney_period() gắn Quarter=5 CHO CẢ báo cáo bán niên đã
        # kiểm toán/soát xét (kỳ THẬT kết thúc tháng 6) lẫn báo cáo năm (kỳ kết thúc tháng 12) — nếu
        # cứ tra thẳng _QUARTER_END_MONTH[5]=12 cho case bán niên bị gắn nhầm này, nó sẽ trông "mới"
        # ngang báo cáo NĂM dù thật ra chỉ mới tới giữa năm, có thể lấn át 1 báo cáo quý 3 thật sự mới
        # hơn (bug phát hiện qua unit test khi mô phỏng kịch bản Quý 3 xuất hiện sau bán niên).
        end_month = 6 if (x["Quarter"] == 5 and is_half) else _QUARTER_END_MONTH[x["Quarter"]]
        reviewed = 1 if _is_reviewed_candidate(x) else 0
        return (x["Year"], end_month, reviewed)

    newest_overall = max(reports, key=_recency_key)

    def _cand_priority(x):
        # Độ ưu tiên trong SỐ CÁC BẢN CÙNG 1 KỲ (period_key) — bản đã soát xét/kiểm toán hơn bản
        # chưa, cùng mức soát xét thì Quarter==6 (gắn đúng từ đầu) hơn Quarter khác — giữ NGUYÊN
        # thứ tự ưu tiên cũ (trước đây dùng để CHỌN 1 bản duy nhất), giờ dùng để SẮP XẾP list cho
        # fetch_bank_risk_gaps_for_period() thử lần lượt.
        return (1 if _is_reviewed_candidate(x) else 0, 1 if x.get("Quarter") == 6 else 0)

    by_period = {}
    for cand in reports:
        pk = _period_key_for_candidate(cand)
        if pk is None:
            continue
        by_period.setdefault(pk, []).append(cand)

    reviewed_reports = []
    for pk, cands in by_period.items():
        cands.sort(key=_cand_priority, reverse=True)
        for cand in cands:
            tagged = dict(cand)
            tagged["period_key"] = pk
            reviewed_reports.append(tagged)
    if not reviewed_reports:
        return newest_overall, None, []
    newest_reviewed = max(reviewed_reports, key=_recency_key)
    return newest_overall, newest_reviewed, reviewed_reports


def latest_reviewed_period(ticker):
    """Kiểm tra RẺ (chỉ gọi API liệt kê danh sách BCTC, KHÔNG tải PDF/KHÔNG OCR — vài giây) xem kỳ
    MỚI NHẤT hiện có của ticker này là gì (vd "2026-Q3") — tên hàm giữ nguyên từ trước khi sửa
    2026-08 (không còn nghĩa "đã soát xét" nữa, xem _period_key_for_candidate: BCTC quý thường CŨNG
    có 2 bảng gap, verify thật qua ảnh chụp MBB/TCB). Dùng cho bước kiểm tra "có dữ liệu mới hơn dữ
    liệu đã lưu chưa" trước khi quyết định có cần OCR lại hay không (xem bank_system_risk.py). Trả
    về None nếu không lấy được danh sách BCTC hoặc không có bản nào map được vào 1 kỳ rõ ràng."""
    ticker = ticker.upper()
    _, newest_reviewed, _ = _select_candidate_reports(ticker)
    return newest_reviewed.get("period_key") if newest_reviewed else None


def fetch_bank_risk_gaps(ticker):
    """Trả về dict {"interest_rate_gap": {bucket: value}, "liquidity_gap": {bucket: value},
    "source_title":, "source_url":, "fetched_year":} hoặc None nếu bất kỳ bước nào thất bại (thiếu
    tesseract, không tìm thấy BCTC nào, không định vị được note, OCR không đọc đủ số...). KHÔNG BAO
    GIỜ raise — template_banking.py gọi hàm này trong try/except nhưng bản thân hàm đã tự an toàn.

    SỬA 2026-08: trước đây giả định chỉ BCTC đã KIỂM TOÁN/SOÁT XÉT (Quarter=5 năm, Quarter=6 bán
    niên) mới có đủ 2 bảng gap, báo cáo quý thường bị coi là "rút gọn thuyết minh" — giả định này
    SAI, verify thật bằng ảnh chụp báo cáo Quý 3/2025 (không soát xét) của CẢ MBB và TCB: đều có đầy
    đủ mục "RỦI RO THỊ TRƯỜNG". Giờ MỌI loại kỳ đều là candidate hợp lệ (xem
    _period_key_for_candidate) — vẫn ưu tiên thử bản MỚI NHẤT theo kỳ dữ liệu thật sự trước, chỉ rơi
    về bản đã kiểm toán/soát xét gần nhất (đáng tin hơn nếu có nghi ngờ chất lượng OCR) khi bản mới
    nhất đó thật sự không đọc được.

    Chỉ lấy được kỳ MỚI NHẤT hiện có — xem fetch_bank_risk_gaps_for_period() để lấy 1 kỳ lịch sử cụ
    thể (dùng cho backfill hệ thống ngân hàng, bank_system_risk.py)."""
    ticker = ticker.upper()
    newest_overall, newest_reviewed, _ = _select_candidate_reports(ticker)
    if newest_overall is None:
        print("  [SKIP] Rui ro lai suat/thanh khoan: khong lay duoc danh sach BCTC hoac khong tim thay BCTC nao")
        return None
    if newest_reviewed is None:
        print("  [SKIP] Rui ro lai suat/thanh khoan: khong xac dinh duoc ky bao cao nao")
        return None

    # Thử bản MỚI NHẤT (bất kể loại kỳ) trước — thường TRÙNG với newest_reviewed (cùng 1 candidate,
    # vì giờ mọi loại kỳ đều hợp lệ), lúc đó chỉ tốn đúng 1 lượt thử. Chỉ thử THÊM bản kiểm toán/
    # soát xét khi bản mới nhất khác nó VÀ đọc không ra (hiếm — vd lỗi OCR trang bị lật/mờ).
    os.makedirs(CACHE_DIR, exist_ok=True)
    candidates_to_try = [newest_overall]
    if newest_reviewed["Link"] != newest_overall["Link"]:
        candidates_to_try.append(newest_reviewed)

    for cand in candidates_to_try:
        try:
            pdf_path = _download_report_pdf(ticker, cand)
        except Exception as e:
            print(f"  [WARN] Rui ro lai suat/thanh khoan: tai '{cand['Name']}' that bai ({e})")
            continue
        status, partial = _extract_gaps_from_pdf(pdf_path)
        if status == "missing_tool":
            print("  [SKIP] Rui ro lai suat/thanh khoan: thieu pytesseract/tesseract-ocr binary "
                  "(cai qua 'winget install UB-Mannheim.TesseractOCR' hoac 'apt install tesseract-ocr')")
            return None
        if status == "no_note":
            print(f"  [DIAG] Rui ro lai suat/thanh khoan: '{cand['Name']}' khong doc du 2 bang nay"
                  + (" - thu ban khac" if cand is newest_overall and len(candidates_to_try) > 1 else ""))
            continue
        partial["source_title"] = cand["Name"]
        partial["source_url"] = cand["Link"]
        partial["fetched_year"] = cand["Year"]
        # period_key ("YYYY-H1"/"YYYY-FY") — None nếu cand tình cờ là 1 báo cáo quý thường đọc được
        # (hiếm, không đúng quy ước lưu trữ per-period) — nơi gọi (fetch_bank_risk_gaps_cached) tự
        # bỏ qua việc lưu vào bank_alm_store khi None, không suy diễn kỳ.
        partial["period_key"] = cand.get("period_key") or _period_key_for_candidate(cand)
        return partial

    print("  [SKIP] Rui ro lai suat/thanh khoan: da thu (các) BCTC gan nhat, khong ban nao doc du so lieu "
          "(OCR chat luong kem hoac dinh dang bang khac chuan)")
    return None


def fetch_bank_risk_gaps_for_period(ticker, period_key):
    """Giống fetch_bank_risk_gaps() nhưng lấy ĐÚNG 1 KỲ LỊCH SỬ cụ thể (vd "2025-H1") thay vì luôn
    lấy kỳ mới nhất — dùng cho backfill lịch sử hệ thống ngân hàng (bank_system_risk.py). Trả về
    None nếu ticker này KHÔNG CÓ báo cáo đã kiểm toán/soát xét đúng kỳ đó (KHÔNG PHẢI lỗi — ngân
    hàng có thể chưa niêm yết lúc đó, hoặc CafeF/24hmoney không còn lưu bản cũ) — không thử "kỳ gần
    đúng nhất", chỉ khớp CHÍNH XÁC period_key hoặc trả về None, để backfill không tự đoán số liệu.
    KHÔNG BAO GIỜ raise, giống fetch_bank_risk_gaps().

    SỬA 2026-08-31 (user cung cấp ảnh chụp thật ACB/BID chứng minh dữ liệu tồn tại dù trước đó báo
    "thiếu"): _select_candidate_reports() giờ trả về NHIỀU ứng viên cho cùng 1 period_key (sắp xếp
    ưu tiên giảm dần) thay vì chỉ 1 — THỬ LẦN LƯỢT từng ứng viên (tải + đọc), dùng bản ĐẦU TIÊN
    thành công, chỉ bỏ cuộc khi TẤT CẢ đều thất bại. Lỗi tải (link hỏng/404) hay "no_note" (tài liệu
    này thật sự không đọc đủ 2 bảng) đều đáng thử bản khác; "missing_tool" (thiếu tesseract) là lỗi
    MÔI TRƯỜNG — thử bản khác cũng vô ích, dừng ngay."""
    ticker = ticker.upper()
    _, _, reviewed_reports = _select_candidate_reports(ticker)
    candidates = [c for c in reviewed_reports if c.get("period_key") == period_key]
    if not candidates:
        print(f"  [SKIP] Rui ro lai suat/thanh khoan ({period_key}): khong tim thay BCTC dung ky nay cho {ticker}")
        return None

    os.makedirs(CACHE_DIR, exist_ok=True)
    for i, cand in enumerate(candidates):
        n_left = len(candidates) - i - 1
        try:
            pdf_path = _download_report_pdf(ticker, cand)
        except Exception as e:
            print(f"  [WARN] Rui ro lai suat/thanh khoan ({period_key}): tai '{cand['Name']}' that bai ({e})"
                  + (f" - thu ban thay the ({n_left} con lai)" if n_left else ""))
            continue
        status, partial = _extract_gaps_from_pdf(pdf_path)
        if status == "missing_tool":
            print("  [SKIP] Rui ro lai suat/thanh khoan: thieu pytesseract/tesseract-ocr binary "
                  "(cai qua 'winget install UB-Mannheim.TesseractOCR' hoac 'apt install tesseract-ocr')")
            return None
        if status == "no_note":
            print(f"  [DIAG] Rui ro lai suat/thanh khoan ({period_key}): '{cand['Name']}' khong doc du 2 bang"
                  + (f" - thu ban thay the ({n_left} con lai)" if n_left else ""))
            continue
        partial["source_title"] = cand["Name"]
        partial["source_url"] = cand["Link"]
        partial["fetched_year"] = cand["Year"]
        partial["period_key"] = period_key
        return partial
    return None


def fetch_bank_risk_gaps_cached(ticker):
    """Wrapper tái sử dụng dữ liệu GIỮA 2 pipeline: phân tích 1 mã lẻ (template_banking.py) và tổng
    hợp toàn hệ thống ngân hàng (bank_system_risk.py) — cả 2 nên gọi hàm này thay vì
    fetch_bank_risk_gaps() trực tiếp, để 1 lần OCR phục vụ được cả 2 nơi.

    Với ticker thuộc bank_universe.BANKING_TICKERS: kiểm tra RẺ (latest_reviewed_period, không OCR)
    xem kỳ mới nhất hiện có là gì; nếu data/bank_alm/<TICKER>.json ĐÃ CÓ đúng kỳ đó với
    status="reported" thì trả thẳng từ store (không tốn mạng/OCR); nếu không, gọi
    fetch_bank_risk_gaps() bình thường rồi LƯU LẠI kết quả vào store cho lần gọi sau (của chính
    pipeline này hoặc pipeline kia). Ticker ngoài bank_universe rơi thẳng về fetch_bank_risk_gaps()
    không qua store (không nên xảy ra vì module này chỉ được gọi cho cổ phiếu ngân hàng, nhưng
    phòng hờ). KHÔNG BAO GIỜ raise — lỗi đọc/ghi store bị nuốt, không làm hỏng kết quả OCR đã fetch
    được."""
    ticker = ticker.upper()
    try:
        from bank_universe import BANKING_TICKERS
    except Exception:
        BANKING_TICKERS = frozenset()
    if ticker not in BANKING_TICKERS:
        return fetch_bank_risk_gaps(ticker)

    import bank_alm_store
    cheap_period = latest_reviewed_period(ticker)
    if cheap_period:
        try:
            bank_alm_store.record_cheap_check(ticker, cheap_period)
            stored = bank_alm_store.get_period_entry(ticker, cheap_period)
        except Exception:
            stored = None
        if stored and stored.get("status") == "reported" and \
                (stored.get("interest_rate_gap") or stored.get("liquidity_gap")):
            result = {"period_key": cheap_period}
            for k in ("interest_rate_gap", "liquidity_gap", "liabilities_by_bucket",
                      "interest_rate_sensitivity_disclosed"):
                if stored.get(k) is not None:
                    result[k] = stored[k]
            src = stored.get("source") or {}
            result["source_title"] = src.get("title")
            result["source_url"] = src.get("url")
            result["fetched_year"] = src.get("fetched_year")
            return result

    partial = fetch_bank_risk_gaps(ticker)
    if partial and partial.get("period_key"):
        source = {"title": partial.get("source_title"), "url": partial.get("source_url"),
                  "fetched_year": partial.get("fetched_year")}
        try:
            bank_alm_store.upsert_reported_period(ticker, partial["period_key"], partial, source)
        except Exception:
            pass  # cache la "co thi tot" — loi ghi khong lam hong ket qua da fetch duoc
    return partial


# ── Tính chỉ số từ gap đã trích ──────────────────────────────────────────────────────────────────

def compute_interest_rate_risk_metrics(gap, total_assets, equity=None, disclosed_sensitivity=None,
                                        stress_bps=(100, 200), nii=None):
    """gap: dict theo INTEREST_RATE_BUCKETS. Trả về dict chỉ số — xem docstring module cho phương
    pháp (static gap, trọng số theo điểm giữa mỗi bucket, chỉ 4 bucket <=12 thang moi anh huong NII
    NAM NAY, bucket >1 nam khong lap lai trong nam danh gia).

    `equity` (VCSH, cùng đơn vị `total_assets`) — tuỳ chọn, nếu có sẽ thêm Gap/Equity (quy mô lệch kỳ
    hạn so với vốn chủ, KHÔNG phải so với tổng tài sản — 1 lệch nhỏ so với tổng tài sản có thể lớn so
    với vốn chủ vì đòn bẩy ngân hàng rất cao).
    `disclosed_sensitivity` — dict {ccy: {...}} từ _extract_rate_sensitivity_table() nếu ngân hàng
    CÓ tự công bố (đáng tin hơn nii_sensitivity_per_shock tự ước tính vì có tính hành vi/giả định nội
    bộ của ngân hàng) — chỉ đính kèm vào output, không dùng để ghi đè số tự tính.
    `stress_bps` — các mức sốc lãi suất (điểm cơ bản) muốn tính kịch bản, mặc định +100/+200bp.
    `nii` (Thu nhập lãi thuần thực tế kỳ gần nhất, cùng đơn vị `total_assets`) — tuỳ chọn, nếu có sẽ
    thêm stress_scenarios_nii_ratio (ΔNII/NII thực tế) — mức độ NGHIÊM TRỌNG của cú sốc so với chính
    dòng thu nhập lãi thuần của ngân hàng, đáng đọc hơn nhiều so với chỉ nhìn số tỷ VND tuyệt đối."""
    b = gap
    horizon_1y_keys = ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang"]
    weights = {"den_1_thang": 11.5 / 12, "tu_1_3_thang": 10 / 12, "tu_3_6_thang": 7.5 / 12, "tu_6_12_thang": 3 / 12}
    cum_1y = sum(b.get(k, 0.0) for k in horizon_1y_keys)
    weighted_gap = sum(b.get(k, 0.0) * weights[k] for k in horizon_1y_keys)
    ratio_1y = cum_1y / total_assets if total_assets else None
    gap_to_equity = cum_1y / equity if equity else None
    # Gap LŨY KẾ theo TỪNG mốc kỳ hạn (không chỉ điểm cuối <=12 tháng) — 1 gap <=12 tháng dương có thể
    # che giấu 1 gap âm rất lớn ở mốc <=1 tháng (xem [[feedback_bank_alm_risk_framework]] — nguyên tắc
    # "never sum all buckets into one number", phải đọc CẢ chuỗi lũy kế theo từng mốc).
    _horizon_keys_progressive = [
        ("1m", ["den_1_thang"]),
        ("3m", ["den_1_thang", "tu_1_3_thang"]),
        ("6m", ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang"]),
        ("12m", horizon_1y_keys),
    ]
    cumulative_gap_by_horizon = {}
    for hz_label, hz_keys in _horizon_keys_progressive:
        hz_val = sum(b.get(k, 0.0) for k in hz_keys)
        cumulative_gap_by_horizon[hz_label] = {
            "value": hz_val,
            "ratio": (hz_val / total_assets if total_assets else None),
        }
    # "level"/"sensitivity_level" CHỈ là MỐC THAM CHIẾU để định hướng đọc số liệu (0-2% cân bằng, 2-5%
    # theo dõi, >5% đáng chú ý) — KHÔNG phải ngưỡng đạt/không đạt kiểu quy định (không có ngưỡng pháp
    # lý nào cho gap nội bộ ngân hàng). Diễn giải đúng đắn LUÔN cần đi kèm: gap này có thực sự ảnh
    # hưởng đáng kể đến NII/LNTT không (xem stress_scenarios_nii), so với biến động NIM lịch sử ra sao,
    # và xu hướng qua các kỳ (cần dữ liệu nhiều năm, hiện tại BCTC risk-note chỉ lấy được kỳ gần nhất).
    # Thang tham chiếu ĐÚNG như user hướng dẫn (0-2% khá cân bằng, 2-5% cần theo dõi, 5-10% khá lớn,
    # >10% đáng chú ý) — chỉ là "ngưỡng phân tích để so sánh các ngân hàng", KHÔNG phải ngưỡng pháp lý.
    if ratio_1y is None:
        level = "Không xác định"
    elif abs(ratio_1y) < 0.02:
        level = "Khá cân bằng"
    elif abs(ratio_1y) < 0.05:
        level = "Cần theo dõi"
    elif abs(ratio_1y) < 0.10:
        level = "Khá lớn"
    else:
        level = "Đáng chú ý"
    # Kich ban stress: +bps (lai suat tang -> gap am se AN, gap duong se LOI) va -bps (nguoc lai) cho
    # MOI muc trong stress_bps — cong thuc tuyen tinh (static gap) nen +200bp = 2x +100bp, van tinh
    # rieng tung muc (khong chi nhan 2) de neu sau nay doi sang mo hinh phi tuyen thi khong phai sua
    # ca cho goi.
    scenarios = {}
    for bps in stress_bps:
        scenarios[f"+{bps}bp"] = round(weighted_gap * (bps / 10000))
        scenarios[f"-{bps}bp"] = round(weighted_gap * (-bps / 10000))
    scenarios_ratio = ({k: (v / nii if nii else None) for k, v in scenarios.items()} if nii else None)
    return {
        "gap_by_bucket": b,
        "gap_ratio_by_bucket": {k: (v / total_assets if total_assets else None) for k, v in b.items()},
        "cumulative_gap_by_horizon": cumulative_gap_by_horizon,
        "cumulative_gap_1y": cum_1y,
        "cumulative_gap_1y_ratio": ratio_1y,
        "gap_to_equity_1y": gap_to_equity,
        "sensitive_type": "Asset-sensitive" if cum_1y > 0 else ("Liability-sensitive" if cum_1y < 0 else "Trung tính"),
        "nii_sensitivity_per_shock": scenarios.get("+100bp", round(weighted_gap * 0.01)),
        "shock_bps": 100,
        "stress_scenarios_nii": scenarios,
        "stress_scenarios_nii_ratio": scenarios_ratio,
        "sensitivity_level": level,
        "disclosed_sensitivity": disclosed_sensitivity or None,
        # Không tự tính ΔEVE (Economic Value of Equity): cần duration/thời lượng còn lại của TỪNG
        # bucket, trong khi thuyết minh BCTC VN chỉ công bố SỐ DƯ gap theo bucket, không có duration —
        # ước lượng duration sẽ là bịa số, không đáng tin hơn việc không tính (xem
        # [[feedback_bank_alm_risk_framework]]).
        "not_computed": ["ΔEVE (cần duration từng bucket, BCTC không công bố)"],
        # Rủi ro định tính KHÔNG thấy được từ riêng bảng gap số — chỉ là lưu ý đi kèm, không phải chỉ
        # số tính toán: (1) basis risk — tài sản/nguồn vốn có thể tham chiếu lãi suất KHÁC NHAU không
        # di chuyển cùng nhịp dù cùng 1 bucket kỳ hạn; (2) rủi ro hành vi — CASA ghi nhận kỳ hạn "không
        # nhạy cảm lãi suất"/qua đêm theo hợp đồng nhưng hành vi thực tế của khách hàng thường ổn định
        # (sticky) hơn nhiều so với kỳ hạn hợp đồng.
        "qualitative_caveats": [
            "Basis risk: tài sản/nguồn vốn có thể tham chiếu lãi suất khác nhau, không di chuyển "
            "cùng nhịp dù cùng 1 bucket kỳ hạn.",
            "Rủi ro hành vi: CASA/tiền gửi không kỳ hạn ghi nhận theo hợp đồng (qua đêm) nhưng hành "
            "vi thực tế thường ổn định (sticky) hơn nhiều.",
        ],
    }


def compute_liquidity_risk_metrics(gap, total_assets, equity=None, liquid_assets=None,
                                    customer_deposits=None, deposit_stress_pct=(5, 10, 20),
                                    liabilities_gap=None):
    """gap: dict theo LIQUIDITY_BUCKETS.

    `liquid_assets` (tiền mặt + tiền gửi NHNN + tiền gửi/cho vay TCTD — LẤY THẲNG từ dữ liệu BCTC đã
    fetch sẵn qua Vietcap trong template_banking.py, KHÔNG cần OCR thêm) — tuỳ chọn, nếu có sẽ tính
    Liquid Assets/Tổng tài sản, dùng làm "đạn" so sánh với kịch bản rút tiền gửi.
    `customer_deposits` — tuỳ chọn, để chạy stress test rút X% tiền gửi khách hàng (deposit run) so
    với `liquid_assets`: thiếu hụt = X%×tiền gửi − liquid_assets (dương = liquid assets KHÔNG đủ bù,
    cần huy động/vay thêm — đúng chuỗi nhân quả rủi ro thanh khoản → áp lực huy động bạn đã nêu).
    `liabilities_gap` — dict theo LIQUIDITY_BUCKETS từ dòng "Tổng nợ phải trả" (nếu trích được, xem
    _extract_gaps_from_pdf) — dùng tính Liquid Assets/Nợ phải trả ngắn hạn, bản ĐƠN GIẢN HÓA của LCR
    (KHÁC LCR Basel chuẩn — không phân loại HQLA/outflow rate — nhưng vẫn hữu ích để phân tích BCTC)."""
    b = gap
    cum_1m = b.get("den_1_thang", 0.0)
    cum_1y = cum_1m + b.get("tu_1_3_thang", 0.0) + b.get("tu_3_12_thang", 0.0)
    ratio_1m = cum_1m / total_assets if total_assets else None
    ratio_1y = cum_1y / total_assets if total_assets else None
    if ratio_1m is None:
        level = "Không xác định"
    elif ratio_1m > -0.03:
        level = "Thấp"
    elif ratio_1m > -0.08:
        level = "Trung bình"
    else:
        level = "Cao"
    # Gap lũy kế theo TỪNG mốc (không chỉ <=1 tháng) — "quá hạn" tính vào mốc <=1 tháng vì thực chất
    # còn cấp bách hơn (đã trễ hạn), khác cumulative_gap_1m ở trên (giữ nguyên định nghĩa cũ để không
    # đổi ý nghĩa của field đã dùng chỗ khác — field mới này bổ sung, không thay thế).
    _st_keys = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang"]
    _horizon_keys_progressive = [
        ("1m", _st_keys),
        ("3m", _st_keys + ["tu_1_3_thang"]),
        ("12m", _st_keys + ["tu_1_3_thang", "tu_3_12_thang"]),
    ]
    cumulative_gap_by_horizon = {}
    for hz_label, hz_keys in _horizon_keys_progressive:
        hz_val = sum(b.get(k, 0.0) for k in hz_keys)
        cumulative_gap_by_horizon[hz_label] = {
            "value": hz_val,
            "ratio": (hz_val / total_assets if total_assets else None),
        }
    result = {
        "gap_by_bucket": b,
        "gap_ratio_by_bucket": {k: (v / total_assets if total_assets else None) for k, v in b.items()},
        "cumulative_gap_by_horizon": cumulative_gap_by_horizon,
        "cumulative_gap_1m": cum_1m,
        "cumulative_gap_1m_ratio": ratio_1m,
        "cumulative_gap_1y": cum_1y,
        "cumulative_gap_1y_ratio": ratio_1y,
        "gap_to_equity_1m": (cum_1m / equity if equity else None),
        "risk_level": level,
        "liquid_assets_ratio": (liquid_assets / total_assets if liquid_assets and total_assets else None),
        # Không tự tính LCR/NSFR chuẩn Basel: cần phân loại HQLA (tài sản thanh khoản chất lượng cao)
        # và nguồn vốn ổn định theo đúng trọng số quy định — BCTC VN không công bố đủ chi tiết để phân
        # loại đúng, một số ước lượng "gần đúng" sẽ SAI LỆCH và tự tin giả tạo hơn là không tính (xem
        # [[feedback_bank_alm_risk_framework]]). liquid_assets_to_st_liabilities_simple bên dưới LÀ 1
        # bản đơn giản hoá KHÁC (không cần phân loại HQLA/outflow rate), không mâu thuẫn với việc này.
        "not_computed": ["LCR chuẩn Basel (cần phân loại HQLA chi tiết)",
                          "NSFR (cần phân loại nguồn vốn ổn định chi tiết)"],
    }
    if liabilities_gap:
        short_term_liab = sum(liabilities_gap.get(k, 0.0) for k in _st_keys + ["tu_1_3_thang", "tu_3_12_thang"])
        result["short_term_liabilities"] = short_term_liab
        result["liquid_assets_to_st_liabilities_simple"] = (
            liquid_assets / short_term_liab if liquid_assets is not None and short_term_liab else None)
    if liquid_assets is not None and customer_deposits:
        result["deposit_run_stress"] = {
            f"-{p}%": round(customer_deposits * (p / 100) - liquid_assets) for p in deposit_stress_pct
        }
        # Ty le "dan" con lai sau stress — vd 65% nghia la liquid assets van con du 65% so voi phan
        # tien gui bi rut, KHONG phai "con thieu 35%" (do CON tai san khac ngoai liquid assets co the
        # xoay xo further, chi la khong "ngay lap tuc" bang liquid assets).
        result["deposit_run_coverage"] = {
            f"-{p}%": (liquid_assets / (customer_deposits * (p / 100)) if customer_deposits * (p / 100) else None)
            for p in deposit_stress_pct
        }
    return result


def compute_balance_sheet_alm_ratios(total_assets, loans, interbank_liab, bonds_issued):
    """2 tỷ lệ cơ cấu bảng cân đối phục vụ ALM tính THẲNG từ dữ liệu BCTC đã fetch sẵn qua Vietcap
    trong template_banking.py (loans_hist/interbank_hist/bonds_hist/total_assets_hist) — KHÔNG cần
    OCR thuyết minh, khác 2 hàm gap ở trên. Mẫu số dùng total_assets (= tổng nguồn vốn, vì
    Assets = Nợ phải trả + VCSH) thay vì tự cộng lại từng khoản mục nguồn vốn — tránh rủi ro
    thiếu/dư 1 khoản mục nào đó khi tính tổng nguồn vốn độc lập.
    - loan_to_assets: tài sản càng dồn vào cho vay (kém thanh khoản hơn) thì dư địa xoay xở càng ít.
    - wholesale_funding_ratio: phụ thuộc vốn liên NH/GTCG (nhạy cảm tâm lý thị trường) càng cao thì
      thanh khoản càng dễ biến động theo khẩu vị thị trường thay vì tiền gửi khách hàng ổn định."""
    return {
        "loan_to_assets": (loans / total_assets if total_assets else None),
        "wholesale_funding_ratio": (((interbank_liab or 0) + (bonds_issued or 0)) / total_assets
                                     if total_assets else None),
    }


def build_risk_narrative_lines(ir_metrics, liq_metrics, ticker, bs_ratios=None, ldr=None, casa=None,
                                nim_current=None):
    """Trả về LIST các câu tiếng Việt CÓ DẤU tóm tắt 2 rủi ro (+ cơ cấu bảng cân đối nếu có) — mỗi
    phần tử là 1 ý riêng, dùng để hiển thị dạng gạch đầu dòng trên web (đọc dễ hơn 1 đoạn văn dài);
    nối lại bằng " ".join() để ra bản văn xuôi cho PDF/narrative cũ. Không tự bịa số nào ngoài các
    dict đầu vào đã tính. Theo đúng phương pháp [[feedback_bank_alm_risk_framework]] — 5 tầng đánh
    giá rủi ro lãi suất (Repricing Gap/Assets, Cumulative Gap theo từng mốc, NII sensitivity, ΔEVE,
    NIM thực tế), và bộ chỉ số thanh khoản (LDR đúng cách, CASA, Liquid Assets/Assets, Liquid
    Assets/Nợ ngắn hạn (LCR đơn giản hoá), Liquidity Gap theo mốc, Wholesale Funding): đọc gap LŨY KẾ
    theo từng mốc (không chỉ điểm cuối), mức tham chiếu chỉ để định hướng chứ không phải ngưỡng
    đạt/không đạt, luôn đọc CẶP chỉ số cùng nhau, và nêu rõ những gì KHÔNG tính được thay vì bỏ qua
    im lặng.

    `ldr`, `casa` (%, 0-100) — LDR/CASA đã tính SẴN ở nơi khác trong template_banking.py (LDR theo
    đúng công thức tín dụng/huy động user đã hướng dẫn, KHÔNG phải Loans/Deposits đơn thuần — công
    thức đó sai vì không phản ánh đúng khác biệt cơ cấu nguồn vốn) — chỉ ĐÍNH KÈM vào đánh giá ALM để
    đọc cặp cùng Liquid Assets/Assets, KHÔNG tính lại.
    `nim_current` (%, kỳ gần nhất) — Tầng 5 của khung 5 tầng: Gap chỉ là mô hình lý thuyết, NIM thực
    tế mới là thứ THỰC SỰ xảy ra."""
    lines = []
    if ir_metrics:
        pct = ir_metrics["cumulative_gap_1y_ratio"]
        pct_s = f"{pct*100:+.2f}%" if pct is not None else "?"
        toe = ir_metrics.get("gap_to_equity_1y")
        toe_s = f", {toe*100:+.0f}% VCSH" if toe is not None else ""
        stype = ir_metrics["sensitive_type"]
        # Diễn giải TRỰC TIẾP ý nghĩa dấu gap (RSA-RSL) theo đúng nguyên lý user đưa ra — không chỉ
        # dán nhãn "Asset/Liability-sensitive" mà nói rõ NẾU lãi suất tăng/giảm thì NII bị ảnh hưởng
        # theo chiều nào, vì đây là phần "đáng tiền" nhất của repricing gap.
        if stype == "Asset-sensitive":
            direction = ("tài sản được định lại lãi suất NHANH HƠN nguồn vốn — lãi suất TĂNG có lợi "
                         "cho NII, lãi suất GIẢM bất lợi cho NII")
        elif stype == "Liability-sensitive":
            direction = ("chi phí vốn được định lại lãi suất NHANH HƠN tài sản — lãi suất TĂNG gây "
                         "áp lực lên NII/NIM, lãi suất GIẢM lại có lợi")
        else:
            direction = "gần như trung tính với biến động lãi suất"
        lines.append(
            f"Repricing Gap (Tầng 1+2): gap lũy kế ≤1 năm = {pct_s} tổng tài sản{toe_s} — Ngân hàng "
            f"đang {stype} ({direction}). Mức tham chiếu: {ir_metrics['sensitivity_level']} (thang so "
            f"sánh 0-2%/2-5%/5-10%/>10%, KHÔNG phải ngưỡng đạt/không đạt)."
        )
        hz = ir_metrics.get("cumulative_gap_by_horizon") or {}
        hz_parts = [f"≤{lbl}: {(hz.get(lbl) or {}).get('ratio')*100:+.2f}%"
                    for lbl in ("1m", "3m", "6m", "12m") if (hz.get(lbl) or {}).get("ratio") is not None]
        if hz_parts:
            lines.append(
                "Cumulative Gap theo từng mốc (không chỉ nhìn điểm cuối ≤12 tháng — 1 gap dương ở "
                "mốc cuối vẫn có thể che 1 gap âm lớn ở mốc gần hơn, cú sốc lãi suất xảy ra HÔM NAY "
                "chứ không phải sau 1 năm): " + "; ".join(hz_parts) + ".")
        sc = ir_metrics.get("stress_scenarios_nii") or {}
        sc_ratio = ir_metrics.get("stress_scenarios_nii_ratio")
        if sc:
            sc_s = "; ".join(
                f"{k}: {v:+,.0f} tỷ VND" + (f" ({sc_ratio[k]*100:+.1f}% NII)" if sc_ratio and sc_ratio.get(k) is not None else "")
                for k, v in sc.items())
            lines.append(f"NII sensitivity (Tầng 3, ước tính static gap): {sc_s}.")
        disclosed = ir_metrics.get("disclosed_sensitivity")
        if disclosed:
            parts = "; ".join(
                f"{ccy} +{v['rate_increase_pct']:.2f}%: LNTT {v['pbt_impact']:+,.0f} tỷ, "
                f"VCSH {v['equity_impact']:+,.0f} tỷ" for ccy, v in disclosed.items()
            )
            lines.append(f"Độ nhạy lãi suất do NGÂN HÀNG TỰ CÔNG BỐ (đáng tin cậy hơn ước tính static gap ở trên): {parts}.")
        if nim_current is not None:
            lines.append(
                f"NIM thực tế kỳ gần nhất (Tầng 5): {nim_current:.2f}% — đây là con số THỰC SỰ xảy "
                f"ra, cần theo dõi khi môi trường lãi suất thay đổi, trong khi Repricing Gap ở trên "
                f"chỉ là mô hình lý thuyết (static gap, chưa tính hành vi khách hàng/prepayment)."
            )
    if liq_metrics:
        r1m = liq_metrics["cumulative_gap_1m_ratio"]
        r1y = liq_metrics["cumulative_gap_1y_ratio"]
        r1m_s = f"{r1m*100:+.2f}%" if r1m is not None else "?"
        r1y_s = f"{r1y*100:+.2f}%" if r1y is not None else "?"
        la_ratio = liq_metrics.get("liquid_assets_ratio")
        la_s = f" Liquid Assets/Tổng TS = {la_ratio*100:.1f}%." if la_ratio is not None else ""
        lines.append(
            f"Liquidity Gap: gap ròng ≤1 tháng = {r1m_s} tổng tài sản, luỹ kế ≤1 năm = {r1y_s} (mức "
            f"tham chiếu: {liq_metrics['risk_level']} — chỉ là mốc định hướng).{la_s} Gap âm nghĩa là "
            f"ngân hàng đang có mismatch kỳ hạn (bình thường với mô hình ngân hàng — huy động ngắn "
            f"cho vay dài), quan trọng là có nguồn thanh khoản thay thế hay không — cần đọc CÙNG "
            f"Liquid Assets buffer ở trên, không đọc riêng lẻ."
        )
        st_ratio = liq_metrics.get("liquid_assets_to_st_liabilities_simple")
        if st_ratio is not None:
            lines.append(
                f"Liquid Assets/Nợ phải trả ngắn hạn (LCR đơn giản hoá, KHÁC LCR Basel chuẩn — không "
                f"phân loại HQLA/outflow rate nhưng vẫn hữu ích để tham khảo): {st_ratio*100:.1f}% — "
                f"tài sản thanh khoản ngay có thể cover được khoảng đó nghĩa vụ ngắn hạn."
            )
        stress = liq_metrics.get("deposit_run_stress")
        coverage = liq_metrics.get("deposit_run_coverage")
        if stress:
            parts = "; ".join(
                f"rút {k} tiền gửi KH: thiếu hụt {v:+,.0f} tỷ VND"
                + (f" (Liquid Assets che phủ {coverage[k]*100:.0f}%)" if coverage and coverage.get(k) is not None else "")
                for k, v in stress.items())
            lines.append(f"Stress test rút tiền gửi (bank run): {parts}.")
    if bs_ratios:
        la = bs_ratios.get("loan_to_assets")
        wf = bs_ratios.get("wholesale_funding_ratio")
        ldr_s = f", LDR = {ldr:.1f}%" if ldr is not None else ""
        casa_s = f", CASA = {casa:.1f}%" if casa is not None else ""
        if la is not None or wf is not None or ldr is not None or casa is not None:
            lines.append(
                f"Cơ cấu bảng cân đối (đọc CÙNG Liquid Assets/Tổng TS ở trên, không đọc riêng lẻ): "
                f"Cho vay/Tổng tài sản = {la*100:.1f}%{ldr_s}{casa_s}" + (
                    f", Wholesale Funding/Tổng tài sản = {wf*100:.1f}%." if wf is not None else "."
                )
            )
    not_computed = []
    for m in (ir_metrics, liq_metrics):
        if m and m.get("not_computed"):
            not_computed.extend(m["not_computed"])
    if not_computed:
        lines.append("KHÔNG tính (BCTC không công bố đủ chi tiết, ước lượng sẽ sai lệch và tự tin "
                      "giả tạo hơn là không tính): " + "; ".join(not_computed) + ".")
    caveats = []
    for m in (ir_metrics, liq_metrics):
        if m and m.get("qualitative_caveats"):
            for c in m["qualitative_caveats"]:
                if c not in caveats:
                    caveats.append(c)
    if caveats:
        lines.append("Lưu ý định tính (không thể hiện qua bảng gap số): " + " ".join(caveats))
    lines.append(
        "Các mức tham chiếu ở trên chỉ dựa trên 1 kỳ báo cáo gần nhất, KHÔNG phải ngưỡng đạt/không "
        "đạt cố định — mỗi ngân hàng có mô hình kinh doanh, tỷ trọng bán lẻ/doanh nghiệp, cơ cấu kỳ "
        "hạn khác nhau; đánh giá đầy đủ cần so sánh với peer cùng ngành và xu hướng qua nhiều kỳ."
    )
    return lines


def build_risk_narrative(ir_metrics, liq_metrics, ticker, bs_ratios=None):
    """Bản văn xuôi (nối các câu của build_risk_narrative_lines bằng khoảng trắng) — giữ lại cho
    tương thích ngược (PDF/JSON cũ dùng 1 chuỗi), dùng build_risk_narrative_lines() trực tiếp ở nơi
    cần hiển thị dạng danh sách (vd web)."""
    return " ".join(build_risk_narrative_lines(ir_metrics, liq_metrics, ticker, bs_ratios=bs_ratios))


def build_risk_summary(ir_metrics, liq_metrics):
    """Trả về dict {"interest_rate_level", "liquidity_level", "summary_text"} hoặc None — bản TÓM TẮT
    2-3 CÂU đặt Ở ĐẦU phần đánh giá ALM chi tiết (build_risk_narrative_lines) VÀ lặp lại trong khối
    "Nhận định nhanh Earning Release" của phần cập nhật quý — người đọc cần biết NGAY rủi ro lãi suất/
    thanh khoản có cao không và khả năng CHỐNG CHỊU biến cố ra sao trước khi đọc chi tiết từng
    tầng/chỉ số dài dòng bên dưới. Khả năng chống chịu dựa trên chính KẾT QUẢ STRESS TEST (NII sensitivity,
    deposit run coverage) — đáng tin hơn nhiều so với chỉ nhìn mức tham chiếu gap thô (1 gap "Cần theo
    dõi" có thể vẫn CHỐNG CHỊU TỐT nếu tác động NII thực tế nhỏ, xem [[feedback_bank_alm_risk_framework]]
    — nguyên tắc luôn đọc cặp Gap + mức độ ảnh hưởng thực tế cùng nhau). Không tự bịa số, chỉ diễn giải
    lại các dict đầu vào đã tính ở compute_interest_rate_risk_metrics/compute_liquidity_risk_metrics."""
    if not ir_metrics and not liq_metrics:
        return None
    parts = []
    ir_level = liq_level = None
    if ir_metrics:
        ir_level = ir_metrics["sensitivity_level"]
        gap_ratio = ir_metrics.get("cumulative_gap_1y_ratio")
        sc_ratio = ir_metrics.get("stress_scenarios_nii_ratio") or {}
        worst = max((abs(v) for v in sc_ratio.values() if v is not None), default=None)
        gap_s = f" (gap {gap_ratio*100:+.2f}% tổng tài sản)" if gap_ratio is not None else ""
        if worst is not None:
            if worst < 0.05:
                resil = "TỐT — NII ít biến động ngay cả khi lãi suất sốc mạnh"
            elif worst < 0.15:
                resil = "VỪA PHẢI — NII biến động đáng kể nhưng chưa ở mức nghiêm trọng"
            else:
                resil = "CẦN LƯU Ý — NII biến động mạnh khi lãi suất sốc"
            parts.append(f"Rủi ro lãi suất: mức tham chiếu {ir_level}{gap_s}, khả năng chống chịu {resil} "
                          f"(kịch bản sốc mạnh nhất làm NII đổi khoảng {worst*100:.1f}%).")
        else:
            parts.append(f"Rủi ro lãi suất: mức tham chiếu {ir_level}{gap_s}.")
    if liq_metrics:
        liq_level = liq_metrics["risk_level"]
        coverage = liq_metrics.get("deposit_run_coverage") or {}
        worst_cov = min((v for v in coverage.values() if v is not None), default=None)
        if worst_cov is not None:
            if worst_cov >= 1.0:
                resil = "TỐT — tài sản thanh khoản đủ bù ngay cả kịch bản rút tiền gửi mạnh nhất đã test"
            elif worst_cov >= 0.5:
                resil = "VỪA PHẢI — có thể cần huy động thêm nếu bị rút tiền gửi mạnh"
            else:
                resil = "CẦN LƯU Ý — tài sản thanh khoản không đủ bù kịch bản rút tiền gửi mạnh đã test"
            parts.append(f"Rủi ro thanh khoản: mức tham chiếu {liq_level}, khả năng chống chịu {resil} "
                          f"(che phủ {worst_cov*100:.0f}% ở kịch bản rút tiền gửi mạnh nhất).")
        else:
            parts.append(f"Rủi ro thanh khoản: mức tham chiếu {liq_level}.")
    return {
        "interest_rate_level": ir_level,
        "liquidity_level": liq_level,
        "summary_text": " ".join(parts),
    }
